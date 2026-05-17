#include "excel_cm.h"
#include <OpenXLSX.hpp>

extern "C" {
#include <xlsxio_read.h>
}

#include <spdlog/spdlog.h>

#include <algorithm>
#include <cctype>
#include <map>
#include <optional>
#include <stdexcept>
#include <string>
#include <vector>

namespace {
// Module-local logger: uses named "excel_cm" if registered, else default
spdlog::logger* log() {
    static spdlog::logger* inst = [] {
        auto l = spdlog::get("excel_cm");
        return l ? l.get() : spdlog::default_logger().get();
    }();
    return inst;
}
} // anonymous namespace

namespace wlcc {

// ============================================================================
// Helpers: OpenXLSX cell → CellValue conversion (for write-mode reads)
// ============================================================================

static CellValue xlCellValueToCellValue(const OpenXLSX::XLCellValueProxy& cellValue) {
    switch (cellValue.type()) {
        case OpenXLSX::XLValueType::String:
            return cellValue.get<std::string>();
        case OpenXLSX::XLValueType::Integer:
            return cellValue.get<int64_t>();
        case OpenXLSX::XLValueType::Float:
            return cellValue.get<double>();
        case OpenXLSX::XLValueType::Boolean:
            return cellValue.get<bool>();
        default:
            return std::monostate{};
    }
}

// ============================================================================
// Helpers: xlsxio string → CellValue inference (for read-mode)
// ============================================================================

static CellValue inferCellValue(const char* raw) {
    if (!raw || raw[0] == '\0') {
        return std::monostate{};
    }

    std::string s(raw);

    // --- Try integer (int64_t) for short numeric strings (≤ 9 digits) ---
    // Longer digit strings (10+ digits) like phone numbers "919447508628"
    // are kept as std::string to avoid misinterpretation.
    {
        const char* p = s.c_str();
        if (*p == '-') p++;
        if (*p != '\0') {
            const char* start = p;
            bool allDigits = true;
            while (*p) {
                if (!std::isdigit(static_cast<unsigned char>(*p))) {
                    allDigits = false;
                    break;
                }
                p++;
            }
            size_t digitCount = static_cast<size_t>(p - start);
            if (allDigits && digitCount >= 1 && digitCount <= 9) {
                try {
                    size_t pos = 0;
                    int64_t val = std::stoll(s, &pos);
                    if (pos == s.size()) {
                        return val;
                    }
                } catch (...) {
                    // Overflow — fall through
                }
            }
        }
    }

    // --- Try double if contains '.' or 'e'/'E' (scientific notation) ---
    // Preserves fractional times (0.791667) while keeping everything else
    // as string.
    {
        bool hasDecimalOrE = false;
        for (char c : s) {
            if (c == '.' || c == 'e' || c == 'E') {
                hasDecimalOrE = true;
                break;
            }
        }
        if (hasDecimalOrE) {
            try {
                size_t pos = 0;
                double val = std::stod(s, &pos);
                if (pos == s.size()) {
                    return val;
                }
            } catch (...) {
                // Fall through to string
            }
        }
    }

    // --- Default: string ---
    return s;
}

// ============================================================================
// Helpers: cell reference parsing
// ============================================================================

// Parse "B384" → {colNumber=2, rowNumber=384} (both 1-based)
static std::pair<uint16_t, uint32_t> parseCellRef(const std::string& cellRef) {
    size_t i = 0;
    while (i < cellRef.size() && std::isalpha(static_cast<unsigned char>(cellRef[i]))) {
        i++;
    }
    if (i == 0 || i == cellRef.size()) {
        throw std::invalid_argument("Invalid cell reference: " + cellRef);
    }
    std::string colPart = cellRef.substr(0, i);
    std::string rowPart = cellRef.substr(i);

    uint16_t col = columnLetterToNumber(colPart);
    uint32_t row = static_cast<uint32_t>(std::stoul(rowPart));
    return {col, row};
}

// ============================================================================
// Impl struct — dual backend
// ============================================================================

// Sheet data cache: rows × cols grid (0-based indexing)
using SheetCache = std::vector<std::vector<CellValue>>;

struct ExcelCM::Impl {
    // --- State ---
    std::string filepath;
    bool open = false;
    bool writeMode = false;  // true once any write/create op occurs

    // --- Read backend (xlsxio) — lazy-loaded per-sheet cache ---
    mutable std::map<std::string, SheetCache> sheetCaches;
    mutable std::optional<std::vector<std::string>> sheetNameCache;

    // --- Write backend (OpenXLSX) — opened lazily on first write ---
    std::optional<OpenXLSX::XLDocument> doc;

    // --- Header state ---
    std::vector<std::string> header_;

    // --- Impl helpers ---

    // Load entire sheet into cache via xlsxio streaming (one fast pass)
    void ensureSheetCached(const std::string& sheetName) const {
        if (sheetCaches.count(sheetName)) return;

        if (filepath.empty()) {
            throw std::runtime_error("No file path set for reading");
        }

        log()->debug("xlsxio: streaming sheet '{}' from '{}'",
                       sheetName, filepath);

        xlsxioreader reader = xlsxioread_open(filepath.c_str());
        if (!reader) {
            log()->error("xlsxio: failed to open file: {}", filepath);
            throw std::runtime_error("xlsxio: failed to open file: " + filepath);
        }

        // Use 0 flags (SKIP_NONE) to preserve column positions
        xlsxioreadersheet sheet = xlsxioread_sheet_open(reader, sheetName.c_str(), 0);
        if (!sheet) {
            xlsxioread_close(reader);
            log()->error("xlsxio: sheet not found: {}", sheetName);
            throw std::runtime_error("xlsxio: sheet not found: " + sheetName);
        }

        SheetCache cache;

        while (xlsxioread_sheet_next_row(sheet)) {
            std::vector<CellValue> row;
            char* cellValue = nullptr;

            while ((cellValue = xlsxioread_sheet_next_cell(sheet)) != nullptr) {
                row.push_back(inferCellValue(cellValue));
                free(cellValue);
            }

            cache.push_back(std::move(row));
        }

        xlsxioread_sheet_close(sheet);
        xlsxioread_close(reader);

        log()->debug("xlsxio: cached sheet '{}' — {} rows",
                       sheetName, cache.size());
        if (!cache.empty()) {
            log()->trace("xlsxio: first row has {} columns", cache[0].size());
        }

        sheetCaches[sheetName] = std::move(cache);
    }

    // Populate sheet name cache via xlsxio
    void ensureSheetNamesCached() const {
        if (sheetNameCache.has_value()) return;

        if (filepath.empty()) {
            throw std::runtime_error("No file path set");
        }

        xlsxioreader reader = xlsxioread_open(filepath.c_str());
        if (!reader) {
            throw std::runtime_error("xlsxio: failed to open file: " + filepath);
        }

        std::vector<std::string> names;
        xlsxioreadersheetlist sheetlist = xlsxioread_sheetlist_open(reader);
        if (sheetlist) {
            const char* name;
            while ((name = xlsxioread_sheetlist_next(sheetlist)) != nullptr) {
                names.emplace_back(name);
            }
            xlsxioread_sheetlist_close(sheetlist);
        }
        xlsxioread_close(reader);

        sheetNameCache = std::move(names);
    }

    // Lazy-open OpenXLSX document for write operations
    void ensureDocOpen() {
        if (doc.has_value()) return;
        if (filepath.empty()) {
            throw std::runtime_error("No file path for opening document");
        }
        log()->debug("OpenXLSX: lazy-opening '{}' for write", filepath);
        doc.emplace();
        doc->open(filepath);
        writeMode = true;
    }
};

// ============================================================================
// Constructor / Destructor / Move
// ============================================================================

ExcelCM::ExcelCM() : m_impl(std::make_unique<Impl>()) {}

ExcelCM::~ExcelCM() {
    if (m_impl && m_impl->open) {
        if (m_impl->doc.has_value()) {
            m_impl->doc->close();
        }
    }
}

ExcelCM::ExcelCM(ExcelCM&&) noexcept = default;
ExcelCM& ExcelCM::operator=(ExcelCM&&) noexcept = default;

// ============================================================================
// File operations
// ============================================================================

void ExcelCM::create(const std::string& filepath) {
    if (m_impl->open) close();

    log()->debug("ExcelCM::create({})", filepath);
    m_impl->filepath = filepath;
    m_impl->doc.emplace();
    m_impl->doc->create(filepath, true);
    m_impl->open = true;
    m_impl->writeMode = true;
}

void ExcelCM::open(const std::string& filepath) {
    if (m_impl->open) close();

    log()->debug("ExcelCM::open({})", filepath);
    m_impl->filepath = filepath;
    m_impl->open = true;
    // Neither backend opened yet — deferred to first read/write operation
}

void ExcelCM::save() {
    if (!m_impl->open) throw std::runtime_error("No document open");
    if (!m_impl->doc.has_value()) throw std::runtime_error("No document to save (read-only)");
    log()->debug("ExcelCM::save() — {}", m_impl->filepath);
    m_impl->doc->save();
}

void ExcelCM::saveAs(const std::string& filepath) {
    if (!m_impl->open) throw std::runtime_error("No document open");
    if (!m_impl->doc.has_value()) throw std::runtime_error("No document to save (read-only)");
    m_impl->doc->saveAs(filepath, true);
}

void ExcelCM::close() {
    if (m_impl->open) {
        log()->debug("ExcelCM::close() — {}", m_impl->filepath);
        if (m_impl->doc.has_value()) {
            m_impl->doc->close();
            m_impl->doc.reset();
        }
        m_impl->sheetCaches.clear();
        m_impl->sheetNameCache.reset();
        m_impl->header_.clear();
        m_impl->filepath.clear();
        m_impl->writeMode = false;
        m_impl->open = false;
    }
}

bool ExcelCM::isOpen() const {
    return m_impl->open;
}

// ============================================================================
// Cell operations
// ============================================================================

CellValue ExcelCM::readCell(const std::string& sheetName,
                            const std::string& cellRef) const {
    if (!m_impl->open) throw std::runtime_error("No document open");

    // Write mode: read via OpenXLSX (preserves exact types)
    if (m_impl->writeMode && m_impl->doc.has_value()) {
        auto wks = m_impl->doc->workbook().worksheet(sheetName);
        auto cell = wks.cell(cellRef);
        return xlCellValueToCellValue(cell.value());
    }

    // Read mode: use xlsxio cache
    m_impl->ensureSheetCached(sheetName);

    auto [colNum, rowNum] = parseCellRef(cellRef);
    uint32_t row0 = rowNum - 1;
    uint16_t col0 = colNum - 1;

    const auto& cache = m_impl->sheetCaches.at(sheetName);
    if (row0 >= cache.size()) return std::monostate{};
    if (col0 >= cache[row0].size()) return std::monostate{};
    return cache[row0][col0];
}

void ExcelCM::writeCell(const std::string& sheetName,
                        const std::string& cellRef,
                        const CellValue& value) {
    if (!m_impl->open) throw std::runtime_error("No document open");

    m_impl->ensureDocOpen();

    auto wks = m_impl->doc->workbook().worksheet(sheetName);
    auto cell = wks.cell(cellRef);

    std::visit([&cell](auto&& v) {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            cell.value().clear();
        } else {
            cell.value() = v;
        }
    }, value);
}

void ExcelCM::writeCell(const std::string& sheetName,
                        const std::string& cellRef,
                        const std::string& data) {
    writeCell(sheetName, cellRef, CellValue{data});
}

void ExcelCM::writeCell(const std::string& sheetName,
                        const std::string& cellRef,
                        double data) {
    writeCell(sheetName, cellRef, CellValue{data});
}

void ExcelCM::writeFormula(const std::string& sheetName,
                           const std::string& cellRef,
                           const std::string& formula) {
    // Delegate to writeFormulaCell so callers get a real <f> formula cell.
    // The historical cell.value() = "=..." path produced cells with
    // dual type t="s" + <f>, which Excel rejects as corrupted (and emits
    // spurious "circular reference" warnings for downstream cells).
    // The calcChain dangling-reference issue this used to dodge is now
    // handled by postProcessXlsx (subtotal_post.{h,cpp}).
    writeFormulaCell(sheetName, cellRef, formula);
}

void ExcelCM::writeFormulaCell(const std::string& sheetName,
                               const std::string& cellRef,
                               const std::string& formula) {
    if (!m_impl->open) throw std::runtime_error("No document open");

    m_impl->ensureDocOpen();

    std::string fml = formula;
    if (!fml.empty() && fml[0] == '=') fml = fml.substr(1);
    if (fml.empty()) return;

    auto wks = m_impl->doc->workbook().worksheet(sheetName);
    wks.cell(cellRef).formula() = fml;
}

// ============================================================================
// Sheet helpers
// ============================================================================

bool ExcelCM::hasSheet(const std::string& sheetName) const {
    if (!m_impl->open) return false;

    if (m_impl->writeMode && m_impl->doc.has_value()) {
        return m_impl->doc->workbook().sheetExists(sheetName);
    }

    // Read-only path: check xlsxio sheet names
    m_impl->ensureSheetNamesCached();
    const auto& names = m_impl->sheetNameCache.value();
    return std::find(names.begin(), names.end(), sheetName) != names.end();
}

void ExcelCM::addSheet(const std::string& sheetName) {
    if (!m_impl->open) throw std::runtime_error("No document open");
    m_impl->ensureDocOpen();
    m_impl->doc->workbook().addWorksheet(sheetName);
}

std::vector<std::string> ExcelCM::getSheetNames() const {
    if (!m_impl->open) throw std::runtime_error("No document open");

    if (m_impl->writeMode && m_impl->doc.has_value()) {
        return m_impl->doc->workbook().worksheetNames();
    }

    m_impl->ensureSheetNamesCached();
    return m_impl->sheetNameCache.value();
}

// ============================================================================
// Row / column operations
// ============================================================================

std::vector<CellValue> ExcelCM::readRow(const std::string& sheetName,
                                        uint32_t rowNumber) const {
    if (!m_impl->open) throw std::runtime_error("No document open");
    if (rowNumber == 0) throw std::invalid_argument("Row number must be >= 1");

    // Write mode: read via OpenXLSX
    if (m_impl->writeMode && m_impl->doc.has_value()) {
        auto wks = m_impl->doc->workbook().worksheet(sheetName);
        auto xlRow = wks.row(rowNumber);
        auto colCount = xlRow.cellCount();
        if (colCount == 0) return {};

        std::vector<CellValue> result;
        result.reserve(colCount);
        for (uint16_t col = 1; col <= colCount; ++col) {
            auto cell = wks.cell(OpenXLSX::XLCellReference(rowNumber, col));
            result.push_back(xlCellValueToCellValue(cell.value()));
        }
        return result;
    }

    // Read mode: use xlsxio cache
    m_impl->ensureSheetCached(sheetName);

    const auto& cache = m_impl->sheetCaches.at(sheetName);
    uint32_t row0 = rowNumber - 1;
    if (row0 >= cache.size()) return {};
    return cache[row0];
}

std::vector<CellValue> ExcelCM::readColumn(const std::string& sheetName,
                                           uint16_t columnNumber) const {
    if (!m_impl->open) throw std::runtime_error("No document open");
    if (columnNumber == 0) throw std::invalid_argument("Column number must be >= 1");

    // Write mode: read via OpenXLSX
    if (m_impl->writeMode && m_impl->doc.has_value()) {
        auto wks = m_impl->doc->workbook().worksheet(sheetName);
        auto rowCount = wks.rowCount();
        if (rowCount == 0) return {};

        std::vector<CellValue> result;
        result.reserve(rowCount);
        for (uint32_t row = 1; row <= rowCount; ++row) {
            auto cell = wks.cell(OpenXLSX::XLCellReference(row, columnNumber));
            result.push_back(xlCellValueToCellValue(cell.value()));
        }
        return result;
    }

    // Read mode: extract column from xlsxio cache
    m_impl->ensureSheetCached(sheetName);

    const auto& cache = m_impl->sheetCaches.at(sheetName);
    uint16_t col0 = columnNumber - 1;

    std::vector<CellValue> result;
    result.reserve(cache.size());
    for (const auto& row : cache) {
        if (col0 < row.size()) {
            result.push_back(row[col0]);
        } else {
            result.push_back(std::monostate{});
        }
    }
    return result;
}

std::vector<CellValue> ExcelCM::readColumn(const std::string& sheetName,
                                           const std::string& columnLetter) const {
    return readColumn(sheetName, columnLetterToNumber(columnLetter));
}

// ============================================================================
// Header detection
// ============================================================================

bool ExcelCM::detectHeader(const std::string& sheetName) {
    if (!m_impl->open) throw std::runtime_error("No document open");

    auto row1 = readRow(sheetName, 1);
    if (row1.empty()) {
        m_impl->header_.clear();
        return false;
    }

    int nonEmptyCount = 0;
    int alphaPassCount = 0;

    for (const auto& cell : row1) {
        if (std::holds_alternative<std::monostate>(cell)) continue;
        nonEmptyCount++;

        std::string str = cellValueToString(cell);
        if (str.empty()) continue;

        int totalChars = 0;
        int alphaChars = 0;
        for (char c : str) {
            totalChars++;
            if (std::isalpha(static_cast<unsigned char>(c))) alphaChars++;
        }

        if (totalChars > 0 &&
            static_cast<double>(alphaChars) / totalChars > 0.80) {
            alphaPassCount++;
        }
    }

    if (nonEmptyCount == 0) {
        m_impl->header_.clear();
        return false;
    }

    bool isHeader = static_cast<double>(alphaPassCount) / nonEmptyCount > 0.80;

    if (isHeader) {
        m_impl->header_.clear();
        m_impl->header_.reserve(row1.size());
        for (const auto& cell : row1) {
            if (std::holds_alternative<std::string>(cell)) {
                m_impl->header_.push_back(std::get<std::string>(cell));
            } else {
                m_impl->header_.push_back(cellValueToString(cell));
            }
        }
        log()->debug("detectHeader: {} columns detected in '{}'",
                       m_impl->header_.size(), sheetName);
        log()->trace("detectHeader: headers = [{}]",
                       [&]() {
                           std::string s;
                           for (size_t i = 0; i < m_impl->header_.size(); ++i) {
                               if (i > 0) s += ", ";
                               s += m_impl->header_[i];
                           }
                           return s;
                       }());
    } else {
        m_impl->header_.clear();
        log()->debug("detectHeader: no header detected in '{}'", sheetName);
    }

    return isHeader;
}

bool ExcelCM::hasHeader() const {
    return !m_impl->header_.empty();
}

const std::vector<std::string>& ExcelCM::getHeader() const {
    return m_impl->header_;
}

void ExcelCM::setHeader(const std::vector<std::string>& header) {
    m_impl->header_ = header;
}

std::string ExcelCM::headerToColumnLetter(const std::string& headerValue) const {
    const auto& hdr = m_impl->header_;
    for (size_t i = 0; i < hdr.size(); ++i) {
        if (hdr[i] == headerValue) {
            return columnNumberToLetter(static_cast<uint16_t>(i + 1));
        }
    }
    throw std::invalid_argument("Header not found: " + headerValue);
}

std::string ExcelCM::columnLetterToHeader(const std::string& columnLetter) const {
    uint16_t colNum = columnLetterToNumber(columnLetter);
    const auto& hdr = m_impl->header_;
    size_t index = static_cast<size_t>(colNum) - 1;
    if (index >= hdr.size()) {
        throw std::out_of_range("Column " + columnLetter + " is beyond header range");
    }
    return hdr[index];
}

// ============================================================================
// Utilities (unchanged from original)
// ============================================================================

std::string cellValueToString(const CellValue& value) {
    return std::visit([](auto&& v) -> std::string {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "(empty)";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v;
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? "true" : "false";
        } else {
            return std::to_string(v);
        }
    }, value);
}

uint16_t columnLetterToNumber(const std::string& columnLetter) {
    if (columnLetter.empty())
        throw std::invalid_argument("Column letter cannot be empty");

    uint32_t result = 0;
    for (char c : columnLetter) {
        if (!std::isalpha(static_cast<unsigned char>(c)))
            throw std::invalid_argument("Invalid column letter: " + columnLetter);
        result = result * 26 + (std::toupper(static_cast<unsigned char>(c)) - 'A' + 1);
    }
    return static_cast<uint16_t>(result);
}

std::string columnNumberToLetter(uint16_t columnNumber) {
    if (columnNumber == 0)
        throw std::invalid_argument("Column number must be >= 1");

    std::string result;
    uint32_t n = columnNumber;
    while (n > 0) {
        n--;  // make 0-based
        result = static_cast<char>('A' + (n % 26)) + result;
        n /= 26;
    }
    return result;
}

} // namespace wlcc
