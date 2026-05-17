#include "dataextraction.h"

#include <spdlog/spdlog.h>

#include <algorithm>
#include <cmath>
#include <ctime>
#include <filesystem>
#include <stdexcept>

namespace fs = std::filesystem;

namespace {
spdlog::logger* log() {
    static spdlog::logger* inst = [] {
        auto l = spdlog::get("dataextraction");
        return l ? l.get() : spdlog::default_logger().get();
    }();
    return inst;
}
} // anonymous namespace

namespace wlcc {

DataExtraction::DataExtraction() = default;
DataExtraction::~DataExtraction() = default;

// --- Utility ---

// Convert Excel serial number to DD-MM-YYYY string.
// Excel serial: 1 = 1900-01-01, with Lotus 1-2-3 bug (Feb 29, 1900 is serial 60).
// For dates after Feb 28, 1900 (serial > 60), we convert via Unix epoch:
//   unix_days = serial - 25569  (25569 = Excel serial for 1970-01-01)
static std::string excelSerialToDate(int serial) {
    if (serial < 61) return "";  // dates before 1900-03-01 — not relevant

    // Convert to Unix days, then to time_t
    int unixDays = serial - 25569;
    time_t t = static_cast<time_t>(unixDays) * 86400;

    struct tm result{};
#ifdef _WIN32
    gmtime_s(&result, &t);
#else
    gmtime_r(&t, &result);
#endif

    char buf[16];
    snprintf(buf, sizeof(buf), "%02d-%02d-%04d",
             result.tm_mday, result.tm_mon + 1, result.tm_year + 1900);
    return buf;
}

static bool isDdMonYyyy(const std::string& s) {
    return std::isalpha(static_cast<unsigned char>(s[3]))
        && std::isalpha(static_cast<unsigned char>(s[4]))
        && std::isalpha(static_cast<unsigned char>(s[5]));
    return true;
}

static std::string convertToDdMmYyyy(const std::string& s) {
    // Must be exactly 11 characters: DD-Mon-YYYY
    /*if (s.size() != 11 || s[2] != '-' || s[6] != '-') return s;

    // DD must be digits
    if (!std::isdigit(static_cast<unsigned char>(s[0])) ||
        !std::isdigit(static_cast<unsigned char>(s[1]))) return s;

    // YYYY must be digits
    for (int i = 7; i < 11; ++i) {
        if (!std::isdigit(static_cast<unsigned char>(s[i]))) return s;
    }

    // Mon must be 3 alpha characters
    if (!std::isalpha(static_cast<unsigned char>(s[3])) ||
        !std::isalpha(static_cast<unsigned char>(s[4])) ||
        !std::isalpha(static_cast<unsigned char>(s[5]))) return s;*/

    // Match month abbreviation (case-insensitive)
    std::string mon;
    mon += static_cast<char>(std::tolower(static_cast<unsigned char>(s[3])));
    mon += static_cast<char>(std::tolower(static_cast<unsigned char>(s[4])));
    mon += static_cast<char>(std::tolower(static_cast<unsigned char>(s[5])));

    constexpr const char* months[] = {
        "jan","feb","mar","apr","may","jun",
        "jul","aug","sep","oct","nov","dec"
    };

    for (int i = 0; i < 12; ++i) {
        if (mon == months[i]) {
            char buf[11];
            std::snprintf(buf, sizeof(buf), "%c%c-%02d-%c%c%c%c",
                          s[0], s[1], i + 1, s[7], s[8], s[9], s[10]);
            return buf;
        }
    }

    // Not a valid month — return unchanged
    return s;
}

std::string DataExtraction::extractDatePart(const std::string& dateTimeStr) {
    // Trim leading/trailing whitespace
    auto start = dateTimeStr.find_first_not_of(" \t");
    if (start == std::string::npos) return "";
    auto end = dateTimeStr.find_last_not_of(" \t");
    std::string trimmed = dateTimeStr.substr(start, end - start + 1);

    // Check if this is a numeric string (Excel serial number).
    // xlsxio returns dates as raw numbers like "45962" (int) or "45962.791667" (double).
    // Detect: starts with digit, contains only digits and optionally one '.'
    if (!trimmed.empty() && std::isdigit(static_cast<unsigned char>(trimmed[0]))) {
        bool isNumeric = true;
        bool hasDot = false;
        for (char c : trimmed) {
            if (c == '.' && !hasDot) { hasDot = true; continue; }
            if (!std::isdigit(static_cast<unsigned char>(c))) { isNumeric = false; break; }
        }
        if (isNumeric) {
            // Take integer part as the serial number (fractional part is time-of-day).
            // Use stoll to avoid overflow on long digit strings (e.g. phone numbers),
            // then reject values outside plausible Excel date range (366 .. 2958465).
            // 2958465 = Excel serial for 9999-12-31.
            try {
                long long val = std::stoll(trimmed);
                if (val > 365 && val <= 2958465) {
                    std::string converted = excelSerialToDate(static_cast<int>(val));
                    if (!converted.empty()) {
                        return converted;
                    }
                }
            } catch (...) {
                // Overflow or parse error — not a date serial, fall through
            }
        }
    }

    // Extract everything before the first space (date portion)
    auto spacePos = trimmed.find(' ');
    if (spacePos != std::string::npos) {
        return trimmed.substr(0, spacePos);
    }
    return trimmed;
}

// --- Excel: load once, query many ---

void DataExtraction::loadExcel(const std::string& filepath,
                               const std::string& sheetName) {
    // Close previous file if any
    if (excel_.isOpen()) {
        excel_.close();
    }
    excelLoaded_ = false;

    log()->info("DataExtraction: loading Excel '{}' sheet '{}'",
                  filepath, sheetName);
    excel_.open(filepath);

    if (!excel_.detectHeader(sheetName)) {
        excel_.close();
        log()->error("DataExtraction: no header row in sheet '{}'", sheetName);
        throw std::runtime_error("No header row detected in sheet: " + sheetName);
    }

    excelSheet_ = sheetName;
    excelLoaded_ = true;
    log()->debug("DataExtraction: Excel loaded, {} header columns",
                   excel_.getHeader().size());
}

bool DataExtraction::isExcelLoaded() const {
    return excelLoaded_;
}

ExcelRecordsByDate DataExtraction::extractExcelRowsByDate(
    const std::string& dateHeaderName) const
{
    if (!excelLoaded_) {
        throw std::runtime_error("No Excel file loaded. Call loadExcel() first.");
    }

    const auto& headers = excel_.getHeader();

    // Resolve the date column letter (throws if header not found)
    std::string dateColLetter = excel_.headerToColumnLetter(dateHeaderName);

    // Read the entire date column to determine row count
    auto dateCol = excel_.readColumn(excelSheet_, dateColLetter);

    ExcelRecordsByDate result;

    log()->debug("extractExcelRowsByDate: date column '{}' has {} rows",
                   dateHeaderName, dateCol.size());

    // dateCol[0] is the header cell (row 1); data starts at index 1
    for (size_t i = 1; i < dateCol.size(); ++i) {
        if (std::holds_alternative<std::monostate>(dateCol[i])) continue;

        // Read the full row (1-based: i+1 because index 0 = row 1 = header)
        uint32_t rowNumber = static_cast<uint32_t>(i + 1);
        auto rowValues = excel_.readRow(excelSheet_, rowNumber);

        // Build ExcelRecord from header names + cell values
        ExcelRecord record;
        for (size_t col = 0; col < headers.size() && col < rowValues.size(); ++col) {
            record[headers[col]] = rowValues[col];
        }

        // Extract date string from this row's date cell
        std::string dateStr = extractDatePart(cellValueToString(dateCol[i]));
        if (isDdMonYyyy(dateStr)) {
            dateStr = convertToDdMmYyyy(dateStr);
        }
        result[dateStr].push_back(std::move(record));
    }

    log()->debug("extractExcelRowsByDate: {} unique dates found",
                   result.size());

    return result;
}

std::vector<ExcelRecord> DataExtraction::getExcelRowsForDate(
    const std::string& dateHeaderName,
    const std::string& date) const
{
    if (!excelLoaded_) {
        throw std::runtime_error("No Excel file loaded. Call loadExcel() first.");
    }

    const auto& headers = excel_.getHeader();
    std::string dateColLetter = excel_.headerToColumnLetter(dateHeaderName);
    auto dateCol = excel_.readColumn(excelSheet_, dateColLetter);

    std::vector<ExcelRecord> result;

    for (size_t i = 1; i < dateCol.size(); ++i) {
        if (std::holds_alternative<std::monostate>(dateCol[i])) continue;

        std::string dateStr = extractDatePart(cellValueToString(dateCol[i]));
        if (dateStr != date) continue;

        uint32_t rowNumber = static_cast<uint32_t>(i + 1);
        auto rowValues = excel_.readRow(excelSheet_, rowNumber);

        ExcelRecord record;
        for (size_t col = 0; col < headers.size() && col < rowValues.size(); ++col) {
            record[headers[col]] = rowValues[col];
        }
        result.push_back(std::move(record));
    }

    return result;
}

// --- CDR: load once, query many ---

void DataExtraction::loadCdrDirectories(
    const std::vector<std::string>& directories,
    const std::string& dateHeaderName)
{
    cdrData_.clear();
    cdrLoaded_ = false;

    int fileCount = 0;
    int totalRecords = 0;

    for (const auto& dirPath : directories) {
        if (!fs::exists(dirPath) || !fs::is_directory(dirPath)) {
            log()->trace("CDR: skipping non-existent directory: {}", dirPath);
            continue;
        }

        for (const auto& entry : fs::directory_iterator(dirPath)) {
            if (!entry.is_regular_file()) continue;
            if (entry.path().extension() != ".txt") continue;

            CdrCC cdr;
            try {
                cdr.load(entry.path().string());
            } catch (...) {
                log()->trace("CDR: failed to parse: {}",
                               entry.path().string());
                continue; // Skip files that fail to parse
            }

            fileCount++;
            for (const auto& record : cdr.getAllRecords()) {
                auto it = record.find(dateHeaderName);
                if (it == record.end() || it->second.empty()) continue;

                std::string dateStr = extractDatePart(it->second);
                cdrData_[dateStr].push_back(record);
                totalRecords++;
            }
        }
    }

    cdrLoaded_ = true;
    log()->debug("CDR: loaded {} records from {} files across {} directories",
                   totalRecords, fileCount, directories.size());
}

bool DataExtraction::isCdrLoaded() const {
    return cdrLoaded_;
}

const CdrRecordsByDate& DataExtraction::getCdrRecordsByDate() const {
    if (!cdrLoaded_) {
        throw std::runtime_error(
            "No CDR data loaded. Call loadCdrDirectories() first.");
    }
    return cdrData_;
}

std::vector<CdrRecord> DataExtraction::getCdrRecordsForDate(
    const std::string& date) const
{
    if (!cdrLoaded_) {
        throw std::runtime_error(
            "No CDR data loaded. Call loadCdrDirectories() first.");
    }

    auto it = cdrData_.find(date);
    if (it != cdrData_.end()) {
        return it->second;
    }
    return {};
}

} // namespace wlcc
