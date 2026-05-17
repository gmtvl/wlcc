// wdatagen.exe — Generate reduced test dataset from full production data.
//
// Reads the full WSCC CDR Excel (~457K rows), samples 100 rows per date
// (30 days → 3000 rows), writes a reduced Excel file, creates 10 matching
// MSC CDR text files, and creates a corresponding reduced CRM Excel file.
//
// Usage:
//   wdatagen.exe -p <path>                  (auto-detect month)
//   wdatagen.exe -p <path> -m nov           (explicit month)
//   wdatagen.exe -p <path> -n 50            (50 rows per date instead of 100)
//   wdatagen.exe -p <path> -o <output_dir>  (output to specific directory)

#include "excel_cm.h"
#include "inputvalidator.h"
#include "cdr_cc.h"

#include <OpenXLSX.hpp>

#include <algorithm>
#include <cstdio>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <numeric>
#include <random>
#include <set>
#include <sstream>
#include <string>
#include <vector>

namespace fs = std::filesystem;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

static std::string cellToString(const wlcc::CellValue& val) {
    return wlcc::cellValueToString(val);
}

// Convert Excel date serial number to DD-MM-YYYY string
static std::string excelSerialToDateStr(int serial) {
    if (serial < 1) return "";
    // Excel serial → Unix days (serial 25569 = 1970-01-01)
    int unix_days = serial - 25569;
    time_t t = static_cast<time_t>(unix_days) * 86400;
    struct tm* tm = gmtime(&t);
    if (!tm) return "";
    char buf[16];
    std::snprintf(buf, sizeof(buf), "%02d-%02d-%04d",
                  tm->tm_mday, tm->tm_mon + 1, tm->tm_year + 1900);
    return buf;
}

// Convert DD-MM-YYYY to DD-Mon-YYYY (for CDR Period line)
static std::string toPeriodDate(const std::string& ddmmyyyy) {
    static const char* months[] = {
        "Jan","Feb","Mar","Apr","May","Jun",
        "Jul","Aug","Sep","Oct","Nov","Dec"
    };
    if (ddmmyyyy.size() < 10) return ddmmyyyy;
    int m = std::stoi(ddmmyyyy.substr(3, 2));
    if (m < 1 || m > 12) return ddmmyyyy;
    return ddmmyyyy.substr(0, 3) + months[m-1] + ddmmyyyy.substr(5);
}

// Extract date string from a CellValue (handles int64_t serial, double serial,
// or string dates)
static std::string extractDate(const wlcc::CellValue& val) {
    if (std::holds_alternative<int64_t>(val)) {
        return excelSerialToDateStr(static_cast<int>(std::get<int64_t>(val)));
    }
    if (std::holds_alternative<double>(val)) {
        return excelSerialToDateStr(static_cast<int>(std::get<double>(val)));
    }
    if (std::holds_alternative<std::string>(val)) {
        const auto& s = std::get<std::string>(val);
        auto sp = s.find(' ');
        return (sp != std::string::npos) ? s.substr(0, sp) : s;
    }
    return "";
}

// Convert Excel time (fractional day as double, or int seconds, or HH:MM:SS
// string) to HH:MM:SS string
static std::string timeToString(const wlcc::CellValue& val) {
    if (std::holds_alternative<double>(val)) {
        double frac = std::get<double>(val);
        if (frac > 1.0) frac -= static_cast<int>(frac);
        int total = static_cast<int>(frac * 86400.0 + 0.5);
        int h = total / 3600;
        int m = (total % 3600) / 60;
        int s = total % 60;
        char buf[16];
        std::snprintf(buf, sizeof(buf), "%02d:%02d:%02d", h, m, s);
        return buf;
    }
    if (std::holds_alternative<int64_t>(val)) {
        // Could be seconds from midnight
        int total = static_cast<int>(std::get<int64_t>(val));
        if (total >= 0 && total < 86400) {
            int h = total / 3600;
            int m = (total % 3600) / 60;
            int s = total % 60;
            char buf[16];
            std::snprintf(buf, sizeof(buf), "%02d:%02d:%02d", h, m, s);
            return buf;
        }
        return std::to_string(total);
    }
    if (std::holds_alternative<std::string>(val)) {
        return std::get<std::string>(val);
    }
    return "00:00:00";
}

// Get integer value from CellValue
static int cellToInt(const wlcc::CellValue& val) {
    if (std::holds_alternative<int64_t>(val))
        return static_cast<int>(std::get<int64_t>(val));
    if (std::holds_alternative<double>(val))
        return static_cast<int>(std::get<double>(val));
    if (std::holds_alternative<std::string>(val)) {
        try { return std::stoi(std::get<std::string>(val)); }
        catch (...) { return 0; }
    }
    return 0;
}

// Lookup helper: find column index by header name
static int findCol(const std::vector<std::string>& headers,
                   const std::string& name) {
    for (size_t i = 0; i < headers.size(); ++i) {
        if (headers[i] == name) return static_cast<int>(i);
    }
    return -1;
}

// ---------------------------------------------------------------------------
// Command-line parsing
// ---------------------------------------------------------------------------

struct Config {
    std::string basePath = ".";
    std::string monthCode;
    int rowsPerDate = 100;
    std::string outputDir;  // empty = same as input

    bool parse(int argc, char* argv[]) {
        for (int i = 1; i < argc; ++i) {
            std::string arg = argv[i];
            if ((arg == "-p" || arg == "--path") && i + 1 < argc)
                basePath = argv[++i];
            else if ((arg == "-m" || arg == "--month") && i + 1 < argc) {
                monthCode = argv[++i];
                monthCode = monthCode.substr(0, 3);
                std::transform(monthCode.begin(), monthCode.end(),
                               monthCode.begin(), ::tolower);
            }
            else if ((arg == "-n" || arg == "--rows") && i + 1 < argc)
                rowsPerDate = std::stoi(argv[++i]);
            else if ((arg == "-o" || arg == "--output") && i + 1 < argc)
                outputDir = argv[++i];
            else if (arg == "-h" || arg == "--help") {
                std::cerr << "Usage: " << argv[0]
                    << " [-p <path>] [-m <month>] [-n <rows_per_date>]"
                    << " [-o <output_dir>]\n";
                return false;
            }
            else {
                std::cerr << "Error: unknown option '" << arg << "'.\n";
                return false;
            }
        }
        return true;
    }
};

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

int main(int argc, char* argv[]) {
    Config cfg;
    if (!cfg.parse(argc, argv)) return 1;

    // Validate directory
    wlcc::InputValidator validator;

    // Auto-detect month if not given
    if (cfg.monthCode.empty()) {
        for (const auto& entry : fs::directory_iterator(cfg.basePath)) {
            if (!entry.is_directory()) continue;
            std::string name = entry.path().filename().string();
            std::string lower = name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (lower.size() >= 13 && lower.substr(0, 10) == "datafiles_") {
                cfg.monthCode = lower.substr(10, 3);
                break;
            }
        }
        if (cfg.monthCode.empty()) {
            std::cerr << "Error: cannot auto-detect month. Use -m <month>.\n";
            return 1;
        }
    }

    if (!validator.validate(cfg.basePath, cfg.monthCode)) {
        std::cerr << "Validation failed:\n";
        for (const auto& err : validator.getErrors())
            std::cerr << "  - " << err << "\n";
        return 1;
    }

    // Determine output directory
    std::string outBase = cfg.outputDir.empty()
        ? validator.getDataFilesPath()
        : cfg.outputDir;

    std::cout << "=== wdatagen — Test Data Generator ===\n";
    std::cout << "Source folder : " << validator.getDataFilesPath() << "\n";
    std::cout << "Output folder : " << outBase << "\n";
    std::cout << "Rows per date : " << cfg.rowsPerDate << "\n\n";

    // ======================================================================
    // Step 1: Read WSCC CDR Excel, sample rows, write reduced file
    // ======================================================================

    std::cout << "--- Step 1: Reduce WSCC CDR Excel ---\n";

    wlcc::ExcelCM wsccExcel;
    wsccExcel.open(validator.getWsccCdrFile().fullPath);
    auto sheetNames = wsccExcel.getSheetNames();
    if (sheetNames.empty()) {
        std::cerr << "Error: WSCC CDR Excel has no sheets.\n";
        return 1;
    }
    std::string wsccSheet = sheetNames[0];
    std::cout << "  Sheet: " << wsccSheet << "\n";

    if (!wsccExcel.detectHeader(wsccSheet)) {
        std::cerr << "Error: no header detected.\n";
        return 1;
    }
    const auto& headers = wsccExcel.getHeader();
    std::cout << "  Headers: " << headers.size() << " columns\n";

    int dateCol = findCol(headers, "Date");
    int cliCol  = findCol(headers, "CLI");
    int ivrStartCol = findCol(headers, "IVRStartTime");
    int totalDurCol = findCol(headers, "TotalDuration");
    int locationCol = findCol(headers, "Location");
    int levelCol    = findCol(headers, "level");
    if (levelCol < 0) levelCol = findCol(headers, "Level");

    if (dateCol < 0 || cliCol < 0) {
        std::cerr << "Error: required columns Date/CLI not found.\n";
        return 1;
    }

    // Read all rows, group by date
    auto dateColumn = wsccExcel.readColumn(wsccSheet,
        static_cast<uint16_t>(dateCol + 1));

    std::map<std::string, std::vector<uint32_t>> rowsByDate;  // date → row numbers (1-based)

    for (size_t i = 1; i < dateColumn.size(); ++i) {
        std::string dateStr = extractDate(dateColumn[i]);
        if (!dateStr.empty()) {
            rowsByDate[dateStr].push_back(static_cast<uint32_t>(i + 1));
        }
    }

    std::cout << "  Unique dates: " << rowsByDate.size() << "\n";

    // Sample rows
    std::mt19937 rng(42);  // fixed seed for reproducibility
    std::vector<uint32_t> sampledRows;

    for (auto& [dateStr, rows] : rowsByDate) {
        int n = std::min(cfg.rowsPerDate, static_cast<int>(rows.size()));
        std::shuffle(rows.begin(), rows.end(), rng);
        rows.resize(n);
        std::sort(rows.begin(), rows.end());
        sampledRows.insert(sampledRows.end(), rows.begin(), rows.end());
        std::cout << "  " << dateStr << ": " << n << " rows sampled\n";
    }

    std::sort(sampledRows.begin(), sampledRows.end());
    std::cout << "  Total sampled: " << sampledRows.size() << "\n";

    // Read sampled rows into memory
    struct SampledRecord {
        std::vector<wlcc::CellValue> cells;
        std::string dateStr;
        std::string cli;       // full CLI string
        std::string last10;    // last 10 digits
        std::string timeStr;   // HH:MM:SS
        int duration = 0;
        std::string location;
    };

    std::vector<SampledRecord> records;
    records.reserve(sampledRows.size());

    for (uint32_t rowNum : sampledRows) {
        auto row = wsccExcel.readRow(wsccSheet, rowNum);
        SampledRecord rec;
        rec.cells = row;

        if (dateCol < static_cast<int>(row.size()))
            rec.dateStr = extractDate(row[dateCol]);
        if (cliCol < static_cast<int>(row.size())) {
            rec.cli = cellToString(row[cliCol]);
            if (rec.cli.size() >= 10)
                rec.last10 = rec.cli.substr(rec.cli.size() - 10);
            else
                rec.last10 = rec.cli;
        }
        if (ivrStartCol >= 0 && ivrStartCol < static_cast<int>(row.size()))
            rec.timeStr = timeToString(row[ivrStartCol]);
        if (totalDurCol >= 0 && totalDurCol < static_cast<int>(row.size()))
            rec.duration = cellToInt(row[totalDurCol]);
        if (locationCol >= 0 && locationCol < static_cast<int>(row.size()))
            rec.location = cellToString(row[locationCol]);

        records.push_back(std::move(rec));
    }

    wsccExcel.close();

    // Write reduced Excel file
    std::string outWsccDir = (fs::path(outBase) / "WSCC CDR").string();
    fs::create_directories(outWsccDir);
    std::string outWsccPath = (fs::path(outWsccDir) /
        ("CDR-" + std::string(1, std::toupper(cfg.monthCode[0]))
         + std::string(1, std::toupper(cfg.monthCode[1]))
         + std::string(1, std::toupper(cfg.monthCode[2]))
         + " 2025 XL-small.xlsx")).string();

    std::cout << "\n  Writing: " << outWsccPath << "\n";

    // Strategy: copy the entire source file (preserves ALL styles, formats,
    // column widths, number formats for dates/times/booleans), then delete
    // the unwanted rows using OpenXLSX::deleteRow() from bottom to top.
    // This is the only approach that truly preserves cell formatting.
    fs::copy_file(validator.getWsccCdrFile().fullPath, outWsccPath,
                  fs::copy_options::overwrite_existing);

    {
        OpenXLSX::XLDocument doc;
        doc.open(outWsccPath);
        auto wks = doc.workbook().worksheet(wsccSheet);

        uint32_t totalRows = wks.rowCount();
        std::cout << "  Source has " << totalRows << " rows\n";

        // Build set of row numbers to KEEP (1-based; row 1 = header always kept)
        std::set<uint32_t> keepSet(sampledRows.begin(), sampledRows.end());
        keepSet.insert(1);  // always keep header

        // Collect rows to DELETE (all data rows NOT in keepSet)
        std::vector<uint32_t> deleteRows;
        for (uint32_t r = 2; r <= totalRows; ++r) {
            if (!keepSet.count(r)) {
                deleteRows.push_back(r);
            }
        }

        std::cout << "  Keeping " << keepSet.size() << " rows (1 header + "
                  << sampledRows.size() << " data), deleting "
                  << deleteRows.size() << " rows...\n";

        // Delete from bottom to top so row numbers above don't shift
        for (auto it = deleteRows.rbegin(); it != deleteRows.rend(); ++it) {
            wks.deleteRow(*it);
        }

        doc.save();
        doc.close();
    }

    std::cout << "  Done. " << sampledRows.size() << " data rows kept "
              << "(header + data, all formats preserved).\n";

    // ======================================================================
    // Step 2: Create 10 matching MSC CDR text files
    // ======================================================================

    std::cout << "\n--- Step 2: Create MSC CDR text files ---\n";

    std::string outCdrDir = (fs::path(outBase) / "MSC CDR" /
        "TestCDR_generated").string();
    fs::create_directories(outCdrDir);

    static const char* calledNumbers[] = {
        "911503", "911507", "919444024365", "9118001801503"
    };

    int numFiles = 10;
    int chunkSize = (static_cast<int>(records.size()) + numFiles - 1) / numFiles;

    // Shuffle a copy for distribution across files
    std::vector<size_t> indices(records.size());
    std::iota(indices.begin(), indices.end(), 0);
    std::shuffle(indices.begin(), indices.end(), rng);

    for (int f = 0; f < numFiles; ++f) {
        // Collect this file's records
        int start = f * chunkSize;
        int end = std::min(start + chunkSize, static_cast<int>(indices.size()));
        if (start >= end) break;

        // Gather and sort by date+time
        std::vector<size_t> chunk(indices.begin() + start,
                                  indices.begin() + end);
        std::sort(chunk.begin(), chunk.end(), [&](size_t a, size_t b) {
            if (records[a].dateStr != records[b].dateStr)
                return records[a].dateStr < records[b].dateStr;
            return records[a].timeStr < records[b].timeStr;
        });

        // Find date range
        std::string firstDate = records[chunk.front()].dateStr;
        std::string lastDate  = records[chunk.back()].dateStr;

        std::string filename = "TestCDR_MTC & MOC_"
            + toPeriodDate(firstDate) + "_"
            + toPeriodDate(lastDate) + "_MSC_"
            + std::to_string(f + 1) + ".txt";
        std::string filepath = (fs::path(outCdrDir) / filename).string();

        std::ofstream ofs(filepath, std::ios::binary);
        if (!ofs) {
            std::cerr << "  Error: cannot create " << filepath << "\n";
            continue;
        }

        // Metadata header
        ofs << "For  MSC            : TEST_GMSC\r\n";
        ofs << "For the Trunk         : TEST TRUNK TO NGC CALL CENTER (TESTTRUNK_NGC_CC)\r\n";
        ofs << "Period              : " << toPeriodDate(firstDate)
            << " To " << toPeriodDate(lastDate) << "\r\n";
        ofs << "Export CDR          : CALL\r\n";
        ofs << "Call Direction      : MTC & MOC\r\n";
        ofs << " \r\n";

        // Column header
        ofs << "Call Type   Calling Number\t\tCalled Number\t\t"
            << "    Date & Time\t\t      Duration\r\n";

        // Separator
        for (int i = 0; i < 150; ++i) ofs << '-';
        ofs << "\r\n";

        // Data rows
        for (size_t idx : chunk) {
            const auto& rec = records[idx];
            // Convert DD-MM-YYYY to DD-MM-YYYY format for CDR
            std::string dtStr = rec.dateStr + " " + rec.timeStr;
            const char* called = calledNumbers[idx % 4];

            // Pad fields to match real CDR format
            char line[256];
            std::snprintf(line, sizeof(line),
                "50\t\t%-16s\t%-16s\t%s\t\t%d\r\n",
                rec.cli.c_str(), called, dtStr.c_str(), rec.duration);
            ofs << line;
        }

        // Footer
        for (int i = 0; i < 150; ++i) ofs << '-';
        ofs << "\r\n";
        ofs << "  \r\n";

        ofs.close();
        std::cout << "  Written: " << filename << " (" << chunk.size()
                  << " records)\n";
    }

    // ======================================================================
    // Step 3: Create reduced CRM Excel file
    // ======================================================================

    std::cout << "\n--- Step 3: Create reduced CRM Excel ---\n";

    // Read CRM file
    wlcc::ExcelCM crmExcel;
    crmExcel.open(validator.getCrmFile().fullPath);
    auto crmSheets = crmExcel.getSheetNames();
    if (crmSheets.empty()) {
        std::cerr << "  Error: CRM Excel has no sheets.\n";
        return 1;
    }
    std::string crmSheet = crmSheets[0];

    if (!crmExcel.detectHeader(crmSheet)) {
        std::cerr << "  Error: no header in CRM file.\n";
        return 1;
    }
    const auto& crmHeaders = crmExcel.getHeader();
    std::cout << "  CRM sheet: " << crmSheet << " ("
              << crmHeaders.size() << " columns)\n";

    int crmMobileCol   = findCol(crmHeaders, "mobileNo");
    int crmCreatedCol  = findCol(crmHeaders, "createdOn");
    int crmSourceCol   = findCol(crmHeaders, "source");

    if (crmMobileCol < 0 || crmCreatedCol < 0) {
        std::cerr << "  Error: mobileNo/createdOn columns not found.\n";
        return 1;
    }

    // Build set of (last10, date) pairs from WSCC sampled records
    // so we can find matching CRM rows
    std::set<std::pair<std::string, std::string>> wsccKeys;
    for (const auto& rec : records) {
        if (!rec.last10.empty() && !rec.dateStr.empty()) {
            wsccKeys.insert({rec.last10, rec.dateStr});
        }
    }
    std::cout << "  WSCC unique (CLI, date) pairs: " << wsccKeys.size() << "\n";

    // Scan CRM rows for matches
    auto crmDateColumn = crmExcel.readColumn(crmSheet,
        static_cast<uint16_t>(crmCreatedCol + 1));
    auto crmMobileColumn = crmExcel.readColumn(crmSheet,
        static_cast<uint16_t>(crmMobileCol + 1));

    std::vector<uint32_t> matchedCrmRows;

    for (size_t i = 1; i < crmDateColumn.size() && i < crmMobileColumn.size(); ++i) {
        std::string dateStr = extractDate(crmDateColumn[i]);

        // CRM dates are in "DD-Mon-YYYY HH:MM AM/PM" format — extract DD-MM-YYYY
        // Try parsing "01-Nov-2025 12:00 AM" format
        if (dateStr.size() >= 11 && std::isalpha(static_cast<unsigned char>(dateStr[3]))) {
            // DD-Mon-YYYY format
            static const char* months[] = {
                "Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"
            };
            std::string monStr = dateStr.substr(3, 3);
            int mon = 0;
            for (int m = 0; m < 12; ++m) {
                if (monStr == months[m]) { mon = m + 1; break; }
            }
            if (mon > 0) {
                dateStr = dateStr.substr(0, 3)
                    + (mon < 10 ? "0" : "") + std::to_string(mon)
                    + dateStr.substr(6, 5);  // -YYYY
            }
        }

        std::string mobile = cellToString(crmMobileColumn[i]);
        std::string last10;
        if (mobile.size() >= 10)
            last10 = mobile.substr(mobile.size() - 10);
        else
            last10 = mobile;

        if (wsccKeys.count({last10, dateStr})) {
            matchedCrmRows.push_back(static_cast<uint32_t>(i + 1));
        }
    }

    std::cout << "  CRM rows matched by (mobile, date): "
              << matchedCrmRows.size() << "\n";

    // If not enough matches, also generate synthetic CRM rows for unmatched
    // WSCC records
    std::set<std::pair<std::string, std::string>> coveredKeys;
    for (uint32_t rowNum : matchedCrmRows) {
        auto row = crmExcel.readRow(crmSheet, rowNum);
        std::string mobile = cellToString(row[crmMobileCol]);
        std::string last10 = (mobile.size() >= 10)
            ? mobile.substr(mobile.size() - 10) : mobile;
        std::string dateStr = extractDate(row[crmCreatedCol]);
        // Re-parse date like above
        if (dateStr.size() >= 11 && std::isalpha(static_cast<unsigned char>(dateStr[3]))) {
            static const char* months[] = {
                "Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"
            };
            std::string monStr = dateStr.substr(3, 3);
            int mon = 0;
            for (int m = 0; m < 12; ++m) {
                if (monStr == months[m]) { mon = m + 1; break; }
            }
            if (mon > 0) {
                dateStr = dateStr.substr(0, 3)
                    + (mon < 10 ? "0" : "") + std::to_string(mon)
                    + dateStr.substr(6, 5);
            }
        }
        coveredKeys.insert({last10, dateStr});
    }

    crmExcel.close();  // done reading CRM via xlsxio

    // Find WSCC records that have no CRM match — we'll generate synthetic CRM rows
    static const char* types[] = {"PREPAID", "POSTPAID"};
    static const char* subTypes[] = {"SERVICES", "BILLING", "NETWORK"};
    static const char* categories[] = {"DATA SERVICES", "RECHARGE", "DND", "VAS"};
    static const char* subCategories[] = {
        "UNABLE TO BROWSE DATA", "RECHARGE ISSUE", "COMPLAINT BOOKING IN DND SITE",
        "VAS DEACTIVATION"
    };
    static const char* states[] = {"KERALA", "TAMILNADU"};

    // Month names for CRM date format
    static const char* monthNames[] = {
        "Jan","Feb","Mar","Apr","May","Jun",
        "Jul","Aug","Sep","Oct","Nov","Dec"
    };

    // Convert DD-MM-YYYY to DD-Mon-YYYY HH:MM AM format
    auto toCrmDate = [&](const std::string& ddmmyyyy,
                         const std::string& timeStr) -> std::string {
        if (ddmmyyyy.size() < 10) return ddmmyyyy;
        int mon = std::stoi(ddmmyyyy.substr(3, 2));
        if (mon < 1 || mon > 12) return ddmmyyyy;
        // Parse time to get hour for AM/PM
        int h = 0, m = 0;
        if (timeStr.size() >= 5) {
            h = std::stoi(timeStr.substr(0, 2));
            m = std::stoi(timeStr.substr(3, 2));
        }
        std::string ampm = (h < 12) ? "AM" : "PM";
        if (h == 0) h = 12;
        else if (h > 12) h -= 12;

        char buf[32];
        std::snprintf(buf, sizeof(buf), "%s%s%s %d:%02d %s",
            ddmmyyyy.substr(0, 3).c_str(),
            monthNames[mon-1],
            ddmmyyyy.substr(5).c_str(),
            h, m, ampm.c_str());
        return buf;
    };

    // Collect synthetic CRM rows
    struct SyntheticCrmRow {
        std::string mobile;
        std::string createdOn;
        std::string state;
        std::string type;
        std::string subType;
        std::string category;
        std::string subCategory;
    };

    std::vector<SyntheticCrmRow> syntheticRows;
    for (const auto& rec : records) {
        if (!coveredKeys.count({rec.last10, rec.dateStr})) {
            SyntheticCrmRow syn;
            syn.mobile = rec.cli;
            syn.createdOn = toCrmDate(rec.dateStr, rec.timeStr);
            syn.state = (rec.location.find("Kerala") != std::string::npos)
                        ? "KERALA" : "TAMILNADU";
            size_t hash = std::hash<std::string>{}(rec.cli + rec.dateStr);
            syn.type = types[hash % 2];
            syn.subType = subTypes[hash % 3];
            syn.category = categories[hash % 4];
            syn.subCategory = subCategories[hash % 4];
            syntheticRows.push_back(std::move(syn));
        }
    }

    std::cout << "  Synthetic CRM rows needed: " << syntheticRows.size() << "\n";

    // Write reduced CRM file
    std::string outCrmPath = (fs::path(outBase) /
        ("CRM-" + std::string(1, std::toupper(cfg.monthCode[0]))
         + std::string(1, std::toupper(cfg.monthCode[1]))
         + std::string(1, std::toupper(cfg.monthCode[2]))
         + "-small.xlsx")).string();

    std::cout << "  Writing: " << outCrmPath << "\n";

    // Same strategy as WSCC: copy source, delete unwanted rows with deleteRow()
    fs::copy_file(validator.getCrmFile().fullPath, outCrmPath,
                  fs::copy_options::overwrite_existing);

    {
        OpenXLSX::XLDocument crmDoc;
        crmDoc.open(outCrmPath);
        auto crmWks = crmDoc.workbook().worksheet(crmSheet);
        uint32_t crmTotalRows = crmWks.rowCount();

        // Build set of rows to keep (header + matched CRM rows)
        std::set<uint32_t> crmKeepSet(matchedCrmRows.begin(),
                                       matchedCrmRows.end());
        crmKeepSet.insert(1);  // header

        // Collect rows to delete
        std::vector<uint32_t> crmDeleteRows;
        for (uint32_t r = 2; r <= crmTotalRows; ++r) {
            if (!crmKeepSet.count(r)) {
                crmDeleteRows.push_back(r);
            }
        }

        std::cout << "  CRM: keeping " << crmKeepSet.size()
                  << " rows, deleting " << crmDeleteRows.size() << "...\n";

        // Delete from bottom to top
        for (auto it = crmDeleteRows.rbegin(); it != crmDeleteRows.rend(); ++it) {
            crmWks.deleteRow(*it);
        }

        // After deletion, matched rows are compacted to rows 2..N+1.
        // Now append synthetic CRM rows at the end.
        uint32_t outRow = static_cast<uint32_t>(matchedCrmRows.size() + 2);

        for (const auto& syn : syntheticRows) {
            if (crmMobileCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    crmMobileCol + 1)).value() = syn.mobile;
            if (crmCreatedCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    crmCreatedCol + 1)).value() = syn.createdOn;
            if (crmSourceCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    crmSourceCol + 1)).value() = std::string("INBOUND");

            int statusCol = findCol(crmHeaders, "status");
            if (statusCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    statusCol + 1)).value() = std::string("CLOSED");

            int stateCol = findCol(crmHeaders, "state");
            if (stateCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    stateCol + 1)).value() = syn.state;

            int circleCol = findCol(crmHeaders, "circle");
            if (circleCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    circleCol + 1)).value() = syn.state;

            int typeCol = findCol(crmHeaders, "type");
            if (typeCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    typeCol + 1)).value() = syn.type;

            int subTypeCol = findCol(crmHeaders, "subType");
            if (subTypeCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    subTypeCol + 1)).value() = syn.subType;

            int catCol = findCol(crmHeaders, "category");
            if (catCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    catCol + 1)).value() = syn.category;

            int subCatCol = findCol(crmHeaders, "subCategory");
            if (subCatCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    subCatCol + 1)).value() = syn.subCategory;

            int callTypeCol = findCol(crmHeaders, "callType");
            if (callTypeCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    callTypeCol + 1)).value() = std::string("QUERY");

            int dueCol = findCol(crmHeaders, "dueDate");
            if (dueCol >= 0)
                crmWks.cell(OpenXLSX::XLCellReference(outRow,
                    dueCol + 1)).value() = syn.createdOn;

            outRow++;
        }

        crmDoc.save();
        crmDoc.close();
    }

    size_t totalCrmRows = matchedCrmRows.size() + syntheticRows.size();
    std::cout << "  Done. " << totalCrmRows << " CRM rows written ("
              << matchedCrmRows.size() << " matched + "
              << syntheticRows.size() << " synthetic).\n";

    // ======================================================================
    // Summary
    // ======================================================================

    std::cout << "\n=== Summary ===\n";
    std::cout << "  WSCC CDR : " << outWsccPath << " (" << records.size()
              << " rows)\n";
    std::cout << "  MSC CDR  : " << outCdrDir << "/ (" << numFiles
              << " files)\n";
    std::cout << "  CRM      : " << outCrmPath << " (" << totalCrmRows
              << " rows)\n";
    std::cout << "\nDone.\n";

    return 0;
}
