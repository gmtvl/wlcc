// wtest.exe — Diagnostic tool for matching WSCC CDR Excel records against
//              MSC CDR text records for specific dates.
//
// Usage:
//   wtest.exe [-p <path>] <date1> [date2] [date3]
//
//   Dates in DD-MM-YYYY format (e.g. 01-11-2025).
//   -p <path> : path containing the datafiles_<mon> folder (default: ".")
//
// Output:
//   For each date:
//     1. Matched records printed in alternating lines (WSCC then MSC),
//        showing the 3 match fields: CLI last-10, time (seconds), duration.
//     2. Unmatched WSCC records (no corresponding MSC CDR found).

#include "ccperformevaluate.h"
#include "excel_cm.h"
#include "inputvalidator.h"

#include <algorithm>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <set>
#include <string>
#include <vector>

namespace fs = std::filesystem;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// Validate DD-MM-YYYY format
static bool isValidDate(const std::string& s) {
    if (s.size() != 10) return false;
    if (s[2] != '-' || s[5] != '-') return false;
    for (int i : {0,1,3,4,6,7,8,9}) {
        if (!std::isdigit(static_cast<unsigned char>(s[i]))) return false;
    }
    int day   = std::stoi(s.substr(0, 2));
    int month = std::stoi(s.substr(3, 2));
    if (month < 1 || month > 12 || day < 1 || day > 31) return false;
    return true;
}

// Extract 3-letter month code from DD-MM-YYYY
static std::string monthCodeFromDate(const std::string& date) {
    static const char* codes[] = {
        "jan","feb","mar","apr","may","jun",
        "jul","aug","sep","oct","nov","dec"
    };
    int m = std::stoi(date.substr(3, 2));
    if (m < 1 || m > 12) return "";
    return codes[m - 1];
}

// Format time in seconds as HH:MM:SS
static std::string formatTime(int secs) {
    if (secs < 0) return "N/A";
    int h = secs / 3600;
    int m = (secs % 3600) / 60;
    int s = secs % 60;
    char buf[16];
    std::snprintf(buf, sizeof(buf), "%02d:%02d:%02d", h, m, s);
    return buf;
}

// Print a CallMatchFields line with a label
static void printMatchFields(const char* label,
                             const wlcc::CallMatchFields& f) {
    std::cout << "  " << std::left << std::setw(5) << label
              << " CLI=" << std::setw(12) << f.callingLast10
              << " Time=" << std::setw(10) << formatTime(f.timeSeconds)
              << " Duration=" << f.durationSeconds << "s"
              << "\n";
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

int main(int argc, char* argv[]) {
    // Parse arguments
    std::string basePath = ".";
    std::vector<std::string> dates;

    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        if ((arg == "-p" || arg == "--path") && i + 1 < argc) {
            basePath = argv[++i];
        } else if (arg == "-h" || arg == "--help") {
            std::cerr << "Usage: " << argv[0]
                      << " [-p <path>] <date1> [date2] [date3]\n"
                      << "  Dates in DD-MM-YYYY format.\n"
                      << "  -p <path>  Path containing datafiles_<mon> folder.\n";
            return 0;
        } else {
            dates.push_back(arg);
        }
    }

    if (dates.empty() || dates.size() > 3) {
        std::cerr << "Error: provide 1 to 3 dates in DD-MM-YYYY format.\n";
        return 1;
    }

    // Validate all dates
    for (const auto& d : dates) {
        if (!isValidDate(d)) {
            std::cerr << "Error: invalid date format '" << d
                      << "'. Expected DD-MM-YYYY.\n";
            return 1;
        }
    }

    // All dates must be in the same month
    std::string monthCode = monthCodeFromDate(dates[0]);
    for (size_t i = 1; i < dates.size(); ++i) {
        if (monthCodeFromDate(dates[i]) != monthCode) {
            std::cerr << "Error: all dates must be in the same month.\n";
            return 1;
        }
    }

    // Validate directory structure
    wlcc::InputValidator validator;
    if (!validator.validate(basePath, monthCode)) {
        std::cerr << "Validation failed:\n";
        for (const auto& err : validator.getErrors()) {
            std::cerr << "  - " << err << "\n";
        }
        return 1;
    }

    std::cout << "Data folder : " << validator.getDataFilesPath() << "\n";
    std::cout << "WSCC CDR    : " << validator.getWsccCdrFile().filename << "\n";
    std::cout << "Dates       :";
    for (const auto& d : dates) std::cout << " " << d;
    std::cout << "\n\n";

    try {

    // --- Step 1: Load MSC CDR text files ---
    const auto& inbound  = validator.getMscInboundCdrFiles();
    const auto& outbound = validator.getMscOutboundCdrFiles();

    std::set<std::string> cdrDirSet;
    for (const auto& f : inbound)
        cdrDirSet.insert(fs::path(f.fullPath).parent_path().string());
    for (const auto& f : outbound)
        cdrDirSet.insert(fs::path(f.fullPath).parent_path().string());
    std::vector<std::string> cdrDirs(cdrDirSet.begin(), cdrDirSet.end());

    wlcc::DataExtraction wsccData;
    wlcc::DataExtraction crmData;  // needed for CCPerformEvaluate ctor, not used here

    std::cout << "Loading MSC CDR data from " << cdrDirs.size()
              << " directories...\n";
    wsccData.loadCdrDirectories(cdrDirs);
    std::cout << "MSC CDR data loaded.\n";

    // --- Step 2: Load WSCC CDR Excel ---
    std::string wsccSheet;
    {
        wlcc::ExcelCM probe;
        probe.open(validator.getWsccCdrFile().fullPath);
        auto sheets = probe.getSheetNames();
        probe.close();
        if (sheets.empty()) {
            std::cerr << "Error: WSCC CDR Excel has no sheets.\n";
            return 1;
        }
        wsccSheet = sheets[0];
    }

    std::cout << "Loading WSCC CDR Excel (sheet: " << wsccSheet << ")...\n";
    wsccData.loadExcel(validator.getWsccCdrFile().fullPath, wsccSheet);
    std::cout << "WSCC CDR Excel loaded.\n\n";

    // --- Step 3: For each date, find and match records ---
    using CCEval = wlcc::CCPerformEvaluate;

    for (const auto& date : dates) {
        std::cout << "========== Date: " << date << " ==========\n";

        // Get WSCC Excel rows for this date
        auto wsccRows = wsccData.getExcelRowsForDate("Date", date);
        std::cout << "  WSCC CDR records: " << wsccRows.size() << "\n";

        // Get MSC CDR text records for this date
        auto cdrRows = wsccData.getCdrRecordsForDate(date);
        std::cout << "  MSC  CDR records: " << cdrRows.size() << "\n\n";

        if (wsccRows.empty()) {
            std::cout << "  No WSCC records for this date.\n\n";
            continue;
        }

        // Extract match fields for all CDR records
        std::vector<wlcc::CallMatchFields> cdrFields;
        cdrFields.reserve(cdrRows.size());
        for (const auto& cdr : cdrRows) {
            cdrFields.push_back(CCEval::extractCdrMatchFields(cdr));
        }

        // Track which CDR records have been matched
        std::vector<bool> cdrMatched(cdrRows.size(), false);

        // Matched and unmatched WSCC records
        int matchCount = 0;

        std::cout << "  --- Matched Records ---\n";

        std::vector<const wlcc::ExcelRecord*> unmatchedWscc;

        for (const auto& wsccRow : wsccRows) {
            wlcc::CallMatchFields wsccFields =
                CCEval::extractExcelMatchFields(wsccRow);

            bool found = false;
            for (size_t i = 0; i < cdrFields.size(); ++i) {
                if (!cdrMatched[i] &&
                    CCEval::fieldsMatch(wsccFields, cdrFields[i])) {
                    cdrMatched[i] = true;
                    found = true;
                    matchCount++;

                    // Print matched pair
                    std::cout << "  Match #" << matchCount << ":\n";
                    printMatchFields("WSCC", wsccFields);
                    printMatchFields("MSC", cdrFields[i]);
                    break;
                }
            }

            if (!found) {
                unmatchedWscc.push_back(&wsccRow);
            }
        }

        std::cout << "\n  Total matched: " << matchCount
                  << " / " << wsccRows.size() << "\n";

        // Print unmatched WSCC records
        if (!unmatchedWscc.empty()) {
            std::cout << "\n  --- Unmatched WSCC Records ("
                      << unmatchedWscc.size() << ") ---\n";
            for (const auto* row : unmatchedWscc) {
                wlcc::CallMatchFields f = CCEval::extractExcelMatchFields(*row);
                printMatchFields("WSCC", f);
            }
        } else {
            std::cout << "  All WSCC records matched.\n";
        }

        std::cout << "\n";
    }

    } catch (const std::exception& e) {
        std::cerr << "\nERROR: " << e.what() << "\n";
        return 2;
    }

    return 0;
}
