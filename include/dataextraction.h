#pragma once

#include "excel_cm.h"
#include "cdr_cc.h"

#include <map>
#include <string>
#include <vector>

namespace wlcc {

// An Excel row as a map from header name to typed cell value
using ExcelRecord = std::map<std::string, CellValue>;

// Rows grouped by date string (e.g., "05-11-2025")
using ExcelRecordsByDate = std::map<std::string, std::vector<ExcelRecord>>;

// CDR records grouped by date string
using CdrRecordsByDate = std::map<std::string, std::vector<CdrRecord>>;

class DataExtraction {
public:
    DataExtraction();
    ~DataExtraction();

    // Non-copyable (ExcelCM is non-copyable)
    DataExtraction(const DataExtraction&) = delete;
    DataExtraction& operator=(const DataExtraction&) = delete;

    // --- Excel: load once, query many ---

    // Open an Excel file and detect its header (one-time setup).
    // Subsequent queries reuse this loaded state.
    void loadExcel(const std::string& filepath, const std::string& sheetName);
    bool isExcelLoaded() const;

    // Extract all rows grouped by date from the loaded Excel.
    // dateHeaderName: the header column containing dates (e.g., "Date")
    ExcelRecordsByDate extractExcelRowsByDate(
        const std::string& dateHeaderName) const;

    // Get rows matching a specific date from the loaded Excel.
    std::vector<ExcelRecord> getExcelRowsForDate(
        const std::string& dateHeaderName,
        const std::string& date) const;

    // --- CDR: load once, query many ---

    // Scan directories for .txt files, parse them via CdrCC, and
    // store all records grouped by date (one-time setup).
    // dateHeaderName: the header containing datetime (default: "Date & Time")
    void loadCdrDirectories(
        const std::vector<std::string>& directories,
        const std::string& dateHeaderName = "Date & Time");
    bool isCdrLoaded() const;

    // Get all CDR records already grouped by date.
    const CdrRecordsByDate& getCdrRecordsByDate() const;

    // Get CDR records matching a specific date.
    std::vector<CdrRecord> getCdrRecordsForDate(const std::string& date) const;

    // --- Utility ---

    // Extract the date portion from a datetime string
    // e.g., "01-01-2025 10:34:08" -> "01-01-2025"
    static std::string extractDatePart(const std::string& dateTimeStr);

private:
    // Excel persistent state
    ExcelCM excel_;
    std::string excelSheet_;
    bool excelLoaded_ = false;

    // CDR persistent state (pre-grouped at load time)
    CdrRecordsByDate cdrData_;
    bool cdrLoaded_ = false;
};

} // namespace wlcc
