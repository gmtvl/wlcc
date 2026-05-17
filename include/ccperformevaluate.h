#pragma once

#include "dataextraction.h"
#include "inputvalidator.h"

#include <map>
#include <string>
#include <vector>

namespace wlcc {

// --- Result structs ---

// Fields extracted from a record for fuzzy matching during daily validation
struct CallMatchFields {
    std::string callingLast10;  // last 10 digits of CLI / Calling Number
    int timeSeconds;            // time of day in seconds from midnight
    int durationSeconds;        // call duration in seconds
};

// Consolidated per-date status — all daily properties under one struct
struct DailyStatus {
    // From validateCDRsForDate — true if all WSCC Excel records matched MSC CDR
    bool recordsValid = false;
    // to track unmatched WSCC rows for logging
    std::vector<ExcelRecord> invalidWsccRows;
    // to track daily untagged WSCC CDR records in CRM.
    std::vector<ExcelRecord> untagged;

    // Repeat call metrics
    int totalNonIvrCalls = 0;
    int totalShortAnsweredCalls = 0;
    int totalShortIvrCalls = 0;
    int repeatCallCount = 0;
    double repeatPercentage = 0.0;
    bool fullPayment = true;       // true when repeatPercentage <= 15.0

    // IVR metrics (FRL == 0, IVRDuration >= 10)
    int totalIvrDuration = 0;
    int totalIvrCallCount = 0;
    double avgIvrHandlingTime = 0.0;         // capped at 120s

    // Agent not answered (FRL == 3)
    int totalAgentNotAnswered    = 0;
    int notAnsweredDueToLowQTime = 0;

    // Entry level metrics (Level == "Entry", TotalTimeAtAgent >= 10)
    int totalEntryAgentTime = 0;
    int totalEntryLevelCount = 0;
    int tcbhEntryLevelCalls = 0;
    int tcbhEntryLevelQWait = 0;
    double avgEntryLevelHandlingTime = 0.0;  // capped at 130s

    // Second level metrics (Level == "Second", TotalTimeAtAgent >= 10)
    int totalSecondAgentTime = 0;
    int totalSecondLevelCount = 0;
    int tcbhSecondLevelCalls = 0;
    int tcbhSecondLevelQWait = 0;
    double avgSecondLevelHandlingTime = 0.0; // capped at 150s

    // Third level metrics (Level == "Third", TotalTimeAtAgent >= 10)
    int totalThirdAgentTime = 0;
    int totalThirdLevelCount = 0;
    int tcbhThirdLevelCalls = 0;
    int tcbhThirdLevelQWait = 0;
    double avgThirdLevelHandlingTime = 0.0;  // capped at 180s

    // --- Repeat-split metrics (for First vs Repeated table population) ---
    // Entry level: repeat calls only (subset of totalEntry*)
    int repeatEntryAgentTime = 0;
    int repeatEntryCount = 0;
    double avgRepeatEntryHandlingTime = 0.0;

    // Second level: repeat calls only (subset of totalSecond*)
    int repeatSecondAgentTime = 0;
    int repeatSecondCount = 0;
    double avgRepeatSecondHandlingTime = 0.0;
};

// SLA configuration values (read from config file)
struct SlaConfig {
    double systemUptime        = 100.0;  // percentage
    double customerSatisfaction = 2.25;  // score (0-3 scale)
    double callQuality          = 86.0;  // score (0-100 scale)
    double ivrPerMinuteCharge   = 0.47;  // Rs per minute
    double agentPerMinuteCharge = 3.11;  // Rs per minute

    // Parse from key=value config file. Returns true on success.
    static bool loadFromFile(const std::string& filepath, SlaConfig& config);
};

// CRM records grouped by mobileNo last 10 digits
using CrmLookup = std::map<std::string, std::vector<ExcelRecord>>;

// --- Evaluation class ---

class CCPerformEvaluate {
public:
    // Construct with references to validated input files and data extraction.
    // wsccData is used for WSCC CDR Excel and MSC CDR text files.
    // crmData is used for the CRM Excel file.
    // The caller owns all objects and must keep them alive.
    CCPerformEvaluate(const InputValidator& validator,
                      DataExtraction& wsccData,
                      DataExtraction& crmData);
    ~CCPerformEvaluate();

    // Non-copyable (holds a reference)
    CCPerformEvaluate(const CCPerformEvaluate&) = delete;
    CCPerformEvaluate& operator=(const CCPerformEvaluate&) = delete;

    // --- Daily Record Validation ---

    // Validate CDRs for a single date: checks that every WSCC CDR Excel record
    // has a matching MSC CDR text record. Returns true if all matched.
    // Records the result in dailyStatus_[date].recordsValid.
    bool validateCDRsForDate(const std::string& date,
                      const std::string& excelDateHeader = "Date");

    // Run validateCDRsForDate() for every calendar day of the month.
    // Month is taken from InputValidator, year from CDR "Period" metadata.
    // Populates dailyStatus_ recordsValid for all days (DD-MM-YYYY format).
    void validateAllDates(const std::string& excelDateHeader = "Date");

    // --- Repeat Call Detection ---

    // Check repeat non-IVR calls for a single date.
    // wsccRows: WSCC CDR rows for this date (pre-fetched).
    // crmByMobile: CRM records indexed by mobileNo last 10 digits.
    // Updates dailyStatus_[date] repeat-call fields.
    void checkRepeatCalls(const std::string& date,
                          const std::vector<ExcelRecord>& wsccRows,
                          const CrmLookup& crmByMobile);

    // Check repeat calls for all dates in the month.
    // Uses wsccData_ for WSCC CDR Excel and crmData_ for CRM Excel.
    void checkAllRepeatCalls(const std::string& wsccSheetName = "Sheet1",
                             const std::string& crmSheetName = "Sheet1",
                             const std::string& excelDateHeader = "Date");

    // --- Access consolidated daily status ---

    const std::map<std::string, DailyStatus>& getDailyStatus() const;

    // Write computed daily metrics to the IBD Data sheet of an existing workbook.
    // workbookPath: path to the calc sheet Excel file (must already exist).
    // Sheet name is constructed internally from month/year as "<MONTHNAME> <YEAR> IBD Data".
    void writeIBDData(const std::string& workbookPath);

    // Write the "Bill Calculation <MONTH> <YEAR>" sheet with Excel formulas.
    // References IBD Data table totals, applies SLA penalties from config.
    void writeBillCalculation(const std::string& workbookPath,
                              const SlaConfig& slaConfig);

    // --- Match helpers (public static for test/diagnostic use) ---

    static std::string extractLast10(const std::string& number);
    static int parseTimeToSeconds(const std::string& timeStr);
    static int parseExcelTime(const CellValue& val);
    static int parseDurationToSeconds(const CellValue& val);
    static CallMatchFields extractExcelMatchFields(const ExcelRecord& record);
    static CallMatchFields extractCdrMatchFields(const CdrRecord& record);
    static bool fieldsMatch(const CallMatchFields& a, const CallMatchFields& b);

    // Parse time component from CRM "createdOn" field.
    // Format: "DD-Mon-YYYY H:MM AM/PM" → seconds from midnight.
    // Returns -1 if unparseable.
    static int parseCrmCreatedOnTime(const CellValue& val);

    // Find WSCC CDR records in callerGroups that have no matching CRM record
    // in crmByMobile. Matching = same mobile (key) AND IVRStartTime within
    // ±30 seconds of CRM createdOn time component.
    // Returns vector of untagged (unmatched) ExcelRecord pointers.
    static std::vector<const ExcelRecord*> findUntaggedRecords(
        const std::map<std::string, std::vector<const ExcelRecord*>>& callerGroups,
        const CrmLookup& crmByMobile,
        int toleranceSecs = 60);

private:
    const InputValidator& validator_;
    DataExtraction& data_;       // WSCC CDR Excel + MSC CDR text
    DataExtraction& crmData_;    // CRM Excel
    std::map<std::string, DailyStatus> dailyStatus_;
    bool cliBased_ = false;  // true if repeat call check is based on CLI due to missing CRM data

    // --- Repeat call helpers ---

    // Build CRM lookup for a single date's CRM rows, group by mobileNo last 10
    static CrmLookup buildCrmLookup(const std::vector<ExcelRecord>& crmRows);

    // Build composite reason from CRM record: "type|subType|category|subCategory"
    static std::string getCompositeReason(const ExcelRecord& crmRecord);

    // Extract month and year from InputValidator + CDR metadata.
    // Returns false if month/year cannot be determined.
    bool getMonthYear(int& month, int& year) const;
};

} // namespace wlcc
