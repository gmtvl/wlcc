#include "ccperformevaluate.h"
#include "cdr_cc.h"
#include "excel_cm.h"
#include "subtotal_post.h"

#include <spdlog/spdlog.h>

#include <algorithm>
#include <cmath>
#include <cstdlib>
#include <fstream>
#include <functional>
#include <set>
#include <stdexcept>
#include <sstream>

namespace {
spdlog::logger* log() {
    static spdlog::logger* inst = [] {
        auto l = spdlog::get("ccperformevaluate");
        return l ? l.get() : spdlog::default_logger().get();
    }();
    return inst;
}
} // anonymous namespace

namespace wlcc {

// --- Constructor / Destructor ---

CCPerformEvaluate::CCPerformEvaluate(const InputValidator& validator,
                                     DataExtraction& wsccData,
                                     DataExtraction& crmData)
    : validator_(validator), data_(wsccData), crmData_(crmData) {}

CCPerformEvaluate::~CCPerformEvaluate() = default;

// --- SLA Config File Parser ---

bool SlaConfig::loadFromFile(const std::string& filepath, SlaConfig& config) {
    std::ifstream file(filepath);
    if (!file.is_open()) {
        return false;
    }

    std::string line;
    while (std::getline(file, line)) {
        // Skip comments and blank lines
        if (line.empty() || line[0] == '#') continue;

        auto eqPos = line.find('=');
        if (eqPos == std::string::npos) continue;

        std::string key = line.substr(0, eqPos);
        std::string val = line.substr(eqPos + 1);

        // Trim whitespace
        auto trim = [](std::string& s) {
            s.erase(0, s.find_first_not_of(" \t\r\n"));
            s.erase(s.find_last_not_of(" \t\r\n") + 1);
        };
        trim(key);
        trim(val);

        try {
            if (key == "system_uptime")         config.systemUptime = std::stod(val);
            else if (key == "customer_satisfaction") config.customerSatisfaction = std::stod(val);
            else if (key == "call_quality")     config.callQuality = std::stod(val);
            else if (key == "ivr_per_minute")   config.ivrPerMinuteCharge = std::stod(val);
            else if (key == "agent_per_minute") config.agentPerMinuteCharge = std::stod(val);
        } catch (...) {
            // Ignore malformed values
        }
    }

    return true;
}

// --- Daily validation helpers ---

std::string CCPerformEvaluate::extractLast10(const std::string& number) {
    if (number.size() >= 10) {
        return number.substr(number.size() - 10);
    }
    return number;
}

int CCPerformEvaluate::parseTimeToSeconds(const std::string& timeStr) {
    // Expected format: "HH:MM:SS"
    int h = 0, m = 0, s = 0;
    char sep1, sep2;
    std::istringstream iss(timeStr);
    iss >> h >> sep1 >> m >> sep2 >> s;
    if (iss.fail()) {
        return -1;
    }
    return h * 3600 + m * 60 + s;
}

int CCPerformEvaluate::parseExcelTime(const CellValue& val) {
    // Excel time can be stored as:
    // - double: fractional day (0.0 = midnight, 0.5 = noon, 1.0 = next midnight)
    // - string: "HH:MM:SS"
    // - int64_t: seconds from midnight
    if (std::holds_alternative<double>(val)) {
        double frac = std::get<double>(val);
        // If value is > 1, it may include a date component; take fractional part
        if (frac > 1.0) {
            frac = frac - static_cast<int>(frac);
        }
        return static_cast<int>(std::round(frac * 86400.0));
    }
    if (std::holds_alternative<std::string>(val)) {
        return parseTimeToSeconds(std::get<std::string>(val));
    }
    if (std::holds_alternative<int64_t>(val)) {
        return static_cast<int>(std::get<int64_t>(val));
    }
    return -1;
}

int CCPerformEvaluate::parseDurationToSeconds(const CellValue& val) {
    if (std::holds_alternative<int64_t>(val)) {
        return static_cast<int>(std::get<int64_t>(val));
    }
    if (std::holds_alternative<double>(val)) {
        return static_cast<int>(std::round(std::get<double>(val)));
    }
    if (std::holds_alternative<std::string>(val)) {
        try {
            return std::stoi(std::get<std::string>(val));
        } catch (...) {
            return 0;
        }
    }
    return 0;
}

CallMatchFields CCPerformEvaluate::extractExcelMatchFields(
    const ExcelRecord& record)
{
    CallMatchFields fields{};

    // CLI — extract last 10 digits
    auto cliIt = record.find("CLI");
    if (cliIt != record.end()) {
        fields.callingLast10 = extractLast10(cellValueToString(cliIt->second));
    }

    // IVRStartTime — read as time
    auto timeIt = record.find("IVRStartTime");
    if (timeIt != record.end()) {
        fields.timeSeconds = parseExcelTime(timeIt->second);
    } else {
        fields.timeSeconds = -1;
    }

    // TotalDuration — read as seconds
    auto durIt = record.find("TotalDuration");
    if (durIt != record.end()) {
        fields.durationSeconds = parseDurationToSeconds(durIt->second);
    } else {
        fields.durationSeconds = 0;
    }

    return fields;
}

CallMatchFields CCPerformEvaluate::extractCdrMatchFields(
    const CdrRecord& record)
{
    CallMatchFields fields{};

    // Calling Number — extract last 10 digits
    auto callingIt = record.find("Calling Number");
    if (callingIt != record.end()) {
        fields.callingLast10 = extractLast10(callingIt->second);
    }

    // Date & Time — extract time portion after space ("DD-MM-YYYY HH:MM:SS")
    auto dtIt = record.find("Date & Time");
    if (dtIt != record.end()) {
        const std::string& dt = dtIt->second;
        auto spacePos = dt.find(' ');
        if (spacePos != std::string::npos) {
            fields.timeSeconds = parseTimeToSeconds(dt.substr(spacePos + 1));
        } else {
            fields.timeSeconds = -1;
        }
    } else {
        fields.timeSeconds = -1;
    }

    // Duration — read as seconds (already a string in CDR)
    auto durIt = record.find("Duration");
    if (durIt != record.end()) {
        try {
            fields.durationSeconds = std::stoi(durIt->second);
        } catch (...) {
            fields.durationSeconds = 0;
        }
    } else {
        fields.durationSeconds = 0;
    }

    return fields;
}

bool CCPerformEvaluate::fieldsMatch(const CallMatchFields& a,
                                    const CallMatchFields& b)
{
    // Calling number last 10 digits must match exactly
    if (a.callingLast10 != b.callingLast10) {
        return false;
    }

    // Time must be within ±5 seconds
 /*/   if (a.timeSeconds < 0 || b.timeSeconds < 0) {
        return false;
    }*/
    if (std::abs(a.timeSeconds - b.timeSeconds) > 60) {
        return false;
    }

    /*// Duration must be within ±5 seconds
    if (std::abs(a.durationSeconds - b.durationSeconds) > 20) {
        return false;
    }*/

    return true;
}

// --- CRM createdOn time parser ---

int CCPerformEvaluate::parseCrmCreatedOnTime(const CellValue& val) {
    // CRM "createdOn" is typically a string like "01-Nov-2025 3:45 PM"
    // or "15-Nov-2025 12:00 AM". We extract the time portion after the date.
    if (std::holds_alternative<std::string>(val)) {
        const std::string& s = std::get<std::string>(val);

        // Find first space — everything after it is the time part
        auto spacePos = s.find(' ');
        if (spacePos == std::string::npos || spacePos + 1 >= s.size()) {
            return -1;
        }
        std::string timePart = s.substr(spacePos + 1);  // "3:45 PM" or "12:00 AM"

        // Parse H:MM or HH:MM
        int h = 0, m = 0;
        char sep;
        std::istringstream iss(timePart);
        iss >> h >> sep >> m;
        if (iss.fail() || sep != ':') {
            return -1;
        }

        // Check for AM/PM suffix
        std::string remainder;
        iss >> remainder;
        // Convert to uppercase for comparison
        std::string ampm;
        for (char c : remainder) {
            ampm += static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
        }

        if (ampm == "PM" && h != 12) {
            h += 12;
        } else if (ampm == "AM" && h == 12) {
            h = 0;
        }

        return h * 3600 + m * 60;
    }

    // If stored as a double (Excel datetime serial), extract time fraction
    if (std::holds_alternative<double>(val)) {
        double v = std::get<double>(val);
        double frac = v - static_cast<int>(v);
        return static_cast<int>(std::round(frac * 86400.0));
    }

    // If stored as int64_t, treat as seconds from midnight
    if (std::holds_alternative<int64_t>(val)) {
        return static_cast<int>(std::get<int64_t>(val));
    }

    return -1;
}

// --- Untagged record finder ---

std::vector<const ExcelRecord*> CCPerformEvaluate::findUntaggedRecords(
    const std::map<std::string, std::vector<const ExcelRecord*>>& callerGroups,
    const CrmLookup& crmByMobile,
    int toleranceSecs)
{
    std::vector<const ExcelRecord*> untagged;

    for (const auto& [mobile, wsccRecords] : callerGroups) {
        // Check if CRM has any records for this mobile number
        auto crmIt = crmByMobile.find(mobile);
        if (crmIt == crmByMobile.end() || crmIt->second.empty()) {
            // No CRM records at all — all WSCC records are untagged
            for (const auto* rec : wsccRecords) {
                untagged.push_back(rec);
            }
            continue;
        }

        const auto& crmRecords = crmIt->second;

        // Track which CRM records have been used (one-to-one matching)
        std::vector<bool> crmUsed(crmRecords.size(), false);

        for (const auto* wsccRec : wsccRecords) {
            // Extract IVRStartTime from WSCC record
            auto ivrIt = wsccRec->find("CallEndTime");
            if (ivrIt == wsccRec->end()) {
                untagged.push_back(wsccRec);
                continue;
            }
            int wsccTimeSecs = parseExcelTime(ivrIt->second);

            // Try to find a matching CRM record within ±toleranceSecs
            bool matched = false;
            for (size_t i = 0; i < crmRecords.size(); ++i) {
                if (crmUsed[i]) continue;

                auto createdIt = crmRecords[i].find("createdOn");
                if (createdIt == crmRecords[i].end()) continue;

                int crmTimeSecs = parseCrmCreatedOnTime(createdIt->second);
                if (crmTimeSecs < 0 || wsccTimeSecs < 0) continue;

                if (std::abs(wsccTimeSecs - crmTimeSecs) <= toleranceSecs) {
                    // Match found — mark CRM record as used and break
                    crmUsed[i] = true;
                    matched = true;
                    break;
                }
            }

            if (!matched) {
                untagged.push_back(wsccRec);
            }
        }
    }

    return untagged;
}

// --- Month/Year helper ---

bool CCPerformEvaluate::getMonthYear(int& month, int& year) const {
    month = validator_.getMonthNumber();
    year = 0;

    const auto& inbound = validator_.getMscInboundCdrFiles();
    const auto& outbound = validator_.getMscOutboundCdrFiles();

    if (!inbound.empty()) {
        CdrCC cdr;
        cdr.load(inbound[0].fullPath);
        year = cdr.getYear();
    } else if (!outbound.empty()) {
        CdrCC cdr;
        cdr.load(outbound[0].fullPath);
        year = cdr.getYear();
    }

    return (month != 0 && year != 0);
}

// --- Calendar helper (local) ---

static int daysInMonth(int month, int year) {
    switch (month) {
        case 4: case 6: case 9: case 11:
            return 30;
        case 2:
            if ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0))
                return 29;
            else
                return 28;
        default:
            return 31;
    }
}

static std::string formatDate(int day, int month, int year) {
    std::ostringstream oss;
    oss << (day < 10 ? "0" : "") << day << "-"
        << (month < 10 ? "0" : "") << month << "-"
        << year;
    return oss.str();
}

// --- IBD Data write helpers (local) ---

static std::string monthNumberToName(int month) {
    static const char* names[] = {
        "", "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
        "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
    };
    if (month < 1 || month > 12) return "";
    return names[month];
}

static std::string cellRef(const std::string& col, int row) {
    return col + std::to_string(row);
}

// --- Daily Record Validation ---

bool CCPerformEvaluate::validateCDRsForDate(
    const std::string& date,
    const std::string& excelDateHeader)
{
    // Get records for this specific date from each source
    auto excelRows = data_.getExcelRowsForDate(excelDateHeader, date);
    auto cdrRows   = data_.getCdrRecordsForDate(date);

    log()->debug("validateCDRsForDate [{}]: {} Excel rows, {} CDR rows",
                  date, excelRows.size(), cdrRows.size());

    // No Excel records for this date — nothing to fail
    if (excelRows.empty()) {
        dailyStatus_[date].recordsValid = true;
        log()->trace("validateCDRsForDate [{}]: no Excel rows, marking valid", date);
        return true;
    }

    // Excel records exist but no CDR records — date is invalid
    if (cdrRows.empty()) {
        dailyStatus_[date].recordsValid = false;
        log()->trace("validateCDRsForDate [{}]: Excel rows exist but no CDR rows — INVALID", date);
        return false;
    }

    // Pre-extract match fields for CDR records of this date
    std::vector<CallMatchFields> cdrFields;
    cdrFields.reserve(cdrRows.size());
    for (const auto& cdr : cdrRows) {
        cdrFields.push_back(extractCdrMatchFields(cdr));
    }

    // Track which CDR records have already been matched
    std::vector<bool> matched(cdrRows.size(), false);
    int matchCount = 0;

    // Every Excel record must find a matching CDR record
    for (const auto& excelRow : excelRows) {
        CallMatchFields excelFields = extractExcelMatchFields(excelRow);
        bool found = false;

        for (size_t i = 0; i < cdrFields.size(); ++i) {
            if (!matched[i] && fieldsMatch(excelFields, cdrFields[i])) {
                matched[i] = true;
                found = true;
                matchCount++;
                break;
            }
        }

        if (!found) {
            dailyStatus_[date].invalidWsccRows.push_back(excelRow);
            log()->debug("validateCDRsForDate [{}]:  unmatched Excel record here:"
                          "(CLI={}, time={}, dur={}).",
                          date, excelFields.callingLast10,
                          excelFields.timeSeconds, excelFields.durationSeconds);
            //return false;
            continue;
        }
    }

    if (matchCount == excelRows.size()) {
        dailyStatus_[date].recordsValid = true;
    } else {
        dailyStatus_[date].recordsValid = false;
    }
    //dailyStatus_[date].recordsValid = true;
    log()->trace("validateCDRsForDate [{}]:  Matched {}/{} Excel rows matched",
                  date, matchCount, excelRows.size());
    return true;
}

void CCPerformEvaluate::validateAllDates(
    const std::string& excelDateHeader)
{
    dailyStatus_.clear();

    int month = 0, year = 0;
    if (!getMonthYear(month, year)) {
        return;  // cannot determine month/year
    }

    int days = daysInMonth(month, year);

    for (int day = 1; day <= days; ++day) {
        validateCDRsForDate(formatDate(day, month, year), excelDateHeader);
    }
}

// --- Repeat call helpers ---

CrmLookup CCPerformEvaluate::buildCrmLookup(
    const std::vector<ExcelRecord>& crmRows)
{
    CrmLookup result;

    for (const auto& row : crmRows) {
        // Only include CRM records where source is "INBOUND"
        auto srcIt = row.find("source");
        if (srcIt == row.end()) continue;
        if (cellValueToString(srcIt->second) != "INBOUND") continue;

        auto it = row.find("mobileNo");
        if (it == row.end()) continue;

        std::string mobile = cellValueToString(it->second);
        std::string last10 = extractLast10(mobile);
        if (!last10.empty()) {
            result[last10].push_back(row);
        }
    }

    return result;
}

std::string CCPerformEvaluate::getCompositeReason(
    const ExcelRecord& crmRecord)
{
    std::string type, subType, category, subCategory;

    auto typeIt = crmRecord.find("type");
    if (typeIt != crmRecord.end()) {
        type = cellValueToString(typeIt->second);
    }

    auto subTypeIt = crmRecord.find("subType");
    if (subTypeIt != crmRecord.end()) {
        subType = cellValueToString(subTypeIt->second);
    }

    auto catIt = crmRecord.find("category");
    if (catIt != crmRecord.end()) {
        category = cellValueToString(catIt->second);
    }

    auto subCatIt = crmRecord.find("subCategory");
    if (subCatIt != crmRecord.end()) {
        subCategory = cellValueToString(subCatIt->second);
    }

    return type + "|" + subType + "|" + category + "|" + subCategory;
}

// --- Repeat Call Detection ---

void CCPerformEvaluate::checkRepeatCalls(
    const std::string& date,
    const std::vector<ExcelRecord>& wsccRows,
       const CrmLookup& crmByMobile)
{
    int totalIvrDuration = 0, totalIvrCallCount = 0;
    int totalAgentNotAnswered = 0, notAnsweredDueToLowQTime =0;
    int totalEntryAgentTime = 0, totalEntryLevelCount = 0;
    int totalSecondAgentTime = 0, totalSecondLevelCount = 0;
    int totalThirdAgentTime = 0, totalThirdLevelCount = 0;
    int tcbhEntryLevelCalls = 0, tcbhEntryLevelQWait = 0;
    int tcbhSecondLevelCalls = 0, tcbhSecondLevelQWait = 0;
    int tcbhThirdLevelCalls = 0, tcbhThirdLevelQWait = 0;
    int totalShortAnsweredCalls = 0, totalShortIvrCalls = 0;
    bool noNonIvrCalls = false; // flag to indicate if there are no non-IVR calls

    cliBased_ = false; // reset to default for each date
    
    // --- Filter non-IVR rows for repeat call logic ---
    // Filter non-IVR rows, excluding short abandonments and third-level calls.
    // Non-IVR: FRL != "0" (and FRL not empty).
    // Short abandonment: TotalTimeAtAgent is blank or < 10.
    // Third-level: "level" header value is "third" (case-insensitive).
    std::vector<const ExcelRecord*> nonIvrRowsforRepCalc;
    for (const auto& row : wsccRows) {
        auto frlIt = row.find("FRL");
        if (frlIt == row.end()) continue;     

        std::string frlVal = cellValueToString(frlIt->second);
        if (frlVal.empty()) { totalShortIvrCalls++; continue; } // since empty FRL is treated as short abandoned IVR calls, skip from non IVR processing 
        // IVR metrics: FRL == 0, IVRDuration >= 10
        if (frlVal == "0") {
            auto ivrDurIt = row.find("IVRDuration");
            if (ivrDurIt != row.end()) {
                int ivrDur = parseDurationToSeconds(ivrDurIt->second);
                if (ivrDur >= 10) {
                    totalIvrDuration += ivrDur;
                    totalIvrCallCount++;
                } 
                else { totalShortIvrCalls++; }
            }
            continue;   // skip IVR calls from non-IVR processing
        }

        auto queueWaitTimeIt = row.find("QueDuration");
        if (queueWaitTimeIt == row.end()) continue;   // should not get to here

        int queueWaitTime = parseDurationToSeconds(queueWaitTimeIt->second);
 
        // Agent not answered: FRL == 3
        if (frlVal == "3") {
            totalAgentNotAnswered++; // Not answered , hence only IVR component billed and it is an IVR call.
            // This is case of abandonement, but happened due to low Qtime and hence avoided from penalty calc.
            if (queueWaitTime < 10) notAnsweredDueToLowQTime++; 
            continue;   //skip agent-not-answered calls from non-IVR processing
        }

        // Exclude short abandonment calls
        // since FRL == 0 has already been skipped, 
        // we can treat missing or blank TotalTimeAtAgent as short abandonment
        // but will we even reach this point for FRL == 0 ? let's keep the logic consistent
        // and check for blank or <10 for all non-IVR calls
        auto agentTimeIt = row.find("TotalTimeAtAgent");
        if (agentTimeIt == row.end()) continue;  // blank → short abandonment

        int agentTime = parseDurationToSeconds(agentTimeIt->second);
        if (agentTime < 10) { totalShortAnsweredCalls++; continue; }   // < 10 seconds → short abandonment, note: it is not added in nonIVR total calls
       
        auto tcbh = row.find("Hour");
        if (tcbh == row.end()) continue;   // should not get to here

        int tcbhValue = parseDurationToSeconds(tcbh->second);

        // Exclude third-level calls
        auto levelIt = row.find("Level");
        if (levelIt != row.end()) {
            std::string levelVal = cellValueToString(levelIt->second);
            std::transform(levelVal.begin(), levelVal.end(),
                           levelVal.begin(), ::tolower);
            if (levelVal == "entry") {
                totalEntryAgentTime += (agentTime >=10) ? agentTime : 0;
                totalEntryLevelCount++;
                if (tcbhValue == 19) {
                    tcbhEntryLevelCalls++;
                    if (queueWaitTime >= 60) tcbhEntryLevelQWait++;
                }
            } else if (levelVal == "second") {
                totalSecondAgentTime += (agentTime >=10) ? agentTime : 0;
                totalSecondLevelCount++;
                if (tcbhValue == 19) {
                    tcbhSecondLevelCalls++;
                    if (queueWaitTime >= 45) tcbhSecondLevelQWait++;
                }
            } else if (levelVal == "third") {
                totalThirdAgentTime += (agentTime >=10) ? agentTime : 0;
                totalThirdLevelCount++;
                if(tcbhValue == 19) {
                    tcbhThirdLevelCalls++;
                    if (queueWaitTime >= 30) tcbhThirdLevelQWait++;
                }
            }
            if (levelVal == "third") continue; // level Three are excluded from repeat calculations 
        }
        nonIvrRowsforRepCalc.push_back(&row);
    }
    DailyStatus& ds = dailyStatus_[date];
    int repeatCount = 0;
    int totalNonIvr = static_cast<int>(nonIvrRowsforRepCalc.size()); /*****review this non sense**** */
    if (totalNonIvr == 0) noNonIvrCalls = true;

    log()->debug("checkRepeatCalls [{}]: {} WSCC rows, {} non-IVR for repeat calculation, "
                  "{} IVR calls, {} Short IVR calls, {} Short answered calls, {} agent-not-answered",
                  date, wsccRows.size(), totalNonIvr, totalIvrCallCount, 
                  totalShortIvrCalls, totalShortAnsweredCalls, totalAgentNotAnswered);

    // Compute averages, capped at maximum values
    if (!noNonIvrCalls) {
        // Group non-IVR rows by CLI last 10 digits
        std::map<std::string, std::vector<const ExcelRecord*>> callerGroups;
        for (const auto* row : nonIvrRowsforRepCalc) {
            auto cliIt = row->find("CLI");
            if (cliIt == row->end()) continue;

            std::string cli = cellValueToString(cliIt->second);
            std::string last10 = extractLast10(cli);
            if (!last10.empty()) {
                callerGroups[last10].push_back(row);
            }
        }
        // checking tagging of qualified CDR rows of the day with the CRM
        // checking equality through containers callerGroups vs crmByMobile.
        // Basis of cliBased_ calculation is changed( 24/03/26) that there will 
        // be such that >=95% to be treated as 100% as per spec. all calls must be tagged 
        // as per this criteria to ensure that we go with reasonBased_ repeat call calculation.

        //record the untagged of WSCC CDR in CRM file for the day.
        auto untaggedPtrs = findUntaggedRecords(callerGroups, crmByMobile, 30*60); // 30 minutes tolerance for repeat call tagging
        ds.untagged.clear();
        ds.untagged.reserve(untaggedPtrs.size());
        for (const auto* ptr : untaggedPtrs) {
            ds.untagged.push_back(*ptr);  // copy the record
        }
        
        // if more than 5% of non-IVR calls are untagged, switch to CLI-based evaluation
        if (static_cast<double>(ds.untagged.size()) / totalNonIvr > 0.05) cliBased_ = true; 

        if(!cliBased_){
            for (const auto& [callingLast10, calls] : callerGroups) {
                // numbers with only 1 call are not repeat candidates
                if (calls.size() <= 1) continue;
                // Look up CRM records for this calling number
                auto crmIt = crmByMobile.find(callingLast10);
                // Collect unique composite reasons from CRM records and record mobiles for each reason
                std::map<std::string, std::vector<std::string>> reasonGroups;
                for (const auto& crm : crmIt->second) {
                    auto mobileIt = crm.find("mobileNo");
                    if (mobileIt != crm.end()) {
                    std::string mobile = extractLast10(cellValueToString(mobileIt->second));
                    reasonGroups[getCompositeReason(crm)].push_back(mobile);
                    }
                }
                int repeatPerMobile = 0;
                for (const auto& [key, vec] : reasonGroups) {
                    repeatPerMobile = static_cast<int>(vec.size());
                    if (repeatPerMobile > 1) {
                        repeatCount += static_cast<int>(vec.size());
                    }   
                }
            }
        }
        else {
            repeatCount = 0;
            // CLI-based evaluation: any caller with >1 non-IVR call is a repeat
            for (const auto& [callingLast10, calls] : callerGroups) {
                if (calls.size() > 1) {
                    repeatCount += static_cast<int>(calls.size());
                }
            }
        }
    }     
    
    // Compute averages, capped at maximum values
    double avgIvrHandlingTime = (totalIvrCallCount > 0)
        ? std::min(static_cast<double>(totalIvrDuration) / totalIvrCallCount, 120.0)
        : 0.0;
    double avgEntryLevelHandlingTime = (totalEntryLevelCount > 0)
        ? std::min(static_cast<double>(totalEntryAgentTime) / totalEntryLevelCount, 130.0)
        : 0.0;
    double avgSecondLevelHandlingTime = (totalSecondLevelCount > 0)
        ? std::min(static_cast<double>(totalSecondAgentTime) / totalSecondLevelCount, 150.0)
        : 0.0;
    double avgThirdLevelHandlingTime = (totalThirdLevelCount > 0)
        ? std::min(static_cast<double>(totalThirdAgentTime) / totalThirdLevelCount, 180.0)
        : 0.0;
    
    // --- Compute repeat-split per-level metrics ---
    // For each non-IVR record, check if its CLI is a repeat caller.
    // If so, accumulate its agent time + count as "repeat" for that level.
    // A record is tagged "repeat" if the caller has >1 non-IVR call.
    int repeatEntryAgentTime = 0, repeatEntryCount = 0;
    int repeatSecondAgentTime = 0, repeatSecondCount = 0;

    if (!noNonIvrCalls && !cliBased_) {
        // Build set of repeat CLIs
        std::map<std::string, std::vector<const ExcelRecord*>> callerGroupsAll;
        for (const auto* row : nonIvrRowsforRepCalc) {
            auto cliIt = row->find("CLI");
            if (cliIt == row->end()) continue;
            std::string last10 = extractLast10(cellValueToString(cliIt->second));
            callerGroupsAll[last10].push_back(row);
        }

        std::set<std::string> repeatCLIs;
        for (const auto& [cli, calls] : callerGroupsAll) {
            if (calls.size() > 1) repeatCLIs.insert(cli);
        }

        // Walk non-IVR rows, accumulate per-level for repeat callers only
        for (const auto* row : nonIvrRowsforRepCalc) {
            auto cliIt = row->find("CLI");
            if (cliIt == row->end()) continue;
            std::string last10 = extractLast10(cellValueToString(cliIt->second));
            if (repeatCLIs.find(last10) == repeatCLIs.end()) continue;

            auto levelIt = row->find("Level");
            if (levelIt == row->end()) continue;
            std::string levelVal = cellValueToString(levelIt->second);
            std::transform(levelVal.begin(), levelVal.end(),
                           levelVal.begin(), ::tolower);

            auto atIt = row->find("TotalTimeAtAgent");
            if (atIt == row->end()) continue;
            int at = parseDurationToSeconds(atIt->second);
            if (at < 10) continue;

            if (levelVal == "entry") {
                repeatEntryAgentTime += at;
                repeatEntryCount++;
            } else if (levelVal == "second") {
                repeatSecondAgentTime += at;
                repeatSecondCount++;
            }
            // Third level excluded from repeat — no split needed
        }
    }

    double avgRepeatEntryHandling = (repeatEntryCount > 0)
        ? std::min(static_cast<double>(repeatEntryAgentTime) / repeatEntryCount, 130.0)
        : 0.0;
    double avgRepeatSecondHandling = (repeatSecondCount > 0)
        ? std::min(static_cast<double>(repeatSecondAgentTime) / repeatSecondCount, 150.0)
        : 0.0;

    // Store call metrics in dailyStatus
    ds.totalIvrDuration           = totalIvrDuration;
    ds.totalIvrCallCount          = totalIvrCallCount;
    ds.avgIvrHandlingTime         = avgIvrHandlingTime;
    ds.totalAgentNotAnswered      = totalAgentNotAnswered;
    ds.totalEntryAgentTime        = totalEntryAgentTime;
    ds.totalEntryLevelCount       = totalEntryLevelCount;
    ds.tcbhEntryLevelCalls        = tcbhEntryLevelCalls;
    ds.tcbhEntryLevelQWait        = tcbhEntryLevelQWait;
    ds.avgEntryLevelHandlingTime  = avgEntryLevelHandlingTime;
    ds.totalSecondAgentTime       = totalSecondAgentTime;
    ds.totalSecondLevelCount      = totalSecondLevelCount;
    ds.tcbhSecondLevelCalls       = tcbhSecondLevelCalls;
    ds.tcbhSecondLevelQWait       = tcbhSecondLevelQWait;
    ds.avgSecondLevelHandlingTime = avgSecondLevelHandlingTime;
    ds.totalThirdAgentTime        = totalThirdAgentTime;
    ds.totalThirdLevelCount       = totalThirdLevelCount;
    ds.tcbhThirdLevelCalls        = tcbhThirdLevelCalls;
    ds.tcbhThirdLevelQWait        = tcbhThirdLevelQWait;
    ds.avgThirdLevelHandlingTime  = avgThirdLevelHandlingTime;
    ds.notAnsweredDueToLowQTime   = notAnsweredDueToLowQTime;
    ds.totalShortAnsweredCalls    = totalShortAnsweredCalls;
    ds.totalShortIvrCalls         = totalShortIvrCalls;

    // Repeat-split metrics
    ds.repeatEntryAgentTime       = repeatEntryAgentTime;
    ds.repeatEntryCount           = repeatEntryCount;
    ds.avgRepeatEntryHandlingTime = avgRepeatEntryHandling;
    ds.repeatSecondAgentTime      = repeatSecondAgentTime;
    ds.repeatSecondCount          = repeatSecondCount;
    ds.avgRepeatSecondHandlingTime = avgRepeatSecondHandling;
   
    int denominator     = cliBased_ ? static_cast<int>(wsccRows.size()) : totalNonIvr;
    ds.totalNonIvrCalls = totalNonIvr;
    ds.repeatCallCount  = repeatCount;
    ds.repeatPercentage =
        (denominator > 0)
            ? (static_cast<double>(repeatCount) * 100.0) / denominator
            : 0.0;
    ds.fullPayment      = (ds.repeatPercentage <= 15.0);

    log()->trace("checkRepeatCalls [{}]: wsccRows={}, repeat={},untagged={}, payment={}, totalNonIvr={}",
                  date, wsccRows.size(), repeatCount, ds.untagged.size(),
                  ds.fullPayment ? "FULL" : "CUT",
                  totalNonIvr);
}

void CCPerformEvaluate::checkAllRepeatCalls(
    const std::string& wsccSheetName,
    const std::string& crmSheetName,
    const std::string& excelDateHeader)
{
    int month = 0, year = 0;
    if (!getMonthYear(month, year)) {
        log()->error("checkAllRepeatCalls: cannot determine month/year");
        return;  // cannot determine month/year
    }

    log()->info("checkAllRepeatCalls: month={}, year={}", month, year);

    // Load both Excel files — each on its own DataExtraction instance
    data_.loadExcel(validator_.getWsccCdrFile().fullPath, wsccSheetName);
    crmData_.loadExcel(validator_.getCrmFile().fullPath, crmSheetName);

    // Extract ALL rows grouped by date in one pass (avoids re-reading 457K rows per day)
    auto wsccByDate = data_.extractExcelRowsByDate(excelDateHeader);
    auto crmByDate  = crmData_.extractExcelRowsByDate("createdOn");

    log()->debug("checkAllRepeatCalls: {} WSCC dates, {} CRM dates",
                   wsccByDate.size(), crmByDate.size());

    int days = daysInMonth(month, year);

    for (int day = 1; day <= days; ++day) {
        std::string date = formatDate(day, month, year);

        auto wsccIt = wsccByDate.find(date);
        const auto& wsccRows = (wsccIt != wsccByDate.end())
                                 ? wsccIt->second
                                 : std::vector<ExcelRecord>{};

        auto crmIt = crmByDate.find(date);
        const auto& crmRows = (crmIt != crmByDate.end())
                                ? crmIt->second
                                : std::vector<ExcelRecord>{};

        CrmLookup crmByMobile = buildCrmLookup(crmRows);
        checkRepeatCalls(date, wsccRows, crmByMobile);
    }
}

// --- Access consolidated daily status ---

const std::map<std::string, DailyStatus>&
CCPerformEvaluate::getDailyStatus() const {
    return dailyStatus_;
}

// --- Write IBD Data to Excel ---

void CCPerformEvaluate::writeIBDData(const std::string& workbookPath) {
    log()->info("writeIBDData: writing to '{}'", workbookPath);

    // Step 1: Determine month, year, sheet name, days
    int month = 0, year = 0;
    if (!getMonthYear(month, year)) {
        log()->error("writeIBDData: cannot determine month/year");
        throw std::runtime_error("writeIBDData: cannot determine month/year");
    }

    std::string sheetName = monthNumberToName(month) + " "
                          + std::to_string(year) + " IBD Data";
    int numDays = daysInMonth(month, year);

    // Step 2: Open workbook, validate sheet
    ExcelCM excel;
    excel.open(workbookPath);

    if (!excel.hasSheet(sheetName)) {
        excel.close();
        throw std::runtime_error("writeIBDData: sheet not found: " + sheetName);
    }

    // Step 3: Write 12 day-wise sub-tables using a reusable lambda.
    // extractFn(ds, avg, count, totalDur) extracts the relevant fields from DailyStatus.
    using ExtractFn = std::function<void(const DailyStatus&,
                                         double&, int&, int&)>;

    auto writeRows = [&](int baseRow, int maxRows, int threshold,
                         const ExtractFn& extractFn) {
        int rowsToWrite = std::min(numDays, maxRows);
        for (int d = 1; d <= rowsToWrite; ++d) {
            int row = baseRow + d;
            std::string dateKey = formatDate(d, month, year);

            double avg = 0.0;
            int count = 0, totalDur = 0;

            auto it = dailyStatus_.find(dateKey);
            if (it != dailyStatus_.end()) {
                extractFn(it->second, avg, count, totalDur);
            }

            // Col A: date
            excel.writeCell(sheetName, cellRef("A", row), dateKey);
            // Col B: average call duration
            excel.writeCell(sheetName, cellRef("B", row), avg);
            // Col C: call count
            excel.writeCell(sheetName, cellRef("C", row),
                            static_cast<double>(count));
            // Col D: total duration (seconds)
            excel.writeCell(sheetName, cellRef("D", row),
                            static_cast<double>(totalDur));
            // Col E: threshold * call count
            double e = static_cast<double>(threshold) * count;
            excel.writeCell(sheetName, cellRef("E", row), e);
            // Col F: Unit (S) = MIN(D, E)
            double f = std::min(static_cast<double>(totalDur), e);
            excel.writeCell(sheetName, cellRef("F", row), f);
            // Col G: Unit (Minutes) = F / 60
            excel.writeCell(sheetName, cellRef("G", row), f / 60.0);
        }
    };

    // --- Extractor functions ---

    // Zero extractor for always-empty tables
    auto zeroExtract = [](const DailyStatus&,
                          double& a, int& c, int& t) {
        a = 0.0; c = 0; t = 0;
    };

    // IVR: always goes to First
    auto ivrExtract = [](const DailyStatus& ds,
                         double& a, int& c, int& t) {
        a = ds.avgIvrHandlingTime;
        c = ds.totalIvrCallCount;
        t = ds.totalIvrDuration;
    };

    // Entry First (unique): total entry minus repeat entry (when !cliBased_)
    auto entryFirstExtract = [this](const DailyStatus& ds,
                                    double& a, int& c, int& t) {
        if (!cliBased_) {
            c = ds.totalEntryLevelCount - ds.repeatEntryCount;
            t = ds.totalEntryAgentTime - ds.repeatEntryAgentTime;
            a = (c > 0) ? std::min(static_cast<double>(t) / c, 130.0) : 0.0;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Entry Repeated: repeat-tagged entry calls only (when !cliBased_)
    auto entryRepeatedExtract = [this](const DailyStatus& ds,
                                       double& a, int& c, int& t) {
        if (!cliBased_) {
            a = ds.avgRepeatEntryHandlingTime;
            c = ds.repeatEntryCount;
            t = ds.repeatEntryAgentTime;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Entry RNA: all entry data when cliBased_ (repeat analysis not applicable)
    auto entryRnaExtract = [this](const DailyStatus& ds,
                                  double& a, int& c, int& t) {
        if (cliBased_) {
            a = ds.avgEntryLevelHandlingTime;
            c = ds.totalEntryLevelCount;
            t = ds.totalEntryAgentTime;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Second First (unique): total second minus repeat second (when !cliBased_)
    auto secondFirstExtract = [this](const DailyStatus& ds,
                                     double& a, int& c, int& t) {
        if (!cliBased_) {
            c = ds.totalSecondLevelCount - ds.repeatSecondCount;
            t = ds.totalSecondAgentTime - ds.repeatSecondAgentTime;
            a = (c > 0) ? std::min(static_cast<double>(t) / c, 150.0) : 0.0;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Second Repeated: repeat-tagged second calls only (when !cliBased_)
    auto secondRepeatedExtract = [this](const DailyStatus& ds,
                                        double& a, int& c, int& t) {
        if (!cliBased_) {
            a = ds.avgRepeatSecondHandlingTime;
            c = ds.repeatSecondCount;
            t = ds.repeatSecondAgentTime;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Second RNA: all second data when cliBased_ (repeat analysis not applicable)
    auto secondRnaExtract = [this](const DailyStatus& ds,
                                   double& a, int& c, int& t) {
        if (cliBased_) {
            a = ds.avgSecondLevelHandlingTime;
            c = ds.totalSecondLevelCount;
            t = ds.totalSecondAgentTime;
        } else {
            a = 0.0; c = 0; t = 0;
        }
    };

    // Third RNA: always goes here (excluded from repeat calc)
    auto thirdRnaExtract = [](const DailyStatus& ds,
                              double& a, int& c, int& t) {
        a = ds.avgThirdLevelHandlingTime;
        c = ds.totalThirdLevelCount;
        t = ds.totalThirdAgentTime;
    };

    // --- Write all 12 sub-tables ---
    // The Excel template already has named Tables with built-in totals rows.
    // Table ranges (from template): header + 30 data rows + 1 totals row.
    // We only write to data cells — do NOT overwrite the table totals rows.
    //       baseRow  maxRows  threshold  extractor
    // maxRows = number of DATA rows in each table (template = header + N data
    // rows + 1 totals row). The 4 small RNA tables have exactly 1 data row;
    // writing 3 here used to spill into the totals row and the row after the
    // table, causing Excel to flag the tables as needing repair and report
    // a circular reference via the auto-computed totals-row SUBTOTAL.
    writeRows(  5,     30,      120,      ivrExtract);            // IVRS First
    writeRows( 40,     30,      120,      zeroExtract);           // IVRS Repeated
    writeRows( 75,     30,      120,      zeroExtract);           // IVRS RNA
    writeRows(113,     30,      130,      entryFirstExtract);     // Entry First
    writeRows(148,     30,      130,      entryRepeatedExtract);  // Entry Repeated
    writeRows(183,      1,      130,      entryRnaExtract);       // Entry RNA   (1 data row only)
    writeRows(191,     30,      150,      secondFirstExtract);    // Second First
    writeRows(226,     30,      150,      secondRepeatedExtract); // Second Repeated
    writeRows(261,      1,      150,      secondRnaExtract);      // Second RNA  (1 data row only)
    writeRows(269,      1,      180,      zeroExtract);           // Third First (1 data row only)
    writeRows(274,      1,      180,      zeroExtract);           // Third Repeated (1 data row only)
    writeRows(279,     30,      180,      thirdRnaExtract);       // Third RNA

    // Step 4: Billed Month table (rows 343-372, cols F-G)
    for (int d = 1; d <= numDays && d <= 30; ++d) {
        int row = 342 + d;
        std::string dateKey = formatDate(d, month, year);
        excel.writeCell(sheetName, cellRef("F", row), dateKey);

        int dailyTotal = 0;
        auto it = dailyStatus_.find(dateKey);
        if (it != dailyStatus_.end()) {
            const auto& ds = it->second;
            dailyTotal = ds.totalIvrCallCount + ds.totalEntryLevelCount
                       + ds.totalSecondLevelCount + ds.totalThirdLevelCount;
        }
        excel.writeCell(sheetName, cellRef("G", row),
                        static_cast<double>(dailyTotal));
    }

    // Step 5: Aggregate monthly SLA sums
    int sumTcbhEntry = 0, sumTcbhEntryQWait = 0;
    int sumTcbhSecond = 0, sumTcbhSecondQWait = 0;
    int sumTcbhThird = 0, sumTcbhThirdQWait = 0;
    int sumIvrCallCount = 0, sumTotalCalls = 0;
    int sumAgentNotAnswered = 0, sumNotAnsweredLowQ = 0;
    int sumNonIvrCalls = 0;

    for (const auto& [date, ds] : dailyStatus_) {
        sumTcbhEntry     += ds.tcbhEntryLevelCalls;
        sumTcbhEntryQWait += ds.tcbhEntryLevelQWait;
        sumTcbhSecond    += ds.tcbhSecondLevelCalls;
        sumTcbhSecondQWait += ds.tcbhSecondLevelQWait;
        sumTcbhThird     += ds.tcbhThirdLevelCalls;
        sumTcbhThirdQWait += ds.tcbhThirdLevelQWait;
        sumIvrCallCount  += ds.totalIvrCallCount;
        sumTotalCalls    += ds.totalIvrCallCount + ds.totalEntryLevelCount
                          + ds.totalSecondLevelCount + ds.totalThirdLevelCount;
        sumAgentNotAnswered += ds.totalAgentNotAnswered;
        sumNotAnsweredLowQ  += ds.notAnsweredDueToLowQTime;
        sumNonIvrCalls   += ds.totalEntryLevelCount + ds.totalSecondLevelCount
                          + ds.totalThirdLevelCount;
    }

    // Step 6: Write SLA summary cells

    // Call Queue Waiting SLA (rows 384-386)
    excel.writeCell(sheetName, "B384",
                    static_cast<double>(sumTcbhEntry));
    excel.writeCell(sheetName, "E384",
                    static_cast<double>(sumTcbhEntry - sumTcbhEntryQWait));
    excel.writeCell(sheetName, "B385",
                    static_cast<double>(sumTcbhSecond));
    excel.writeCell(sheetName, "E385",
                    static_cast<double>(sumTcbhSecond - sumTcbhSecondQWait));
    excel.writeCell(sheetName, "B386",
                    static_cast<double>(sumTcbhThird));
    excel.writeCell(sheetName, "E386",
                    static_cast<double>(sumTcbhThird - sumTcbhThirdQWait));

    // IVR Efficiency KPI (rows 389-390)
    excel.writeCell(sheetName, "F389",
                    static_cast<double>(sumIvrCallCount));
    excel.writeCell(sheetName, "F390",
                    static_cast<double>(sumTotalCalls));

    // Queue Waiting SLA percentages (F = E / B, if B > 0)
    excel.writeFormula(sheetName, "F384", "IF(B384=0,0,E384/B384)");
    excel.writeFormula(sheetName, "F385", "IF(B385=0,0,E385/B385)");
    excel.writeFormula(sheetName, "F386", "IF(B386=0,0,E386/B386)");

    // Call Abandonment SLA (rows 394-395)
    excel.writeCell(sheetName, "F394",
                    static_cast<double>(sumAgentNotAnswered - sumNotAnsweredLowQ));
    excel.writeCell(sheetName, "F395",
                    static_cast<double>(sumNonIvrCalls));

    // Call Abandonment percentage (F396 = F394 / F395)
    excel.writeFormula(sheetName, "F396", "IF(F395=0,0,F394/F395)");

    // Step 7: Save and close
    log()->debug("writeIBDData: saving workbook");
    excel.save();
    excel.close();
    log()->info("writeIBDData: complete");
}

// ============================================================================
// Bill Calculation Sheet Generation
// ============================================================================

// Helper: write a string label to a cell
static void writeStr(ExcelCM& excel, const std::string& sheet,
                     const std::string& cell, const std::string& val) {
    excel.writeCell(sheet, cell, val);
}

// Helper: write a double to a cell
static void writeDbl(ExcelCM& excel, const std::string& sheet,
                     const std::string& cell, double val) {
    excel.writeCell(sheet, cell, val);
}

// Helper: write an Excel formula to a cell (formula without leading '=')
static void writeFml(ExcelCM& excel, const std::string& sheet,
                     const std::string& cell, const std::string& formula) {
    excel.writeFormula(sheet, cell, formula);
}

// Helper: write IBD row labels (column A-B)
static void writeIbdLabels(ExcelCM& excel, const std::string& sheet) {
    writeStr(excel, sheet, "A3",  "IBD");
    writeStr(excel, sheet, "B3",  "A  Unique Calls (Minutes)");
    writeStr(excel, sheet, "B4",  "B  Repeat Calls (Minutes)");
    writeStr(excel, sheet, "B5",  "C  Repeat Non-applicable Calls (Minutes)");
    writeStr(excel, sheet, "B6",  "D  Actual Bill (Minutes)");
    writeStr(excel, sheet, "B7",  "E  Per Minute Charge");
    writeStr(excel, sheet, "B8",  "F  Actual Bill (Rs)");
    writeStr(excel, sheet, "B9",  "G  15% of Unique Calls");
    writeStr(excel, sheet, "B10", "H  Whether (B > G)");
    writeStr(excel, sheet, "B11", "I  Net Connect Minutes after deducting Repeat Calls");
    writeStr(excel, sheet, "B12", "J  Net Amount after deducting Repeat Calls (Rs)");
    writeStr(excel, sheet, "B13", "K  Repeat Call KPI Deduction (Rs)");
}

// Helper: write IBD data references (rows 3-5) using direct cell refs
// totalsRow* = the row in IBD Data sheet where SUM totals are
static void writeIbdDataRefs(ExcelCM& excel, const std::string& sheet,
                              const std::string& ibdSheet,
                              const std::string& col,
                              int firstTotalsRow,
                              int repeatTotalsRow,
                              int rnaTotalsRow) {
    // Cross-sheet reference: ='NOVEMBER 2025 IBD Data'!G<row>
    std::string prefix = "'" + ibdSheet + "'!G";
    writeFml(excel, sheet, col + "3", prefix + std::to_string(firstTotalsRow));
    writeFml(excel, sheet, col + "4", prefix + std::to_string(repeatTotalsRow));
    writeFml(excel, sheet, col + "5", prefix + std::to_string(rnaTotalsRow));
}

// Helper: write IBD calculation formulas (rows 6-13) for one column
static void writeIbdCalcFormulas(ExcelCM& excel, const std::string& sheet,
                                  const std::string& col) {
    // D = A + B + C
    writeFml(excel, sheet, col + "6",
             col + "3+" + col + "4+" + col + "5");
    // F = D * E
    writeFml(excel, sheet, col + "8",
             col + "6*" + col + "7");
    // G = A * 0.15
    writeFml(excel, sheet, col + "9",
             col + "3*0.15");
    // H = IF(B > G, "Yes", "No")
    writeFml(excel, sheet, col + "10",
             "IF(" + col + "4>" + col + "9,\"Yes\",\"No\")");
    // I = IF(H="Yes", A*1.15+C, D)
    writeFml(excel, sheet, col + "11",
             "IF(" + col + "10=\"Yes\"," + col + "3*1.15+" + col + "5,"
             + col + "6)");
    // J = I * E
    writeFml(excel, sheet, col + "12",
             col + "11*" + col + "7");
    // K = F - J
    writeFml(excel, sheet, col + "13",
             col + "8-" + col + "12");
}

// Helper: write SLA penalty pair (value row + penalty row)
static void writeSlaUptimePenalty(ExcelCM& excel, const std::string& sheet,
                                   double uptimeValue) {
    writeStr(excel, sheet, "A18", "5.1.1");
    writeStr(excel, sheet, "B18", "Call Centre System Uptime");
    writeDbl(excel, sheet, "G18", uptimeValue / 100.0);
    writeStr(excel, sheet, "B19", "Penalty (Rs) P");
    writeFml(excel, sheet, "G19",
        "IF(G18>=0.997,0,IF(G18>=0.987,G17*0.01,"
        "IF(G18>=0.97,G17*0.03,IF(G18>=0.95,G17*0.05,"
        "IF(G18>=0.9,G17*0.10,G17*0.15)))))");
}

static void writeSlaAccessibilityPenalty(ExcelCM& excel,
                                          const std::string& sheet,
                                          const std::string& ibdSheet) {
    writeStr(excel, sheet, "A20", "5.1.2");
    writeStr(excel, sheet, "B20", "Call Centre Accessibility @TCBH");
    writeFml(excel, sheet, "G20",
             "'" + ibdSheet + "'!F380");
    writeStr(excel, sheet, "B21", "Penalty (Rs) Q");
    writeFml(excel, sheet, "G21",
        "IF(G20>=0.95,0,IF(G20>=0.9,G17*0.05,"
        "IF(G20>=0.85,G17*0.10,G17*0.15)))");
}

static void writeSlaQueueWaitingPenalty(ExcelCM& excel,
                                         const std::string& sheet,
                                         const std::string& ibdSheet) {
    writeStr(excel, sheet, "A22", "5.1.3");
    writeStr(excel, sheet, "B22", "Call Queue Waiting SLA");
    writeFml(excel, sheet, "D22", "'" + ibdSheet + "'!F384");
    writeFml(excel, sheet, "E22", "'" + ibdSheet + "'!F385");
    writeFml(excel, sheet, "F22", "'" + ibdSheet + "'!F386");
    writeStr(excel, sheet, "B23", "Penalty (Rs) R");
    for (const auto& col : {"D", "E", "F"}) {
        std::string c(col);
        writeFml(excel, sheet, c + "23",
            "IF(" + c + "22>=0.95,0,IF(" + c + "22>=0.85," + c + "8*0.05,"
            + c + "8*0.10))");
    }
    writeFml(excel, sheet, "G23", "D23+E23+F23");
}

static void writeSlaSatisfactionPenalty(ExcelCM& excel,
                                         const std::string& sheet,
                                         double satisfactionScore) {
    writeStr(excel, sheet, "A24", "5.1.4");
    writeStr(excel, sheet, "B24", "Customer Satisfaction");
    writeDbl(excel, sheet, "G24", satisfactionScore);
    writeStr(excel, sheet, "B25", "Penalty (Rs) S");
    writeFml(excel, sheet, "G25",
        "IF(G24>=2,0,IF(G24>=1.1,G17*0.03,G17*0.05))");
}

static void writeSlaAbandonmentPenalty(ExcelCM& excel,
                                        const std::string& sheet,
                                        const std::string& ibdSheet) {
    writeStr(excel, sheet, "A26", "5.1.5");
    writeStr(excel, sheet, "B26", "Call Abandonment");
    writeFml(excel, sheet, "G26",
             "'" + ibdSheet + "'!F396");
    writeStr(excel, sheet, "B27", "Penalty (Rs) T");
    writeFml(excel, sheet, "G27",
        "IF(G26<0.05,0,IF(G26<=0.1,G17*0.05,G17*0.10))");
}

static void writeSlaQualityPenalty(ExcelCM& excel, const std::string& sheet,
                                    double qualityScore) {
    writeStr(excel, sheet, "A28", "5.1.6");
    writeStr(excel, sheet, "B28", "Call Quality");
    writeDbl(excel, sheet, "G28", qualityScore);
    writeStr(excel, sheet, "B29", "Penalty (Rs) U");
    writeFml(excel, sheet, "G29",
        "IF(G28>85,0,IF(G28>80,G17*0.02,IF(G28>=75,G17*0.05,G17*0.10)))");
}

static void writePenaltySummary(ExcelCM& excel, const std::string& sheet) {
    writeStr(excel, sheet, "B30", "V  Total Penalty Amount Calculated (Rs)");
    writeFml(excel, sheet, "G30", "G19+G21+G23+G25+G27+G29");

    writeStr(excel, sheet, "B31", "W  % with respect to Total Bill amount");
    writeFml(excel, sheet, "G31", "IF(G17=0,0,G30/G17)");

    writeStr(excel, sheet, "B32", "X  20% of Total Bill Amount (Rs)");
    writeFml(excel, sheet, "G32", "G17*0.2");

    writeStr(excel, sheet, "B33", "Y  Total Amount excluding Penalties (Rs)");
    writeFml(excel, sheet, "G33", "G17-MIN(G30,G32)");
}

// --- Main writeBillCalculation function ---

void CCPerformEvaluate::writeBillCalculation(
    const std::string& workbookPath,
    const SlaConfig& slaConfig)
{
    log()->info("writeBillCalculation: writing to '{}'", workbookPath);

    int month = 0, year = 0;
    if (!getMonthYear(month, year)) {
        log()->error("writeBillCalculation: cannot determine month/year");
        throw std::runtime_error("writeBillCalculation: cannot determine month/year");
    }

    std::string billSheet = "Bill Calculation "
                          + monthNumberToName(month) + " " + std::to_string(year);
    std::string ibdSheet  = monthNumberToName(month) + " "
                          + std::to_string(year) + " IBD Data";

    ExcelCM excel;
    excel.open(workbookPath);

    if (!excel.hasSheet(billSheet)) {
        excel.addSheet(billSheet);
    }

    // Row 1: Title
    writeStr(excel, billSheet, "A1",
             "Calculation Summary - " + monthNumberToName(month)
             + " " + std::to_string(year));

    // Row 2: Column headers
    writeStr(excel, billSheet, "A2", "Item");
    writeStr(excel, billSheet, "C2", "IVR");
    writeStr(excel, billSheet, "D2", "Agent Seg 1");
    writeStr(excel, billSheet, "E2", "Agent Seg 2");
    writeStr(excel, billSheet, "F2", "Agent Seg 3");
    writeStr(excel, billSheet, "G2", "Total");
    writeStr(excel, billSheet, "H2", "Remarks");

    // IBD section labels
    writeIbdLabels(excel, billSheet);

    // IBD data references (rows 3-5) — direct cell refs to table totals rows.
    // The template has named Tables whose last row is the totals row:
    //   Table418: A5:G37  → totals at G37    Table519: A40:G72  → G72
    //   Table620: A75:G107 → G107            Table721: A113:G145 → G145
    //   Table822: A148:G180 → G180           Table1125: A183:G185 → G185
    //   Table923: A191:G223 → G223           Table1024: A226:G258 → G258
    //   Table1326: A261:G263 → G263          Table1427: A269:G271 → G271
    //   Table1528: A274:G276 → G276          Table1629: A279:G311 → G311
    writeIbdDataRefs(excel, billSheet, ibdSheet, "C",  37,  72, 107);  // IVR
    writeIbdDataRefs(excel, billSheet, ibdSheet, "D", 145, 180, 185);  // Entry
    writeIbdDataRefs(excel, billSheet, ibdSheet, "E", 223, 258, 263);  // Second
    writeIbdDataRefs(excel, billSheet, ibdSheet, "F", 271, 276, 311);  // Third

    // Totals column G (rows 3-5)
    writeFml(excel, billSheet, "G3", "C3+D3+E3+F3");
    writeFml(excel, billSheet, "G4", "C4+D4+E4+F4");
    writeFml(excel, billSheet, "G5", "C5+D5+E5+F5");

    // Per-minute charges (row 7)
    writeDbl(excel, billSheet, "C7", slaConfig.ivrPerMinuteCharge);
    writeDbl(excel, billSheet, "D7", slaConfig.agentPerMinuteCharge);
    writeDbl(excel, billSheet, "E7", slaConfig.agentPerMinuteCharge);
    writeDbl(excel, billSheet, "F7", slaConfig.agentPerMinuteCharge);

    // Calculation formulas for each segment column (rows 6, 8-13)
    for (const auto& col : {"C", "D", "E", "F"}) {
        writeIbdCalcFormulas(excel, billSheet, col);
    }

    // Totals column G (rows 6, 8, 9, 11, 12, 13)
    writeFml(excel, billSheet, "G6",  "C6+D6+E6+F6");
    writeFml(excel, billSheet, "G8",  "C8+D8+E8+F8");
    writeFml(excel, billSheet, "G9",  "C9+D9+E9+F9");
    writeFml(excel, billSheet, "G11", "C11+D11+E11+F11");
    writeFml(excel, billSheet, "G12", "C12+D12+E12+F12");
    writeFml(excel, billSheet, "G13", "C13+D13+E13+F13");

    // Row 17: Total Bill O = J + N (IBD net + OBD net)
    // OBD not implemented yet, so just IBD for now
    writeStr(excel, billSheet, "A17", "Total Bill amount (Rs) O");
    writeFml(excel, billSheet, "G17", "G12");

    // SLA Penalties (rows 18-29)
    writeSlaUptimePenalty(excel, billSheet, slaConfig.systemUptime);
    writeSlaAccessibilityPenalty(excel, billSheet, ibdSheet);
    writeSlaQueueWaitingPenalty(excel, billSheet, ibdSheet);
    writeSlaSatisfactionPenalty(excel, billSheet, slaConfig.customerSatisfaction);
    writeSlaAbandonmentPenalty(excel, billSheet, ibdSheet);
    writeSlaQualityPenalty(excel, billSheet, slaConfig.callQuality);

    // Penalty summary (rows 30-33)
    writePenaltySummary(excel, billSheet);

    // Write SUBTOTAL(109,...) formulas to the totals row of every named table
    // on the IBD Data sheet. The template ships with these formulas already,
    // but OpenXLSX's read-modify-save cycle can corrupt the inherited shared
    // formulas (orphan ref=, stale <v> cache, dangling calcChain). Overwriting
    // them here with clean independent <f> elements gives us a known-good
    // input for the post-process step that follows excel.close().
    wlcc::writeAllIbdSubtotals(excel, ibdSheet);

    excel.save();
    excel.close();

    // Post-process the saved file: drop calcChain part + references, inject
    // fullCalcOnLoad="1", strip orphan ref= attrs from worksheet <f> tags.
    if (!wlcc::postProcessXlsx(workbookPath)) {
        log()->error("writeBillCalculation: postProcessXlsx failed for '{}'",
                     workbookPath);
        throw std::runtime_error("writeBillCalculation: postProcessXlsx failed");
    }
    log()->info("writeBillCalculation: complete");
}

} // namespace wlcc
