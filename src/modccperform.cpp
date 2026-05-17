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
    bool noNonIvrCalls = false; // flag to indicate if there are no non-IVR calls

    // --- Filter non-IVR rows for repeat call logic ---
    // Filter non-IVR rows, excluding short abandonments and third-level calls.
    // Non-IVR: FRL != "0" (and FRL not empty).
    // Short abandonment: TotalTimeAtAgent is blank or < 10.
    // Third-level: "level" header value is "third" (case-insensitive).
    std::vector<const ExcelRecord*> nonIvrRows;
    for (const auto& row : wsccRows) {
        auto frlIt = row.find("FRL");
        if (frlIt == row.end()) continue;

        std::string frlVal = cellValueToString(frlIt->second);
        if (frlVal.empty()) continue; // since empty FRL is treated as short abandonment, skip it
        // IVR metrics: FRL == 0, IVRDuration >= 10
        if (frlVal == "0") {
            auto ivrDurIt = row.find("IVRDuration");
            if (ivrDurIt != row.end()) {
                int ivrDur = parseDurationToSeconds(ivrDurIt->second);
                if (ivrDur >= 10) {
                    totalIvrDuration += ivrDur;
                    totalIvrCallCount++;
                }
            }
            continue;  // skip IVR calls from non-IVR processing
        }

        auto queueWaitTimeIt = row.find("QueDuration");
        if (queueWaitTimeIt == row.end()) continue;  // should not get to here

        int queueWaitTime = parseDurationToSeconds(queueWaitTimeIt->second);
 
        // Agent not answered: FRL == 3
        if (frlVal == "3") {
            totalAgentNotAnswered++;
            if (queueWaitTime < 10) notAnsweredDueToLowQTime++;
        }

        // Exclude short abandonment calls
        // since FRL == 0 has already been skipped, 
        // we can treat missing or blank TotalTimeAtAgent as short abandonment
        // but will we even reach this point for FRL == 0 ? let's keep the logic consistent
        // and check for blank or <10 for all non-IVR calls
        auto agentTimeIt = row.find("TotalTimeAtAgent");
        if (agentTimeIt == row.end()) continue;  // blank → short abandonment

        int agentTime = parseDurationToSeconds(agentTimeIt->second);
        if (agentTime < 10) continue;  // < 10 seconds → short abandonment
       
        auto tcbh = row.find("Hour");
        if (tcbh == row.end()) continue;  // should not get to here

        int tcbhValue = parseDurationToSeconds(tcbh->second);

        // Exclude third-level calls
        auto levelIt = row.find("level");
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
            if (levelVal == "third") continue;
        }
        nonIvrRows.push_back(&row);
    }
    DailyStatus& ds = dailyStatus_[date];
    int repeatCount = 0;
    int totalNonIvr = static_cast<int>(nonIvrRows.size());
    if (totalNonIvr == 0) noNonIvrCalls = true;

    // Compute averages, capped at maximum values
    if (!noNonIvrCalls) {
        // Group non-IVR rows by CLI last 10 digits
        std::map<std::string, std::vector<const ExcelRecord*>> callerGroups;
        for (const auto* row : nonIvrRows) {
            auto cliIt = row->find("CLI");
            if (cliIt == row->end()) continue;

            std::string cli = cellValueToString(cliIt->second);
            std::string last10 = extractLast10(cli);
            if (!last10.empty()) {
                callerGroups[last10].push_back(row);
            }
        }
        //checking tagging of qualified CDR rows of the day with the CRM
        //checking equality through containers callerGroups vs crmByMobile
        if (callerGroups.size() != crmByMobile.size()) cliBased_ = true;

        if (!cliBased_) {
            for (const auto& [mobile, cdrVec] : callerGroups) {
                auto it = crmByMobile.find(mobile);
                if (it == crmByMobile.end() || it->second.empty()){
                    cliBased_ = true;
                    break;
                }
                const auto& crmVec = it->second;
                if (cdrVec.size() > crmVec.size()) {
                    cliBased_ = true;
                    break;
                }
            }
        }

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
   
    int denominator     = cliBased_ ? static_cast<int>(wsccRows.size()) : totalNonIvr;
    ds.totalNonIvrCalls = denominator;
    ds.repeatCallCount  = repeatCount;
    ds.repeatPercentage =
        (denominator > 0)
            ? (static_cast<double>(repeatCount) * 100.0) / denominator
            : 0.0;
    ds.fullPayment      = (ds.repeatPercentage <= 15.0);
}
