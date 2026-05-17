#include "ccperformevaluate.h"
#include "excel_cm.h"
#include "inputvalidator.h"

#include <spdlog/spdlog.h>
#include <spdlog/sinks/stdout_color_sinks.h>
#include <spdlog/sinks/basic_file_sink.h>

#include <algorithm>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <map>
#include <set>
#include <string>
#include <vector>

namespace fs = std::filesystem;

// --- Command-line parser ---------------------------------------------------

class AppConfig {
public:
    std::string monthCode;  // 3-letter lowercase (e.g. "nov")
    std::string basePath = ".";

    // SLA configuration
    std::string slaConfigPath;  // --sla-config <path>

    // Logging configuration
    bool logConsole = false;       // --log-console: enable console logging
    bool logFile    = false;       // --log-file: enable file logging
    std::string logFilePath = "wlcc.log";  // default log file name
    bool logOff     = false;       // --log-off: disable all logging
    int  logLevel   = 1;           // 1=info, 2=debug, 3=trace

    // Per-module log level overrides: module_name → level (1/2/3/0)
    // Module names: excel_cm, cdr_cc, dataextraction,
    //               inputvalidator, ccperformevaluate, wlcc
    std::map<std::string, int> moduleLogLevels;

    bool parse(int argc, char* argv[]) {
        bool monthGiven = false;

        for (int i = 1; i < argc; ++i) {
            std::string arg = argv[i];

            if (arg == "-h" || arg == "--help") {
                printUsage(argv[0]);
                return false;
            } else if ((arg == "-m" || arg == "--month") && i + 1 < argc) {
                std::string raw = argv[++i];
                if (raw.size() < 3) {
                    std::cerr << "Error: month argument must be at least "
                              << "3 characters (e.g. nov, november).\n";
                    return false;
                }
                monthCode = raw.substr(0, 3);
                std::transform(monthCode.begin(), monthCode.end(),
                               monthCode.begin(), ::tolower);
                if (!isValidMonth(monthCode)) {
                    std::cerr << "Error: '" << raw
                              << "' is not a recognised month.\n";
                    return false;
                }
                monthGiven = true;
            } else if ((arg == "-p" || arg == "--path") && i + 1 < argc) {
                basePath = argv[++i];
            } else if (arg == "--log-console") {
                logConsole = true;
            } else if (arg == "--log-file") {
                logFile = true;
                // Optional: next arg is file path (if it doesn't start with '-')
                if (i + 1 < argc && argv[i + 1][0] != '-') {
                    logFilePath = argv[++i];
                }
            } else if (arg == "--log-off") {
                logOff = true;
            } else if (arg == "--log-level" && i + 1 < argc) {
                std::string lvl = argv[++i];
                if (lvl == "info" || lvl == "1")       logLevel = 1;
                else if (lvl == "debug" || lvl == "2") logLevel = 2;
                else if (lvl == "trace" || lvl == "3") logLevel = 3;
                else {
                    std::cerr << "Error: invalid log level '" << lvl
                              << "'. Use info, debug, or trace.\n";
                    return false;
                }
            } else if (arg == "--log-module" && i + 1 < argc) {
                // Format: module:level  e.g. excel_cm:off, cdr_cc:debug
                std::string spec = argv[++i];
                auto colonPos = spec.find(':');
                if (colonPos == std::string::npos || colonPos == 0) {
                    std::cerr << "Error: --log-module format is name:level "
                              << "(e.g. excel_cm:off, cdr_cc:debug).\n";
                    return false;
                }
                std::string modName = spec.substr(0, colonPos);
                std::string modLvl  = spec.substr(colonPos + 1);
                int lvl = -1;
                if (modLvl == "off" || modLvl == "0")        lvl = 0;
                else if (modLvl == "info" || modLvl == "1")  lvl = 1;
                else if (modLvl == "debug" || modLvl == "2") lvl = 2;
                else if (modLvl == "trace" || modLvl == "3") lvl = 3;
                else {
                    std::cerr << "Error: invalid module log level '" << modLvl
                              << "'. Use off, info, debug, or trace.\n";
                    return false;
                }
                moduleLogLevels[modName] = lvl;
            } else if (arg == "--sla-config" && i + 1 < argc) {
                slaConfigPath = argv[++i];
            } else if (arg == "-v") {
                logConsole = true;
                logLevel = 2;  // debug
            } else if (arg == "-vv") {
                logConsole = true;
                logLevel = 3;  // trace
            } else {
                std::cerr << "Error: unknown option '" << arg << "'.\n";
                printUsage(argv[0]);
                return false;
            }
        }

        // Verify base path exists
        if (!fs::exists(basePath) || !fs::is_directory(basePath)) {
            std::cerr << "Error: path '" << basePath
                      << "' does not exist or is not a directory.\n";
            return false;
        }

        // If month not given, auto-detect from folder
        if (!monthGiven) {
            if (!resolveMonthFromFolder()) {
                return false;
            }
        }

        return true;
    }

    // Initialize spdlog based on parsed flags
    void initLogging() const {
        // All known module names
        static const char* moduleNames[] = {
            "excel_cm", "cdr_cc", "dataextraction",
            "inputvalidator", "ccperformevaluate", "wlcc"
        };

        if (logOff || (!logConsole && !logFile)) {
            spdlog::set_level(spdlog::level::off);
            return;
        }

        std::vector<spdlog::sink_ptr> sinks;

        if (logConsole) {
            sinks.push_back(
                std::make_shared<spdlog::sinks::stderr_color_sink_mt>());
        }

        if (logFile) {
            sinks.push_back(
                std::make_shared<spdlog::sinks::basic_file_sink_mt>(
                    logFilePath, true));
        }

        // Map int level → spdlog level
        auto toSpdLevel = [](int lvl) -> spdlog::level::level_enum {
            switch (lvl) {
                case 0:  return spdlog::level::off;
                case 3:  return spdlog::level::trace;
                case 2:  return spdlog::level::debug;
                default: return spdlog::level::info;
            }
        };

        spdlog::level::level_enum globalLevel = toSpdLevel(logLevel);
        std::string pattern = "[%H:%M:%S.%e] [%^%l%$] [%n] %v";

        // Set the default logger (used by main and as fallback)
        auto defaultLogger = std::make_shared<spdlog::logger>("default",
            sinks.begin(), sinks.end());
        defaultLogger->set_level(globalLevel);
        defaultLogger->set_pattern(pattern);
        spdlog::set_default_logger(defaultLogger);

        // Create a named logger for each module, sharing the same sinks
        for (const auto* modName : moduleNames) {
            auto modLogger = std::make_shared<spdlog::logger>(modName,
                sinks.begin(), sinks.end());
            modLogger->set_pattern(pattern);

            // Apply per-module override if specified, else use global level
            auto it = moduleLogLevels.find(modName);
            if (it != moduleLogLevels.end()) {
                modLogger->set_level(toSpdLevel(it->second));
            } else {
                modLogger->set_level(globalLevel);
            }

            spdlog::register_logger(modLogger);
        }
    }

    void printUsage(const char* prog) {
        std::cerr
            << "Usage: " << prog
            << " [-m <month>] [-p <path>] [logging options]\n\n"
            << "Options:\n"
            << "  -m, --month <month>  Month to process (at least 3 letters,\n"
            << "                       e.g. nov, november). If omitted, the\n"
            << "                       folder must contain exactly one\n"
            << "                       datafiles_<mon> directory.\n"
            << "  -p, --path  <path>   Folder containing the datafiles_<mon>\n"
            << "                       directory. Defaults to current folder.\n"
            << "  -h, --help           Show this help message.\n"
            << "  --sla-config <path>  SLA config file (key=value format).\n"
            << "\nLogging:\n"
            << "  --log-console        Log to console (stderr).\n"
            << "  --log-file [path]    Log to file (default: wlcc.log).\n"
            << "  --log-off            Disable all logging.\n"
            << "  --log-level <level>  Set log level: info, debug, trace.\n"
            << "  --log-module <m:l>   Set per-module level. Modules:\n"
            << "                       excel_cm, cdr_cc, dataextraction,\n"
            << "                       inputvalidator, ccperformevaluate, wlcc\n"
            << "                       Levels: off, info, debug, trace.\n"
            << "                       Example: --log-module excel_cm:off\n"
            << "  -v                   Shortcut for --log-console --log-level debug.\n"
            << "  -vv                  Shortcut for --log-console --log-level trace.\n";
    }

private:
    bool resolveMonthFromFolder() {
        std::vector<std::string> matches;

        for (const auto& entry : fs::directory_iterator(basePath)) {
            if (!entry.is_directory()) continue;
            std::string name = entry.path().filename().string();
            std::string nameLower = name;
            std::transform(nameLower.begin(), nameLower.end(),
                           nameLower.begin(), ::tolower);
            if (nameLower.size() >= 13 &&
                nameLower.substr(0, 10) == "datafiles_") {
                matches.push_back(name);
            }
        }

        if (matches.empty()) {
            std::cerr << "Error: no datafiles_<mon> folder found in '"
                      << basePath << "'.\n";
            return false;
        }
        if (matches.size() > 1) {
            std::cerr << "Error: multiple datafiles folders found in '"
                      << basePath << "':\n";
            for (const auto& m : matches) {
                std::cerr << "  " << m << "\n";
            }
            std::cerr << "Please specify the month with -m <month>.\n";
            return false;
        }

        monthCode = extractMonthCode(matches[0]);
        if (!isValidMonth(monthCode)) {
            std::cerr << "Error: could not extract a valid month code from '"
                      << matches[0] << "'.\n";
            return false;
        }
        return true;
    }

    static std::string extractMonthCode(const std::string& folderName) {
        // folderName is e.g. "datafiles_nov" → extract after "datafiles_"
        std::string code = folderName.substr(10, 3);
        std::transform(code.begin(), code.end(), code.begin(), ::tolower);
        return code;
    }

    static bool isValidMonth(const std::string& code) {
        static const char* months[] = {
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec"
        };
        for (const auto* m : months) {
            if (code == m) return true;
        }
        return false;
    }
};

// --- Main ------------------------------------------------------------------

int main(int argc, char* argv[]) {
    AppConfig config;
    if (!config.parse(argc, argv)) {
        return 1;
    }

    // Initialize logging based on parsed flags
    config.initLogging();

    // Get module logger for main (falls back to default if not registered)
    auto log = spdlog::get("wlcc");
    if (!log) log = spdlog::default_logger();

    log->info("wlcc starting — month={}, path={}",
              config.monthCode, fs::absolute(config.basePath).string());

    std::cout << "Month : " << config.monthCode << "\n";
    std::cout << "Path  : " << fs::absolute(config.basePath).string() << "\n\n";

    // Validate directory structure
    wlcc::InputValidator validator;
    if (!validator.validate(config.basePath, config.monthCode)) {
        std::cerr << "Validation failed:\n";
        for (const auto& err : validator.getErrors()) {
            std::cerr << "  - " << err << "\n";
            log->error("Validation error: {}", err);
        }
        return 1;
    }

    log->info("Validation passed");

    // Print discovered files
    std::cout << "Data folder : " << validator.getDataFilesPath() << "\n";
    std::cout << "CRM file    : " << validator.getCrmFile().filename << "\n";
    std::cout << "WSCC CDR    : " << validator.getWsccCdrFile().filename << "\n";

    if (validator.hasWsccObFile()) {
        std::cout << "WSCC OB     : "
                  << validator.getWsccObFile().filename << "\n";
    }

    const auto& inbound = validator.getMscInboundCdrFiles();
    const auto& outbound = validator.getMscOutboundCdrFiles();
    std::cout << "MSC CDR     : " << inbound.size() << " inbound, "
              << outbound.size() << " outbound text files\n";

    log->debug("CRM file: {}", validator.getCrmFile().fullPath);
    log->debug("WSCC CDR file: {}", validator.getWsccCdrFile().fullPath);
    log->debug("MSC CDR: {} inbound, {} outbound files",
               inbound.size(), outbound.size());

    std::cout << "\nValidation passed. Ready for processing.\n";

    // Wire up processing pipeline
    wlcc::DataExtraction wsccData;   // for WSCC CDR Excel + MSC CDR text files
    wlcc::DataExtraction crmData;    // for CRM Excel file
    wlcc::CCPerformEvaluate evaluator(validator, wsccData, crmData);

    try {

    // Step 1: Collect unique parent directories of all MSC CDR text files
    //         (loadCdrDirectories uses non-recursive directory_iterator)
    std::set<std::string> cdrDirSet;
    for (const auto& f : inbound) {
        cdrDirSet.insert(fs::path(f.fullPath).parent_path().string());
    }
    for (const auto& f : outbound) {
        cdrDirSet.insert(fs::path(f.fullPath).parent_path().string());
    }
    std::vector<std::string> cdrDirs(cdrDirSet.begin(), cdrDirSet.end());
    log->debug("CDR directories to scan: {}", cdrDirs.size());

    // Step 2: Discover WSCC CDR Excel sheet name (varies by month)
    std::string wsccSheet;
    {
        wlcc::ExcelCM probe;
        probe.open(validator.getWsccCdrFile().fullPath);
        auto sheets = probe.getSheetNames();
        probe.close();
        if (sheets.empty()) {
            std::cerr << "Error: WSCC CDR Excel has no sheets.\n";
            log->error("WSCC CDR Excel has no sheets");
            return 1;
        }
        wsccSheet = sheets[0];
        log->debug("WSCC CDR sheet names: {} (using first: {})",
                   sheets.size(), wsccSheet);
    }
    std::cout << "WSCC CDR sheet: " << wsccSheet << std::endl;

    // Step 3: Load WSCC CDR Excel
    std::cout << "Loading WSCC CDR Excel (sheet: " << wsccSheet << ")..."
              << std::endl;
    log->info("Loading WSCC CDR Excel (sheet: {})...", wsccSheet);
    wsccData.loadExcel(validator.getWsccCdrFile().fullPath, wsccSheet);
    std::cout << "WSCC CDR Excel loaded." << std::endl;
    log->info("WSCC CDR Excel loaded");

    std::cout << "Loading MSC CDR data from " << cdrDirs.size()
              << " directories..." << std::endl;
    log->info("Loading MSC CDR data from {} directories...", cdrDirs.size());
    wsccData.loadCdrDirectories(cdrDirs);
    std::cout << "MSC CDR data loaded." << std::endl;
    log->info("MSC CDR data loaded");

    evaluator.validateAllDates("Date");

    // Step 4: Check repeat calls for all dates (loads WSCC+CRM Excel internally)
    std::cout << "Checking repeat calls for all dates..." << std::endl;
    log->info("Checking repeat calls for all dates...");
    evaluator.checkAllRepeatCalls(wsccSheet);
    std::cout << "Repeat call check complete." << std::endl;
    log->info("Repeat call check complete");

    // Step 5: Print daily status report
    const auto& dailyStatus = evaluator.getDailyStatus();
    std::cout << "\n========== Daily Status Report ==========\n";
    std::cout << std::left << std::setw(10) << "Date"
              << std::right
              << std::setw(7)  << "Untag"
              << std::setw(12) << "NonIVR"
              << std::setw(6)  << "Rpt"
              << std::setw(8)  << "Rpt%"
              << std::setw(6)  << "Pay"
              << std::setw(6)  << "IVR#"
              << std::setw(8)  << "IVRAvg"
              << std::setw(7)  << "Entry"
              << std::setw(8)  << "EntAvg"
              << std::setw(7)  << "Sec"
              << std::setw(8)  << "SecAvg"
              << std::setw(7)  << "Third"
              << std::setw(8)  << "ThdAvg"
              << std::setw(6)  << "NAns"
              << "\n";
    std::cout << std::string(113, '-') << "\n";

    for (const auto& [date, ds] : dailyStatus) {
        std::cout << std::left << std::setw(10) << date
                  << std::right
                  << std::setw(7)  << ds.untagged.size()
                  << std::setw(12) << (ds.totalNonIvrCalls + ds.totalThirdLevelCount)
                  << std::setw(6)  << ds.repeatCallCount
                  << std::setw(7)  << std::fixed << std::setprecision(1)
                                   << ds.repeatPercentage << "%"
                  << std::setw(6)  << (ds.fullPayment ? "FULL" : "CUT")
                  << std::setw(6)  << ds.totalIvrCallCount
                  << std::setw(8)  << std::setprecision(1) << ds.avgIvrHandlingTime
                  << std::setw(7)  << ds.totalEntryLevelCount
                  << std::setw(8)  << std::setprecision(1) << ds.avgEntryLevelHandlingTime
                  << std::setw(7)  << ds.totalSecondLevelCount
                  << std::setw(8)  << std::setprecision(1) << ds.avgSecondLevelHandlingTime
                  << std::setw(7)  << ds.totalThirdLevelCount
                  << std::setw(8)  << std::setprecision(1) << ds.avgThirdLevelHandlingTime
                  << std::setw(6)  << ds.totalAgentNotAnswered
                  << "\n";
    }

    // Print TCBH summary
    std::cout << "\n========== TCBH Summary ==========\n";
    int sumTcbhEntry = 0, sumTcbhEntryQ = 0;
    int sumTcbhSecond = 0, sumTcbhSecondQ = 0;
    int sumTcbhThird = 0, sumTcbhThirdQ = 0;
    for (const auto& [date, ds] : dailyStatus) {
        sumTcbhEntry   += ds.tcbhEntryLevelCalls;
        sumTcbhEntryQ  += ds.tcbhEntryLevelQWait;
        sumTcbhSecond  += ds.tcbhSecondLevelCalls;
        sumTcbhSecondQ += ds.tcbhSecondLevelQWait;
        sumTcbhThird   += ds.tcbhThirdLevelCalls;
        sumTcbhThirdQ  += ds.tcbhThirdLevelQWait;
    }
    std::cout << "Entry  TCBH calls: " << sumTcbhEntry
              << "  QWait >= 60s: " << sumTcbhEntryQ << "\n";
    std::cout << "Second TCBH calls: " << sumTcbhSecond
              << "  QWait >= 45s: " << sumTcbhSecondQ << "\n";
    std::cout << "Third  TCBH calls: " << sumTcbhThird
              << "  QWait >= 30s: " << sumTcbhThirdQ << "\n";

    std::cout << "\nDone.\n";

    // Step 6: Write IBD Data sheet (must run before Bill Calculation)
    std::string calcFile = "datafiles_" + config.monthCode
                         + "\\Centre wise_Calc_Sheet_2025-11 (1).xlsx";
    fs::path calcPath = fs::path(config.basePath) / calcFile;

    // Refresh the working file from the pristine template every run. wlcc
    // opens-modifies-saves calcPath in place; without this refresh, residual
    // data from a previous (possibly buggy) run carries over and produces
    // corrupted output. The template lives alongside the working file,
    // differing only by the " (1)" suffix.
    fs::path templatePath = calcPath;
    {
        std::string stem = calcPath.stem().string();           // "...11 (1)"
        const std::string suffix = " (1)";
        if (stem.size() >= suffix.size()
            && stem.compare(stem.size() - suffix.size(),
                            suffix.size(), suffix) == 0) {
            stem.resize(stem.size() - suffix.size());
            templatePath = calcPath.parent_path()
                         / (stem + calcPath.extension().string());
        }
    }
    if (fs::exists(templatePath)) {
        std::error_code ec;
        fs::copy_file(templatePath, calcPath,
                      fs::copy_options::overwrite_existing, ec);
        if (ec) {
            std::cerr << "WARNING: could not refresh '" << calcPath.string()
                      << "' from template '" << templatePath.string()
                      << "': " << ec.message() << "\n";
        } else {
            std::cout << "Refreshed working file from template: "
                      << templatePath.filename().string() << "\n";
        }
    } else {
        std::cerr << "WARNING: template '" << templatePath.string()
                  << "' not found — using existing working file as-is.\n";
    }

    std::cout << "\nWriting IBD Data to: " << calcPath.string() << std::endl;
    evaluator.writeIBDData(calcPath.string());
    std::cout << "IBD Data written." << std::endl;

    // Step 7: Write Bill Calculation sheet (references IBD Data tables)
    wlcc::SlaConfig slaConfig;
    if (!config.slaConfigPath.empty()) {
        wlcc::SlaConfig::loadFromFile(config.slaConfigPath, slaConfig);
    }
    std::cout << "Writing Bill Calculation..." << std::endl;
    evaluator.writeBillCalculation(calcPath.string(), slaConfig);
    std::cout << "Bill Calculation written." << std::endl;

    log->info("Processing complete");

    } catch (const std::exception& e) {
        std::cerr << "\nERROR: " << e.what() << "\n";
        log->error("Fatal exception: {}", e.what());
        return 2;
    } catch (...) {
        std::cerr << "\nERROR: unknown exception\n";
        log->error("Fatal unknown exception");
        return 2;
    }

    return 0;
}
