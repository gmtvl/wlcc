#include "inputvalidator.h"

#include <spdlog/spdlog.h>

#include <algorithm>
#include <cctype>
#include <filesystem>

namespace fs = std::filesystem;

namespace {
spdlog::logger* log() {
    static spdlog::logger* inst = [] {
        auto l = spdlog::get("inputvalidator");
        return l ? l.get() : spdlog::default_logger().get();
    }();
    return inst;
}
} // anonymous namespace

namespace wlcc {

InputValidator::InputValidator() = default;
InputValidator::~InputValidator() = default;

// --- Helpers ---

bool InputValidator::containsIgnoreCase(const std::string& str,
                                        const std::string& sub) {
    if (sub.empty()) return true;
    if (str.size() < sub.size()) return false;

    std::string lowerStr = str;
    std::string lowerSub = sub;
    std::transform(lowerStr.begin(), lowerStr.end(), lowerStr.begin(),
                   [](unsigned char c) { return std::tolower(c); });
    std::transform(lowerSub.begin(), lowerSub.end(), lowerSub.begin(),
                   [](unsigned char c) { return std::tolower(c); });

    return lowerStr.find(lowerSub) != std::string::npos;
}

bool InputValidator::isExcelFile(const std::string& filename) {
    return containsIgnoreCase(filename, ".xlsx")
        || containsIgnoreCase(filename, ".xls");
}

bool InputValidator::isTempFile(const std::string& filename) {
    return filename.size() >= 2
        && filename[0] == '~'
        && filename[1] == '$';
}

// --- Reset ---

void InputValidator::reset() {
    monthCode_.clear();
    dataFilesPath_.clear();
    crmFile_ = {};
    wsccCdrFile_ = {};
    wsccObFile_ = {};
    hasWsccOb_ = false;
    mscInboundFiles_.clear();
    mscOutboundFiles_.clear();
    errors_.clear();
    validated_ = false;
}

// --- Main validation entry point ---

bool InputValidator::validate(const std::string& basePath,
                              const std::string& monthCode) {
    reset();

    // Build the expected directory name: datafiles_<mon>
    std::string lowerMonth = monthCode;
    std::transform(lowerMonth.begin(), lowerMonth.end(), lowerMonth.begin(),
                   [](unsigned char c) { return std::tolower(c); });

    monthCode_ = lowerMonth;
    dataFilesPath_ = (fs::path(basePath) / ("datafiles_" + lowerMonth)).string();

    if (!fs::exists(dataFilesPath_) || !fs::is_directory(dataFilesPath_)) {
        errors_.push_back("Data directory not found: " + dataFilesPath_);
        log()->error("Data directory not found: {}", dataFilesPath_);
        return false;
    }

    log()->debug("Validating directory structure: {}", dataFilesPath_);

    bool rootOk   = validateRoot();
    bool wsccOk   = validateWsccCdr();
    bool mscOk    = validateMscCdr();

    validated_ = rootOk && wsccOk && mscOk;
    log()->debug("Validation result: root={}, wscc={}, msc={} => {}",
                   rootOk, wsccOk, mscOk, validated_ ? "PASS" : "FAIL");
    return validated_;
}

// --- Root validation ---
// Look for exactly one CRM xlsx, and verify MSC CDR / WSCC CDR folders exist.

bool InputValidator::validateRoot() {
    bool ok = true;

    // Scan root for CRM Excel files
    std::vector<FileInfo> crmFiles;

    for (const auto& entry : fs::directory_iterator(dataFilesPath_)) {
        if (!entry.is_regular_file()) continue;

        std::string fname = entry.path().filename().string();
        if (isTempFile(fname)) continue;
        if (!isExcelFile(fname)) continue;

        if (containsIgnoreCase(fname, "crm")) {
            crmFiles.push_back({fname, entry.path().string()});
        }
    }

    if (crmFiles.empty()) {
        errors_.push_back("No CRM Excel file found in: " + dataFilesPath_);
        ok = false;
    } else if (crmFiles.size() > 1) {
        errors_.push_back("Multiple CRM Excel files found in: " + dataFilesPath_
                          + " (expected exactly one)");
        for (const auto& f : crmFiles) {
            errors_.push_back("  - " + f.filename);
        }
        ok = false;
    } else {
        crmFile_ = crmFiles[0];
        log()->trace("Found CRM file: {}", crmFile_.filename);
    }

    // Check for required subdirectories
    fs::path mscPath  = fs::path(dataFilesPath_) / "MSC CDR";
    fs::path wsccPath = fs::path(dataFilesPath_) / "WSCC CDR";

    if (!fs::exists(mscPath) || !fs::is_directory(mscPath)) {
        errors_.push_back("Required folder not found: " + mscPath.string());
        ok = false;
    }

    if (!fs::exists(wsccPath) || !fs::is_directory(wsccPath)) {
        errors_.push_back("Required folder not found: " + wsccPath.string());
        ok = false;
    }

    return ok;
}

// --- WSCC CDR validation ---
// Exactly one xlsx with "CDR" (excluding "OB" matches).
// At most one xlsx with "OB" (optional).

bool InputValidator::validateWsccCdr() {
    fs::path wsccPath = fs::path(dataFilesPath_) / "WSCC CDR";
    if (!fs::exists(wsccPath) || !fs::is_directory(wsccPath)) {
        return false; // already reported in validateRoot
    }

    bool ok = true;

    // Collect all non-temp Excel files
    std::vector<FileInfo> obFiles;
    std::vector<FileInfo> cdrFiles;

    for (const auto& entry : fs::directory_iterator(wsccPath)) {
        if (!entry.is_regular_file()) continue;

        std::string fname = entry.path().filename().string();
        if (isTempFile(fname)) continue;
        if (!isExcelFile(fname)) continue;

        // Check "OB" first so "OB CDR.xlsx" goes to OB, not CDR
        if (containsIgnoreCase(fname, "ob")) {
            obFiles.push_back({fname, entry.path().string()});
        } else if (containsIgnoreCase(fname, "cdr")) {
            cdrFiles.push_back({fname, entry.path().string()});
        }
    }

    // Validate CDR file (exactly one required)
    if (cdrFiles.empty()) {
        errors_.push_back("No CDR Excel file found in: " + wsccPath.string());
        ok = false;
    } else if (cdrFiles.size() > 1) {
        errors_.push_back("Multiple CDR Excel files found in: "
                          + wsccPath.string() + " (expected exactly one)");
        for (const auto& f : cdrFiles) {
            errors_.push_back("  - " + f.filename);
        }
        ok = false;
    } else {
        wsccCdrFile_ = cdrFiles[0];
        log()->trace("Found WSCC CDR file: {}", wsccCdrFile_.filename);
    }

    // Validate OB file (at most one, optional)
    if (obFiles.size() > 1) {
        errors_.push_back("Multiple OB Excel files found in: "
                          + wsccPath.string() + " (expected at most one)");
        for (const auto& f : obFiles) {
            errors_.push_back("  - " + f.filename);
        }
        ok = false;
    } else if (obFiles.size() == 1) {
        wsccObFile_ = obFiles[0];
        hasWsccOb_ = true;
    }
    // size == 0 is fine — OB file is optional

    return ok;
}

// --- MSC CDR validation ---
// Recursively scan for .txt files, skip __MACOSX.
// Classify by "ob" substring: outbound vs inbound.

bool InputValidator::validateMscCdr() {
    fs::path mscPath = fs::path(dataFilesPath_) / "MSC CDR";
    if (!fs::exists(mscPath) || !fs::is_directory(mscPath)) {
        return false; // already reported in validateRoot
    }

    for (const auto& entry : fs::recursive_directory_iterator(mscPath)) {
        if (!entry.is_regular_file()) continue;

        // Skip anything under __MACOSX
        std::string pathStr = entry.path().string();
        if (containsIgnoreCase(pathStr, "__MACOSX")) continue;

        std::string fname = entry.path().filename().string();

        // Skip temp files and non-txt files
        if (isTempFile(fname)) continue;
        if (!containsIgnoreCase(fname, ".txt")) continue;

        FileInfo info{fname, entry.path().string()};

        if (containsIgnoreCase(fname, "ob")) {
            mscOutboundFiles_.push_back(std::move(info));
        } else {
            mscInboundFiles_.push_back(std::move(info));
        }
    }

    log()->trace("MSC CDR: {} inbound, {} outbound files found",
                   mscInboundFiles_.size(), mscOutboundFiles_.size());
    return true;
}

// --- Accessors ---

const std::vector<std::string>& InputValidator::getErrors() const {
    return errors_;
}

const std::string& InputValidator::getMonthCode() const {
    return monthCode_;
}

int InputValidator::getMonthNumber() const {
    static const char* months[] = {
        "jan", "feb", "mar", "apr", "may", "jun",
        "jul", "aug", "sep", "oct", "nov", "dec"
    };
    for (int i = 0; i < 12; ++i) {
        if (monthCode_ == months[i]) return i + 1;
    }
    return 0;
}

const std::string& InputValidator::getDataFilesPath() const {
    return dataFilesPath_;
}

const FileInfo& InputValidator::getCrmFile() const {
    return crmFile_;
}

const FileInfo& InputValidator::getWsccCdrFile() const {
    return wsccCdrFile_;
}

const FileInfo& InputValidator::getWsccObFile() const {
    return wsccObFile_;
}

bool InputValidator::hasWsccObFile() const {
    return hasWsccOb_;
}

const std::vector<FileInfo>& InputValidator::getMscInboundCdrFiles() const {
    return mscInboundFiles_;
}

const std::vector<FileInfo>& InputValidator::getMscOutboundCdrFiles() const {
    return mscOutboundFiles_;
}

} // namespace wlcc
