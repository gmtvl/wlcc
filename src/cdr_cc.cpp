#include "cdr_cc.h"

#include <spdlog/spdlog.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <stdexcept>

namespace {
spdlog::logger* log() {
    static spdlog::logger* inst = [] {
        auto l = spdlog::get("cdr_cc");
        return l ? l.get() : spdlog::default_logger().get();
    }();
    return inst;
}
} // anonymous namespace

namespace wlcc {

// --- Helpers ---

static std::string trim(const std::string& s) {
    auto start = s.find_first_not_of(" \t\r\n");
    if (start == std::string::npos) return "";
    auto end = s.find_last_not_of(" \t\r\n");
    return s.substr(start, end - start + 1);
}

static bool isSeparatorLine(const std::string& line) {
    std::string trimmed = trim(line);
    if (trimmed.empty()) return false;
    return trimmed.find_first_not_of('-') == std::string::npos;
}

// Split a line by tab, trim each token, discard empty tokens
static std::vector<std::string> splitByTab(const std::string& line) {
    std::vector<std::string> tokens;
    std::istringstream ss(line);
    std::string token;
    while (std::getline(ss, token, '\t')) {
        std::string trimmed = trim(token);
        if (!trimmed.empty()) {
            tokens.push_back(trimmed);
        }
    }
    return tokens;
}

// Hardwire for now 
static std::vector<std::string> hardWiredSplit(const std::string& line) {
    std::vector<std::string> tokens;

    tokens.push_back("Call Type"); 
    tokens.push_back("Calling Number");
    tokens.push_back("Called Number");
    tokens.push_back("Date & Time");
    tokens.push_back("Duration");

    return tokens;  
}
// --- CdrCC ---

CdrCC::CdrCC() = default;
CdrCC::~CdrCC() = default;

void CdrCC::load(const std::string& filepath) {
    log()->trace("CdrCC::load({})", filepath);

    // Use absolute path with \\?\ prefix to bypass Windows MAX_PATH (260) limit
    auto abspath = std::filesystem::absolute(filepath);
    std::wstring wpath = abspath.wstring();
    if (wpath.size() > 250 && wpath.substr(0, 4) != L"\\\\?\\") {
        wpath = L"\\\\?\\" + wpath;
    }
    std::ifstream file(wpath);
    if (!file.is_open()) {
        log()->error("CdrCC: cannot open file: {}", filepath);
        throw std::runtime_error("Cannot open file: " + filepath);
    }

    // Clear any previous data
    cdrInfo_.clear();
    headers_.clear();
    records_.clear();

    enum class ParseState { Metadata, HeaderFound, FirstData, Data };
    ParseState state = ParseState::Metadata;
    std::string prevLine;  // buffer to capture the header line (line before separator)

    std::string line;
    while (std::getline(file, line)) {
        if (isSeparatorLine(line)) {
            if (state == ParseState::Metadata) {
                // First separator: the previous non-blank line is the header
                std::string headerLine = trim(prevLine);
                if (!headerLine.empty()) {
                    headers_ = hardWiredSplit(headerLine);
                }
                state = ParseState::HeaderFound;
                continue;
            } else if (state == ParseState::HeaderFound) {
                // Second separator: footer, stop parsing data
                break;
            }
        }

        if (state == ParseState::Metadata) {
            std::string trimmed = trim(line);
            if (trimmed.empty()) {
                // Keep prevLine as-is (don't overwrite with blank)
                continue;
            }
            // Check if this is a metadata line (contains ':')
            auto colonPos = trimmed.find(':');
            if (colonPos != std::string::npos) {
                std::string key = trim(trimmed.substr(0, colonPos));
                std::string value = trim(trimmed.substr(colonPos + 1));
                if (!key.empty()) {
                    cdrInfo_[key] = value;
                }
            }
            prevLine = line;
        } else if (state == ParseState::HeaderFound) {
            // Data row
            std::string trimmed = trim(line);
            if (trimmed.empty()) continue;

            auto fields = splitByTab(line);
            if (fields.empty()) continue;

            CdrRecord record;
            for (size_t i = 0; i < headers_.size(); ++i) {
                if (i < fields.size()) {
                    record[headers_[i]] = fields[i];
                } else {
                    record[headers_[i]] = "";
                }
            }
            records_.push_back(std::move(record));
        }
    }

    log()->trace("CdrCC: parsed {} records, {} headers, {} metadata fields",
                   records_.size(), headers_.size(), cdrInfo_.size());
}

const std::map<std::string, std::string>& CdrCC::getCdrInfo() const {
    return cdrInfo_;
}

std::string CdrCC::getCdrInfoValue(const std::string& key) const {
    auto it = cdrInfo_.find(key);
    if (it != cdrInfo_.end()) {
        return it->second;
    }
    return "";
}

const std::vector<std::string>& CdrCC::getHeaders() const {
    return headers_;
}

size_t CdrCC::recordCount() const {
    return records_.size();
}

const CdrRecord& CdrCC::getRecord(size_t index) const {
    if (index >= records_.size()) {
        throw std::out_of_range("Record index out of range: " + std::to_string(index));
    }
    return records_[index];
}

const std::vector<CdrRecord>& CdrCC::getAllRecords() const {
    return records_;
}

std::vector<std::string> CdrCC::getColumn(const std::string& headerName) const {
    std::vector<std::string> column;
    column.reserve(records_.size());
    for (const auto& record : records_) {
        auto it = record.find(headerName);
        if (it != record.end()) {
            column.push_back(it->second);
        } else {
            column.push_back("");
        }
    }
    return column;
}

int CdrCC::getYear() const {
    // Period format: "01-Nov-2025 To 01-Nov-2025"
    // Extract first word "01-Nov-2025", then parse year after last '-'
    std::string period = getCdrInfoValue("Period");
    if (period.empty()) return 0;

    // Get first word (up to first space)
    auto spacePos = period.find(' ');
    std::string firstWord = (spacePos != std::string::npos)
                            ? period.substr(0, spacePos)
                            : period;

    // Find year after the last '-' in "DD-Mon-YYYY"
    auto lastDash = firstWord.rfind('-');
    if (lastDash == std::string::npos || lastDash + 1 >= firstWord.size()) {
        return 0;
    }

    try {
        return std::stoi(firstWord.substr(lastDash + 1));
    } catch (...) {
        return 0;
    }
}

} // namespace wlcc
