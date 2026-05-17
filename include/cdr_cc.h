#pragma once

#include <map>
#include <string>
#include <vector>

namespace wlcc {

// A single CDR record — maps header name to value
using CdrRecord = std::map<std::string, std::string>;

class CdrCC {
public:
    CdrCC();
    ~CdrCC();

    // Load and parse a CDR text file
    void load(const std::string& filepath);

    // Metadata (key:value lines at the top of the file)
    const std::map<std::string, std::string>& getCdrInfo() const;
    std::string getCdrInfoValue(const std::string& key) const;

    // Column headers
    const std::vector<std::string>& getHeaders() const;

    // Data rows
    size_t recordCount() const;
    const CdrRecord& getRecord(size_t index) const;
    const std::vector<CdrRecord>& getAllRecords() const;

    // Extract all values for a given header across all records
    std::vector<std::string> getColumn(const std::string& headerName) const;

    // Extract the year from the "Period" metadata value.
    // Period format: "DD-Mon-YYYY To DD-Mon-YYYY" — parses year from first word.
    int getYear() const;

private:
    std::map<std::string, std::string> cdrInfo_;
    std::vector<std::string> headers_;
    std::vector<CdrRecord> records_;
};

} // namespace wlcc
