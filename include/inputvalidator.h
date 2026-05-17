#pragma once

#include <string>
#include <vector>

namespace wlcc {

// Holds a discovered file's name and full path
struct FileInfo {
    std::string filename;   // just the filename (e.g., "NOV CRM.xlsx")
    std::string fullPath;   // absolute path     (e.g., "C:/data/datafiles_nov/NOV CRM.xlsx")
};

class InputValidator {
public:
    InputValidator();
    ~InputValidator();

    // Validate the datafiles_<monthCode> directory under basePath.
    // basePath:  parent directory (e.g., "C:/Dev/wlcc")
    // monthCode: 3-letter month code (e.g., "nov", "jul")
    // Returns true if all checks pass, false otherwise.
    bool validate(const std::string& basePath, const std::string& monthCode);

    // Errors collected during the last validate() call
    const std::vector<std::string>& getErrors() const;

    // --- Discovered file accessors ---

    // 3-letter month code extracted from the datafiles_<mon> folder name
    const std::string& getMonthCode() const;

    // Month as an integer (1-12), derived from monthCode
    int getMonthNumber() const;

    const std::string& getDataFilesPath() const;

    // CRM Excel file (root of datafiles_<mon>)
    const FileInfo& getCrmFile() const;

    // WSCC CDR Excel file (must have "CDR" in name)
    const FileInfo& getWsccCdrFile() const;

    // WSCC OB Excel file (has "OB" in name, optional)
    const FileInfo& getWsccObFile() const;
    bool hasWsccObFile() const;

    // MSC CDR inbound .txt files (no "ob" in filename)
    const std::vector<FileInfo>& getMscInboundCdrFiles() const;

    // MSC CDR outbound .txt files ("ob" in filename)
    const std::vector<FileInfo>& getMscOutboundCdrFiles() const;

private:
    std::string monthCode_;
    std::string dataFilesPath_;
    FileInfo crmFile_;
    FileInfo wsccCdrFile_;
    FileInfo wsccObFile_;
    bool hasWsccOb_ = false;
    std::vector<FileInfo> mscInboundFiles_;
    std::vector<FileInfo> mscOutboundFiles_;
    std::vector<std::string> errors_;
    bool validated_ = false;

    // Clear all state before a new validation run
    void reset();

    // Sub-validation steps (each appends to errors_)
    bool validateRoot();
    bool validateWsccCdr();
    bool validateMscCdr();

    // Helpers
    static bool containsIgnoreCase(const std::string& str, const std::string& sub);
    static bool isExcelFile(const std::string& filename);
    static bool isTempFile(const std::string& filename);
};

} // namespace wlcc
