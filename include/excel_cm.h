#pragma once

#include <cstdint>
#include <memory>
#include <string>
#include <variant>
#include <vector>

namespace wlcc {

// Cell values can be: empty, string, int, double, or bool
using CellValue = std::variant<std::monostate, std::string, int64_t, double, bool>;

class ExcelCM {
public:
    ExcelCM();
    ~ExcelCM();

    ExcelCM(const ExcelCM&) = delete;
    ExcelCM& operator=(const ExcelCM&) = delete;
    ExcelCM(ExcelCM&&) noexcept;
    ExcelCM& operator=(ExcelCM&&) noexcept;

    // File operations
    void create(const std::string& filepath);
    void open(const std::string& filepath);
    void save();
    void saveAs(const std::string& filepath);
    void close();
    bool isOpen() const;

    // Cell operations
    CellValue readCell(const std::string& sheetName,
                       const std::string& cellRef) const;
    void writeCell(const std::string& sheetName,
                   const std::string& cellRef,
                   const CellValue& value);
    void writeCell(const std::string& sheetName,
                   const std::string& cellRef,
                   const std::string& data);
    void writeCell(const std::string& sheetName,
                   const std::string& cellRef,
                   double data);

    // Write an Excel formula to a cell (e.g. "SUM(A1:A10)" without leading '=')
    void writeFormula(const std::string& sheetName,
                      const std::string& cellRef,
                      const std::string& formula);

    // Write a real computing Excel formula cell using cell.formula() — this
    // creates an <f> XML element (NOT a string cell). Required for formulas
    // that Excel must compute (SUBTOTAL etc.). The saved xlsx must afterwards
    // be post-processed to drop the dangling xl/calcChain.xml reference that
    // OpenXLSX leaves behind. formula: text without leading '='.
    void writeFormulaCell(const std::string& sheetName,
                          const std::string& cellRef,
                          const std::string& formula);

    // Sheet helpers
    bool hasSheet(const std::string& sheetName) const;
    void addSheet(const std::string& sheetName);
    std::vector<std::string> getSheetNames() const;

    // Row / column operations
    std::vector<CellValue> readRow(const std::string& sheetName,
                                   uint32_t rowNumber) const;
    std::vector<CellValue> readColumn(const std::string& sheetName,
                                      uint16_t columnNumber) const;
    std::vector<CellValue> readColumn(const std::string& sheetName,
                                      const std::string& columnLetter) const;

    // Header detection and storage
    bool detectHeader(const std::string& sheetName);
    bool hasHeader() const;
    const std::vector<std::string>& getHeader() const;
    void setHeader(const std::vector<std::string>& header);

    // Map a header value to its Excel column letter (e.g. "Age" -> "B")
    std::string headerToColumnLetter(const std::string& headerValue) const;

    // Map an Excel column letter to its header value (e.g. "B" -> "Age")
    std::string columnLetterToHeader(const std::string& columnLetter) const;

private:
    struct Impl;
    std::unique_ptr<Impl> m_impl;
};

std::string cellValueToString(const CellValue& value);
uint16_t columnLetterToNumber(const std::string& columnLetter);
std::string columnNumberToLetter(uint16_t columnNumber);

} // namespace wlcc
