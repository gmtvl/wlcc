#pragma once

// ================================================================
//  include/ibd_subtotals.h
//
//  IBD sub-table SUBTOTAL formula writer.
//  Uses ExcelCM (the project's confirmed-working Excel wrapper)
//  instead of raw OpenXLSX, matching the pattern in
//  ccperformevaluate.cpp exactly.
//
//  Formula strings are passed WITHOUT a leading '=' — that is
//  how ExcelCM::writeFormula() expects them (see lines 1154-1165
//  in ccperformevaluate.cpp for reference).
// ================================================================

#include "excel_cm.h"

#include <array>
#include <string>

namespace wlcc {
namespace ibd {

// ----------------------------------------------------------------
//  ColRef — how a totals-row cell references its data column
// ----------------------------------------------------------------
enum class ColRef {
    None,       // leave cell blank (segment produces no data here)
    TableRef,   // structured ref:  TableXXX[Column Name]
    RangeRef,   // plain A1 range:  ColLetter{start}:ColLetter{end}
};

// ----------------------------------------------------------------
//  SubtotalSpec — full description of one sub-table's totals row
// ----------------------------------------------------------------
struct SubtotalSpec {
    int         totalRow;       // Excel row that holds the SUBTOTAL formulas
    const char* tableName;      // Named table object, e.g. "Table418"
    int         dataStartRow;   // First data row (only used for RangeRef cols)
    int         threshold;      // Cap in seconds; drives col-E label:
                                //   "{threshold} * Call Count (S)"
    ColRef colC = ColRef::None;
    ColRef colD = ColRef::None;
    ColRef colE = ColRef::None;
    ColRef colF = ColRef::None;
    // col G is always TableRef — handled implicitly
};

// ----------------------------------------------------------------
//  kSpecs — master config for all 12 IBD sub-tables
// ----------------------------------------------------------------
inline constexpr std::array<SubtotalSpec, 12> kSpecs = {{

    // IVR (120 s) — D/E/F range refs, C/G table refs
    { 37,  "Table418",  6,   120, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },
    { 72,  "Table519",  41,  120, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },
    { 107, "Table620",  76,  120, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },

    // Entry (130 s)
    { 145, "Table721",  114, 130, ColRef::TableRef, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef },
    { 180, "Table822",  149, 130, ColRef::TableRef, ColRef::TableRef, ColRef::None,     ColRef::None     },
    { 185, "Table1125", 183, 130, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    // Second (150 s)
    { 223, "Table923",  193, 150, ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None     },
    { 258, "Table1024", 228, 150, ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None     },
    { 263, "Table1326", 262, 150, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    // Third (180 s)
    { 271, "Table1427", 270, 180, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },
    { 276, "Table1528", 275, 180, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },
    { 311, "Table1629", 280, 180, ColRef::TableRef, ColRef::TableRef, ColRef::TableRef, ColRef::TableRef },
}};

// ----------------------------------------------------------------
//  writeSubtotalRow()
// ----------------------------------------------------------------
inline void writeSubtotalRow(ExcelCM&            excel,
                             const std::string&  sheet,
                             const SubtotalSpec& spec)
{
    const int   row      = spec.totalRow;
    const int   rangeEnd = spec.totalRow - 1;
    const char* tbl      = spec.tableName;

    // No leading '=' — ExcelCM::writeFormula convention
    auto formula = [](const std::string& ref) {
        return "SUBTOTAL(109," + ref + ")";
    };
    auto tRef = [&](const std::string& colName) {
        return std::string(tbl) + "[" + colName + "]";
    };
    auto rRef = [&](const std::string& colLetter) {
        return colLetter + std::to_string(spec.dataStartRow)
             + ":" + colLetter + std::to_string(rangeEnd);
    };
    auto writeCol = [&](const std::string& colLetter,
                        ColRef             mode,
                        const std::string& colName)
    {
        if (mode == ColRef::None) return;
        const std::string ref = (mode == ColRef::TableRef)
                              ? tRef(colName) : rRef(colLetter);
        excel.writeFormula(sheet,
                           colLetter + std::to_string(row),
                           formula(ref));
    };

    writeCol("C", spec.colC, "Call Count");
    writeCol("D", spec.colD, "Total Duration (S)");
    writeCol("E", spec.colE, std::to_string(spec.threshold) + " * Call Count (S)");
    writeCol("F", spec.colF, "Unit (S)");
    excel.writeFormula(sheet, "G" + std::to_string(row),
                       formula(tRef("Unit (Minutes)")));
}

// ----------------------------------------------------------------
//  writeAllSubtotalRows() — writes all 12 in one call
//
//  Integration in writeIBDData():
//      wlcc::ibd::writeAllSubtotalRows(excel, sheetName);
// ----------------------------------------------------------------
inline void writeAllSubtotalRows(ExcelCM&           excel,
                                 const std::string& sheet)
{
    for (const auto& spec : kSpecs) {
        writeSubtotalRow(excel, sheet, spec);
    }
}

} // namespace ibd
} // namespace wlcc
