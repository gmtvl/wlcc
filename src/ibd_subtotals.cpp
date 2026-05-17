#include "ibd_subtotals.h"

// ============================================================
//  IBD Data sheet — SUBTOTAL formula writer
//  File: excerpt to be integrated into ccperformevaluate.cpp
//
//  Writes =SUBTOTAL(109,...) formulas to the totals row of
//  every one of the 12 IBD sub-tables (IVR / Entry / Second
//  / Third × First / Repeat / RNA).
//
//  Dependencies: OpenXLSX (already linked via excel_cm)
// ============================================================

// ── Place this block in the anonymous namespace at the top
//    of ccperformevaluate.cpp, alongside any other helpers ──

namespace {

// How each column in a totals row should be referenced.
//
//  TableRef  →  TableXXX[Column Name]          (structured table column)
//  RangeRef  →  ColLetter{start}:ColLetter{end} (plain A1 range)
//  None      →  column is left blank (data is zero / not applicable)
enum class ColRef { None, TableRef, RangeRef };

// Describes the totals row of one IBD sub-table.
struct SubtotalSpec {
    int         totalRow;       // Excel row number of the totals row
    std::string tableName;      // OpenXLSX table object name, e.g. "Table418"
    int         dataStartRow;   // First data row (for RangeRef columns)
    int         threshold;      // Seconds cap; used in col-E label:
                                //   "{threshold} * Call Count (S)"
    // Per-column mode (col G is always TableRef — handled implicitly)
    ColRef colC = ColRef::None;
    ColRef colD = ColRef::None;
    ColRef colE = ColRef::None;
    ColRef colF = ColRef::None;
};

// ─────────────────────────────────────────────────────────────
//  Master config — all 12 IBD sub-tables
//
//  Layout verified against NOVEMBER 2025 IBD Data sheet:
//    • rangeEnd is always totalRow − 1  (no exception across all 12)
//    • Column G always uses a table-column reference
//    • Table names (Table418 … Table1629) are fixed in the template
// ─────────────────────────────────────────────────────────────
const std::array<SubtotalSpec, 12> kIbdSubtotals = {{

    //  ── IVR segment  (threshold = 120 s) ──────────────────────────────
    //
    //  D / E / F use plain range refs in the IVR tables because those
    //  column names in the named table object don't include the threshold
    //  label (the template designer chose range refs here).
    //  C and G use structured table-column refs throughout.

    { 37,  "Table418",  6,   120,
        ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },

    { 72,  "Table519",  41,  120,
        ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },

    { 107, "Table620",  76,  120,
        ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef },

    //  ── Entry segment  (threshold = 130 s) ───────────────────────────
    //
    //  Entry First: C and D use table-column refs (the template table
    //  exposes "Total Duration (S)" as a named column); E and F use
    //  range refs.
    //  Entry Repeat: only C, D, G — E and F are zero (not applicable).
    //  Entry RNA:    only G — the single-row aggregate table.

    { 145, "Table721",  114, 130,
        ColRef::TableRef, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef },

    { 180, "Table822",  149, 130,
        ColRef::TableRef, ColRef::TableRef, ColRef::None,     ColRef::None     },

    { 185, "Table1125", 183, 130,
        ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    //  ── Second segment  (threshold = 150 s) ──────────────────────────
    //
    //  All three Second tables write only C and G (or just G for RNA).
    //  The remaining columns are zero in the current billing model and
    //  the template leaves them blank.

    { 223, "Table923",  193, 150,
        ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None     },

    { 258, "Table1024", 228, 150,
        ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None     },

    { 263, "Table1326", 262, 150,
        ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    //  ── Third segment  (threshold = 180 s) ───────────────────────────
    //
    //  Third First and Third Repeat: only G (single-row tables, always
    //  zero in the current billing model).
    //  Third RNA:  all five columns via table-column refs — this is the
    //  only sub-table where every column is a structured table reference.

    { 271, "Table1427", 270, 180,
        ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    { 276, "Table1528", 275, 180,
        ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None     },

    { 311, "Table1629", 280, 180,
        ColRef::TableRef, ColRef::TableRef, ColRef::TableRef, ColRef::TableRef },
}};

// ─────────────────────────────────────────────────────────────
//  writeSubtotalRow()
//
//  Writes the five SUBTOTAL(109,...) formulas for one sub-table
//  into the given OpenXLSX worksheet.
//
//  Formula anatomy:
//    =SUBTOTAL(109, <ref>)
//    function 109 = SUM ignoring hidden/filtered rows
//
//  ref is either:
//    TableName[Column Name]           ← structured table-column ref
//    ColLetter{dataStart}:ColLetter{totalRow-1}  ← A1-style range ref
// ─────────────────────────────────────────────────────────────
void writeSubtotalRow(OpenXLSX::XLWorksheet& ws, const SubtotalSpec& spec)
{
    const int   row      = spec.totalRow;
    const int   rangeEnd = spec.totalRow - 1;   // always holds for this template
    const auto& tbl      = spec.tableName;

    // Build =SUBTOTAL(109,<ref>)
    auto formula = [](std::string_view ref) {
        return std::string("=SUBTOTAL(109,") + std::string(ref) + ")";
    };

    // Structured reference:  TableXXX[Column Name]
    auto tRef = [&](std::string_view colName) {
        return tbl + "[" + std::string(colName) + "]";
    };

    // A1-style range:  e.g.  D6:D36
    auto rRef = [&](std::string_view colLetter) {
        return std::string(colLetter) + std::to_string(spec.dataStartRow)
             + ":" + std::string(colLetter) + std::to_string(rangeEnd);
    };

    // Write one cell: skip if mode is None
    auto writeCell = [&](std::string_view colLetter,
                         ColRef           mode,
                         std::string_view colName)
    {
        if (mode == ColRef::None) return;

        std::string ref = (mode == ColRef::TableRef)
                        ? tRef(colName)
                        : rRef(colLetter);

        ws.cell(std::string(colLetter) + std::to_string(row))
          .setFormula(formula(ref));
    };

    // ── Columns C–F ──────────────────────────────────────────
    writeCell("C", spec.colC, "Call Count");
    writeCell("D", spec.colD, "Total Duration (S)");

    // Column E label is threshold-specific:  "120 * Call Count (S)" etc.
    const std::string eColName =
        std::to_string(spec.threshold) + " * Call Count (S)";
    writeCell("E", spec.colE, eColName);

    writeCell("F", spec.colF, "Unit (S)");

    // ── Column G — always a table-column ref ─────────────────
    ws.cell("G" + std::to_string(row))
      .setFormula(formula(tRef("Unit (Minutes)")));
}

} // anonymous namespace


// ─────────────────────────────────────────────────────────────
//  Integration point inside CCPerformEvaluate::writeIBDData()
//
//  After the day-wise data rows have been written for all 12
//  sub-tables, add this single loop to stamp the totals row
//  formulas into the worksheet:
//
//      // ... existing writeRows(...) calls for all 12 segments ...
//
//      // Write SUBTOTAL formulas to all 12 sub-table totals rows
//      for (const auto& spec : kIbdSubtotals) {
//          writeSubtotalRow(ws, spec);
//      }
//
//  'ws' is the OpenXLSX::XLWorksheet& obtained from ExcelCM for
//  the "<MONTHNAME> <YEAR> IBD Data" sheet.
// ─────────────────────────────────────────────────────────────
