#pragma once

// ============================================================================
// subtotal_post.h
// ----------------------------------------------------------------------------
// Reusable SUBTOTAL formula writer + xlsx post-processor.
//
// Two functions:
//
//   writeAllIbdSubtotals(excel, sheet)
//     Writes SUBTOTAL(109,...) formula cells into the totals row of every
//     named table on the IBD Data sheet (Table418..Table1629). Uses
//     ExcelCM::writeFormulaCell to produce real <f> XML elements (not
//     string cells). Per-table per-column ColRef specs mirror the template's
//     totalsRowFunction values (TableRef / RangeRef / None) to avoid the
//     "We found a problem" repair dialog.
//
//   postProcessXlsx(xlsxPath)
//     ZIP-rebuilds a saved xlsx via minizip to defuse the four corruption
//     traps that OpenXLSX leaves behind when it touches formula cells:
//       1. Drops xl/calcChain.xml entirely (the part itself).
//       2. Strips <Relationship .../calcChain.../> from
//          xl/_rels/workbook.xml.rels.
//       3. Strips <Override PartName="/xl/calcChain.xml" .../> from
//          [Content_Types].xml.
//       4. Adds fullCalcOnLoad="1" to <calcPr/> in xl/workbook.xml so Excel
//          recomputes on first open instead of trusting the cached <v>0</v>
//          values OpenXLSX writes for un-evaluated formulas.
//       5. Strips orphan ref="..." from <f> tags in worksheet XML left
//          behind when OpenXLSX overwrites a t="shared" formula cell.
//     Reads each ZIP entry chunked (uncompressed_size==0 entries with data
//     descriptors would otherwise truncate to 1 byte).
//     Returns true on success.
// ============================================================================

#include "excel_cm.h"

#include <string>

namespace wlcc {

void writeAllIbdSubtotals(ExcelCM& excel, const std::string& sheet);

bool postProcessXlsx(const std::string& xlsxPath);

} // namespace wlcc
