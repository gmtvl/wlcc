// ============================================================================
// subtotal_post.cpp
// Implementation of writeAllIbdSubtotals() and postProcessXlsx().
// Helpers were lifted verbatim from the standalone wsubtotal.cpp test program;
// the only design change is that the sheet name is now a parameter instead
// of being baked into kIbdSpecs (production sheet name varies by month/year).
// ============================================================================

#include "subtotal_post.h"

#include <minizip/unzip.h>
#include <minizip/zip.h>

#include <array>
#include <cstdio>
#include <cstring>
#include <iostream>
#include <string>
#include <vector>

namespace wlcc {
namespace {

// ----------------------------------------------------------------------------
//                         IBD Data subtotal specs
// ----------------------------------------------------------------------------

enum class ColRef {
    None,       // leave cell blank
    TableRef,   // structured reference: TableXXX[Column Name]
    RangeRef,   // A1 range: ColLetter{start}:ColLetter{end}
};

struct SubtotalSpec {
    int         totalRow;       // row that holds the SUBTOTAL formulas
    int         dataStartRow;   // first data row (for RangeRef columns)
    const char* tableName;      // structured-table object name
    int         threshold;      // for "{threshold} * Call Count (S)" col E label
    ColRef      colC;
    ColRef      colD;
    ColRef      colE;
    ColRef      colF;
    ColRef      colG;
};

// 12 IBD tables. Specs derived from xl/tables/tableN.xml in the template
// (which columns are totalsRowFunction="custom" / "sum" / absent).
constexpr std::array<SubtotalSpec, 12> kIbdSpecs = {{
    // IVR (120 s)
    {  37,   6, "Table418",  120,
       ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::TableRef },
    {  72,  41, "Table519",  120,
       ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::TableRef },
    { 107,  76, "Table620",  120,
       ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::TableRef },
    // Entry (130 s)
    { 145, 114, "Table721",  130,
       ColRef::TableRef, ColRef::TableRef, ColRef::RangeRef, ColRef::RangeRef, ColRef::TableRef },
    { 180, 149, "Table822",  130,
       ColRef::TableRef, ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::TableRef },
    { 185, 183, "Table1125", 130,
       ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    // Second (150 s)
    { 223, 193, "Table923",  150,
       ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    { 258, 228, "Table1024", 150,
       ColRef::TableRef, ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    { 263, 262, "Table1326", 150,
       ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    // Third (180 s)
    { 271, 270, "Table1427", 180,
       ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    { 276, 275, "Table1528", 180,
       ColRef::None,     ColRef::None,     ColRef::None,     ColRef::None,     ColRef::TableRef },
    { 311, 280, "Table1629", 180,
       ColRef::TableRef, ColRef::TableRef, ColRef::TableRef, ColRef::TableRef, ColRef::TableRef },
}};

// ----------------------------------------------------------------------------
//                              Formula helpers
// ----------------------------------------------------------------------------

std::string colName(char col, int threshold) {
    switch (col) {
        case 'C': return "Call Count";
        case 'D': return "Total Duration (S)";
        case 'E': return std::to_string(threshold) + " * Call Count (S)";
        case 'F': return "Unit (S)";
        case 'G': return "Unit (Minutes)";
    }
    return "";
}

std::string buildFormula(char col, ColRef mode, const SubtotalSpec& s) {
    if (mode == ColRef::None) return "";
    std::string ref;
    if (mode == ColRef::TableRef) {
        ref = std::string(s.tableName) + "[" + colName(col, s.threshold) + "]";
    } else { // RangeRef
        const int rangeEnd = s.totalRow - 1;
        ref = std::string(1, col) + std::to_string(s.dataStartRow)
            + ":" + std::string(1, col) + std::to_string(rangeEnd);
    }
    return "SUBTOTAL(109," + ref + ")";
}

void writeOneSpec(ExcelCM& excel,
                  const std::string& sheet,
                  const SubtotalSpec& s) {
    auto write = [&](char col, ColRef mode) {
        if (mode == ColRef::None) return;
        const std::string ref = std::string(1, col) + std::to_string(s.totalRow);
        const std::string fml = buildFormula(col, mode, s);
        excel.writeFormulaCell(sheet, ref, fml);
    };
    write('C', s.colC);
    write('D', s.colD);
    write('E', s.colE);
    write('F', s.colF);
    write('G', s.colG);
}

// ----------------------------------------------------------------------------
//                       XML post-processing helpers
// ----------------------------------------------------------------------------

// Remove ref="..." from <f> tags that lack t="shared" / t="array".
// OpenXLSX strips t but keeps ref when overwriting a shared-formula cell.
std::string fixOrphanedRefs(const std::string& xml) {
    std::string out;
    out.reserve(xml.size());
    size_t pos = 0;
    while (pos < xml.size()) {
        size_t fStart = xml.find("<f ", pos);
        if (fStart == std::string::npos) {
            out += xml.substr(pos);
            break;
        }
        out += xml.substr(pos, fStart - pos);
        size_t fEnd = xml.find('>', fStart);
        if (fEnd == std::string::npos) {
            out += xml.substr(fStart);
            break;
        }
        std::string tag = xml.substr(fStart, fEnd - fStart + 1);
        bool hasType = tag.find("t=\"shared\"") != std::string::npos
                    || tag.find("t=\"array\"")  != std::string::npos;
        if (!hasType) {
            size_t r = tag.find(" ref=\"");
            while (r != std::string::npos) {
                size_t rEnd = tag.find('"', r + 6);
                if (rEnd == std::string::npos) break;
                tag.erase(r, rEnd - r + 1);
                r = tag.find(" ref=\"", r);
            }
        }
        out += tag;
        pos = fEnd + 1;
    }
    return out;
}

// Strip <Relationship .../calcChain.../> from xl/_rels/workbook.xml.rels.
std::string stripCalcChainRel(const std::string& xml) {
    std::string out = xml;
    size_t mPos = out.find("calcChain");
    while (mPos != std::string::npos) {
        size_t start = out.rfind("<Relationship", mPos);
        if (start == std::string::npos) break;
        size_t end = out.find("/>", start);
        if (end == std::string::npos) break;
        end += 2;
        if (mPos < end) {
            out.erase(start, end - start);
            mPos = out.find("calcChain", start);
        } else {
            mPos = out.find("calcChain", end);
        }
    }
    return out;
}

// Inject fullCalcOnLoad="1" into the <calcPr/> element of xl/workbook.xml.
// Without this, Excel uses the cached <v>0</v> we serialise and never
// recomputes the SUBTOTAL formulas (since we also dropped the calcChain
// reference). With fullCalcOnLoad set, Excel recalculates the whole workbook
// on first open, which is exactly what we want when the calc chain has been
// invalidated.
std::string forceFullCalcOnLoad(const std::string& xml) {
    const size_t pos = xml.find("<calcPr");
    if (pos == std::string::npos) return xml;            // no calcPr — leave alone
    const size_t end = xml.find('>', pos);
    if (end == std::string::npos) return xml;
    std::string tag = xml.substr(pos, end - pos + 1);    // e.g. <calcPr calcId="191028"/>
    if (tag.find("fullCalcOnLoad=") != std::string::npos) return xml; // already set
    const bool selfClose = (tag.size() >= 2 && tag[tag.size() - 2] == '/');
    const size_t insAt = selfClose ? tag.size() - 2 : tag.size() - 1;
    tag.insert(insAt, " fullCalcOnLoad=\"1\"");
    std::string out = xml;
    out.replace(pos, end - pos + 1, tag);
    return out;
}

// Strip <Override PartName="/xl/calcChain.xml" .../> from [Content_Types].xml.
std::string stripCalcChainOverride(const std::string& xml) {
    std::string out = xml;
    size_t nPos = out.find("calcChain.xml");
    while (nPos != std::string::npos) {
        size_t start = out.rfind("<Override", nPos);
        if (start == std::string::npos) break;
        size_t end = out.find("/>", start);
        if (end == std::string::npos) break;
        end += 2;
        if (nPos < end) {
            out.erase(start, end - start);
            nPos = out.find("calcChain.xml", start);
        } else {
            nPos = out.find("calcChain.xml", end);
        }
    }
    return out;
}

bool isWorksheetXml(const std::string& name) {
    if (name.rfind("xl/worksheets/", 0) != 0) return false;
    if (name.find("_rels") != std::string::npos) return false;
    if (name.size() < 4) return false;
    return name.compare(name.size() - 4, 4, ".xml") == 0;
}

} // anonymous namespace

// ----------------------------------------------------------------------------
//                           Public API entry points
// ----------------------------------------------------------------------------

void writeAllIbdSubtotals(ExcelCM& excel, const std::string& sheet) {
    for (const auto& s : kIbdSpecs) {
        writeOneSpec(excel, sheet, s);
    }
}

bool postProcessXlsx(const std::string& xlsxPath) {
    const std::string tmp = xlsxPath + ".tmp";

    unzFile uf = unzOpen(xlsxPath.c_str());
    if (!uf) {
        std::cerr << "[postProcessXlsx] cannot open source\n";
        return false;
    }
    zipFile zf = zipOpen(tmp.c_str(), APPEND_STATUS_CREATE);
    if (!zf) {
        std::cerr << "[postProcessXlsx] cannot create temp\n";
        unzClose(uf);
        return false;
    }

    int ret = unzGoToFirstFile(uf);
    while (ret == UNZ_OK) {
        char nameBuf[1024];
        unz_file_info fi;
        unzGetCurrentFileInfo(uf, &fi, nameBuf, sizeof(nameBuf), nullptr, 0, nullptr, 0);
        std::string name(nameBuf);

        // Drop calcChain entirely
        if (name == "xl/calcChain.xml") {
            ret = unzGoToNextFile(uf);
            continue;
        }

        // Read entry — chunked, NEVER trust uncompressed_size
        if (unzOpenCurrentFile(uf) != UNZ_OK) {
            std::cerr << "[postProcessXlsx] cannot open entry " << name << "\n";
            break;
        }
        std::vector<char> buf;
        buf.reserve(fi.uncompressed_size > 0 ? fi.uncompressed_size : 65536);
        std::vector<char> chunk(65536);
        int n;
        while ((n = unzReadCurrentFile(uf, chunk.data(), 65536)) > 0) {
            buf.insert(buf.end(), chunk.begin(), chunk.begin() + n);
        }
        unzCloseCurrentFile(uf);
        if (n < 0) {
            std::cerr << "[postProcessXlsx] read error on " << name << "\n";
            break;
        }

        std::string text(buf.data(), buf.size());
        if (name == "xl/_rels/workbook.xml.rels") {
            text = stripCalcChainRel(text);
        } else if (name == "[Content_Types].xml") {
            text = stripCalcChainOverride(text);
        } else if (name == "xl/workbook.xml") {
            text = forceFullCalcOnLoad(text);
        } else if (isWorksheetXml(name)) {
            text = fixOrphanedRefs(text);
        }

        zip_fileinfo zfi{};
        zipOpenNewFileInZip(zf, name.c_str(), &zfi,
                            nullptr, 0, nullptr, 0, nullptr,
                            Z_DEFLATED, Z_DEFAULT_COMPRESSION);
        zipWriteInFileInZip(zf, text.data(), static_cast<unsigned>(text.size()));
        zipCloseFileInZip(zf);

        ret = unzGoToNextFile(uf);
    }

    unzClose(uf);
    zipClose(zf, nullptr);

    if (std::remove(xlsxPath.c_str()) != 0) {
        std::cerr << "[postProcessXlsx] cannot remove original\n";
        return false;
    }
    if (std::rename(tmp.c_str(), xlsxPath.c_str()) != 0) {
        std::cerr << "[postProcessXlsx] cannot rename tmp\n";
        return false;
    }
    return true;
}

} // namespace wlcc
