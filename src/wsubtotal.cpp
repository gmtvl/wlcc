// wsubtotal.cpp
// ----------------------------------------------------------------------------
// Test driver for subtotal_post (writeAllIbdSubtotals + postProcessXlsx).
//
// Usage:  wsubtotal <input.xlsx> <output.xlsx>
//
// All the real logic now lives in src/subtotal_post.cpp; this file just
// drives it for ad-hoc testing on a "Centre wise Calc Sheet".
// ----------------------------------------------------------------------------

#include "subtotal_post.h"

#include <iostream>
#include <string>

int main(int argc, char* argv[]) {
    if (argc < 3) {
        std::cout << "Usage: wsubtotal <input.xlsx> <output.xlsx>\n";
        return 1;
    }
    const std::string inputPath  = argv[1];
    const std::string outputPath = argv[2];
    const std::string ibdSheet   = "NOVEMBER 2025 IBD Data";

    wlcc::ExcelCM excel;
    excel.open(inputPath);

    wlcc::writeAllIbdSubtotals(excel, ibdSheet);
    std::cout << "Wrote subtotal formulas to all 12 IBD tables.\n";

    excel.saveAs(outputPath);
    excel.close();
    std::cout << "Saved: " << outputPath << "\n";

    if (!wlcc::postProcessXlsx(outputPath)) {
        std::cerr << "ERROR: postProcessXlsx failed.\n";
        return 1;
    }
    std::cout << "Done.\n";
    return 0;
}
