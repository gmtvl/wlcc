// ================================================================
//  src/test_ibd_subtotals.cpp
//  Standalone test for ibd_subtotals.h
//
//  Usage:
//    test_ibd_subtotals.exe <path\to\calc_sheet.xlsx> [sheet-name]
//    sheet-name defaults to "NOVEMBER 2025 IBD Data"
// ================================================================

#include "ibd_subtotals.h"   // brings in excel_cm.h → wlcc::ExcelCM

#include <cstdio>            // _popen / _pclose / FILE
#include <cstdlib>
#include <filesystem>
#include <fstream>           // std::ofstream
#include <iostream>
#include <string>
#include <vector>

namespace fs = std::filesystem;

// Pull ExcelCM and ibd helpers into scope without full qualification
using namespace wlcc;
using namespace wlcc::ibd;

// ----------------------------------------------------------------
//  Ground-truth formula strings (openpyxl returns WITH leading '=')
// ----------------------------------------------------------------
struct ExpectedCell { std::string address; std::string formula; };

static const std::vector<ExpectedCell> kExpected = {
    // IVR First (37)
    { "C37",  "=SUBTOTAL(109,Table418[Call Count])"            },
    { "D37",  "=SUBTOTAL(109,D6:D36)"                          },
    { "E37",  "=SUBTOTAL(109,E6:E36)"                          },
    { "F37",  "=SUBTOTAL(109,F6:F36)"                          },
    { "G37",  "=SUBTOTAL(109,Table418[Unit (Minutes)])"         },
    // IVR Repeat (72)
    { "C72",  "=SUBTOTAL(109,Table519[Call Count])"            },
    { "D72",  "=SUBTOTAL(109,D41:D71)"                         },
    { "E72",  "=SUBTOTAL(109,E41:E71)"                         },
    { "F72",  "=SUBTOTAL(109,F41:F71)"                         },
    { "G72",  "=SUBTOTAL(109,Table519[Unit (Minutes)])"         },
    // IVR RNA (107)
    { "C107", "=SUBTOTAL(109,Table620[Call Count])"            },
    { "D107", "=SUBTOTAL(109,D76:D106)"                        },
    { "E107", "=SUBTOTAL(109,E76:E106)"                        },
    { "F107", "=SUBTOTAL(109,F76:F106)"                        },
    { "G107", "=SUBTOTAL(109,Table620[Unit (Minutes)])"         },
    // Entry First (145)
    { "C145", "=SUBTOTAL(109,Table721[Call Count])"            },
    { "D145", "=SUBTOTAL(109,Table721[Total Duration (S)])"    },
    { "E145", "=SUBTOTAL(109,E114:E144)"                       },
    { "F145", "=SUBTOTAL(109,F114:F144)"                       },
    { "G145", "=SUBTOTAL(109,Table721[Unit (Minutes)])"         },
    // Entry Repeat (180)
    { "C180", "=SUBTOTAL(109,Table822[Call Count])"            },
    { "D180", "=SUBTOTAL(109,Table822[Total Duration (S)])"    },
    { "G180", "=SUBTOTAL(109,Table822[Unit (Minutes)])"         },
    // Entry RNA (185) — G only
    { "G185", "=SUBTOTAL(109,Table1125[Unit (Minutes)])"        },
    // Second First (223)
    { "C223", "=SUBTOTAL(109,Table923[Call Count])"            },
    { "G223", "=SUBTOTAL(109,Table923[Unit (Minutes)])"         },
    // Second Repeat (258)
    { "C258", "=SUBTOTAL(109,Table1024[Call Count])"           },
    { "G258", "=SUBTOTAL(109,Table1024[Unit (Minutes)])"        },
    // Second RNA (263) — G only
    { "G263", "=SUBTOTAL(109,Table1326[Unit (Minutes)])"        },
    // Third First (271) — G only
    { "G271", "=SUBTOTAL(109,Table1427[Unit (Minutes)])"        },
    // Third Repeat (276) — G only
    { "G276", "=SUBTOTAL(109,Table1528[Unit (Minutes)])"        },
    // Third RNA (311) — all columns
    { "C311", "=SUBTOTAL(109,Table1629[Call Count])"           },
    { "D311", "=SUBTOTAL(109,Table1629[Total Duration (S)])"   },
    { "E311", "=SUBTOTAL(109,Table1629[180 * Call Count (S)])" },
    { "F311", "=SUBTOTAL(109,Table1629[Unit (S)])"             },
    { "G311", "=SUBTOTAL(109,Table1629[Unit (Minutes)])"        },
};

// ----------------------------------------------------------------
//  Make a working copy — original is never mutated
// ----------------------------------------------------------------
static fs::path makeTempCopy(const fs::path& src)
{
    fs::path dst = src.parent_path()
                 / (src.stem().string() + "_subtotals_test"
                    + src.extension().string());
    fs::copy_file(src, dst, fs::copy_options::overwrite_existing);
    return dst;
}

// ----------------------------------------------------------------
//  Verify via Python/openpyxl
//  Writes an inline Python script, runs it, reads stdout.
//  Returns number of failures, or -1 if Python unavailable.
// ----------------------------------------------------------------
static int verifyWithPython(const fs::path&    xlsxPath,
                             const std::string& sheet)
{
    const fs::path pyScript =
        xlsxPath.parent_path() / "verify_subtotals_tmp.py";

    // Write the Python script
    {
        std::ofstream py(pyScript);
        if (!py) {
            std::cerr << "ERROR: cannot write temp Python script to "
                      << pyScript << "\n";
            return -1;
        }
        py << "import sys\n"
           << "from openpyxl import load_workbook\n"
           << "wb = load_workbook(r'" << xlsxPath.string() << "')\n"
           << "ws = wb['" << sheet << "']\n";

        for (const auto& e : kExpected) {
            py << "v = ws['" << e.address << "'].value\n"
               << "got = str(v) if v else ''\n"
               << "want = r'" << e.formula << "'\n"
               << "print(('PASS ' if got==want else 'FAIL ') + '"
               << e.address << "|' + want + '|' + got)\n";
        }
        py << "wb.close()\n";
    }

    // Run it
    const std::string cmd =
        "python \"" + pyScript.string() + "\" 2>&1";

    FILE* pipe = _popen(cmd.c_str(), "r");
    if (!pipe) {
        std::cerr << "ERROR: could not launch Python.\n"
                  << "Ensure 'python' is on PATH and openpyxl is installed.\n";
        fs::remove(pyScript);
        return -1;
    }

    int passed = 0, failed = 0;
    char buf[4096];
    while (fgets(buf, sizeof(buf), pipe)) {
        std::string line(buf);
        while (!line.empty() &&
               (line.back() == '\n' || line.back() == '\r'))
            line.pop_back();

        if (line.rfind("PASS ", 0) == 0) {
            // "PASS C37|=SUBTOTAL(...)|=SUBTOTAL(...)"
            const std::string addr = line.substr(5, line.find('|') - 5);
            std::cout << "  PASS  " << addr << "\n";
            ++passed;
        } else if (line.rfind("FAIL ", 0) == 0) {
            const std::string rest = line.substr(5);
            const auto p1 = rest.find('|');
            const auto p2 = rest.find('|', p1 + 1);
            const std::string addr = rest.substr(0, p1);
            const std::string want = rest.substr(p1 + 1, p2 - p1 - 1);
            const std::string got  = rest.substr(p2 + 1);
            std::cout << "  FAIL  " << addr << "\n"
                      << "         want: " << want << "\n"
                      << "         got:  "
                      << (got.empty() ? "<empty/None>" : got) << "\n";
            ++failed;
        } else {
            // Python error or unexpected output
            std::cout << "  [py]  " << line << "\n";
        }
    }
    _pclose(pipe);
    fs::remove(pyScript);

    std::cout << "\n"
              << "─────────────────────────────────────────\n"
              << "  Total cells checked : " << (passed + failed) << "\n"
              << "  Passed              : " << passed             << "\n"
              << "  Failed              : " << failed             << "\n"
              << "─────────────────────────────────────────\n";

    return failed;
}

// ----------------------------------------------------------------
//  main
// ----------------------------------------------------------------
int main(int argc, char* argv[])
{
    if (argc < 2) {
        std::cerr
            << "Usage: test_ibd_subtotals <path\\to\\calc_sheet.xlsx>"
               " [sheet-name]\n"
            << "       sheet-name defaults to \"NOVEMBER 2025 IBD Data\"\n";
        return 1;
    }

    const fs::path    srcPath = argv[1];
    const std::string sheet   = (argc >= 3)
                                  ? argv[2]
                                  : "NOVEMBER 2025 IBD Data";

    if (!fs::exists(srcPath)) {
        std::cerr << "ERROR: File not found: " << srcPath << "\n";
        return 1;
    }

    // ── Step 1: make a working copy ───────────────────────────
    const fs::path workPath = makeTempCopy(srcPath);
    std::cout << "Working copy: " << workPath << "\n\n";

    // ── Step 2: write formulas via ExcelCM ───────────────────
    {
        ExcelCM excel;
        excel.open(workPath.string());

        if (!excel.hasSheet(sheet)) {
            std::cerr << "ERROR: Sheet not found: \"" << sheet << "\"\n";
            excel.close();
            fs::remove(workPath);
            return 1;
        }

        writeAllSubtotalRows(excel, sheet);   // ← function under test

        excel.save();
        excel.close();
    }
    std::cout << "Formulas written and saved.\n\n";

    // ── Step 3: verify with Python/openpyxl ──────────────────
    std::cout << "Verifying with openpyxl...\n\n";
    const int failed = verifyWithPython(workPath, sheet);

    // ── Step 4: clean up ─────────────────────────────────────
    fs::remove(workPath);

    if (failed < 0) {
        std::cerr << "\nVerification skipped — Python unavailable.\n"
                  << "Install Python + openpyxl and ensure 'python'"
                     " is on PATH.\n";
        return 2;
    }

    return (failed == 0) ? 0 : 1;
}
