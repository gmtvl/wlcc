/**
 * pattern_match.cpp
 *
 * Standalone utility to demonstrate pattern encoding and column inference.
 *
 * Three functions:
 *   1. encodeString    — encode a raw string into '+'/'-' pattern
 *   2. countWords      — count words in a '+'/'-' encoded string
 *   3. inferColumns    — compare header and dataline patterns to find column boundaries
 *
 * Build:
 *   cl /EHsc /std:c++17 src/pattern_match.cpp /Fe:pattern_match.exe
 *   -- or --
 *   g++ -std=c++17 -o pattern_match src/pattern_match.cpp
 */

#include <iostream>
#include <string>
#include <vector>

// ---------------------------------------------------------------------------
// 1. encodeString
//    Every non-space, non-tab character → '+'
//    Space → '-'
//    Tab   → '---' (three dashes)
// ---------------------------------------------------------------------------
std::string encodeString(const std::string& input) {
    std::string result;
    result.reserve(input.size() + input.size() / 2);

    for (char c : input) {
        if (c == '\t') {
            result += "---";
        } else if (c == ' ') {
            result += '-';
        } else {
            result += '+';
        }
    }

    return result;
}

// ---------------------------------------------------------------------------
// 2. countWords
//    A "word" = contiguous block of '+' characters.
//    Words are separated by '-' characters (from spaces or tabs).
// ---------------------------------------------------------------------------
int countWords(const std::string& input) {
    int count = 0;
    bool inWord = false;

    for (char c : input) {
        if (c == '+') {
            if (!inWord) {
                ++count;
                inWord = true;
            }
        } else {
            inWord = false;
        }
    }

    return count;
}

// ---------------------------------------------------------------------------
// 3. inferColumns
//    Compare header and dataline patterns to determine real column boundaries.
//    Returns the header string with '|' inserted at each column boundary.
//
//    Parameters:
//      header          — encoded header string ('+' and '-')
//      headerWordCount — number of '+' blocks in header
//      dataline        — encoded data string   ('+' and '-')
//      dataWordCount   — number of '+' blocks in dataline
// ---------------------------------------------------------------------------
std::string inferColumns(const std::string& header, int headerWordCount,
                         const std::string& dataline, int dataWordCount) {

    // A "word" = contiguous block of '+' characters
    struct Word { size_t start; size_t end; };  // [start, end)

    auto parseWords = [](const std::string& s) -> std::vector<Word> {
        std::vector<Word> words;
        size_t i = 0;
        while (i < s.size()) {
            if (s[i] == '+') {
                size_t start = i;
                while (i < s.size() && s[i] == '+') i++;
                words.push_back({start, i});
            } else {
                i++;
            }
        }
        return words;
    };

    auto hWords = parseWords(header);
    auto dWords = parseWords(dataline);

    // --- Case 1: word counts match → every header word is its own column ---
    if (headerWordCount == dataWordCount) {
        std::string result;
        size_t pos = 0;
        for (size_t i = 0; i < hWords.size(); i++) {
            // Copy up to end of this word
            while (pos < hWords[i].end)
                result += header[pos++];
            // Copy trailing gap
            size_t nextStart = (i + 1 < hWords.size())
                                 ? hWords[i + 1].start : header.size();
            while (pos < nextStart)
                result += header[pos++];
            // Insert column delimiter
            result += '|';
        }
        return result;
    }

    // --- Case 2: counts differ → infer boundaries from dataline pattern ---
    std::vector<bool> isBoundary;
    for (size_t i = 0; i + 1 < hWords.size(); i++) {
        size_t gapStart = hWords[i].end;
        size_t gapEnd   = hWords[i + 1].start;
        size_t gapLen   = gapEnd - gapStart;

        // Sample the midpoint of the gap in the dataline
        size_t mid = (gapStart + gapEnd) / 2;

        if (mid >= dataline.size()) {
            isBoundary.push_back(true);
        } else if (dataline[mid] == '-') {
            isBoundary.push_back(true);
        } else if (gapLen >= 3) {
            // Wide gap — check if the entire gap region is '+' in dataline
            bool allPlus = true;
            for (size_t p = gapStart; p < gapEnd && p < dataline.size(); p++) {
                if (dataline[p] == '-') { allPlus = false; break; }
            }
            isBoundary.push_back(!allPlus);
        } else {
            // Short gap, dataline has '+' → intra-column space
            isBoundary.push_back(false);
        }
    }

    // Build result with '|' at column boundaries
    std::string result;
    size_t pos = 0;

    for (size_t i = 0; i < hWords.size(); i++) {
        // Copy up to end of this word
        while (pos < hWords[i].end)
            result += header[pos++];

        bool boundary = (i < isBoundary.size()) ? isBoundary[i] : true;

        // Copy trailing gap
        size_t nextStart = (i + 1 < hWords.size())
                             ? hWords[i + 1].start : header.size();
        while (pos < nextStart)
            result += header[pos++];

        if (boundary)
            result += '|';
    }

    // Any trailing characters
    while (pos < header.size())
        result += header[pos++];

    return result;
}

// ---------------------------------------------------------------------------
// Main — demonstrate with CDR-style examples
// ---------------------------------------------------------------------------
int main() {
    std::cout << "=== Pattern Match Demo ===\n\n";

    // Example: CDR header and data line
    std::string rawHeader  = "Call Type\tCalling Number\t\tCalled Number\t\t    Date & Time\t\t      Duration";
    std::string rawData    = "50\t\t919726075989    \t911503          \t01-11-2025 05:28:23\t\t119";

    // Step 1: Encode
    std::string encHeader  = encodeString(rawHeader);
    std::string encData    = encodeString(rawData);

    std::cout << "Raw header : " << rawHeader << "\n";
    std::cout << "Encoded    : " << encHeader << "\n\n";

    std::cout << "Raw data   : " << rawData << "\n";
    std::cout << "Encoded    : " << encData << "\n\n";

    // Step 2: Count words
    int hCount = countWords(encHeader);
    int dCount = countWords(encData);

    std::cout << "Header word count  : " << hCount << "\n";
    std::cout << "Data word count    : " << dCount << "\n\n";

    // Step 3: Infer columns
    std::string inferred = inferColumns(encHeader, hCount, encData, dCount);

    std::cout << "Inferred columns   : " << inferred << "\n";
    std::cout << "                     (|  marks column boundaries)\n\n";

    // Count inferred columns
    int colCount = 0;
    for (char c : inferred) {
        if (c == '|') colCount++;
    }
    std::cout << "Number of columns  : " << colCount << "\n\n";

    // --- Simple test cases ---
    std::cout << "=== Additional Tests ===\n\n";

    // Test: counts match — every word is a column
    {
        std::string h = encodeString("Name\tAge\tCity");
        std::string d = encodeString("John\t25\tDelhi");
        int hc = countWords(h);
        int dc = countWords(d);
        std::string result = inferColumns(h, hc, d, dc);
        std::cout << "Test (counts match):\n";
        std::cout << "  Header : " << h << " (" << hc << " words)\n";
        std::cout << "  Data   : " << d << " (" << dc << " words)\n";
        std::cout << "  Result : " << result << "\n\n";
    }

    // Test: multi-word column name
    {
        std::string h = encodeString("First Name\tLast Name\tAge");
        std::string d = encodeString("Ravi Kumar\tNair\t30");
        int hc = countWords(h);
        int dc = countWords(d);
        std::string result = inferColumns(h, hc, d, dc);
        std::cout << "Test (multi-word data spans header gap):\n";
        std::cout << "  Header : " << h << " (" << hc << " words)\n";
        std::cout << "  Data   : " << d << " (" << dc << " words)\n";
        std::cout << "  Result : " << result << "\n";

        int cols = 0;
        for (char c : result) if (c == '|') cols++;
        std::cout << "  Columns: " << cols << "\n\n";
    }

    return 0;
}
