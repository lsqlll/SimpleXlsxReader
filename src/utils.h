#ifndef UTILS_H
#define UTILS_H

#include "Exceptions.h"

#include <algorithm>
#include <filesystem>
#include <set>

extern "C"
{
#include <xls.h>
}

namespace fs = std::filesystem;

const std::vector<std::string> formats = { "xls", "xlsx", "csv" };
inline bool
isExcelFormat (const std::string& format)
{
    return std::any_of (formats->begin (), formats->end (),
                        [ format ] (const std::string& f)
                            { return format == f; });
}

inline bool
isValide (const fs::path& p)
{
    if (!fs::exists (p))
        {
            throw ExcelReader::FileNotFoundException (p.string ());
        }

    if (fs::is_directory (p))
        {
            throw ExcelReader::PathNotFileException (p.string ());
        }

    if (!fs::is_regular_file (p))
        {
            throw ExcelReader::UnsupportedException ("format " + p.string ());
        }

    if (!isExcelFormat (p.extension ().string ()))
        {
            throw ExcelReader::UnsupportedException (
                "format " + p.extension ().string ());
        }
    return true;
}

inline std::string
trim (const std::string& str)
{
    auto start = str.find_first_not_of (" \t");
    auto end = str.find_last_not_of (" \t");
    return str.substr (start, end - start + 1);
}

inline bool
isEmpty (const std::string& raw_value, bool trims = false)
{
    if (raw_value.empty ())
        {
            return true;
        }
    if (trims)
        {
            auto raw_string = trim (raw_value);
            // 找到第一个非空字符的位置
            auto start =
                std::find_if_not (raw_string.begin (), raw_string.end (),
                                  [] (unsigned char c)
                                      { return std::isspace (c); });
            // 找到最后一个非空字符的位置
            auto end =
                std::find_if_not (raw_string.rbegin (), raw_string.rend (),
                                  [] (unsigned char c)
                                      { return std::isspace (c); })
                    .base ();

            // 如果 start >= end，则表示去除空白后字符串为空
            return start >= end;
        }

    return std::all_of (raw_value.begin (), raw_value.end (),
                        [] (unsigned char c) { return std::isspace (c); });
}

inline std::string
tolower (const std::string& raw_value)
{
    std::string tmp{};
    std::transform (raw_value.begin (), raw_value.end (), tmp.begin (),
                    [] (char c) { return std::tolower (c); });
    return tmp;
}

inline bool
isDateTime (int id)
{
    // Page and section numbers below refer to
    // ECMA-376 (version, date, and download URL given in XlsxCell.h)
    //
    // Example from L.2.7.4.4 p4698 for hypothetical cell D2
    // Cell D2 contains the text "Q1" and is defined in the cell table of sheet1
    // as:
    //
    // <c r="D2" s="7" t="s">
    //     <v>0</v>
    // </c>
    //
    // On this cell, the attribute value s="7" indicates that the 7th
    // (zero-based) <xf> definition of <cellXfs> holds the formatting
    // information for the cell. The 7th <xf> of <cellXfs> is defined as:
    //
    // <xf numFmtId="0" fontId="4" fillId="2" borderId="2" xfId="1"
    // applyBorder="1"/>
    //
    // The number formatting information cannot be found in a <numFmt>
    // definition because it is a built-in format; instead, it is implicitly
    // understood to be the 0th built-in number format.
    //
    // This function stores knowledge about these built-in number formats.
    //
    // 18.8.30 numFmt (Number Format) p1786
    // Date times: 14-22, 27-36, 45-47, 50-58, 71-81 (inclusive)
    if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36)
        || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)
        || (id >= 71 && id <= 81))
        return true;

    // Built-in format that's not a date
    if (id < 164)
        return false;
    std::set<int> dateFormats = { 14, 15, 16, 17, 18, 19, 20, 21, 22,
                                  27, 28, 29, 30, 31, 32, 33, 34, 35,
                                  36, 50, 51, 52, 53, 54, 55, 56, 57,
                                  58, 59, 60, 61, 62, 67, 68, 69 };

    return dateFormats.count (id) > 0;
}

inline std::pair<std::size_t, std::size_t>
parseAddress (const std::string& addr)
{
    if (addr.empty ())
        {
            throw std::invalid_argument ("address can't be empty");
        }
    if (!std::isalpha (addr[ 0 ]))
        throw std::invalid_argument (
            "The first character of the address is not alpha");

    std::size_t row{ 0 }, col{ 0 }, idx{ 0 };

    auto it = addr.begin ();
    while (it != addr.end () && std::isalpha (*it))
        {
            col = col * 26 + (*it - 'A' + 1);
            ++it, ++idx;
        }

    std::string num = addr.substr (idx, addr.size ());
    try
        {
            row = std::stoi (num);
        }
    catch (...)
        {
            throw ExcelReader::parseAddrException (addr);
        }

    return std::make_pair (row, col);
}

#endif