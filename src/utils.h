#ifndef UTILS_H
#define UTILS_H

#include "Exceptions.h"

#include <algorithm>
#include <cstring>
#include <filesystem>
#include <optional>
#include <set>
#include <vector>

extern "C"
{
#include <xls.h>
}

namespace fs = std::filesystem;
inline std::optional<std::string_view>
getStringView (const char *str)
{
    if (str == nullptr)
    {
        return std::nullopt;
    }
    return std::string_view (str);
}

inline std::optional<std::string_view>
getStringView (const std::string &str)
{
    if (str.empty ())
    {
        return std::nullopt;
    }
    return std::string_view (str);
}

inline bool
startsWith (std::string_view str, std::string_view prefix)
{
    return str.substr (0, prefix.size ()) == prefix;
};

inline bool
isExcelFormat (const std::string &format)
{
    const std::vector<std::string> formats = { "xls", "xlsx", "csv" };
    return std::any_of (formats.begin (), formats.end (),
                        [format] (const std::string &fmt)
                        { return format == fmt; });
}

inline bool
isValide (const fs::path &path)
{
    if (!fs::exists (path))
    {
        throw ExcelReader::FileNotFoundException (path.string ());
    }

    if (fs::is_directory (path))
    {
        throw ExcelReader::PathNotFileException (path.string ());
    }

    if (!fs::is_regular_file (path))
    {
        throw ExcelReader::UnsupportedException ("format " + path.string ());
    }

    if (!isExcelFormat (path.extension ().string ()))
    {
        throw ExcelReader::UnsupportedException (
            "format " + path.extension ().string ());
    }
    return true;
}

inline std::string
trim (const std::string &str)
{
    auto start = str.find_first_not_of (" \t");
    auto end = str.find_last_not_of (" \t");
    return str.substr (start, end - start + 1);
}

inline bool
isEmpty (const std::string &raw_value, bool trims = false)
{
    if (raw_value.empty ())
    {
        return true;
    }
    if (trims)
    {
        auto raw_string = trim (raw_value);
        // 找到第一个非空字符的位置
        auto start = std::find_if_not (raw_string.begin (), raw_string.end (),
                                       [] (unsigned char chr)
                                       { return std::isspace (chr); });
        // 找到最后一个非空字符的位置
        auto end = std::find_if_not (raw_string.rbegin (), raw_string.rend (),
                                     [] (unsigned char chr)
                                     { return std::isspace (chr); })
                       .base ();

        // 如果 start >= end，则表示去除空白后字符串为空
        return start >= end;
    }

    return std::all_of (raw_value.begin (), raw_value.end (),
                        [] (unsigned char chr) { return std::isspace (chr); });
}

inline std::string
tolower (const std::string &raw_value)
{
    std::string tmp = raw_value;
    std::transform (tmp.begin (), tmp.end (), tmp.begin (),
                    [] (char chr) { return std::tolower (chr); });
    return tmp;
}

inline bool
isDateTime (int Did)
{
    // Page and section numbers below refer to
    // ECMA-376 (version, date, and download URL given in XlsxCell.h)
    //
    // Example from L.2.7.4.4 p4698 for hypothetical cell D2
    // Cell D2 contains the text "Q1" and is defined in the cell table of
    // sheet1 as:
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
    if ((Did >= 14 && Did <= 22) || (Did >= 27 && Did <= 36)
        || (Did >= 45 && Did <= 47) || (Did >= 50 && Did <= 58)
        || (Did >= 71 && Did <= 81))
    {
        return true;
    }

    // Built-in format that's not a date
    if (Did < 164)
    {
        return false;
    }
    std::set<int> dateFormats
        = { 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29,
            30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54,
            55, 56, 57, 58, 59, 60, 61, 62, 67, 68, 69 };

    return dateFormats.count (Did) > 0;
}

inline std::pair<std::size_t, std::size_t>
parseAddress (const std::string &addr)
{
    if (addr.empty ())
    {
        throw std::invalid_argument ("address can't be empty");
    }
    if (std::isalpha (addr[0]) == 0)
    {
        throw std::invalid_argument (
            "The first character of the address is not alpha");
    }

    std::size_t row{ 0 };
    std::size_t col{ 0 };
    std::size_t idx{ 0 };

    auto iter = addr.begin ();
    while (iter != addr.end () && (std::isalpha (*iter) != 0))
    {
        col = col * 26 + (*iter - 'A' + 1);
        ++iter, ++idx;
    }

    std::string num = addr.substr (idx, addr.size ());
    try
    {
        row = std::stoi (num);
    }
    catch (...)
    {
        throw ExcelReader::ParseAddrException (addr);
    }

    return std::make_pair (row, col);
}

std::wstring
_transformString2Wstring (const std::string &s)
{
    setlocale (LC_CTYPE, "en_US.UTF-8");
    const char *_Source = s.c_str ();
    size_t len = std::strlen (_Source) + 1;
    size_t converted = 0;
    auto *wStr = new wchar_t[len];
    mbstowcs_s (&converted, wStr, len, _Source, _TRUNCATE);
    std::wstring result (wStr);
    delete[] wStr;
    return result;
}

fs::path
_getPath (const std::string &p)
{
    std::wstring wdirname = _transformString2Wstring (p);

    // 使用 std::filesystem 和 std::wstring 来创建目录
    std::filesystem::path path (wdirname);
    return path;
}

const auto xlsDeleter = [] (xls::xlsWorkBook *wb)
{
    if (wb)
        xls::xls_close_WB (wb);
};
#endif
