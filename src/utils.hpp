#ifndef UTILS_H
#define UTILS_H

#include "Exceptions.hpp"

#include <algorithm>
#include <cstring>
#include <filesystem>
#include <optional>
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

inline bool
startsWith (const std::string &str, const std::string &prefix)
{
    return str.substr (0, prefix.size ()) == prefix;
};

inline bool
isExcelFormat (const std::string &format)
{
    const std::vector<std::string> formats = { "xls", "xlsx", "csv" };
    return std::ranges::any_of (formats, [format] (const std::string &fmt)
                                { return format == fmt; });
}

inline bool
isValid (const fs::path &path)
{
    if (!fs::exists (path))
    {
        throw ExcelReader::FileNotFoundException (path.string ());
    }

    if (fs::is_directory (path))
    {
        throw ExcelReader::FileNotFoundException (path.string ());
    }

    if (!fs::is_regular_file (path))
    {
        throw ExcelReader::UnsupportedFormatException ("format "
                                                       + path.string ());
    }

    if (!isExcelFormat (path.extension ().string ()))
    {
        throw ExcelReader::UnsupportedFormatException (
            "format " + path.extension ().string ());
    }
    return true;
}

std::string
trim (const std::string &str)
{
    auto start = str.find_first_not_of (" \t\r\n\v");
    if (start == std::string::npos)
    {
        return ""; // String contains only whitespace
    }
    auto end = str.find_last_not_of (" \t\r\n\v");
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
        auto start = std::ranges::find_if_not (
            raw_string, [] (unsigned char chr) { return std::isspace (chr); });
        // 找到最后一个非空字符的位置
        auto end = std::ranges::find_if_not (
                       raw_string.rbegin (), raw_string.rend (),
                       [] (unsigned char chr) { return std::isspace (chr); })
                       .base ();

        // 如果 start >= end，则表示去除空白后字符串为空
        return start >= end;
    }

    return std::ranges::all_of (raw_value, [] (unsigned char chr)
                                { return std::isspace (chr); });
}

inline std::string
tolower (const std::string &raw_value)
{
    std::string tmp = raw_value;
    std::ranges::transform (tmp, tmp.begin (),
                            [] (char chr) { return std::tolower (chr); });
    return tmp;
}

inline bool
isDateTime (int formatId)
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
    // Built-in date/time format IDs according to ECMA-376:
    // 14-22: Date and time formats
    // 27-36: More date formats (Chinese, Japanese, Korean locales)
    // 45-47: Time formats
    // 50-58: More date/time formats
    // 71-81: Additional date/time formats

    // Check for built-in date/time format IDs
    return (formatId >= 14 && formatId <= 22) || // Standard date/time formats
           (formatId >= 27 && formatId <= 36) || // CJK date formats
           (formatId >= 45 && formatId <= 47) || // Time formats
           (formatId >= 50 && formatId <= 58) || // More date/time formats
           (formatId >= 71 && formatId <= 81); // Additional date/time formats

    // Note: For format IDs >= 164, they are custom formats.
    // We would need to parse the format string to determine if they contain
    // date/time formatting codes (y, m, d, h, s, etc.).
    // The original code had a set of specific IDs, but that approach was
    // incomplete and potentially incorrect for custom formats.
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
        throw ExcelReader::AddressParseException (addr);
    }

    return std::make_pair (row, col);
}

auto
_transformString2Wstring (const std::string &s)
{
    setlocale (LC_CTYPE, "en_US.UTF-8");
    const size_t len = s.length () + 1;
    size_t converted = 0;
    std::vector<wchar_t> wStr (len);
    mbstowcs_s (&converted, wStr.data (), len, s.c_str (), _TRUNCATE);
    return std::wstring (wStr.data ());
}

fs::path
_getPath (const std::string &p)
{
    std::wstring wdirname = _transformString2Wstring (p);

    // 使用 std::filesystem 和 std::wstring 来创建目录
    std::filesystem::path path (wdirname);
    return path;
}

#endif
