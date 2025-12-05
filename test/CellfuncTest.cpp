#include "../src/Exceptions.h"

#include <cctype>
#include <cmath>
#include <cstring>
#include <filesystem>
#include <gtest/gtest.h>
#include <iomanip>
#include <iostream>
#include <limits>
#include <optional>
#include <sstream>
#include <string>
#include <variant>
#include <vector>

// 自定义Cell结构体
struct Cell
{
    int row;
    int col;
    int id;    // 记录类型标识
    int xf;    // 格式信息
    int l;     // 长度或其他标志
    double d;  // 数值
    char* str; // 字符串内容
};

// 模拟xls.h中的常量定义
#define XLS_RECORD_LABELSST 0x00FD
#define XLS_RECORD_LABEL 0x0204
#define XLS_RECORD_RSTRING 0x00D6
#define XLS_RECORD_FORMULA 0x0006
#define XLS_RECORD_FORMULA_ALT 0x0406
#define XLS_RECORD_MULRK 0x00BD
#define XLS_RECORD_NUMBER 0x0203
#define XLS_RECORD_RK 0x027E
#define XLS_RECORD_MULBLANK 0x00BE
#define XLS_RECORD_BLANK 0x0201
#define XLS_RECORD_BOOLERR 0x0205

enum class CellType
{
    STRING,
    NUMBER,
    BOOL,
    UNKNOWN,
    BLANK,
    DATE
};
// 辅助函数声明
namespace fs = std::filesystem;

std::vector<std::string> formats = { "xls", "xlsx", "csv" };

inline bool
isExcelFormat (const std::string& format)
{
    return std::any_of (formats.begin (), formats.end (),
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
    if (start == std::string::npos)
        {
            return "";
        }
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
    std::string tmp = raw_value;
    std::transform (tmp.begin (), tmp.end (), tmp.begin (),
                    [] (unsigned char c) { return std::tolower (c); });
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
struct CellPosition
{
    std::size_t row;
    std::size_t col;
    std::string addr; // Like A1

    template<typename T>
    explicit CellPosition (T row, T col)
        : row (static_cast<std::size_t> (row))
        , col (static_cast<std::size_t> (col))
    {
        // 计算Excel地址 (A1, B2, etc.)
        int temp_col = col;
        std::string column_str = "";
        while (temp_col >= 0)
            {
                column_str = char ('A' + (temp_col % 26)) + column_str;
                temp_col = (temp_col / 26) - 1;
                if (temp_col < 0)
                    break;
            }
        addr = column_str + std::to_string (row + 1);
    }

    template<typename T>
    explicit CellPosition (std::pair<T, T> loc)
        : row (static_cast<std::size_t> (loc.first))
        , col (static_cast<std::size_t> (loc.second))
    {
        // 计算Excel地址
        int temp_col = col;
        std::string column_str = "";
        while (temp_col >= 0)
            {
                column_str = char ('A' + (temp_col % 26)) + column_str;
                temp_col = (temp_col / 26) - 1;
                if (temp_col < 0)
                    break;
            }
        addr = column_str + std::to_string (row + 1);
    }

    explicit CellPosition (const std::string& addr)
        : addr (addr)
    {
        // 解析地址 A1 -> (0,0), B2 -> (1,1)
        std::string col_part = "";
        std::string row_part = "";
        size_t i = 0;
        while (i < addr.length () && std::isalpha (addr[ i ]))
            {
                col_part += addr[ i ];
                i++;
            }
        while (i < addr.length () && std::isdigit (addr[ i ]))
            {
                row_part += addr[ i ];
                i++;
            }

        int col_num = 0;
        for (char c : col_part)
            {
                col_num = col_num * 26 + (c - 'A' + 1);
            }

        if (!row_part.empty ())
            {
                row = std::stoi (row_part) - 1;
            }
        else
            {
                row = 0;
            }

        col = col_num >= 0 ? static_cast<std::size_t> (col_num - 1) : 0;
    };
};

class XlsCell
{
private:
    Cell* cell_;
    CellPosition location_;
    std::optional<CellType> type_;
    std::variant<std::monostate, std::string, double, bool> value_;

    void
    inferTypeFromStringCell (bool trimWs)
    {
        std::string s = cell_->str == nullptr ? "" : std::string (cell_->str);
        if (isEmpty (s, trimWs))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        try
            {
                auto tmp = std::stod (s);
                if (tmp)
                    {
                        value_ = tmp;
                        type_ = CellType::NUMBER;
                    }
            }
        catch (...)
            {
            }

        try
            {
                auto strlower = tolower (s);
                auto tmp =
                    std::equal (strlower.begin (), strlower.end (), "bool");
                if (tmp)
                    {
                        value_ = tmp;
                        type_ = CellType::BOOL;
                    }
            }
        catch (...)
            {
            }
        type_ = CellType::STRING;
        value_ = trimWs ? trim (s) : s;
    };
    void
    inferTypeFromFormulaCell ()
    {
        if (cell_->l == 0)
            {
                auto strVal = std::to_string (cell_->d);
                if (isEmpty (strVal))
                    {
                        type_ = CellType::BLANK;
                        value_ = std::monostate{};
                        return;
                    }
                int format = cell_->xf;
                type_ =
                    (isDateTime (format)) ? CellType::DATE : CellType::NUMBER;
                value_ = cell_->d;
                return;
            }

        // 处理布尔公式
        if (cell_->str && strncmp ((char*)cell_->str, "bool", 4) == 0)
            {
                bool isValidBool =
                    (cell_->d == 0 && cell_->str
                     && strcasecmp ((char*)cell_->str, "false") == 0)
                    || (cell_->d == 1 && cell_->str
                        && strcasecmp ((char*)cell_->str, "true") == 0);
                if (isValidBool)
                    {
                        type_ = CellType::BLANK;
                        value_ = std::monostate{};
                    }
                else
                    {
                        type_ = CellType::BOOL;
                        value_ = cell_->d != 0;
                    }
                return;
            }

        // 处理错误公式
        if (cell_->str && strncmp ((char*)cell_->str, "error", 5) == 0
            && cell_->d > 0)
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
                return;
            }

        // 处理字符串公式
        std::string str = cell_->str ? std::string (cell_->str) : "";
        if (isEmpty (str))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        else
            {
                type_ = CellType::STRING;
                value_ = str;
            }
    };
    void
    inferTypeFromNumberCell ()
    {
        if (isEmpty (std::to_string (cell_->col)))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
                return;
            }
        int format = cell_->xf;
        type_ = (isDateTime (format)) ? CellType::DATE : CellType::NUMBER;
        value_ = cell_->d;
    };
    void
    inferTypeFromBoolErrCell (bool trimWs)
    {
        std::string tmp = trimWs ? trim (cell_->str) : cell_->str;
        if (tmp.size () == 0)
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        std::string strVal = cell_->str ? std::string (cell_->str) : "";
        std::string lowerStrVal = tolower (strVal);

        if (lowerStrVal == "false" || lowerStrVal == "true")
            {
                type_ = CellType::BOOL;
                value_ = cell_->d != 0;
            }
        else
            {
                type_ = CellType::STRING;
                value_ = tmp;
            }
    };
    void
    inferBlankCell (bool trimWs)
    {
        if (!cell_ || !cell_->str)
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
                return;
            }

        std::string tmp = trimWs ? trim (std::string (cell_->str)) : cell_->str;
        if (isEmpty (tmp))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        type_ = CellType::STRING;
        value_ = tmp;
    };
    void
    inferUnknownCell ()
    {
        type_ = CellType::UNKNOWN;
        value_ = std::monostate{};
    };
    std::string
    asBoolString () const
    {
        bool boolValue = false;

        if (auto* b = std::get_if<bool> (&value_))
            {
                boolValue = *b;
            }
        else if (cell_)
            {
                boolValue = cell_->d != 0.0;
            }

        return boolValue ? "TRUE" : "FALSE";
    }

    std::string
    asNumberString () const
    {
        double value = 0.0;

        if (auto* d = std::get_if<double> (&value_))
            {
                value = *d;
            }
        else if (cell_)
            {
                value = cell_->d;
            }

        return formatDouble (value);
    }

    std::string
    asString (bool trimWs) const
    {
        std::string result;

        if (auto* s = std::get_if<std::string> (&value_))
            {
                result = *s;
            }
        else if (cell_ && cell_->str)
            {
                result = std::string (cell_->str);
            }

        return trimWs ? trim (result) : result;
    }

    std::string
    formatDouble (double value) const
    {
        std::ostringstream oss;

        // 检查是否为整数
        double intpart;
        if (std::modf (value, &intpart) == 0.0)
            {
                // 整数值：使用整数格式化
                if (value >= std::numeric_limits<int64_t>::min ()
                    && value <= double (std::numeric_limits<int64_t>::max ()))
                    {
                        oss << static_cast<int64_t> (value);
                    }
                else
                    {
                        oss << std::fixed << std::setprecision (0) << value;
                    }
            }
        else
            {
                // 浮点值：智能格式化
                oss << std::setprecision (15) << value;

                std::string str = oss.str ();

                // 移除尾随的0
                if (str.find ('.') != std::string::npos)
                    {
                        // 移除尾随的0
                        str.erase (str.find_last_not_of ('0') + 1,
                                   std::string::npos);
                        // 如果最后是小数点，也移除
                        if (!str.empty () && str.back () == '.')
                            {
                                str.pop_back ();
                            }
                    }

                return str;
            }

        return oss.str ();
    }

public:
    explicit XlsCell (Cell* cell)
        : cell_ (cell ? cell : nullptr)
        , type_ (std::nullopt)
        , location_ (CellPosition (0, 0))
        , value_ (std::monostate{})
    {
        if (cell)
            {
                this->inferValue (false);
                this->location_ =
                    CellPosition (std::make_pair (cell_->row, cell_->col));
            }
    };

    int
    row () const
    {
        return location_.row;
    }

    int
    col () const
    {
        return location_.col;
    }

    CellType
    type ()
    {
        if (!type_.has_value () && cell_ != nullptr)
            {
                inferValue (false);
            }
        if (!type_.has_value () && cell_ == nullptr)
            {
                throw ExcelReader::NullCellException ("");
            }
        return type_.value ();
    }


    void
    inferValue (const bool trimWs) // 添加 const 修饰符
    {
        // 如果已经有类型，不需要重新推断（可以修改这个逻辑如果需要强制重新推断）
        if (type_.has_value () && type_ != CellType::UNKNOWN
            && cell_ != nullptr)
            {
                return;
            }

        if (!cell_)
            {
                inferBlankCell (trimWs);
                return;
            }

        // 使用 switch 分发到具体的处理函数
        switch (cell_->id)
            {
                case XLS_RECORD_LABELSST:
                case XLS_RECORD_LABEL:
                case XLS_RECORD_RSTRING:
                    inferTypeFromStringCell (trimWs);
                    break;

                case XLS_RECORD_FORMULA:
                case XLS_RECORD_FORMULA_ALT:
                    inferTypeFromFormulaCell ();
                    break;

                case XLS_RECORD_MULRK:
                case XLS_RECORD_NUMBER:
                case XLS_RECORD_RK:
                    inferTypeFromNumberCell ();
                    break;

                case XLS_RECORD_MULBLANK:
                case XLS_RECORD_BLANK:
                    inferBlankCell (trimWs);
                    break;

                case XLS_RECORD_BOOLERR:
                    inferTypeFromBoolErrCell (trimWs);
                    break;

                default:
                    inferUnknownCell ();
                    break;
            }
    }

    std::string
    asStdString (const bool trimWs,
                 const std::vector<std::string>& stringTable) const
    {
        // 处理未初始化的情况
        if (!type_.has_value ())
            {
                return "";
            }

        CellType cellType = type_.value ();

        switch (cellType)
            {
                case CellType::UNKNOWN:
                case CellType::BLANK:
                    return "";

                case CellType::BOOL:
                    return asBoolString ();

                case CellType::DATE:
                case CellType::NUMBER:
                    return asNumberString ();

                case CellType::STRING:
                    return asString (trimWs);

                default:
                    return "";
            }
    }

    bool
    asLogical () const
    {
        if (!cell_ && type_.value () == CellType::BLANK)
            {
                return false;
            }

        switch (type_.value ())
            {

                case CellType::UNKNOWN:
                case CellType::BLANK:
                case CellType::DATE:
                case CellType::STRING:
                    return false;

                case CellType::BOOL:
                    if (auto* b = std::get_if<bool> (&value_))
                        {
                            return *b;
                        }
                    return cell_->d != 0;

                case CellType::NUMBER:
                    if (auto* d = std::get_if<double> (&value_))
                        {
                            return *d != 0.0;
                        }
                    return cell_->d != 0;

                default:
                    return false;
            }
    }

    double
    asDouble () const
    {
        if (!cell_ && type_.value () == CellType::BLANK)
            {
                return 0.0;
            }

        switch (type_.value ())
            {

                case CellType::UNKNOWN:
                case CellType::BLANK:
                case CellType::STRING:
                    return 0.0;

                case CellType::BOOL:
                    if (auto* b = std::get_if<bool> (&value_))
                        {
                            return *b ? 1.0 : 0.0;
                        }
                    return cell_->d != 0 ? 1.0 : 0.0;

                case CellType::DATE:
                case CellType::NUMBER:
                    if (auto* d = std::get_if<double> (&value_))
                        {
                            return *d;
                        }
                    return cell_->d;

                default:
                    return 0.0;
            }
    }

    // 新增value()方法，返回实际存储的值
    const std::variant<std::monostate, std::string, double, bool>&
    value () const
    {
        return value_;
    }

    // 便捷方法：获取实际值的类型
    std::string
    valueType () const
    {
        return std::visit (
            [] (const auto& val) -> std::string
                {
                    using T = std::decay_t<decltype (val)>;
                    if constexpr (std::is_same_v<T, std::monostate>)
                        {
                            return "monostate";
                        }
                    else if constexpr (std::is_same_v<T, std::string>)
                        {
                            return "string";
                        }
                    else if constexpr (std::is_same_v<T, double>)
                        {
                            return "double";
                        }
                    else if constexpr (std::is_same_v<T, bool>)
                        {
                            return "bool";
                        }
                    return "unknown";
                },
            value_);
    }
};

// 用于单元测试的辅助函数
Cell*
createStringCell (int row, int col, const std::string& content)
{
    Cell* cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_LABEL;
    cell->xf = 0;
    cell->l = 0;
    cell->d = 0.0;
    cell->str = new char[ content.length () + 1 ];
    strcpy (cell->str, content.c_str ());
    return cell;
}

Cell*
createNumberCell (int row, int col, double value)
{
    Cell* cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_NUMBER;
    cell->xf = 0;
    cell->l = 0;
    cell->d = value;
    cell->str = nullptr;
    return cell;
}

Cell*
createBoolCell (int row, int col, bool value)
{
    Cell* cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_BOOLERR;
    cell->xf = 0;
    cell->l = 0;
    cell->d = value ? 1.0 : 0.0;
    cell->str =
        value ? const_cast<char*> ("true") : const_cast<char*> ("false");
    return cell;
}

Cell*
createBlankCell (int row, int col)
{
    Cell* cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_BLANK;
    cell->xf = 0;
    cell->l = 0;
    cell->d = 0.0;
    cell->str = nullptr;
    return cell;
}

void
deleteCell (Cell* cell)
{
    if (cell)
        {
            if (cell->str)
                {
                    delete[] cell->str;
                }
            delete cell;
        }
}

// 测试XlsCell的基本功能
class XlsCellTest : public ::testing::Test
{
protected:
    void
    SetUp () override
    {
    }
    void
    TearDown () override
    {
    }

    // Helper functions to create test cells
    Cell*
    createTestStringCell (int row, int col, const std::string& content)
    {
        return createStringCell (row, col, content);
    }

    Cell*
    createTestNumberCell (int row, int col, double value)
    {
        return createNumberCell (row, col, value);
    }

    Cell*
    createTestBoolCell (int row, int col, bool value)
    {
        return createBoolCell (row, col, value);
    }

    Cell*
    createTestBlankCell (int row, int col)
    {
        return createBlankCell (row, col);
    }
};

// 测试构造函数
TEST_F (XlsCellTest, ConstructorWithValidCell)
{
    // Given: 一个有效的字符串单元格
    Cell* cell = createTestStringCell (0, 0, "Hello World");

    // When: 创建XlsCell对象
    XlsCell xlsCell (cell);

    // Then: 应该正确初始化
    EXPECT_EQ (xlsCell.row (), 0);
    EXPECT_EQ (xlsCell.col (), 0);
    EXPECT_EQ (xlsCell.type (), CellType::STRING);

    // Cleanup
    // deleteCell (cell);
}

TEST_F (XlsCellTest, ConstructorWithNullCell)
{
    // When: 使用nullptr创建XlsCell
    // 应当抛出NullCellException
    EXPECT_THROW (XlsCell xlsCell (nullptr), ExcelReader::NullCellException);
}

// 测试字符串单元格
TEST_F (XlsCellTest, StringCellHandling)
{
    // Given: 一个包含内容的字符串单元格
    Cell* cell = createTestStringCell (0, 0, "Test Content");
    XlsCell xlsCell (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCell.type (), CellType::STRING);
    EXPECT_EQ (xlsCell.asStdString (false, {}), "Test Content");
    EXPECT_FALSE (xlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 0.0);
    EXPECT_EQ (xlsCell.valueType (), "string");

    // Cleanup
    // deleteCell (cell);
}

// 测试带空格的字符串单元格
TEST_F (XlsCellTest, StringCellWithWhitespace)
{
    // Given: 一个包含前后空格的字符串单元格
    Cell* cell = createTestStringCell (0, 0, "  Test Content  ");
    XlsCell xlsCell (cell);

    EXPECT_EQ (xlsCell.asStdString (true, {}), "Test Content"); // trim=true
    EXPECT_EQ (xlsCell.asStdString (false, {}),
               "  Test Content  "); // trim=false

    // Cleanup
    // deleteCell (cell);
}

// 测试数字单元格
TEST_F (XlsCellTest, NumberCellHandling)
{
    // Given: 一个数字单元格
    Cell* cell = createTestNumberCell (0, 0, 123.456);
    XlsCell xlsCell (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCell.type (), CellType::NUMBER);
    EXPECT_EQ (xlsCell.asStdString (false, {}), "123.456");
    EXPECT_TRUE (xlsCell.asLogical ()); // 非零数字应为true
    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 123.456);
    EXPECT_EQ (xlsCell.valueType (), "double");

    // Cleanup
    // deleteCell (cell);
}

// 测试整数单元格
TEST_F (XlsCellTest, IntegerCellHandling)
{
    // Given: 一个整数单元格
    Cell* cell = createTestNumberCell (0, 0, 42.0);
    XlsCell xlsCell (cell);

    // When & Then: 检查字符串转换是否为整数形式
    EXPECT_EQ (xlsCell.asStdString (true, {}), "42");

    // Cleanup
    // deleteCell (cell);
}

// 测试布尔单元格
TEST_F (XlsCellTest, BoolCellHandling)
{
    // Given: 一个TRUE布尔单元格
    Cell* trueCell = createTestBoolCell (0, 0, true);
    XlsCell trueXlsCell (trueCell);

    // When & Then: 检查TRUE值的各种转换
    EXPECT_EQ (trueXlsCell.type (), CellType::BOOL);
    EXPECT_EQ (trueXlsCell.asStdString (false, {}), "TRUE");
    EXPECT_TRUE (trueXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (trueXlsCell.asDouble (), 1.0);
    EXPECT_EQ (trueXlsCell.valueType (), "bool");

    // Cleanup
    // deleteCell (trueCell);

    // Given: 一个FALSE布尔单元格
    Cell* falseCell = createTestBoolCell (0, 0, false);
    XlsCell falseXlsCell (falseCell);

    // When & Then: 检查FALSE值的各种转换
    EXPECT_EQ (falseXlsCell.type (), CellType::BOOL);
    EXPECT_EQ (falseXlsCell.asStdString (false, {}), "FALSE");
    EXPECT_FALSE (falseXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (falseXlsCell.asDouble (), 0.0);

    // Cleanup
    // deleteCell (falseCell);
}

// 测试空白单元格
TEST_F (XlsCellTest, BlankCellHandling)
{
    // Given: 一个空白单元格
    Cell* cell = createTestBlankCell (0, 0);
    XlsCell xlsCell (cell);

    // When & Then: 检查空白单元格的行为
    EXPECT_EQ (xlsCell.type (), CellType::BLANK);
    EXPECT_EQ (xlsCell.asStdString (false, {}), "");
    EXPECT_FALSE (xlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 0.0);
    EXPECT_EQ (xlsCell.valueType (), "monostate");

    // Cleanup
    // deleteCell (cell);
}

// 测试空指针单元格
TEST_F (XlsCellTest, NullCellHandling)
{
    // Given: 一个null单元格
    XlsCell xlsCell (nullptr);

    // When & Then: 检查其行为
    EXPECT_EQ (xlsCell.type (), CellType::BLANK);
    EXPECT_EQ (xlsCell.asStdString (false, {}), "");
    EXPECT_FALSE (xlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 0.0);
    EXPECT_EQ (xlsCell.valueType (), "monostate");
}

// 测试CellPosition构造函数
TEST (CellPositionTest, FromRowCol)
{
    // Given: 行列号
    CellPosition pos (0, 0);

    // When & Then: 检查地址计算
    EXPECT_EQ (pos.row, 0);
    EXPECT_EQ (pos.col, 0);
    EXPECT_EQ (pos.addr, "A1");
}

TEST (CellPositionTest, FromPair)
{
    // Given: pair形式的位置
    CellPosition pos (std::make_pair (1, 2));

    // When & Then: 检查行列和地址
    EXPECT_EQ (pos.row, 1);
    EXPECT_EQ (pos.col, 2);
    EXPECT_EQ (pos.addr, "C2");
}

TEST (CellPositionTest, FromAddress)
{
    // Given: 地址字符串
    CellPosition pos ("B3");

    // When & Then: 检查解析结果
    EXPECT_EQ (pos.row, 2); // Excel行从1开始，内部从0开始
    EXPECT_EQ (pos.col, 1); // Excel列从A开始，内部从0开始
    EXPECT_EQ (pos.addr, "B3");
}

TEST (CellPositionTest, ComplexColumnAddress)
{
    // Given: 复杂列地址（超过Z）
    CellPosition pos ("AA1");

    // When & Then: 检查解析结果
    EXPECT_EQ (pos.row, 0);
    EXPECT_EQ (pos.col, 26);
    EXPECT_EQ (pos.addr, "AA1");
}

// 测试Excel格式检查函数
TEST (FormatTest, isExcelFormatTest)
{
    EXPECT_TRUE (isExcelFormat ("xls"));
    EXPECT_TRUE (isExcelFormat ("xlsx"));
    EXPECT_TRUE (isExcelFormat ("csv"));
    EXPECT_FALSE (isExcelFormat ("txt"));
    EXPECT_FALSE (isExcelFormat ("doc"));
    EXPECT_FALSE (isExcelFormat ("pdf"));
}

// 测试trim函数
TEST (StringTest, TrimTest)
{
    EXPECT_EQ (trim ("  hello  "), "hello");
    EXPECT_EQ (trim ("hello"), "hello");
    EXPECT_EQ (trim ("   "), "");
    EXPECT_EQ (trim (""), "");
    EXPECT_EQ (trim (" hello world "), "hello world");
}

// 测试isEmpty函数
TEST (StringTest, IsEmptyTest)
{
    EXPECT_TRUE (isEmpty (""));
    EXPECT_TRUE (isEmpty ("   "));
    EXPECT_TRUE (isEmpty (" \t "));
    EXPECT_TRUE (isEmpty (" \n "));
    EXPECT_FALSE (isEmpty ("hello"));
    EXPECT_FALSE (isEmpty ("  hello  "));
    EXPECT_TRUE (isEmpty ("   ", true));
    EXPECT_FALSE (isEmpty ("  hello  ", true));
    EXPECT_TRUE (isEmpty ("  \t  \n  ", true));
}

int
main (int argc, char** argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}
