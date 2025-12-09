#include "../src/Exceptions.h"
#include "gtest/gtest.h"

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstring>
#include <filesystem>
#include <iomanip>
#include <iostream>
#include <limits>
#include <optional>
#include <sstream>
#include <string>
#include <variant>
#include <vector>

struct Cell
{
    int row;
    int col;
    int id;    // 记录类型标识
    int xf;    // 格式信息
    int l;     // 长度或其他标志
    double d;  // 数值
    char *str; // 字符串内容
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

namespace fs = std::filesystem;

// 辅助函数声明
std::vector<std::string> formats = { "xls", "xlsx", "csv" };
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
    std::string tmp{};
    std::transform (raw_value.begin (), raw_value.end (), tmp.begin (),
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

// 用于单元测试的辅助函数
Cell *
createStringCell (int row, int col, const std::string &content)
{
    Cell *cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_LABEL;
    cell->xf = 0;
    cell->l = 0;
    cell->d = 0.0;
    cell->str = new char[content.length () + 1];
    strcpy (cell->str, content.c_str ());
    return cell;
}

Cell *
createNumberCell (int row, int col, double value)
{
    Cell *cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_NUMBER;
    cell->xf = 0;
    cell->l = 0;
    cell->d = value;
    cell->str = nullptr;
    return cell;
}

Cell *
createBoolCell (int row, int col, bool value)
{
    Cell *cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_BOOLERR;
    cell->xf = 0;
    cell->l = 0;
    cell->d = value ? 1.0 : 0.0;
    cell->str
        = value ? const_cast<char *> ("true") : const_cast<char *> ("false");
    return cell;
}

Cell *
createBlankCell (int row, int col)
{
    Cell *cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_BLANK;
    cell->xf = 0;
    cell->l = 0;
    cell->d = 0.0;
    cell->str = nullptr;
    return cell;
}

Cell *
createFormulaCell (int row, int col, double value,
                   const std::string &formula = "")
{
    Cell *cell = new Cell ();
    cell->row = row;
    cell->col = col;
    cell->id = XLS_RECORD_FORMULA;
    cell->xf = 0;
    cell->l = formula.length ();
    cell->d = value;
    if (!formula.empty ())
    {
        cell->str = new char[formula.length () + 1];
        strcpy (cell->str, formula.c_str ());
    }
    else
    {
        cell->str = nullptr;
    }
    return cell;
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
    static Cell *
    createTestStringCell (int row, int col, const std::string &content)
    {
        return createStringCell (row, col, content);
    }

    static Cell *
    createTestNumberCell (int row, int col, double value)
    {
        return createNumberCell (row, col, value);
    }

    static Cell *
    createTestBoolCell (int row, int col, bool value)
    {
        return createBoolCell (row, col, value);
    }

    static Cell *
    createTestBlankCell (int row, int col)
    {
        return createBlankCell (row, col);
    }

    static Cell *
    createTestFormulaCell (int row, int col, double value,
                           const std::string &formula = "")
    {
        return createFormulaCell (row, col, value, formula);
    }
};

enum class CellType : uint8_t
{
    STRING = 0,
    NUMBER,
    BOOL,
    UNKNOWN,
    BLANK,
    DATE
};

struct CellPosition
{
  public:
    std::optional<std::size_t> row;
    std::optional<std::size_t> col;
    std::optional<std::string> addr; // Like A1

    explicit CellPosition (std::nullopt_t /*unused*/)
        : row (std::nullopt), col (std::nullopt), addr (std::nullopt) {};

    explicit CellPosition (std::size_t rowVal, std::size_t colVal)
        : row (rowVal), col (colVal), addr (std::nullopt)
    {
        calculateExcelAddress ();
    }

    template <typename T1, typename T2>
    explicit CellPosition (T1 rowVal, T2 colVal)
    {
        row = convertToSizef (rowVal);
        col = convertToSizef (colVal);
        if (row.has_value () && col.has_value ())
        {
            calculateExcelAddress ();
        }
    }

    template <typename T1, typename T2>
    explicit CellPosition (std::pair<T1, T2> loc)
        : row (static_cast<std::size_t> (loc.first)),
          col (static_cast<std::size_t> (loc.second)), addr (std::nullopt)
    {
        calculateExcelAddress ();
    }

    explicit CellPosition (const std::string &addr) : addr (addr)
    {
        // 解析地址 A1 -> (0,0), B2 -> (1,1)
        std::string col_part;
        std::string row_part;
        size_t idx = 0;
        while (idx < addr.length () && (std::isalpha (addr[idx]) != 0))
        {
            col_part += addr[idx];
            idx++;
        }
        while (idx < addr.length () && (std::isdigit (addr[idx]) != 0))
        {
            row_part += addr[idx];
            idx++;
        }

        int col_num = 0;
        for (char chr : col_part)
        {
            col_num = col_num * 26 + (chr - 'A' + 1);
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

    ~CellPosition () = default;

  private:
    void
    calculateExcelAddress ()
    {
        if (!row.has_value () || !col.has_value ())
        {
            addr = std::nullopt;
            return;
        }

        std::string colPart;
        std::size_t colNum = col.value ();
        while (true)
        {
            colPart += static_cast<char> ('A' + (colNum % 26));
            if (colNum < 26)
            {
                break;
            }
            colNum = colNum / 26 - 1;
        }
        std::reverse (colPart.begin (), colPart.end ());
        this->addr = colPart + std::to_string (row.value () + 1);
    }

    template <typename T>
    static std::optional<std::size_t>
    convertToSizef (T value)
    {
        if constexpr (std::is_same_v<T, std::nullopt_t>)
        {
            return std::nullopt;
        }
        else if constexpr (std::is_integral_v<T>)
        {
            if (value >= 0)
            {
                return static_cast<std::size_t> (value);
            }

            return std::nullopt;
        }

        return std::nullopt;
    }
};

class XlsCell
{
  private:
    std::shared_ptr<Cell> cell_;
    CellPosition location_;
    std::optional<CellType> type_;
    std::variant<std::monostate, std::string, double, bool> value_;

    void
    inferValueFromStringCell (bool trimWs)
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
            if (tmp < std::numeric_limits<double>::epsilon ())
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
            auto tmp = std::equal (strlower.begin (), strlower.end (), "bool");
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
    inferValueFromFormulaCell ()
    {
        // 添加缺失的变量定义
        auto rawStringView = getStringView (cell_->str);
        constexpr std::string_view BoolStringView = "bool";

        if (cell_->l == 0)
        {
            auto strVal = std::to_string (cell_->d);
            if (isEmpty (strVal))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
                return;
            }
            auto format = cell_->xf;
            type_ = (isDateTime (format)) ? CellType::DATE : CellType::NUMBER;
            value_ = cell_->d;
            return;
        }

        // 处理布尔公式
        if (rawStringView.has_value ()
            && startsWith (rawStringView.value (), BoolStringView))
        {
            bool isValidBool
                = (cell_->d == 0
                   && rawStringView.value ().substr (0, 5) == "false")
                  || (cell_->d == 1
                      && rawStringView.value ().substr (0, 4) == "true");
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
        }

        // 处理错误公式
        if (rawStringView.has_value ()
            && rawStringView->substr (0, 5) == "error")
        {
            type_ = CellType::BLANK;
            value_ = std::monostate{};
            return;
        }

        // 处理字符串公式
        if (isEmpty (cell_->str))
        {
            type_ = CellType::BLANK;
            value_ = std::monostate{};
        }
        else
        {
            type_ = CellType::STRING;
            value_ = std::string (cell_->str);
        }
    };
    void
    inferValueFromNumberCell ()
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
    inferValueFromBoolErrCell (bool trimWs)
    {
        std::string tmp = trimWs ? trim (cell_->str) : cell_->str;
        if (tmp.size () == 0)
        {
            type_ = CellType::BLANK;
            value_ = std::monostate{};
        }
        std::string strVal
            = (cell_->str != nullptr) ? std::string (cell_->str) : "";
        std::string lowerStrVal = tolower (strVal);

        if (lowerStrVal == "false" || lowerStrVal == "true")
        {
            type_ = CellType::BOOL;
            value_ = cell_->d != 0;
        }
        else
        {
            type_ = CellType::BLANK;
            value_ = std::monostate{};
        }
    };

    void
    inferBlankCell (bool trimWs)
    {
        if (cell_->str == nullptr)
        {
            type_ = CellType::BLANK;
            value_ = std::monostate{};
            return;
        }

        std::string tmp
            = trimWs ? trim (std::string (cell_->str)) : cell_->str;
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
    void
    inferValue (const bool trimWs) // 添加 const 修饰符
    {
        // 如果已经有类型，不需要重新推断（可以修改这个逻辑如果需要强制重新推断）
        if (type_.has_value () && type_ != CellType::UNKNOWN
            && cell_ != nullptr)
        {
            return;
        }

        // 使用 switch 分发到具体的处理函数
        switch (cell_->id)
        {
        case XLS_RECORD_LABELSST:
        case XLS_RECORD_LABEL:
        case XLS_RECORD_RSTRING:
            inferValueFromStringCell (trimWs);
            break;

        case XLS_RECORD_FORMULA:
        case XLS_RECORD_FORMULA_ALT:
            inferValueFromFormulaCell (); // 默认公式计算的值不需要改变，不需要我们做额外推理
            break;

        case XLS_RECORD_MULRK:
        case XLS_RECORD_NUMBER:
        case XLS_RECORD_RK:
            inferValueFromNumberCell ();
            break;

        case XLS_RECORD_MULBLANK:
        case XLS_RECORD_BLANK:
            inferBlankCell (trimWs);
            // 根据需要决定是否为逻辑上的空而非字面值上的空， 例如：
            // " "逻辑上是空，但值不是字面值上的空
            break;

        case XLS_RECORD_BOOLERR:
            inferValueFromBoolErrCell (trimWs);
            break;

        default:
            inferUnknownCell ();
            break;
        }
    }

    [[nodiscard]] std::string
    asBoolString () const
    {
        bool boolValue = false;

        if (const auto *b = std::get_if<bool> (&value_))
        {
            boolValue = *b;
        }
        boolValue = cell_->d != 0.0;

        return boolValue ? "TRUE" : "FALSE";
    }

    [[nodiscard]] std::string
    asNumberString () const
    {
        double value = 0.0;

        if (const auto *d = std::get_if<double> (&value_))
        {
            value = *d;
        }
        value = cell_->d;

        return formatDouble (value);
    }

    [[nodiscard]] std::string
    asString (bool trimWs) const
    {
        std::string result;

        if (const auto *s = std::get_if<std::string> (&value_))
        {
            result = *s;
        }
        else if (cell_->str != nullptr)
        {
            result = std::string (cell_->str);
        }

        return trimWs ? trim (result) : result;
    }

    [[nodiscard]] static std::string
    formatDouble (double value)
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
                str.erase (str.find_last_not_of ('0') + 1, std::string::npos);
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

    void
    initialize (Cell *cell)
    {
        if (cell == nullptr)
        {
            throw ExcelReader::NullCellException ("");
        }

        cell_ = std::make_shared<Cell> (*cell);
        try
        {
            location_ = CellPosition (cell->row, cell->col);
            inferValue (false);
        }
        catch (...)
        {
            type_ = CellType::UNKNOWN;
            value_ = std::monostate{};
        }
    }

  public:
    XlsCell (const XlsCell &) = default;
    XlsCell (XlsCell &&) = default;
    XlsCell &operator= (const XlsCell &) = default;
    XlsCell &operator= (XlsCell &&) = default;
    explicit XlsCell (Cell *cell) : location_ (CellPosition (std::nullopt))
    {
        initialize (cell);
    }

    [[nodiscard]] int
    row () const
    {
        if (location_.row.has_value ())
        {
            return location_.row.value ();
        }
        return -1;
    }

    [[nodiscard]] int
    col () const
    {
        if (location_.col.has_value ())
        {

            return location_.col.value ();
        }
        return -1;
    }

    CellType
    type ()
    {
        if (!type_.has_value () && cell_ != nullptr)
        {
            inferValue (false);
        }
        return type_.value ();
    }

    // 如果trim那么意味着推到字符串可能代码的值，例如" 123"代表123
    // 否则只推到字面值

    [[nodiscard]] std::string
    asStdString (const bool trimWs) const
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

    [[nodiscard]] bool
    asLogical () const
    {
        switch (type_.value ())
        {

        case CellType::UNKNOWN:
        case CellType::BLANK:
        case CellType::DATE:
        case CellType::STRING:
            return false;

        case CellType::BOOL:
            if (const auto *b = std::get_if<bool> (&value_))
            {
                return *b;
            }
            return cell_->d != 0;

        case CellType::NUMBER:
            if (const auto *d = std::get_if<double> (&value_))
            {
                return *d != 0.0;
            }
            return cell_->d != 0;

        default:
            return false;
        }
    }

    [[nodiscard]] double
    asDouble () const
    {
        switch (type_.value ())
        {

        case CellType::UNKNOWN:
        case CellType::BLANK:
        case CellType::STRING:
            return 0.0;

        case CellType::BOOL:
            if (const auto *b = std::get_if<bool> (&value_))
            {
                return *b ? 1.0 : 0.0;
            }
            return cell_->d != 0 ? 1.0 : 0.0;

        case CellType::DATE:
        case CellType::NUMBER:
            if (const auto *d = std::get_if<double> (&value_))
            {
                return *d;
            }
            return cell_->d;

        default:
            return 0.0;
        }
    }

    // 新增value()方法，返回实际存储的值
    [[nodiscard]] const auto &
    value () const
    {
        return value_;
    }

    // 便捷方法：获取实际值的类型
    [[nodiscard]] std::string
    valueType () const
    {
        return std::visit (
            [] (const auto &val) -> std::string
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
// 测试构造函数
TEST_F (XlsCellTest, ConstructorWithValidCell)
{
    // Given: 一个有效的字符串单元格
    Cell *cell = createTestStringCell (0, 0, "Hello World");

    // When: 创建XlsCell对象
    XlsCell xlsCellObj (cell);

    // Then: 应该正确初始化
    EXPECT_EQ (xlsCellObj.row (), 0);
    EXPECT_EQ (xlsCellObj.col (), 0);
    EXPECT_EQ (xlsCellObj.type (), CellType::STRING);

    // Cleanup
}

TEST_F (XlsCellTest, ConstructorWithNullCell)
{
    // When: 使用nullptr创建XlsCell
    // 应当抛出NullCellException
    EXPECT_THROW (XlsCell xlsCellObj (nullptr),
                  ExcelReader::NullCellException);
}

// 测试字符串单元格
TEST_F (XlsCellTest, StringCellHandling)
{
    // Given: 一个包含内容的字符串单元格
    Cell *cell = createTestStringCell (0, 0, "Test Content");
    XlsCell xlsCellObj (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCellObj.type (), CellType::STRING);
    EXPECT_EQ (xlsCellObj.asStdString (false), "Test Content");
    EXPECT_FALSE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    EXPECT_EQ (xlsCellObj.valueType (), "string");

    // Cleanup
}

// 测试带空格的字符串单元格
TEST_F (XlsCellTest, StringCellWithWhitespace)
{
    // Given: 一个包含前后空格的字符串单元格
    Cell *cell = createTestStringCell (0, 0, "  Test Content  ");
    XlsCell xlsCellObj (cell);

    EXPECT_EQ (xlsCellObj.asStdString (true), "Test Content"); // trim=true
    EXPECT_EQ (xlsCellObj.asStdString (false),
               "  Test Content  "); // trim=false

    // Cleanup
}

// 测试数字单元格
TEST_F (XlsCellTest, NumberCellHandling)
{
    // Given: 一个数字单元格
    Cell *cell = createTestNumberCell (0, 0, 123.456);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
    EXPECT_EQ (xlsCellObj.asStdString (false), "123.456");
    EXPECT_TRUE (xlsCellObj.asLogical ()); // 非零数字应为true
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 123.456);
    EXPECT_EQ (xlsCellObj.valueType (), "double");

    // Cleanup
}

// 测试整数单元格
TEST_F (XlsCellTest, IntegerCellHandling)
{
    // Given: 一个整数单元格
    Cell *cell = createTestNumberCell (0, 0, 42.0);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查字符串转换是否为整数形式
    EXPECT_EQ (xlsCellObj.asStdString (true), "42");

    // Cleanup
}

// 测试布尔单元格
TEST_F (XlsCellTest, BoolCellHandling)
{
    // Given: 一个TRUE布尔单元格
    Cell *trueCell = createTestBoolCell (0, 0, true);
    XlsCell trueXlsCell (trueCell);

    // When & Then: 检查TRUE值的各种转换
    EXPECT_EQ (trueXlsCell.type (), CellType::BOOL);
    EXPECT_EQ (trueXlsCell.asStdString (false), "TRUE");
    EXPECT_TRUE (trueXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (trueXlsCell.asDouble (), 1.0);
    EXPECT_EQ (trueXlsCell.valueType (), "bool");

    // Given: 一个FALSE布尔单元格
    Cell *falseCell = createTestBoolCell (0, 0, false);
    XlsCell falseXlsCell (falseCell);

    // When & Then: 检查FALSE值的各种转换
    EXPECT_EQ (falseXlsCell.type (), CellType::BOOL);
    EXPECT_EQ (falseXlsCell.asStdString (false), "FALSE");
    EXPECT_FALSE (falseXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (falseXlsCell.asDouble (), 0.0);
}

// 测试空白单元格
TEST_F (XlsCellTest, BlankCellHandling)
{
    // Given: 一个空白单元格
    Cell *cell = createTestBlankCell (0, 0);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查空白单元格的行为
    EXPECT_EQ (xlsCellObj.type (), CellType::BLANK);
    EXPECT_EQ (xlsCellObj.asStdString (false), "");
    EXPECT_FALSE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    EXPECT_EQ (xlsCellObj.valueType (), "monostate");

    // Cleanup
}

// 测试公式单元格
TEST_F (XlsCellTest, FormulaCellHandling)
{
    // Given: 一个公式单元格
    Cell *cell = createTestFormulaCell (0, 0, 100.0, "SUM(A1:A10)");
    XlsCell xlsCellObj (cell);

    // When & Then: 检查公式单元格的行为
    // 注意：根据XlsCell.h的实现，公式单元格会被推断为NUMBER类型
    EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
    EXPECT_EQ (xlsCellObj.asStdString (false), "100");
    EXPECT_TRUE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 100.0);

    // Cleanup
}

// 测试CellPosition构造函数
TEST (CellPositionTest, FromRowCol)
{
    // Given: 行列号
    CellPosition pos (0, 0);

    // When & Then: 检查地址计算
    EXPECT_EQ (pos.row.value (), 0);
    EXPECT_EQ (pos.col.value (), 0);
    EXPECT_EQ (pos.addr.value (), "A1");
}

TEST (CellPositionTest, FromPair)
{
    // Given: pair形式的位置
    CellPosition pos (std::make_pair (1, 2));

    // When & Then: 检查行列和地址
    EXPECT_EQ (pos.row.value (), 1);
    EXPECT_EQ (pos.col.value (), 2);
    EXPECT_EQ (pos.addr.value (), "C2");
}

TEST (CellPositionTest, FromAddress)
{
    // Given: 地址字符串
    CellPosition pos ("B3");

    // When & Then: 检查解析结果
    EXPECT_EQ (pos.row.value (), 2); // Excel行从1开始，内部从0开始
    EXPECT_EQ (pos.col.value (), 1); // Excel列从A开始，内部从0开始
    EXPECT_EQ (pos.addr.value (), "B3");
}

TEST (CellPositionTest, ComplexColumnAddress)
{
    // Given: 复杂列地址（超过Z）
    CellPosition pos ("AA1");

    // When & Then: 检查解析结果
    EXPECT_EQ (pos.row.value (), 0);
    EXPECT_EQ (pos.col.value (), 26);
    EXPECT_EQ (pos.addr.value (), "AA1");
}

// 测试nullopt构造函数
TEST (CellPositionTest, NulloptConstructor)
{
    // Given: nullopt构造函数
    CellPosition pos (std::nullopt);

    // When & Then: 检查所有字段都是nullopt
    EXPECT_FALSE (pos.row.has_value ());
    EXPECT_FALSE (pos.col.has_value ());
    EXPECT_FALSE (pos.addr.has_value ());
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

// 测试XlsCell的value()方法
TEST_F (XlsCellTest, ValueMethodTest)
{
    // 测试字符串单元格的value()
    {
        Cell *cell = createTestStringCell (0, 0, "Test");
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<std::string> (value));
        EXPECT_EQ (std::get<std::string> (value), "Test");
    }

    // 测试数字单元格的value()
    {
        Cell *cell = createTestNumberCell (0, 0, 123.45);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<double> (value));
        EXPECT_DOUBLE_EQ (std::get<double> (value), 123.45);
    }

    // 测试布尔单元格的value()
    {
        Cell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<bool> (value));
        EXPECT_TRUE (std::get<bool> (value));
    }

    // 测试空白单元格的value()
    {
        Cell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<std::monostate> (value));
    }
}

// 测试XlsCell的valueType()方法
TEST_F (XlsCellTest, ValueTypeMethodTest)
{
    // 测试各种单元格类型的valueType()
    {
        Cell *cell = createTestStringCell (0, 0, "Test");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "string");
    }

    {
        Cell *cell = createTestNumberCell (0, 0, 123.45);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "double");
    }

    {
        Cell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "bool");
    }

    {
        Cell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "monostate");
    }
}

// 测试XlsCell的复制构造函数和赋值运算符
TEST_F (XlsCellTest, CopyAndMoveOperations)
{
    // 测试复制构造函数
    {
        Cell *cell = createTestStringCell (0, 0, "Original");
        XlsCell original (cell);
        XlsCell copy (original); // 复制构造函数

        EXPECT_EQ (original.row (), copy.row ());
        EXPECT_EQ (original.col (), copy.col ());
        EXPECT_EQ (original.type (), copy.type ());
        EXPECT_EQ (original.asStdString (false), copy.asStdString (false));
    }

    // 测试移动构造函数
    {
        Cell *cell = createTestNumberCell (0, 0, 42.0);
        XlsCell original (cell);
        XlsCell moved (std::move (original)); // 移动构造函数

        EXPECT_EQ (moved.row (), 0);
        EXPECT_EQ (moved.col (), 0);
        EXPECT_EQ (moved.type (), CellType::NUMBER);
        EXPECT_EQ (moved.asStdString (false), "42");
    }

    // 测试复制赋值运算符
    {
        Cell *cell1 = createTestStringCell (0, 0, "First");
        Cell *cell2 = createTestNumberCell (1, 1, 99.0);

        XlsCell cellA (cell1);
        XlsCell cellB (cell2);

        cellB = cellA; // 复制赋值

        EXPECT_EQ (cellB.row (), 0);
        EXPECT_EQ (cellB.col (), 0);
        EXPECT_EQ (cellB.type (), CellType::STRING);
        EXPECT_EQ (cellB.asStdString (false), "First");
    }

    // 测试移动赋值运算符
    {
        Cell *cell1 = createTestBoolCell (0, 0, true);
        Cell *cell2 = createTestBlankCell (1, 1);

        XlsCell cellA (cell1);
        XlsCell cellB (cell2);

        cellB = std::move (cellA); // 移动赋值

        EXPECT_EQ (cellB.row (), 0);
        EXPECT_EQ (cellB.col (), 0);
        EXPECT_EQ (cellB.type (), CellType::BOOL);
        EXPECT_EQ (cellB.asStdString (false), "TRUE");
    }
}

// 测试边缘情况
TEST_F (XlsCellTest, EdgeCases)
{
    // 测试空字符串单元格
    {
        Cell *cell = createTestStringCell (0, 0, "");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::BLANK);
        EXPECT_EQ (xlsCellObj.asStdString (false), "");
    }

    // 测试只包含空格的字符串单元格
    {
        Cell *cell = createTestStringCell (0, 0, "   ");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::BLANK);
        EXPECT_EQ (xlsCellObj.asStdString (true), "");
    }

    // 测试零值数字单元格
    {
        Cell *cell = createTestNumberCell (0, 0, 0.0);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
        EXPECT_EQ (xlsCellObj.asStdString (false), "0");
        EXPECT_FALSE (xlsCellObj.asLogical ());
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    }

    // 测试负值数字单元格
    {
        Cell *cell = createTestNumberCell (0, 0, -123.456);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
        EXPECT_EQ (xlsCellObj.asStdString (false), "-123.456");
        EXPECT_TRUE (xlsCellObj.asLogical ()); // 非零值应为true
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), -123.456);
    }

    // 测试非常大的数字
    {
        Cell *cell = createTestNumberCell (0, 0, 1.23456789e15);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
        // 检查字符串表示是否正确
        std::string str = xlsCellObj.asStdString (false);
        EXPECT_TRUE (str.find ("1.23456789") != std::string::npos);
    }

    // 测试非常小的数字
    {
        Cell *cell = createTestNumberCell (0, 0, 1.23456789e-15);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::NUMBER);
        std::string str = xlsCellObj.asStdString (false);
        EXPECT_TRUE (str.find ("1.23456789") != std::string::npos);
    }
}

// 测试日期格式检测
TEST (DateTimeTest, IsDateTimeTest)
{
    // 测试日期格式ID
    EXPECT_TRUE (isDateTime (14)); // 内置日期格式
    EXPECT_TRUE (isDateTime (15)); // 内置日期格式
    EXPECT_TRUE (isDateTime (22)); // 内置日期格式
    EXPECT_TRUE (isDateTime (27)); // 内置日期格式
    EXPECT_TRUE (isDateTime (36)); // 内置日期格式
    EXPECT_TRUE (isDateTime (50)); // 内置日期格式
    EXPECT_TRUE (isDateTime (58)); // 内置日期格式

    // 测试非日期格式ID
    EXPECT_FALSE (isDateTime (0));  // 通用格式
    EXPECT_FALSE (isDateTime (1));  // 0
    EXPECT_FALSE (isDateTime (2));  // 0.00
    EXPECT_FALSE (isDateTime (3));  // #,##0
    EXPECT_FALSE (isDateTime (4));  // #,##0.00
    EXPECT_FALSE (isDateTime (9));  // 0%
    EXPECT_FALSE (isDateTime (10)); // 0.00%
    EXPECT_FALSE (isDateTime (11)); // 0.00E+00
    EXPECT_FALSE (isDateTime (12)); // # ?/?
    EXPECT_FALSE (isDateTime (13)); // # ??/??

    // 测试自定义格式ID（大于等于164）
    EXPECT_FALSE (isDateTime (164)); // 自定义格式，不是日期
}

// 测试地址解析函数
TEST (AddressTest, ParseAddressTest)
{
    // 测试有效地址
    EXPECT_NO_THROW (parseAddress ("A1"));
    EXPECT_NO_THROW (parseAddress ("B2"));
    EXPECT_NO_THROW (parseAddress ("Z26"));
    EXPECT_NO_THROW (parseAddress ("AA100"));
    EXPECT_NO_THROW (parseAddress ("ZZ1000"));

    // 测试无效地址
    EXPECT_THROW (parseAddress (""), std::invalid_argument);
    EXPECT_THROW (parseAddress ("1A"), std::invalid_argument);
    EXPECT_THROW (parseAddress ("A"), std::invalid_argument);
    EXPECT_THROW (parseAddress ("1"), std::invalid_argument);
    EXPECT_THROW (parseAddress ("A1B"), std::invalid_argument);

    // 测试解析结果
    {
        auto result = parseAddress ("A1");
        EXPECT_EQ (result.first, 1);  // 行号
        EXPECT_EQ (result.second, 1); // 列号
    }

    {
        auto result = parseAddress ("B2");
        EXPECT_EQ (result.first, 2);  // 行号
        EXPECT_EQ (result.second, 2); // 列号
    }

    {
        auto result = parseAddress ("AA100");
        EXPECT_EQ (result.first, 100); // 行号
        EXPECT_EQ (result.second, 27); // 列号（A=1, AA=27）
    }
}

// 测试文件验证函数
TEST (FileTest, IsValideTest)
{
    // 注意：这些测试依赖于实际文件系统，可能需要调整
    // 这里主要测试异常抛出

    // 测试不存在的文件
    EXPECT_THROW (isValide ("nonexistent_file.xls"),
                  ExcelReader::FileNotFoundException);

    // 测试目录（假设当前目录存在）
    EXPECT_THROW (isValide ("."), ExcelReader::PathNotFileException);

    // 测试不支持的文件格式
    // 注意：需要创建一个实际的文件来测试
    // 这里只是演示测试结构
}

// 测试tolower函数
TEST (StringTest, ToLowerTest)
{
    EXPECT_EQ (tolower ("HELLO"), "hello");
    EXPECT_EQ (tolower ("Hello World"), "hello world");
    EXPECT_EQ (tolower ("123ABC"), "123abc");
    EXPECT_EQ (tolower (""), "");
    EXPECT_EQ (tolower (" "), " ");
    EXPECT_EQ (tolower ("HELLO123WORLD"), "hello123world");
}

// 测试XlsCell的类型推断
TEST_F (XlsCellTest, TypeInferenceTest)
{
    // 测试字符串单元格的类型推断
    {
        Cell *cell = createTestStringCell (0, 0, "123");
        XlsCell xlsCellObj (cell);
        // 注意：根据XlsCell.h的实现，"123"会被推断为STRING类型，而不是NUMBER
        EXPECT_EQ (xlsCellObj.type (), CellType::STRING);
    }

    // 测试看起来像数字的字符串
    {
        Cell *cell = createTestStringCell (0, 0, "123.456");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::STRING);
    }

    // 测试看起来像布尔的字符串
    {
        Cell *cell = createTestStringCell (0, 0, "true");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::STRING);
    }

    // 测试布尔单元格的类型推断
    {
        Cell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.type (), CellType::BOOL);
    }
}

// 测试XlsCell的asLogical()方法
TEST_F (XlsCellTest, AsLogicalTest)
{
    // 测试各种类型的逻辑值
    {
        Cell *cell = createTestNumberCell (0, 0, 0.0);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestNumberCell (0, 0, 1.0);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestNumberCell (0, 0, -1.0);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestBoolCell (0, 0, false);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestStringCell (0, 0, "anything");
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
    }

    {
        Cell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
    }
}

// 测试XlsCell的asDouble()方法
TEST_F (XlsCellTest, AsDoubleTest)
{
    // 测试各种类型的双精度值
    {
        Cell *cell = createTestNumberCell (0, 0, 123.456);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 123.456);
    }

    {
        Cell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 1.0);
    }

    {
        Cell *cell = createTestBoolCell (0, 0, false);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    }

    {
        Cell *cell = createTestStringCell (0, 0, "hello");
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    }

    {
        Cell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    }
}

int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}
