#pragma once

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <iomanip>
#include <limits>
#include <memory>
#include <optional>
#include <sstream>
#include <string>
#include <string_view>
#include <variant>

#include "utils.h"

extern "C"
{
#include "xls.h"
}

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

    template <typename T> explicit CellPosition (T row, T col)
    {
        if constexpr (std::is_same_v<T, std::nullopt_t>)
        {
            row = std::nullopt;
            col = std::nullopt;
            addr = std::nullopt;
        }
        else
        {
            row = static_cast<std::size_t> (row);
            col = static_cast<std::size_t> (col);
            addr = std::nullopt;
        }
        calculateExcelAddress ();
    }

    template <typename T> explicit CellPosition (std::pair<T, T> loc)
    {
        if constexpr (std::is_same_v<T, std::nullopt_t>)
        {
            row = std::nullopt;
            col = std::nullopt;
            addr = std::nullopt;
        }
        else
        {
            row = static_cast<std::size_t> (loc.first);
            col = static_cast<std::size_t> (loc.second);
            addr = std::nullopt;
        }
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
};

class XlsCell
{
  private:
    xls::xlsCell *cell_{};
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

    static std::shared_ptr<xls::xlsCell>
    create (xls::xlsCell *cell)
    {
        if (cell == nullptr)
        {
            return nullptr;
        }
        try
        {
            return std::make_shared<xls::xlsCell> (cell);
        }
        catch (...)
        {
            return nullptr;
        }
    }
    void
    initialize (xls::xlsCell *cell)
    {
        if (cell == nullptr)
        {
            throw ExcelReader::NullCellException ("");
        }

        // 先设置 cell_ 指针
        cell_ = cell;
        location_ = CellPosition (cell->row, cell->col);
        inferValue (false);
    }

  public:
    XlsCell (const XlsCell &) = default;
    XlsCell (XlsCell &&) = default;
    XlsCell &operator= (const XlsCell &) = default;
    XlsCell &operator= (XlsCell &&) = default;
    explicit XlsCell (xls::xlsCell *cell)
        : location_ (CellPosition (std::nullopt, std::nullopt))
    {
        initialize (cell);
    }

    [[nodiscard]] int
    row () const
    {
        return location_.row.value ();
    }

    [[nodiscard]] int
    col () const
    {
        return location_.col.value ();
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
    [[nodiscard]] const std::variant<std::monostate, std::string, double,
                                     bool> &
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