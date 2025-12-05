#ifndef TYPE_H
#define TYPE_H

#include "utils.h"

#include <cctype>
#include <cmath>
#include <cstring>
#include <iomanip>
#include <limits>
#include <optional>
#include <sstream>
#include <string>
#include <variant>
#include <vector>

extern "C"
{
#include "xls.h"
}

enum class CellType
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
    xls::xlsCell* cell_;
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
    inferValueFromFormulaCell ()
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
        std::string strVal = cell_->str ? std::string (cell_->str) : "";
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
    explicit XlsCell (xls::xlsCell* cell)
        : cell_ (cell)
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
        throw ExcelReader::NullCellException ("");
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

using XLSheets = std::vector<xls::xlsWorkSheet*>;
using XLSheetsName = std::vector<std::string>;

const auto xlsDeleter = [] (xls::xlsWorkBook* wb)
    {
        if (wb)
            xls::xls_close_WB (wb);
    };

#endif
