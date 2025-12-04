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
    inferTypeFromStringCell (bool trimWs)
    {
        std::string s = cell_->str == nullptr ? "" : std::string (cell_->str);
        if (isEmpty (s, trimWs))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        else
            {
                type_ = CellType::STRING;
                value_ = trimWs ? trim (s) : s;
            }
    };
    void
    inferTypeFromFormulaCell (bool trimWs)
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
        if (isEmpty (str, trimWs))
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
        else
            {
                type_ = CellType::STRING;
                value_ = trimWs ? trim (str) : str;
            }
    };
    void
    inferTypeFromNumberCell (bool trimWs)
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
        if (cell_->str && strncmp ((char*)cell_->str, "bool", 4) == 0)
            {
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
            }
        else
            {
                type_ = CellType::BLANK;
                value_ = std::monostate{};
            }
    };
    void
    inferBlankCell ()
    {
        type_ = CellType::BLANK;
        value_ = std::monostate{};
    };
    void
    inferUnknownCell ()
    {
        type_ = CellType::UNKNOWN;
        value_ = std::monostate{};
    };

public:
    explicit XlsCell (xls::xlsCell* cell)
        : cell_ (cell)
        , location_ (CellPosition (std::make_pair (cell_->row, cell_->col)))
        , type_ (CellType::UNKNOWN)
        , value_ (std::monostate{})
    {
        if (cell)
            {
                inferType (true);
            }
    };

    explicit XlsCell (std::pair<int, int> loc)
        : cell_ (NULL)
        , location_ (loc)
        , type_ (CellType::BLANK)
        , value_ (std::monostate{}) {};

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
        if (type_.has_value ())
            {
                return type_.value ();
            }
        inferType (true);
        return type_.value ();
    }

    void
    inferType (const bool trimWs) // 添加 const 修饰符
    {
        if (type_ != CellType::UNKNOWN)
            {
                return;
            }

        CellType ct;
        switch (cell_->id)
            {
                case XLS_RECORD_LABELSST:
                case XLS_RECORD_LABEL:
                case XLS_RECORD_RSTRING:
                    {
                        std::string s = cell_->str == NULL ? "" : cell_->str;
                        if (isEmpty (s, trimWs))
                            {
                                ct = CellType::BLANK;
                                value_ = std::monostate{};
                            }
                        else
                            {
                                ct = CellType::STRING;
                                value_ = trimWs ? trim (s) : s;
                            }
                        break;
                    }

                case XLS_RECORD_FORMULA:
                case XLS_RECORD_FORMULA_ALT:
                    if (cell_->l == 0)
                        {
                            auto strVal = std::to_string (cell_->d);

                            if (isEmpty (strVal))
                                {
                                    ct = CellType::BLANK;
                                    value_ = std::monostate{};
                                    break;
                                }
                            int format = cell_->xf;
                            ct = (isDateTime (format)) ? CellType::DATE
                                                       : CellType::NUMBER;
                            value_ = cell_->d;
                            break;
                        }
                    else
                        {
                            if (strncmp ((char*)cell_->str, "bool", 4) == 0)
                                {
                                    if ((cell_->d == 0
                                         && strcasecmp ((char*)cell_->str,
                                                        "false")
                                                == 0)
                                        || (cell_->d == 1
                                            && strcasecmp ((char*)cell_->str,
                                                           "true")
                                                   == 0))
                                        {
                                            ct = CellType::BLANK;
                                            value_ = std::monostate{};
                                        }
                                    else
                                        {
                                            ct = CellType::BOOL;
                                            value_ = cell_->d != 0;
                                        }
                                    break;
                                }

                            if (strncmp ((char*)cell_->str, "error", 5) == 0
                                && cell_->d > 0)
                                {
                                    ct = CellType::BLANK;
                                    value_ = std::monostate{};
                                    break;
                                }

                            std::string str = cell_->str;
                            if (isEmpty (str, trimWs))
                                {
                                    ct = CellType::BLANK;
                                    value_ = std::monostate{};
                                }
                            else
                                {
                                    ct = CellType::STRING;
                                    value_ = trimWs ? trim (str) : str;
                                }
                        }
                    break;

                case XLS_RECORD_MULRK:
                case XLS_RECORD_NUMBER:
                case XLS_RECORD_RK:
                    {
                        if (isEmpty (std::to_string (cell_->col)))
                            {
                                ct = CellType::BLANK;
                                value_ = std::monostate{};
                                break;
                            }
                        int format = cell_->xf;
                        ct = (isDateTime (format)) ? CellType::DATE
                                                   : CellType::NUMBER;
                        value_ = cell_->d;
                    }
                    break;

                case XLS_RECORD_MULBLANK:
                case XLS_RECORD_BLANK:
                    ct = CellType::BLANK;
                    value_ = std::monostate{};
                    break;

                case XLS_RECORD_BOOLERR:
                    if (strncmp ((char*)cell_->str, "bool", 4) == 0)
                        {
                            if ((cell_->d == 0
                                 && tolower (std::string (cell_->str)
                                             == "false"))
                                || (cell_->d == 1
                                    && tolower (std::string (cell_->str)
                                                == "false")))
                                {
                                    ct = CellType::BLANK;
                                    value_ = std::monostate{};
                                }
                            else
                                {
                                    ct = CellType::BOOL;
                                    value_ = cell_->d != 0;
                                }
                        }
                    else
                        {
                            ct = CellType::BLANK;
                            value_ = std::monostate{};
                        }
                    break;

                default:
                    ct = CellType::UNKNOWN;
                    value_ = std::monostate{};
            }

        type_ = ct; // 修改 type_ 的值需要移除 const 修饰符或使用 mutable
    }

    std::string
    asStdString (const bool trimWs,
                 const std::vector<std::string>& stringTable) const
    {
        switch (type_.value ())
            {
                case CellType::UNKNOWN:
                case CellType::BLANK:
                    return "";

                case CellType::BOOL:
                    {
                        if (auto* b = std::get_if<bool> (&value_))
                            {
                                return *b ? "TRUE" : "FALSE";
                            }
                        return cell_->d ? "TRUE" : "FALSE";
                    }

                case CellType::DATE:
                case CellType::NUMBER:
                    {
                        if (auto* d = std::get_if<double> (&value_))
                            {
                                std::ostringstream strs;
                                double intpart;
                                if (std::modf (*d, &intpart) == 0.0)
                                    {
                                        strs << std::fixed
                                             << static_cast<int64_t> (*d);
                                    }
                                else
                                    {
                                        strs << std::setprecision (
                                            std::numeric_limits<
                                                double>::digits10
                                            + 2)
                                             << *d;
                                    }
                                std::string out_string = strs.str ();
                                return out_string;
                            }
                        // fallback to original
                        std::ostringstream strs;
                        double intpart;
                        if (std::modf (cell_->d, &intpart) == 0.0)
                            {
                                strs << std::fixed
                                     << static_cast<int64_t> (cell_->d);
                            }
                        else
                            {
                                strs << std::setprecision (
                                    std::numeric_limits<double>::digits10 + 2)
                                     << cell_->d;
                            }
                        std::string out_string = strs.str ();
                        return out_string;
                    }

                case CellType::STRING:
                    {
                        if (auto* s = std::get_if<std::string> (&value_))
                            {
                                return trimWs ? trim (*s) : *s;
                            }
                        std::string out_string = cell_->str;
                        return trimWs ? trim (out_string) : out_string;
                    }

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
