#pragma once

#include <algorithm>
#include <cstdint>
#include <optional>
#include <string>

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
        if (row.has_value () || col.has_value ())
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

    [[nodiscard]] std::size_t
    getRow () const
    {
        if (this->row.has_value ())
        {
            return this->row.value ();
        }
        return -1;
    }

    [[nodiscard]] std::size_t
    getCol () const
    {
        if (this->col.has_value ())
        {
            return this->col.value ();
        }
        return -1;
    }
    [[nodiscard]] std::string
    getAddr () const
    {
        if (this->addr.has_value ())
        {
            return this->addr.value ();
        }
        return "";
    }

    ~CellPosition () = default;

  private:
    void
    calculateExcelAddress ()
    {
        if (!row.has_value () && !col.has_value ())
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
            return static_cast<std::size_t> (value);
        }

        return std::nullopt;
    }
};
