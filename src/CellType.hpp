#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>

#include "utils.hpp"

// 基础整数概念
template <typename T = std::size_t> concept IntegralIndex = requires (T t)
{
    requires std::integral<T>;
    requires !std::same_as<T, bool>;
    requires !std::same_as<T, char>;
    { t >= 0 }->std::convertible_to<bool>; // 确保可以比较
    { t + 1 }->std::same_as<T>;            // 确保可以做算术
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

class CellPosition
{
  private:
    std::optional<std::size_t> row;
    std::optional<std::size_t> col;
    std::optional<std::string> addr; // Like A1

    void
    calculateExcelAddress ()
    {
        if (!row.has_value () && !col.has_value ())
        {
            addr = std::nullopt;
            return;
        }

        std::string colPart;
        std::size_t colNum{ 0 };

        while (true)
        {
            colPart += static_cast<char> ('A' + (colNum % 26));
            if (colNum < 26)
            {
                break;
            }
            colNum = colNum / 26 - 1;
        }
        std::ranges::reverse (colPart);
        this->addr = colPart + std::to_string (row.value () + 1);
    }

  public:
    explicit CellPosition (std::nullopt_t /*unused*/)
        : row (std::nullopt), col (std::nullopt), addr (std::nullopt) {};

    explicit CellPosition (std::size_t row, std::size_t col)
        : row (row), col (col)
    {
        calculateExcelAddress ();
    }

    template <IntegralIndex T1, IntegralIndex T2>
    explicit CellPosition (T1 row, T2 col) : row (row), col (col)
    {
        calculateExcelAddress ();
    }

    template <IntegralIndex T1, IntegralIndex T2>
    explicit CellPosition (std::pair<T1, T2> loc)
        : row (loc.first), col (loc.second), addr (std::nullopt)
    {
        calculateExcelAddress ();
    }

    explicit CellPosition (const std::string &addr) : addr (addr)
    {
        // 解析地址 A1 -> (0,0), B2 -> (1,1)
        auto [rowNum, colNum] = parseAddress (addr);
        this->row = rowNum;
        this->col = colNum;
    };

    [[nodiscard]] constexpr std::size_t
    getRow () const
    {
        return this->row.value_or (-1);
    }

    [[nodiscard]] constexpr std::size_t
    getCol () const
    {
        return this->col.value_or (-1);
    }

    [[nodiscard]] std::string
    getAddr () const
    {
        return this->addr.value_or ("");
    }

    [[nodiscard]] bool
    hasRow () const
    {
        return this->row.has_value ();
    }

    [[nodiscard]] bool
    hasCol () const
    {
        return this->col.has_value ();
    }

    ~CellPosition () = default;

  private:
};
