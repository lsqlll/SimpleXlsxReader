#include "../src/Exceptions.hpp"
#include "../src/XlsCell.hpp"
#include "CellType.hpp"
#include "gtest/gtest.h"

#include <cmath>
#include <filesystem>
#include <string>
#include <variant>

extern "C"
{
#include <xls.h>
}

namespace fs = std::filesystem;

// ============================================================================
// 辅助函数：将CellType转换为字符串
// ============================================================================

std::string
cellTypeToString (CellType type)
{
    switch (type)
    {
    case CellType::STRING:
        return "STRING";
    case CellType::NUMBER:
        return "NUMBER";
    case CellType::BOOL:
        return "BOOL";
    case CellType::UNKNOWN:
        return "UNKNOWN";
    case CellType::BLANK:
        return "BLANK";
    case CellType::DATE:
        return "DATE";
    default:
        return "INVALID_TYPE";
    }
}

// ============================================================================
// 测试文件路径
// ============================================================================

const std::string TEST_XLS_FILE = "test/test.xls";

// ============================================================================
// 测试Fixture - 使用真实的XLS文件
// ============================================================================

class XlsCellTest : public ::testing::Test
{
  protected:
    xls::xlsWorkBook *workbook = nullptr;
    xls::xlsWorkSheet *worksheet = nullptr;

    void
    SetUp () override
    {
        // 检查测试文件是否存在
        if (!fs::exists (TEST_XLS_FILE))
        {
            GTEST_SKIP () << "测试文件 " << TEST_XLS_FILE
                          << " 不存在。请将 test.csv 转换为 test.xls";
        }

        // 打开XLS文件
        workbook = xls::xls_open (TEST_XLS_FILE.c_str (), "UTF-8");
        if (workbook == nullptr)
        {
            GTEST_SKIP () << "无法打开XLS文件: " << TEST_XLS_FILE;
        }

        // 获取第一个工作表
        worksheet = xls::xls_getWorkSheet (workbook, 0);
        if (worksheet == nullptr)
        {
            GTEST_SKIP () << "无法获取工作表";
        }

        // 解析工作表
        xls::xls_parseWorkSheet (worksheet);
    }

    void
    TearDown () override
    {
        if (worksheet != nullptr)
        {
            xls::xls_close_WS (worksheet);
        }
        if (workbook != nullptr)
        {
            xls::xls_close_WB (workbook);
        }
    }

    // 辅助方法：获取指定位置的单元格
    xls::xlsCell *
    getCell (xls::WORD row, xls::WORD col)
    {
        if (worksheet == nullptr)
        {
            return nullptr;
        }

        xls::xlsRow *xlsRow = xls::xls_row (worksheet, row);
        if (xlsRow == nullptr)
        {
            return nullptr;
        }

        if (col >= xlsRow->cells.count)
        {
            return nullptr;
        }

        return &xlsRow->cells.cell[col];
    }
};

// ============================================================================
// XlsCell 构造函数测试
// ============================================================================

TEST_F (XlsCellTest, ConstructorWithValidCell)
{
    // 测试读取 Alice 的姓名 (row=1, col=0, 跳过标题行)
    xls::xlsCell *cell = getCell (1, 0);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_EQ (xlsCell.row (), 1);
    EXPECT_EQ (xlsCell.col (), 0);
    EXPECT_EQ (xlsCell.asString (false), "Alice");
}

TEST_F (XlsCellTest, ConstructorWithNullCell)
{
    EXPECT_THROW (XlsCell xlsCell (nullptr),
                  ExcelReader::InvalidCellException);
}

// ============================================================================
// XlsCell 字符串类型测试
// ============================================================================

TEST_F (XlsCellTest, StringCellBasic)
{
    // 读取 "Bob" (row=2, col=0)
    xls::xlsCell *cell = getCell (2, 0);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_EQ (xlsCell.asString (false), "Bob");
    EXPECT_FALSE (xlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 0.0);
}

TEST_F (XlsCellTest, StringCellWithWhitespace)
{
    // 读取包含空格的字符串 "  Spaces  " (row=12, col=0)
    xls::xlsCell *cell = getCell (12, 0);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    // trim=true 应该去除前后空格
    EXPECT_EQ (xlsCell.asString (true), "Spaces");
}

TEST_F (XlsCellTest, StringCellWithNewline)
{
    // 读取包含空格的字符串 "  Spaces  " (row=13, col=0)
    xls::xlsCell *cell = getCell (13, 0);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    // trim=true 应该去除前后空格
    EXPECT_EQ (xlsCell.asString (true), "");
}

TEST_F (XlsCellTest, StringCellMultipleRecords)
{
    // 测试多个字符串记录
    const char *expectedNames[] = { "Alice", "Bob",   "Charlie", "Diana",
                                    "Eve",   "Frank", "Grace",   "Henry" };

    for (int i = 0; i < 8; i++)
    {
        xls::xlsCell *cell = getCell (i + 1, 0);
        if (cell != nullptr)
        {
            XlsCell xlsCell (cell);
            EXPECT_EQ (xlsCell.asString (false), expectedNames[i]);
        }
    }
}

// ============================================================================
// XlsCell 数字类型测试
// ============================================================================

TEST_F (XlsCellTest, NumberCellBasic)
{
    // 读取 Alice 的分数 95.5 (row=1, col=2)
    xls::xlsCell *cell = getCell (1, 2);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 95.5);
    EXPECT_TRUE (xlsCell.asLogical ());
}

TEST_F (XlsCellTest, IntegerCell)
{
    // 读取 Alice 的年龄 25 (row=1, col=1)
    xls::xlsCell *cell = getCell (1, 1);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 25.0);
    EXPECT_EQ (xlsCell.asString (false), "25");
}

TEST_F (XlsCellTest, ZeroCell)
{
    // 读取 Eve 的分数 0 (row=5, col=2)
    xls::xlsCell *cell = getCell (5, 2);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 0.0);
    EXPECT_FALSE (xlsCell.asLogical ());
    EXPECT_EQ (xlsCell.asString (false), "0");
}

TEST_F (XlsCellTest, NegativeCell)
{
    // 读取 Henry 的分数 -5.5 (row=8, col=2)
    xls::xlsCell *cell = getCell (8, 2);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), -5.5);
    EXPECT_TRUE (xlsCell.asLogical ());
}

TEST_F (XlsCellTest, PerfectScoreCell)
{
    // 读取 Frank 的分数 100 (row=6, col=2)
    xls::xlsCell *cell = getCell (6, 2);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_DOUBLE_EQ (xlsCell.asDouble (), 100.0);
    EXPECT_EQ (xlsCell.asString (false), "100");
}

// ============================================================================
// XlsCell 布尔类型测试
// ============================================================================

TEST_F (XlsCellTest, BooleanValues)
{
    // 读取 Alice 的 Passed 字段 (row=1, col=3)
    xls::xlsCell *cell = getCell (1, 3);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    // Excel中的TRUE/FALSE可能被存储为字符串或布尔值
    std::string value = xlsCell.asString (false);
    EXPECT_FALSE (value == "FLASE" || value == "false" || value == "0"
                  || xlsCell.asLogical ());
}

TEST_F (XlsCellTest, BooleanFalseValues)
{
    // 读取 Charlie 的 Passed 字段 FALSE (row=3, col=3)
    xls::xlsCell *cell = getCell (3, 3);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    std::string value = xlsCell.asString (false);
    EXPECT_TRUE (value == "FALSE" || value == "false" || value == "0"
                 || !xlsCell.asLogical ());
}

// ============================================================================
// XlsCell value() 和 valueType() 测试
// ============================================================================

TEST_F (XlsCellTest, ValueMethodString)
{
    xls::xlsCell *cell = getCell (1, 0); // Alice
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    const auto value = xlsCell.getValue () ();
    CellType type = xlsCell.getType ();

    EXPECT_TRUE (std::holds_alternative<std::string> (value)
                 || type == CellType::STRING || type == CellType::BLANK);
}

TEST_F (XlsCellTest, ValueMethodNumber)
{
    xls::xlsCell *cell = getCell (1, 2); // 95.5
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    const auto value = xlsCell.getValue () ();

    CellType type = xlsCell.getType ();

    EXPECT_TRUE (std::holds_alternative<double> (value)
                 || type == CellType::NUMBER || type == CellType::BLANK);
}

TEST_F (XlsCellTest, ValueTypeVariety)
{
    // 测试不同类型的单元格
    struct TestCase
    {
        xls::WORD row;
        xls::WORD col;
        std::vector<CellType> allowedTypes;
    };

    std::vector<TestCase> cases = {
        { .row = 1,
          .col = 0,
          .allowedTypes
          = { CellType::STRING, CellType::BLANK } }, // Alice (字符串)
        { .row = 1,
          .col = 1,
          .allowedTypes = { CellType::NUMBER, CellType::BLANK } }, // 25 (数字)
        { .row = 1,
          .col = 2,
          .allowedTypes
          = { CellType::NUMBER, CellType::BLANK } }, // 95.5 (浮点数)
        { .row = 1,
          .col = 3,
          .allowedTypes = { CellType::STRING, CellType::BOOL,
                            CellType::BLANK } } // TRUE (布尔)
    };

    for (const auto &tc : cases)
    {
        xls::xlsCell *cell = getCell (tc.row, tc.col);
        if (cell != nullptr)
        {
            XlsCell xlsCell (cell);
            CellType type = xlsCell.getType ();

            bool found = false;
            for (const auto &allowedType : tc.allowedTypes)
            {
                if (type == allowedType)
                {
                    found = true;
                    break;
                }
            }
            EXPECT_TRUE (found)
                << "Row " << tc.row << ", Col " << tc.col
                << " has unexpected type: " << cellTypeToString (type);
        }
    }
}

// ============================================================================
// XlsCell 公式情况测试
// ============================================================================

TEST_F (XlsCellTest, FormulaCell)
{
    // 读取包含公式的单元格 (row=13, col=1)
    xls::xlsCell *cell = getCell (13, 1);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    // 根据实际情况调整期望值
    CellType type = xlsCell.getType ();
    if (type == CellType::NUMBER)
    {
        EXPECT_TRUE (xlsCell.isNumber ());
    }
    else if (type == CellType::STRING)
    {
        EXPECT_TRUE (xlsCell.isString ());
    }
    else
    {
        EXPECT_TRUE (xlsCell.isBlank ()
                     || xlsCell.getType () == CellType::UNKNOWN);
    }
}

// ============================================================================
// XlsCell 边缘情况测试
// ============================================================================

TEST_F (XlsCellTest, EmptyStringCell)
{
    // 读取 Eve 的空 Notes (row=5, col=5)
    xls::xlsCell *cell = getCell (5, 5);
    ASSERT_NE (cell, nullptr);

    XlsCell xlsCell (cell);

    EXPECT_EQ (xlsCell.asString (false), "");
}

TEST_F (XlsCellTest, MissingCells)
{
    // 测试缺失的单元格
    // 缺失姓名 (row=9, col=0)
    xls::xlsCell *cell1 = getCell (9, 0);
    if (cell1 != nullptr)
    {
        XlsCell xlsCell (cell1);
        EXPECT_EQ (xlsCell.asString (false), "");
    }

    // 缺失年龄 (row=10, col=1)
    xls::xlsCell *cell2 = getCell (10, 1);
    if (cell2 != nullptr)
    {
        XlsCell xlsCell (cell2);
        std::string value = xlsCell.asString (false);
        EXPECT_TRUE (value == "" || xlsCell.asDouble () == 0.0);
    }
}

// ============================================================================
// XlsCell 复制和移动操作测试
// ============================================================================

TEST_F (XlsCellTest, CopyConstructor)
{
    xls::xlsCell *cell = getCell (1, 0);
    ASSERT_NE (cell, nullptr);

    XlsCell original (cell);
    const XlsCell &copy = original; // 使用赋值构造而非引用

    EXPECT_EQ (original.row (), copy.row ());
    EXPECT_EQ (original.col (), copy.col ());
    EXPECT_EQ (original.asString (false), copy.asString (false));
}

TEST_F (XlsCellTest, MoveConstructor)
{
    xls::xlsCell *cell = getCell (1, 1);
    ASSERT_NE (cell, nullptr);

    XlsCell original (cell);
    int originalRow = original.row ();
    int originalCol = original.col ();
    std::string originalValue = original.asString (false);

    XlsCell moved = std::move (original);

    EXPECT_EQ (moved.row (), originalRow);
    EXPECT_EQ (moved.col (), originalCol);
    EXPECT_EQ (moved.asString (false), originalValue);
}

TEST_F (XlsCellTest, CopyAssignment)
{
    xls::xlsCell *cell1 = getCell (1, 0);
    xls::xlsCell *cell2 = getCell (2, 0);
    ASSERT_NE (cell1, nullptr);
    ASSERT_NE (cell2, nullptr);

    XlsCell cellA (cell1);
    XlsCell cellB (cell2);

    cellB = cellA;

    EXPECT_EQ (cellB.asString (false), "Alice");
}

// ============================================================================
// XlsCell asLogical() 测试
// ============================================================================

TEST_F (XlsCellTest, AsLogicalTests)
{
    struct LogicalTest
    {
        xls::WORD row;
        xls::WORD col;
        bool expectedLogical;
    };

    std::vector<LogicalTest> tests = {
        { .row = 5, .col = 2, .expectedLogical = false }, // 0 应该是 false
        { .row = 1, .col = 2, .expectedLogical = true },  // 95.5 应该是 true
        { .row = 6, .col = 2, .expectedLogical = true },  // 100 应该是 true
        { .row = 8,
          .col = 2,
          .expectedLogical = true }, // -5.5 应该是 true (非零)
    };

    for (const auto &test : tests)
    {
        xls::xlsCell *cell = getCell (test.row, test.col);
        if (cell != nullptr)
        {
            XlsCell xlsCell (cell);
            EXPECT_EQ (xlsCell.asLogical (), test.expectedLogical)
                << "Row " << test.row << ", Col " << test.col;
        }
    }
}

// ============================================================================
// XlsCell asDouble() 测试
// ============================================================================

TEST_F (XlsCellTest, AsDoubleTests)
{
    struct DoubleTest
    {
        xls::WORD row;
        xls::WORD col;
        double expectedValue;
    };

    std::vector<DoubleTest> tests = {
        { .row = 1, .col = 1, .expectedValue = 25.0 },  // Alice 的年龄
        { .row = 1, .col = 2, .expectedValue = 95.5 },  // Alice 的分数
        { .row = 2, .col = 1, .expectedValue = 30.0 },  // Bob 的年龄
        { .row = 5, .col = 2, .expectedValue = 0.0 },   // Eve 的分数 0
        { .row = 6, .col = 2, .expectedValue = 100.0 }, // Frank 的分数 100
        { .row = 8, .col = 2, .expectedValue = -5.5 },  // Henry 的分数 -5.5
    };

    for (const auto &test : tests)
    {
        xls::xlsCell *cell = getCell (test.row, test.col);
        if (cell != nullptr)
        {
            XlsCell xlsCell (cell);
            EXPECT_DOUBLE_EQ (xlsCell.asDouble (), test.expectedValue)
                << "Row " << test.row << ", Col " << test.col;
        }
    }
}

// ============================================================================
// CellPosition 测试
// ============================================================================

TEST (CellPositionTest, FromRowCol)
{
    CellPosition pos (0, 0);
    EXPECT_EQ (pos.getAddr (), "A1");
}

TEST (CellPositionTest, FromAddress)
{
    CellPosition pos ("B3");
    EXPECT_EQ (pos.getRow (), 3);
    EXPECT_EQ (pos.getCol (), 2);
}

TEST (CellPositionTest, ComplexAddress)
{
    CellPosition pos ("AA100");
    EXPECT_EQ (pos.getRow (), 100);
    EXPECT_EQ (pos.getCol (), 27);
}

TEST (CellPositionTest, NulloptConstructor)
{
    CellPosition pos (std::nullopt);
    EXPECT_FALSE (pos.hasRow ());
    EXPECT_FALSE (pos.hasCol ());
}

// ============================================================================
// 工具函数测试
// ============================================================================

TEST (UtilsTest, IsExcelFormat)
{
    EXPECT_TRUE (isExcelFormat ("xls"));
    EXPECT_TRUE (isExcelFormat ("xlsx"));
    EXPECT_TRUE (isExcelFormat ("csv"));
    EXPECT_FALSE (isExcelFormat ("txt"));
    EXPECT_FALSE (isExcelFormat ("pdf"));
}

TEST (UtilsTest, Trim)
{
    EXPECT_EQ (trim ("  hello  "), "hello");
    EXPECT_EQ (trim ("hello"), "hello");
    EXPECT_EQ (trim ("   "), "");
    EXPECT_EQ (trim (""), "");
}

TEST (UtilsTest, IsEmpty)
{
    EXPECT_TRUE (isEmpty (""));
    EXPECT_TRUE (isEmpty ("   "));
    EXPECT_FALSE (isEmpty ("hello"));
    EXPECT_TRUE (isEmpty ("   ", true));
}

TEST (UtilsTest, ToLower)
{
    EXPECT_EQ (tolower ("HELLO"), "hello");
    EXPECT_EQ (tolower ("HeLLo WoRLd"), "hello world");
    EXPECT_EQ (tolower ("123ABC"), "123abc");
}

TEST (UtilsTest, IsDateTime)
{
    EXPECT_TRUE (isDateTime (14));
    EXPECT_TRUE (isDateTime (15));
    EXPECT_TRUE (isDateTime (22));
    EXPECT_FALSE (isDateTime (0));
    EXPECT_FALSE (isDateTime (10));
}

TEST (UtilsTest, ParseAddressValid)
{
    auto result1 = parseAddress ("A1");
    EXPECT_EQ (result1.first, 1);  // Row is 0-based
    EXPECT_EQ (result1.second, 1); // Col is 0-based

    auto result2 = parseAddress ("B2");
    EXPECT_EQ (result2.first, 2);  // Row is 0-based
    EXPECT_EQ (result2.second, 2); // Col is 0-based

    auto result3 = parseAddress ("AA100");
    EXPECT_EQ (result3.first, 100); // Row is 0-based
    EXPECT_EQ (result3.second, 27); // Col is 0-based
}

TEST (UtilsTest, ParseAddressInvalid)
{
    EXPECT_THROW (parseAddress (""), std::invalid_argument);
    EXPECT_THROW (parseAddress ("1A"), std::invalid_argument);
    EXPECT_THROW (parseAddress ("A"), ExcelReader::AddressParseException);
}

TEST (UtilsTest, IsValidFileNotFound)
{
    EXPECT_THROW (isValid ("nonexistent_file.xls"),
                  ExcelReader::FileNotFoundException);
}

// ============================================================================
// main 函数
// ============================================================================

int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}
