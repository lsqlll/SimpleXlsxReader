#include "../src/Exceptions.h"
#include "../src/XlsCell.h"
#include "gtest/gtest.h"

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstring>
#include <filesystem>
#include <optional>
#include <string>
#include <variant>

namespace fs = std::filesystem;

inline xls::xlsCell *
createXlsCell (int row, int col, int id, int xf, int l, double d,
               const char *str)
{
    xls::xlsCell *cell = new xls::xlsCell ();
    cell->row = row;
    cell->col = col;
    cell->id = id;
    cell->xf = xf;
    cell->l = l;
    cell->d = d;
    if (str != nullptr)
    {
        cell->str = new char[strlen (str) + 1];
        strcpy (cell->str, str);
    }
    else
    {
        cell->str = nullptr;
    }
    return cell;
}

inline void
freeXlsCell (xls::xlsCell *cell)
{
    if (cell != nullptr)
    {
        if (cell->str != nullptr)
        {
            delete[] cell->str;
        }
        delete cell;
    }
}

xls::xlsCell *
createStringCell (int row, int col, const std::string &content)
{
    return createXlsCell (row, col, XLS_RECORD_LABEL, 0, 0, 0.0,
                          content.c_str ());
}

xls::xlsCell *
createNumberCell (int row, int col, double value)
{
    return createXlsCell (row, col, XLS_RECORD_NUMBER, 0, 0, value, nullptr);
}

xls::xlsCell *
createBoolCell (int row, int col, bool value)
{
    const char *str = value ? "true" : "false";
    return createXlsCell (row, col, XLS_RECORD_BOOLERR, 0, 0,
                          value ? 1.0 : 0.0, str);
}

xls::xlsCell *
createBlankCell (int row, int col)
{
    return createXlsCell (row, col, XLS_RECORD_BLANK, 0, 0, 0.0, nullptr);
}

xls::xlsCell *
createFormulaCell (int row, int col, double value,
                   const std::string &formula = "")
{
    return createXlsCell (row, col, XLS_RECORD_FORMULA, 0, formula.length (),
                          value,
                          formula.empty () ? nullptr : formula.c_str ());
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
    static xls::xlsCell *
    createTestStringCell (int row, int col, const std::string &content)
    {
        return createStringCell (row, col, content);
    }

    static xls::xlsCell *
    createTestNumberCell (int row, int col, double value)
    {
        return createNumberCell (row, col, value);
    }

    static xls::xlsCell *
    createTestBoolCell (int row, int col, bool value)
    {
        return createBoolCell (row, col, value);
    }

    static xls::xlsCell *
    createTestBlankCell (int row, int col)
    {
        return createBlankCell (row, col);
    }

    static xls::xlsCell *
    createTestFormulaCell (int row, int col, double value,
                           const std::string &formula = "")
    {
        return createFormulaCell (row, col, value, formula);
    }
};

// 测试构造函数
TEST_F (XlsCellTest, ConstructorWithValidCell)
{
    // Given: 一个有效的字符串单元格
    xls::xlsCell *cell = createTestStringCell (0, 0, "Hello World");

    // When: 创建XlsCell对象
    XlsCell xlsCellObj (cell);

    // Then: 应该正确初始化
    EXPECT_EQ (xlsCellObj.row (), 0);
    EXPECT_EQ (xlsCellObj.col (), 0);

    // Cleanup
    freeXlsCell (cell);
}

TEST_F (XlsCellTest, ConstructorWithNullCell)
{
    // When: 使用nullptr创建XlsCell
    // 应当抛出InvalidCellException
    EXPECT_THROW (XlsCell xlsCellObj (nullptr),
                  ExcelReader::InvalidCellException);
}

// 测试字符串单元格
TEST_F (XlsCellTest, StringCellHandling)
{
    // Given: 一个包含内容的字符串单元格
    xls::xlsCell *cell = createTestStringCell (0, 0, "Test Content");
    XlsCell xlsCellObj (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCellObj.asStdString (false), "Test Content");
    EXPECT_FALSE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    EXPECT_EQ (xlsCellObj.valueType (), "string");

    // Cleanup
    freeXlsCell (cell);
}

// 测试带空格的字符串单元格
TEST_F (XlsCellTest, StringCellWithWhitespace)
{
    // Given: 一个包含前后空格的字符串单元格
    xls::xlsCell *cell = createTestStringCell (0, 0, "  Test Content  ");
    XlsCell xlsCellObj (cell);

    EXPECT_EQ (xlsCellObj.asStdString (true), "Test Content"); // trim=true
    EXPECT_EQ (xlsCellObj.asStdString (false),
               "  Test Content  "); // trim=false

    // Cleanup
    freeXlsCell (cell);
}

// 测试数字单元格
TEST_F (XlsCellTest, NumberCellHandling)
{
    // Given: 一个数字单元格
    xls::xlsCell *cell = createTestNumberCell (0, 0, 123.456);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查各种转换
    EXPECT_EQ (xlsCellObj.asStdString (false), "123.456");
    EXPECT_TRUE (xlsCellObj.asLogical ()); // 非零数字应为true
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 123.456);
    EXPECT_EQ (xlsCellObj.valueType (), "double");

    // Cleanup
    freeXlsCell (cell);
}

// 测试整数单元格
TEST_F (XlsCellTest, IntegerCellHandling)
{
    // Given: 一个整数单元格
    xls::xlsCell *cell = createTestNumberCell (0, 0, 42.0);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查字符串转换是否为整数形式
    EXPECT_EQ (xlsCellObj.asStdString (true), "42");

    // Cleanup
    freeXlsCell (cell);
}

// 测试布尔单元格
TEST_F (XlsCellTest, BoolCellHandling)
{
    // Given: 一个TRUE布尔单元格
    xls::xlsCell *trueCell = createTestBoolCell (0, 0, true);
    XlsCell trueXlsCell (trueCell);

    // When & Then: 检查TRUE值的各种转换
    EXPECT_TRUE (trueXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (trueXlsCell.asDouble (), 1.0);

    // Given: 一个FALSE布尔单元格
    xls::xlsCell *falseCell = createTestBoolCell (0, 0, false);
    XlsCell falseXlsCell (falseCell);

    // When & Then: 检查FALSE值的各种转换
    EXPECT_FALSE (falseXlsCell.asLogical ());
    EXPECT_DOUBLE_EQ (falseXlsCell.asDouble (), 0.0);

    // Cleanup
    freeXlsCell (trueCell);
    freeXlsCell (falseCell);
}

// 测试空白单元格
TEST_F (XlsCellTest, BlankCellHandling)
{
    // Given: 一个空白单元格
    xls::xlsCell *cell = createTestBlankCell (0, 0);
    XlsCell xlsCellObj (cell);

    // When & Then: 检查空白单元格的行为
    EXPECT_EQ (xlsCellObj.asStdString (false), "");
    EXPECT_FALSE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
    EXPECT_EQ (xlsCellObj.valueType (), "monostate");

    // Cleanup
    freeXlsCell (cell);
}

// 测试公式单元格
TEST_F (XlsCellTest, FormulaCellHandling)
{
    // Given: 一个公式单元格
    xls::xlsCell *cell = createTestFormulaCell (0, 0, 100.0, "SUM(A1:A10)");
    XlsCell xlsCellObj (cell);

    // When & Then: 检查公式单元格的行为
    EXPECT_EQ (xlsCellObj.asStdString (false), "100");
    EXPECT_TRUE (xlsCellObj.asLogical ());
    EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 100.0);

    // Cleanup
    freeXlsCell (cell);
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
        xls::xlsCell *cell = createTestStringCell (0, 0, "Test");
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<std::string> (value));
        EXPECT_EQ (std::get<std::string> (value), "Test");
        freeXlsCell (cell);
    }

    // 测试数字单元格的value()
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 123.45);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<double> (value));
        EXPECT_DOUBLE_EQ (std::get<double> (value), 123.45);
        freeXlsCell (cell);
    }

    // 测试布尔单元格的value()
    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<bool> (value));
        EXPECT_TRUE (std::get<bool> (value));
        freeXlsCell (cell);
    }

    // 测试空白单元格的value()
    {
        xls::xlsCell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        const auto &value = xlsCellObj.value ();
        EXPECT_TRUE (std::holds_alternative<std::monostate> (value));
        freeXlsCell (cell);
    }
}

// 测试XlsCell的valueType()方法
TEST_F (XlsCellTest, ValueTypeMethodTest)
{
    // 测试各种单元格类型的valueType()
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "Test");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "string");
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 123.45);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "double");
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "bool");
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.valueType (), "monostate");
        freeXlsCell (cell);
    }
}

// 测试XlsCell的复制构造函数和赋值运算符
TEST_F (XlsCellTest, CopyAndMoveOperations)
{
    // 测试复制构造函数
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "Original");
        XlsCell original (cell);
        XlsCell copy (original); // 复制构造函数

        EXPECT_EQ (original.row (), copy.row ());
        EXPECT_EQ (original.col (), copy.col ());
        EXPECT_EQ (original.asStdString (false), copy.asStdString (false));
        freeXlsCell (cell);
    }

    // 测试移动构造函数
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 42.0);
        XlsCell original (cell);
        XlsCell moved (std::move (original)); // 移动构造函数

        EXPECT_EQ (moved.row (), 0);
        EXPECT_EQ (moved.col (), 0);
        EXPECT_EQ (moved.asStdString (false), "42");
        freeXlsCell (cell);
    }

    // 测试复制赋值运算符
    {
        xls::xlsCell *cell1 = createTestStringCell (0, 0, "First");
        xls::xlsCell *cell2 = createTestNumberCell (1, 1, 99.0);

        XlsCell cellA (cell1);
        XlsCell cellB (cell2);

        cellB = cellA; // 复制赋值

        EXPECT_EQ (cellB.row (), 0);
        EXPECT_EQ (cellB.col (), 0);
        EXPECT_EQ (cellB.asStdString (false), "First");
        freeXlsCell (cell1);
        freeXlsCell (cell2);
    }

    // 测试移动赋值运算符
    {
        xls::xlsCell *cell1 = createTestBoolCell (0, 0, true);
        xls::xlsCell *cell2 = createTestBlankCell (1, 1);

        XlsCell cellA (cell1);
        XlsCell cellB (cell2);

        cellB = std::move (cellA); // 移动赋值

        EXPECT_EQ (cellB.row (), 0);
        EXPECT_EQ (cellB.col (), 0);
        EXPECT_EQ (cellB.asStdString (false), "TRUE");
        freeXlsCell (cell1);
        freeXlsCell (cell2);
    }
}

// 测试边缘情况
TEST_F (XlsCellTest, EdgeCases)
{
    // 测试空字符串单元格
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "");
        freeXlsCell (cell);
    }

    // 测试只包含空格的字符串单元格
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "   ");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (true), "");
        freeXlsCell (cell);
    }

    // 测试零值数字单元格
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 0.0);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "0");
        EXPECT_FALSE (xlsCellObj.asLogical ());
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
        freeXlsCell (cell);
    }

    // 测试负值数字单元格
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, -123.456);
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "-123.456");
        EXPECT_TRUE (xlsCellObj.asLogical ()); // 非零值应为true
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), -123.456);
        freeXlsCell (cell);
    }

    // 测试非常大的数字
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 1.23456789e15);
        XlsCell xlsCellObj (cell);
        std::string str = xlsCellObj.asStdString (false);
        EXPECT_TRUE (str.find ("1.23456789") != std::string::npos);
        freeXlsCell (cell);
    }

    // 测试非常小的数字
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 1.23456789e-15);
        XlsCell xlsCellObj (cell);
        std::string str = xlsCellObj.asStdString (false);
        EXPECT_TRUE (str.find ("1.23456789") != std::string::npos);
        freeXlsCell (cell);
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
    EXPECT_THROW (parseAddress ("A"), ExcelReader::AddressParseException);
    EXPECT_THROW (parseAddress ("1"), std::invalid_argument);
    EXPECT_THROW (parseAddress ("A1B"), ExcelReader::AddressParseException);

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
    // 注意：isValide对目录会抛出FileNotFoundException而不是NotAFileException
    EXPECT_THROW (isValide ("."), ExcelReader::FileNotFoundException);

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
        xls::xlsCell *cell = createTestStringCell (0, 0, "123");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "123");
        freeXlsCell (cell);
    }

    // 测试看起来像数字的字符串
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "123.456");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "123.456");
        freeXlsCell (cell);
    }

    // 测试看起来像布尔的字符串
    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "true");
        XlsCell xlsCellObj (cell);
        EXPECT_EQ (xlsCellObj.asStdString (false), "true");
        freeXlsCell (cell);
    }

    // 测试布尔单元格
    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }
}

// 测试XlsCell的asLogical()方法
TEST_F (XlsCellTest, AsLogicalTest)
{
    // 测试各种类型的逻辑值
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 0.0);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 1.0);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, -1.0);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, false);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_TRUE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "anything");
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_FALSE (xlsCellObj.asLogical ());
        freeXlsCell (cell);
    }
}

// 测试XlsCell的asDouble()方法
TEST_F (XlsCellTest, AsDoubleTest)
{
    // 测试各种类型的双精度值
    {
        xls::xlsCell *cell = createTestNumberCell (0, 0, 123.456);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 123.456);
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, true);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 1.0);
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBoolCell (0, 0, false);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestStringCell (0, 0, "hello");
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
        freeXlsCell (cell);
    }

    {
        xls::xlsCell *cell = createTestBlankCell (0, 0);
        XlsCell xlsCellObj (cell);
        EXPECT_DOUBLE_EQ (xlsCellObj.asDouble (), 0.0);
        freeXlsCell (cell);
    }
}

int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}