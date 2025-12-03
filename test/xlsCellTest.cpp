#include "../src/type.h" // 包含 XlsCell 类定义

#include <gmock/gmock.h>
#include <gtest/gtest.h>

using ::testing::_;
using ::testing::Return;

// Mock xlsCell structure using a class wrapper for easier mocking
class MockXlsCell
{
public:
    MOCK_METHOD (int, getId, (), ());
    MOCK_METHOD (int, getRow, (), ());
    MOCK_METHOD (int, getCol, (), ());
    MOCK_METHOD (const char*, getStr, (), ());
    MOCK_METHOD (double, getD, (), ());
    MOCK_METHOD (long, getL, (), ());
    MOCK_METHOD (int, getXf, (), ());

    operator xls::xlsCell*()
    {
	static xls::xlsCell cell;
	cell.id = getId ();
	cell.row = getRow ();
	cell.col = getCol ();
	cell.str = const_cast<char*> (getStr ());
	cell.d = getD ();
	cell.l = getL ();
	cell.xf = getXf ();
	return &cell;
    }
};

TEST (XlsCellTest, ConstructorWithPointerInitializesCorrectly)
{
    MockXlsCell mock_cell;
    EXPECT_CALL (mock_cell, getRow ()).WillOnce (Return (5));
    EXPECT_CALL (mock_cell, getCol ()).WillOnce (Return (3));

    XlsCell cell (new xls::xlsCell (3, &mock_cell));

    EXPECT_EQ (cell.row (), 5);
    EXPECT_EQ (cell.col (), 3);
    EXPECT_EQ (cell.type (),
	       CellType::UNKNOWN); // Should trigger inferType internally
}

TEST (XlsCellTest, ConstructorWithLocationSetsBlankType)
{
    XlsCell cell ({ 7, 9 });

    EXPECT_EQ (cell.row (), 7);
    EXPECT_EQ (cell.col (), 9);
    EXPECT_EQ (cell.type (), CellType::BLANK);
}

TEST (XlsCellTest, TypeInferenceForStringRecordReturnsStringOrBlank)
{
    MockXlsCell mock_cell;
    EXPECT_CALL (mock_cell, getId ())
	.WillRepeatedly (Return (XLS_RECORD_LABEL));
    EXPECT_CALL (mock_cell, getStr ()).WillRepeatedly (Return ("Hello"));
    EXPECT_CALL (mock_cell, getRow ()).WillOnce (Return (0));
    EXPECT_CALL (mock_cell, getCol ()).WillOnce (Return (0));

    XlsCell cell (static_cast<xls::xlsCell*> (&mock_cell));
    EXPECT_EQ (cell.type (), CellType::STRING);

    // Test blank string
    EXPECT_CALL (mock_cell, getStr ()).WillRepeatedly (Return (""));
    XlsCell blank_cell (static_cast<xls::xlsCell*> (&mock_cell));
    EXPECT_EQ (blank_cell.type (), CellType::BLANK);
}

TEST (XlsCellTest, AsStdStringHandlesAllTypes)
{
    MockXlsCell mock_cell;
    EXPECT_CALL (mock_cell, getId ())
	.WillRepeatedly (Return (XLS_RECORD_NUMBER));
    EXPECT_CALL (mock_cell, getD ()).WillRepeatedly (Return (42.5));
    EXPECT_CALL (mock_cell, getRow ()).WillOnce (Return (0));
    EXPECT_CALL (mock_cell, getCol ()).WillOnce (Return (0));

    XlsCell num_cell (static_cast<xls::xlsCell*> (&mock_cell));
    EXPECT_EQ (num_cell.asStdString (false, {}), "42.5");

    // Test logical conversion
    EXPECT_CALL (mock_cell, getD ()).WillRepeatedly (Return (1.0));
    EXPECT_EQ (num_cell.asLogical (), true);

    // Test double conversion
    EXPECT_DOUBLE_EQ (num_cell.asDouble (), 42.5);
}

// Add more tests covering all branches in inferType and other methods...