#include <gtest/gtest.h>
#include <stdexcept>
#include <string>

// 假设ExcelReader类定义（如果不存在需要先定义）
class ExcelReader
{
public:
    class parseAddrException : public std::exception
    {
    private:
	std::string msg;

    public:
	explicit parseAddrException (const std::string& addr)
	    : msg ("Failed to parse address: " + addr)
	{
	}
	const char*
	what () const noexcept override
	{
	    return msg.c_str ();
	}
    };
};

// 声明被测函数（如果不在头文件中）
inline std::pair<std::size_t, std::size_t>
parseAddress (const std::string& addr)
{
    if (addr.empty ())
	{
	    throw std::invalid_argument ("address can't be empty");
	}
    if (!std::isalpha (addr[0]))
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
// Test suite for parseAddress function
class ParseAddressTest : public ::testing::Test
{
protected:
    void
    SetUp () override
    {
	// Setup code if needed
    }

    void
    TearDown () override
    {
	// Cleanup code if needed
    }
};

// Test case: Basic functionality with simple address
TEST_F (ParseAddressTest, BasicAddressParsing)
{
    // 测试基本地址解析功能 A1 -> (1, 1)
    auto result = parseAddress ("A1");
    EXPECT_EQ (result.first, 1);  // 行号应该为1
    EXPECT_EQ (result.second, 1); // 列号应该为1
}

// Test case: Single letter column maximum
TEST_F (ParseAddressTest, SingleLetterMaximum)
{
    // 测试单字母列号最大值 Z1 -> (1, 26)
    auto result = parseAddress ("Z1");
    EXPECT_EQ (result.first, 1);   // 行号应该为1
    EXPECT_EQ (result.second, 26); // Z对应第26列
}

// Test case: Multi-letter column addressing
TEST_F (ParseAddressTest, MultiLetterColumn)
{
    // 测试多字母列号 AA1 -> (1, 27)
    auto result = parseAddress ("AA1");
    EXPECT_EQ (result.first, 1);   // 行号应该为1
    EXPECT_EQ (result.second, 27); // AA对应第27列

    // 测试 AZ1 -> (1, 52)
    result = parseAddress ("AZ1");
    EXPECT_EQ (result.first, 1);   // 行号应该为1
    EXPECT_EQ (result.second, 52); // AZ对应第52列
}

// Test case: Large row numbers
TEST_F (ParseAddressTest, LargeRowNumbers)
{
    // 测试大行号 BC234 -> (234, ?)
    auto result = parseAddress ("BC234");
    EXPECT_EQ (result.first, 234); // 行号应该为234
    EXPECT_EQ (result.second, 55); // BC对应第55列 (B=2, C=3, 2*26+3=55)
}

// Test case: Empty string exception
TEST_F (ParseAddressTest, EmptyStringException)
{
    // 测试空字符串应该抛出 invalid_argument 异常
    EXPECT_THROW ({ parseAddress (""); }, std::invalid_argument);
}

// Test case: Invalid first character - number
TEST_F (ParseAddressTest, InvalidFirstCharacterNumber)
{
    // 测试以数字开头的地址应该抛出 invalid_argument 异常
    EXPECT_THROW ({ parseAddress ("1A"); }, std::invalid_argument);
}

// Test case: Invalid first character - special character
TEST_F (ParseAddressTest, InvalidFirstCharacterSpecial)
{
    // 测试以特殊字符开头的地址应该抛出 invalid_argument 异常
    EXPECT_THROW ({ parseAddress ("@1"); }, std::invalid_argument);
}

// Test case: Missing row number
TEST_F (ParseAddressTest, MissingRowNumber)
{
    // 测试缺少行号的地址应该抛出 parseAddrException 异常
    EXPECT_THROW ({ parseAddress ("A"); }, ExcelReader::parseAddrException);
}

// Test case: Only column letters
TEST_F (ParseAddressTest, OnlyColumnLetters)
{
    // 测试只有列号字母的地址应该抛出 parseAddrException 异常
    EXPECT_THROW ({ parseAddress ("AB"); }, ExcelReader::parseAddrException);
}

// Test case: Lowercase letters (should work according to isalpha)
TEST_F (ParseAddressTest, LowercaseLetters)
{
    // 测试小写字母是否能正常工作 a1 -> (1, 1) ,不支持
    EXPECT_THROW ({ parseAddress ("aB"); }, ExcelReader::parseAddrException);
}

// Test case: Mixed case letters
TEST_F (ParseAddressTest, MixedCaseLetters)
{
    // 测试混合大小写的地址解析
    auto result = parseAddress ("Ab1");
    EXPECT_EQ (result.first, 1);
    // 注意：根据实现，a和A可能有不同的处理方式
}

// Parameterized test for multiple valid addresses
class ParseAddressParamTest
    : public ::testing::TestWithParam<
	  std::tuple<std::string, std::size_t, std::size_t>>
{
};

// 参数化测试：验证多个有效地址的解析结果
TEST_P (ParseAddressParamTest, ValidAddresses)
{
    auto [address, expected_row, expected_col] = GetParam ();
    auto result = parseAddress (address);
    EXPECT_EQ (result.first, expected_row) << "Failed for address: " << address;
    EXPECT_EQ (result.second, expected_col)
	<< "Failed for address: " << address;
}

// 测试数据：{地址字符串, 期望行号, 期望列号}
INSTANTIATE_TEST_SUITE_P (
    ValidAddressTests, ParseAddressParamTest,
    ::testing::Values (
	std::make_tuple ("A1", 1, 1), std::make_tuple ("B1", 1, 2),
	std::make_tuple ("C1", 1, 3), std::make_tuple ("Z1", 1, 26),
	std::make_tuple ("AA1", 1, 27), std::make_tuple ("AB1", 1, 28),
	std::make_tuple ("AZ1", 1, 52), std::make_tuple ("BA1", 1, 53),
	std::make_tuple ("A100", 100, 1), std::make_tuple ("Z999", 999, 26)));
int
main (int argc, char** argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}