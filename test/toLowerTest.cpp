#include < gtest.h>
#include <algorithm>
#include <cctype>
#include <string>

// 被测函数实现
inline std::string
tolower (const std::string &raw_value)
{
    std::string tmp{};
    // 修复原代码的问题：需要预分配空间或使用back_inserter
    tmp.resize (raw_value.length ());
    std::transform (raw_value.begin (), raw_value.end (), tmp.begin (),
                    [] (char c) { return std::tolower (c); });
    return tmp;
}

// 或者使用更安全的实现方式
inline std::string
tolower_safe (const std::string &raw_value)
{
    std::string tmp{};
    tmp.reserve (raw_value.length ()); // 预分配空间提高性能
    std::transform (raw_value.begin (), raw_value.end (),
                    std::back_inserter (tmp),
                    [] (char c) { return std::tolower (c); });
    return tmp;
}

// 单元测试类
class StringToLowerTest : public ::testing::Test
{
  protected:
    void
    SetUp () override
    {
        // 测试前准备
    }

    void
    TearDown () override
    {
        // 测试后清理
    }
};

// TC001: 测试空字符串处理
TEST_F (StringToLowerTest, EmptyString_ReturnsEmptyString)
{
    // Arrange
    std::string input = "";
    std::string expected = "";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// TC002: 测试大写转小写
TEST_F (StringToLowerTest, UppercaseString_ReturnsLowercase)
{
    // Arrange
    std::string input = "HELLO";
    std::string expected = "hello";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// TC003: 测试小写保持不变
TEST_F (StringToLowerTest, LowercaseString_ReturnsSameString)
{
    // Arrange
    std::string input = "world";
    std::string expected = "world";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// TC004: 测试混合大小写
TEST_F (StringToLowerTest, MixedCaseString_ReturnsLowercase)
{
    // Arrange
    std::string input = "Hello World";
    std::string expected = "hello world";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// TC005: 测试包含数字和特殊字符
TEST_F (StringToLowerTest,
        StringWithNumbersAndSpecialChars_PreservesNonLetters)
{
    // Arrange
    std::string input = "123ABC!@#";
    std::string expected = "123abc!@#";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// TC006: 测试实际应用场景
TEST_F (StringToLowerTest, RealWorldScenario_CPlusPlusProgramming)
{
    // Arrange
    std::string input = "C++ Programming";
    std::string expected = "c++ programming";

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

// 边界测试：长字符串
TEST_F (StringToLowerTest, LongString_HandlesCorrectly)
{
    // Arrange
    std::string input (1000, 'A');    // 1000个'A'
    std::string expected (1000, 'a'); // 期望1000个'a'

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
    EXPECT_EQ (result.length (), 1000);
}

// 性能对比测试：比较两种实现方式
TEST_F (StringToLowerTest, SafeImplementation_CompareWithOriginal)
{
    // Arrange
    std::string input = "Hello WORLD 123!";

    // Act
    std::string result1 = tolower (input);
    std::string result2 = tolower_safe (input);

    // Assert
    EXPECT_EQ (result1, result2);
}

// 参数化测试：批量测试不同输入
class StringToLowerParamTest
    : public ::testing::TestWithParam<std::pair<std::string, std::string> >
{
};

TEST_P (StringToLowerParamTest, BatchTest_Parametrized)
{
    // Arrange
    auto test_data = GetParam ();
    std::string input = test_data.first;
    std::string expected = test_data.second;

    // Act
    std::string result = tolower (input);

    // Assert
    EXPECT_EQ (result, expected);
}

INSTANTIATE_TEST_SUITE_P (
    StringToLowerTests, StringToLowerParamTest,
    ::testing::Values (std::make_pair ("", ""), std::make_pair ("ABC", "abc"),
                       std::make_pair ("xyz", "xyz"),
                       std::make_pair ("MiXeD CaSe", "mixed case"),
                       std::make_pair ("123!@#", "123!@#")));

int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}