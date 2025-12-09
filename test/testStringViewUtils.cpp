#include <gtest/gtest.h>
#include <optional>
#include <string>
#include <string_view>

// 被测函数声明（实际项目中应包含头文件）
inline std::optional<std::string_view>
getStringView (const char *str)
{
    if (str == nullptr)
    {
        return std::nullopt;
    }
    return std::string_view (str);
}

inline std::optional<std::string_view>
getStringView (const std::string &str)
{
    if (str.empty ())
    {
        return std::nullopt;
    }
    return std::string_view (str);
}

inline bool
startsWith (std::string_view str, std::string_view prefix)
{
    return str.substr (0, prefix.size ()) == prefix;
};

// Test suite for getStringView(const char*) function
// getStringView(const char*)函数测试套件
class GetStringViewCharPtrTest : public ::testing::Test
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

// Test case: When input is nullptr, should return std::nullopt
// 测试用例：当输入为nullptr时，应返回std::nullopt
TEST_F (GetStringViewCharPtrTest, ReturnsNulloptWhenInputIsNullptr)
{
    const char *null_str = nullptr;
    auto result = getStringView (null_str);
    EXPECT_FALSE (result.has_value ());
}

// Test case: When input is valid string, should return optional with string
// view 测试用例：当输入为有效字符串时，应返回包含字符串视图的optional
TEST_F (GetStringViewCharPtrTest,
        ReturnsOptionalWithStringViewWhenInputIsValid)
{
    const char *valid_str = "hello";
    auto result = getStringView (valid_str);
    ASSERT_TRUE (result.has_value ());
    EXPECT_EQ (result.value (), "hello");
}

// Test case: When input is empty string, should return optional with empty
// string view 测试用例：当输入为空字符串时，应返回包含空字符串视图的optional
TEST_F (GetStringViewCharPtrTest,
        ReturnsOptionalWithEmptyStringViewWhenInputIsEmptyString)
{
    const char *empty_str = "";
    auto result = getStringView (empty_str);
    ASSERT_TRUE (result.has_value ());
    EXPECT_EQ (result.value (), "");
}

// Test suite for getStringView(const std::string&) function
// getStringView(const std::string&)函数测试套件
class GetStringViewStdStringTest : public ::testing::Test
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

// Test case: When input string is empty, should return std::nullopt
// 测试用例：当输入字符串为空时，应返回std::nullopt
TEST_F (GetStringViewStdStringTest, ReturnsNulloptWhenInputStringIsEmpty)
{
    std::string empty_str;
    auto result = getStringView (empty_str);
    EXPECT_FALSE (result.has_value ());
}

// Test case: When input string is not empty, should return optional with
// string view 测试用例：当输入字符串非空时，应返回包含字符串视图的optional
TEST_F (GetStringViewStdStringTest,
        ReturnsOptionalWithStringViewWhenInputStringIsNotEmpty)
{
    std::string valid_str = "world";
    auto result = getStringView (valid_str);
    ASSERT_TRUE (result.has_value ());
    EXPECT_EQ (result.value (), "world");
}

// Test suite for startsWith function
// startsWith函数测试套件
class StartsWithTest : public ::testing::Test
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

// Test case: When string starts with given prefix, should return true
// 测试用例：当字符串以前缀开头时，应返回true
TEST_F (StartsWithTest, ReturnsTrueWhenStringStartsWithPrefix)
{
    std::string_view str = "hello world";
    std::string_view prefix = "hello";
    EXPECT_TRUE (startsWith (str, prefix));
}

// Test case: When string does not start with given prefix, should return false
// 测试用例：当字符串不以前缀开头时，应返回false
TEST_F (StartsWithTest, ReturnsFalseWhenStringDoesNotStartWithPrefix)
{
    std::string_view str = "hello world";
    std::string_view prefix = "world";
    EXPECT_FALSE (startsWith (str, prefix));
}

// Test case: When string equals to prefix, should return true
// 测试用例：当字符串等于前缀时，应返回true
TEST_F (StartsWithTest, ReturnsTrueWhenStringEqualsToPrefix)
{
    std::string_view str = "hello";
    std::string_view prefix = "hello";
    EXPECT_TRUE (startsWith (str, prefix));
}

// Test case: When prefix is longer than string, should return false
// 测试用例：当前缀长度超过字符串时，应返回false
TEST_F (StartsWithTest, ReturnsFalseWhenPrefixIsLongerThanString)
{
    std::string_view str = "hi";
    std::string_view prefix = "hello";
    EXPECT_FALSE (startsWith (str, prefix));
}

// Test case: When prefix is empty, should return true
// 测试用例：当前缀为空时，应返回true
TEST_F (StartsWithTest, ReturnsTrueWhenPrefixIsEmpty)
{
    std::string_view str = "hello";
    std::string_view prefix;
    EXPECT_TRUE (startsWith (str, prefix));
}

// Test case: When both string and prefix are empty, should return true
// 测试用例：当字符串和前缀都为空时，应返回true
TEST_F (StartsWithTest, ReturnsTrueWhenBothStringAndPrefixAreEmpty)
{
    std::string_view str;
    std::string_view prefix;
    EXPECT_TRUE (startsWith (str, prefix));
}

int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}