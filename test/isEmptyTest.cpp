#include < gtest.h>
#include <algorithm>
#include <cctype>
#include <string>

// 被测试的函数实现
inline std::string
trim (const std::string &str)
{
    if (str.empty ())
        return str;

    auto start = str.find_first_not_of (" \t");
    if (start == std::string::npos)
        return "";

    auto end = str.find_last_not_of (" \t");
    return str.substr (start, end - start + 1);
}

inline bool
isEmpty (const std::string &raw_value, bool trims = false)
{
    if (raw_value.empty ())
    {
        return true;
    }
    if (trims)
    {
        auto raw_string = trim (raw_value);
        return raw_string.empty ();
    }

    return std::all_of (raw_value.begin (), raw_value.end (),
                        [] (unsigned char c) { return std::isspace (c); });
}

// trim函数测试用例
TEST (TrimTest, EmptyString) { EXPECT_EQ (trim (""), ""); }

TEST (TrimTest, OnlySpaces)
{
    EXPECT_EQ (trim ("   "), "");
    EXPECT_EQ (trim (" "), "");
    EXPECT_EQ (trim ("          "), "");
}

TEST (TrimTest, OnlyTabs)
{
    EXPECT_EQ (trim ("\t\t"), "");
    EXPECT_EQ (trim ("\t"), "");
}

TEST (TrimTest, MixedWhitespaceOnly)
{
    EXPECT_EQ (trim (" \t \t "), "");
    EXPECT_EQ (trim ("\t \t"), "");
}

TEST (TrimTest, NoLeadingOrTrailingWhitespace)
{
    EXPECT_EQ (trim ("hello"), "hello");
    EXPECT_EQ (trim ("hello world"), "hello world");
    EXPECT_EQ (trim ("test123"), "test123");
}

TEST (TrimTest, LeadingWhitespaceOnly)
{
    EXPECT_EQ (trim ("   hello"), "hello");
    EXPECT_EQ (trim ("\t\tworld"), "world");
    EXPECT_EQ (trim (" \t test"), "test");
}

TEST (TrimTest, TrailingWhitespaceOnly)
{
    EXPECT_EQ (trim ("hello   "), "hello");
    EXPECT_EQ (trim ("world\t\t"), "world");
    EXPECT_EQ (trim ("test \t "), "test");
}

TEST (TrimTest, BothLeadingAndTrailingWhitespace)
{
    EXPECT_EQ (trim ("   hello   "), "hello");
    EXPECT_EQ (trim ("\t\tworld\t\t"), "world");
    EXPECT_EQ (trim (" \t test \t "), "test");
    EXPECT_EQ (trim ("  hello world  "), "hello world");
}

TEST (TrimTest, SingleCharacter)
{
    EXPECT_EQ (trim ("a"), "a");
    EXPECT_EQ (trim (" "), "");
    EXPECT_EQ (trim ("\t"), "");
    EXPECT_EQ (trim (" x "), "x");
}

TEST (TrimTest, InternalWhitespacePreserved)
{
    EXPECT_EQ (trim (" hello world "), "hello world");
    EXPECT_EQ (trim ("  one   two  "), "one   two");
    EXPECT_EQ (trim ("\t tab \t separated \t"), "tab \t separated");
}

// isEmpty函数测试用例
TEST (IsEmptyTest, TrulyEmptyString)
{
    EXPECT_TRUE (isEmpty (""));
    EXPECT_TRUE (isEmpty ("", false));
    EXPECT_TRUE (isEmpty ("", true));
}

TEST (IsEmptyTest, WhitespaceOnlyWithoutTrim)
{
    EXPECT_TRUE (isEmpty ("   ", false));
    EXPECT_TRUE (isEmpty ("\t\t", false));
    EXPECT_TRUE (isEmpty (" \t ", false));
    EXPECT_TRUE (isEmpty ("          ", false));
}

TEST (IsEmptyTest, WhitespaceOnlyWithTrim)
{
    EXPECT_TRUE (isEmpty ("   ", true));
    EXPECT_TRUE (isEmpty ("\t\t", true));
    EXPECT_TRUE (isEmpty (" \t ", true));
    EXPECT_TRUE (isEmpty ("          ", true));
    EXPECT_TRUE (isEmpty ("\t \t\t \t", true));
}

TEST (IsEmptyTest, NonEmptyContentWithoutTrim)
{
    EXPECT_FALSE (isEmpty ("hello", false));
    EXPECT_FALSE (isEmpty ("a", false));
    EXPECT_FALSE (isEmpty (" test ", false));
    EXPECT_FALSE (isEmpty ("\tcontent", false));
    EXPECT_FALSE (isEmpty ("content\t", false));
}

TEST (IsEmptyTest, NonEmptyContentWithTrim)
{
    EXPECT_FALSE (isEmpty ("hello", true));
    EXPECT_FALSE (isEmpty ("a", true));
    EXPECT_FALSE (isEmpty (" test ", true)); // Still has content after trim
    EXPECT_FALSE (isEmpty ("\tcontent", true));
    EXPECT_FALSE (isEmpty ("content\t", true));
}

TEST (IsEmptyTest, BecomesEmptyAfterTrim)
{
    EXPECT_TRUE (isEmpty ("   ", true));
    EXPECT_TRUE (isEmpty (" \t ", true));
    EXPECT_TRUE (isEmpty ("\t\t\t", true));
    EXPECT_TRUE (isEmpty ("  \t\t  ", true));
}

TEST (IsEmptyTest, ContainsNonWhitespaceCharacters)
{
    EXPECT_FALSE (isEmpty ("a", false));
    EXPECT_FALSE (isEmpty (" a ", false));
    EXPECT_FALSE (isEmpty (" a ", true)); // Has actual content
    EXPECT_FALSE (isEmpty ("hello world", false));
    EXPECT_FALSE (isEmpty (" hello world ", true)); // Has content after trim
}

TEST (IsEmptyTest, EdgeCases)
{
    EXPECT_TRUE (isEmpty (" ", true));
    EXPECT_TRUE (isEmpty ("\t", true));
    EXPECT_FALSE (isEmpty (" a", false));
    EXPECT_FALSE (isEmpty ("a ", false));
    EXPECT_FALSE (isEmpty (" a", true));
    EXPECT_FALSE (isEmpty ("a ", true));
}

// 必须的main函数
int
main (int argc, char **argv)
{
    ::testing::InitGoogleTest (&argc, argv);
    return RUN_ALL_TESTS ();
}