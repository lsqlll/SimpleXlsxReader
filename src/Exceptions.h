#pragma once

#include <string>
#include <utility>
namespace ExcelReader
{

class ExcelException : public std::exception
{
  private:
    std::string _message;

  public:
    explicit ExcelException (std::string msg) : _message (std::move (msg)) {};

    [[nodiscard]] const char *
    what () const noexcept override
    {
        return _message.c_str ();
    }
};

class FileNotFoundException : public ExcelException
{
  public:
    explicit FileNotFoundException (const std::string &msg)
        : ExcelException ("File not found: " + msg) {};
};

class UnsupportedException : public ExcelException
{
  public:
    explicit UnsupportedException (const std::string &msg)
        : ExcelException (msg + "not supported") {};
};

class PathNotFileException : public ExcelException
{
  public:
    explicit PathNotFileException (const std::string &msg)
        : ExcelException (msg + "is not a file") {};
};

class FailedOpenException : public ExcelException
{
  public:
    explicit FailedOpenException (const std::string &msg)
        : ExcelException ("Error reading file: " + msg) {};
};

class IndexOutException : public ExcelException
{
  public:
    explicit IndexOutException (const std::string &msg)
        : ExcelException ("Index" + msg + "out of Range") {};
};

class NullCellException : public ExcelException
{
  public:
    explicit NullCellException (const std::string &msg)
        : ExcelException ("Null cell: " + msg) {};
};

class ParseAddrException : public ExcelException
{
  public:
    explicit ParseAddrException (const std::string &msg)
        : ExcelException ("Parse error: " + msg) {};
};

} // namespace ExcelReader
