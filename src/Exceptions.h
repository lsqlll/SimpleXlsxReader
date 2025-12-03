#pragma once


#include <string>
namespace ExcelReader
{

class ExcelException : public std::exception
{
private:
    std::string _message;

public:
    explicit ExcelException (const std::string& msg)
	: _message (msg) {};

    const char*
    what () const noexcept override
    {
	return _message.c_str ();
    }
};

class FileNotFoundException : public ExcelException
{
public:
    explicit FileNotFoundException (const std::string& msg)
	: ExcelException ("File not found: " + msg) {};
};

class UnsupportedException : public ExcelException
{
public:
    explicit UnsupportedException (const std::string& msg)
	: ExcelException (msg + "not supported") {};
};

class PathNotFileException : public ExcelException
{
public:
    explicit PathNotFileException (const std::string& msg)
	: ExcelException (msg + "is not a file") {};
};

class FailedOpenException : public ExcelException
{
public:
    explicit FailedOpenException (const std::string& msg)
	: ExcelException ("Error reading file: " + msg) {};
};

class IndexOutException : public ExcelException
{
public:
    explicit IndexOutException (const std::string& msg)
	: ExcelException ("Index" + msg + "out of Range") {};
};
class parseAddrException : public ExcelException
{
public:
    explicit parseAddrException (const std::string& msg)
	: ExcelException ("Parse error: " + msg) {};
};

class parseErrorException : public ExcelException
{
public:
    explicit parseErrorException (const std::string& file,
				  const std::string& msg)
	: ExcelException ("Parse error in " + file + ": " + msg) {};
};
} // namespace ExcelReader