#pragma once

#include <sstream>
#include <stdexcept>
#include <string>

namespace ExcelReader
{

// ============================================================================
// 基础异常类
// ============================================================================

/**
 * @brief Excel处理基础异常类
 *
 * 继承自std::runtime_error的理由：
 * 1. std::runtime_error是标准库中表示运行时错误的异常基类
 * 2. 它提供了what()方法的实现，返回错误信息
 * 3. 适合表示程序运行时发生的可恢复错误
 * 4. 与std::logic_error区分，后者表示程序逻辑错误
 *
 * 使用场景：所有Excel处理相关的异常基类
 */
class ExcelException : public std::runtime_error
{
  public:
    /**
     * @brief 构造函数：仅包含错误信息
     * @param message 错误描述信息
     *
     * 使用场景：当异常没有明确的底层原因时使用
     */
    explicit ExcelException (const std::string &message)
        : std::runtime_error (message)
    {
    }

    /**
     * @brief 构造函数：包含错误信息和底层原因
     * @param message 错误描述信息
     * @param cause 底层异常原因
     *
     * 使用场景：当异常是由其他异常引起时，保留异常链信息
     */
    ExcelException (const std::string &message, const std::exception &cause)
        : std::runtime_error (message
                              + " [原因: " + std::string (cause.what ()) + "]")
    {
    }

    ~ExcelException () override = default;
};

// ============================================================================
// 文件相关异常
// ============================================================================

/**
 * @brief 文件未找到异常
 *
 * 使用场景：
 * 1. 尝试打开不存在的Excel文件时
 * 2. 文件路径错误或文件被移动/删除时
 * 3. 权限不足无法访问文件时
 */
class FileNotFoundException : public ExcelException
{
  public:
    explicit FileNotFoundException (const std::string &filePath)
        : ExcelException ("文件未找到: '" + filePath + "'")
    {
    }

    FileNotFoundException (const std::string &filePath,
                           const std::exception &cause)
        : ExcelException ("文件未找到: '" + filePath + "'", cause)
    {
    }
};

/**
 * @brief 不支持的文件格式异常
 *
 * 使用场景：
 * 1. 尝试打开非Excel格式文件时（如.txt, .pdf等）
 * 2. 文件扩展名不是.xls, .xlsx, .csv时
 * 3. 文件内容损坏或格式不正确时
 */
class UnsupportedFormatException : public ExcelException
{
  public:
    explicit UnsupportedFormatException (const std::string &format)
        : ExcelException ("不支持的文件格式: '" + format + "'")
    {
    }

    UnsupportedFormatException (const std::string &format,
                                const std::exception &cause)
        : ExcelException ("不支持的文件格式: '" + format + "'", cause)
    {
    }
};

/**
 * @brief 路径不是文件异常
 *
 * 使用场景：
 * 1. 提供的路径指向目录而不是文件时
 * 2. 路径是符号链接但指向目录时
 * 3. 路径是设备文件或其他特殊文件时
 */
class NotAFileException : public ExcelException
{
  public:
    explicit NotAFileException (const std::string &path)
        : ExcelException ("路径不是文件: '" + path + "'")
    {
    }

    NotAFileException (const std::string &path, const std::exception &cause)
        : ExcelException ("路径不是文件: '" + path + "'", cause)
    {
    }
};

/**
 * @brief 文件访问异常
 *
 * 使用场景：
 * 1. 文件被其他进程锁定无法打开时
 * 2. 磁盘空间不足无法读取文件时
 * 3. 文件损坏或格式错误无法解析时
 * 4. 网络文件访问超时或中断时
 */
class FileAccessException : public ExcelException
{
  public:
    explicit FileAccessException (const std::string &filePath)
        : ExcelException ("无法访问文件: '" + filePath + "'")
    {
    }

    FileAccessException (const std::string &filePath,
                         const std::exception &cause)
        : ExcelException ("无法访问文件: '" + filePath + "'", cause)
    {
    }
};

// ============================================================================
// 数据访问异常
// ============================================================================

/**
 * @brief 索引越界异常
 *
 * 使用场景：
 * 1. 访问不存在的行或列时（如row=100但只有50行）
 * 2. 工作表索引超出范围时
 * 3. 单元格引用超出有效范围时
 */
class IndexOutOfBoundsException : public ExcelException
{
  public:
    IndexOutOfBoundsException (const std::string &indexName, int indexValue,
                               int minValue, int maxValue)
        : ExcelException (
              buildMessage (indexName, indexValue, minValue, maxValue))
    {
    }

    IndexOutOfBoundsException (const std::string &indexName, int indexValue,
                               int minValue, int maxValue,
                               const std::exception &cause)
        : ExcelException (
              buildMessage (indexName, indexValue, minValue, maxValue), cause)
    {
    }

  private:
    static std::string
    buildMessage (const std::string &indexName, int indexValue, int minValue,
                  int maxValue)
    {
        std::ostringstream oss;
        oss << "索引越界: '" << indexName << "' 值 " << indexValue
            << " 超出范围 [最小值: " << minValue << ", 最大值: " << maxValue
            << "]";
        return oss.str ();
    }
};

/**
 * @brief 无效单元格异常
 *
 * 使用场景：
 * 1. 单元格指针为nullptr时
 * 2. 单元格数据损坏或格式错误时
 * 3. 尝试访问已释放或无效的单元格时
 * 4. 单元格坐标无效时（如负值或超出范围）
 */
class InvalidCellException : public ExcelException
{
  public:
    explicit InvalidCellException (const std::string &cellInfo)
        : ExcelException ("无效的单元格: " + cellInfo)
    {
    }

    InvalidCellException (const std::string &cellInfo,
                          const std::exception &cause)
        : ExcelException ("无效的单元格: " + cellInfo, cause)
    {
    }

    InvalidCellException (int row, int col)
        : ExcelException (buildMessage (row, col))
    {
    }

    InvalidCellException (int row, int col, const std::exception &cause)
        : ExcelException (buildMessage (row, col), cause)
    {
    }

  private:
    static std::string
    buildMessage (int row, int col)
    {
        std::ostringstream oss;
        oss << "无效的单元格位置 [行: " << row << ", 列: " << col << "]";
        return oss.str ();
    }
};

/**
 * @brief 地址解析异常
 *
 * 使用场景：
 * 1. Excel地址格式错误时（如"A", "1", "AA1B"等）
 * 2. 地址包含非法字符时
 * 3. 行号或列号超出有效范围时
 * 4. 地址字符串为空时
 */
class AddressParseException : public ExcelException
{
  public:
    explicit AddressParseException (const std::string &address)
        : ExcelException ("无法解析地址: '" + address + "'")
    {
    }

    AddressParseException (const std::string &address,
                           const std::exception &cause)
        : ExcelException ("无法解析地址: '" + address + "'", cause)
    {
    }

    AddressParseException (const std::string &address,
                           const std::string &reason)
        : ExcelException ("无法解析地址: '" + address + "' - " + reason)
    {
    }
};

/**
 * @brief 工作表访问异常
 *
 * 使用场景：
 * 1. 工作表索引超出范围时
 * 2. 工作表名不存在时
 * 3. 工作表数据损坏无法读取时
 * 4. 工作表被隐藏或保护无法访问时
 */
class WorksheetAccessException : public ExcelException
{
  public:
    explicit WorksheetAccessException (const std::string &sheetInfo)
        : ExcelException ("无法访问工作表: " + sheetInfo)
    {
    }

    WorksheetAccessException (const std::string &sheetInfo,
                              const std::exception &cause)
        : ExcelException ("无法访问工作表: " + sheetInfo, cause)
    {
    }

    WorksheetAccessException (int sheetIndex, int totalSheets)
        : ExcelException (buildMessage (sheetIndex, totalSheets))
    {
    }

  private:
    static std::string
    buildMessage (int sheetIndex, int totalSheets)
    {
        std::ostringstream oss;
        oss << "工作表索引 " << sheetIndex
            << " 超出范围 [工作表总数: " << totalSheets << "]";
        return oss.str ();
    }
};

/**
 * @brief 数据类型转换异常
 *
 * 使用场景：
 * 1. 尝试将字符串转换为数字失败时
 * 2. 布尔值格式不正确时
 * 3. 日期格式无法解析时
 * 4. 单元格类型与期望类型不匹配时
 */
class DataTypeConversionException : public ExcelException
{
  public:
    explicit DataTypeConversionException (const std::string &conversionInfo)
        : ExcelException ("数据类型转换失败: " + conversionInfo)
    {
    }

    DataTypeConversionException (const std::string &conversionInfo,
                                 const std::exception &cause)
        : ExcelException ("数据类型转换失败: " + conversionInfo, cause)
    {
    }

    DataTypeConversionException (const std::string &fromType,
                                 const std::string &toType,
                                 const std::string &value)
        : ExcelException (buildMessage (fromType, toType, value))
    {
    }

  private:
    static std::string
    buildMessage (const std::string &fromType, const std::string &toType,
                  const std::string &value)
    {
        std::ostringstream oss;
        oss << "无法将 " << fromType << " 类型值 '" << value << "' 转换为 "
            << toType << " 类型";
        return oss.str ();
    }
};

/**
 * @brief 内存分配异常
 *
 * 使用场景：
 * 1. 分配大文件内存失败时
 * 2. 系统内存不足时
 * 3. 内存碎片导致分配失败时
 */
class MemoryAllocationException : public ExcelException
{
  public:
    explicit MemoryAllocationException (const std::string &allocationInfo)
        : ExcelException ("内存分配失败: " + allocationInfo)
    {
    }

    MemoryAllocationException (const std::string &allocationInfo,
                               const std::exception &cause)
        : ExcelException ("内存分配失败: " + allocationInfo, cause)
    {
    }

    MemoryAllocationException (size_t requestedSize,
                               const std::string &purpose)
        : ExcelException (buildMessage (requestedSize, purpose))
    {
    }

  private:
    static std::string
    buildMessage (size_t requestedSize, const std::string &purpose)
    {
        std::ostringstream oss;
        oss << "无法分配 " << requestedSize << " 字节内存用于: " << purpose;
        return oss.str ();
    }
};

} // namespace ExcelReader