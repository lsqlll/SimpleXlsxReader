#include "Exceptions.h"
#include "type.h"
#include "utils.h"

#include <filesystem>
#include <memory>
extern "C"
{
#include "xls.h"
}

namespace fs = std::filesystem;

class ReadStrategy
{
public:
    virtual ~ReadStrategy () = default;

    virtual CellType readCell (const std::size_t pos, const std::size_t row,
                               const std::size_t col) = 0;

    virtual CellType readCell (const std::size_t pos,
                               const std::string& addr) = 0;

    virtual CellType readCell (const std::size_t pos,
                               const CellPosition& cpos) = 0;
};

class XLSReadStrategy : public ReadStrategy
{
private:
    XLSheets sheets_;
    std::vector<bool> parsedSheets_;
    XLSheetsName names_;

public:
    explicit XLSReadStrategy (const fs::path& path,
                              std::unique_ptr<xls::xlsWorkBook> workbook_)
    {
        auto err = xls::LIBXLS_OK;

        if (!workbook_)
            throw ExcelReader::FailedOpenException (path.string ());

        for (std::size_t i = 0; i < workbook_.get ()->sheets.count; ++i)
            {
                auto sheet = xls::xls_getWorkSheet (workbook_.get (), i);
                sheets_.push_back (sheet);
                names_[ i ] = workbook_.get ()->sheets.sheet[ i ].name;
                parsedSheets_[ i ] = false;
            }
    }

    ~XLSReadStrategy () override = default;

    CellType
    readCell (const std::size_t pos, const std::size_t row,
              const std::size_t col) override
    {
        if (pos > names_.size ())
            throw ExcelReader::IndexOutException ("sheets["
                                                  + std::to_string (pos) + "]");

        if (!parsedSheets_[ pos ])
            {
                auto sheet = sheets_[ pos ];
                xls::xls_parseWorkSheet (sheet);
                parsedSheets_[ pos ] = true;
            }

        auto xls_cell = xls::xls_cell (sheets_[ pos ], row, col);
        auto c = XlsCell (xls_cell);
        return c.type ();
    }

    CellType
    readCell (std::size_t pos, const std::string& addr) override
    {
        auto loc = parseAddress (addr);
        return readCell (pos, loc.first, loc.second);
    }

    CellType
    readCell (std::size_t pos, const CellPosition& cpos) override
    {
        if (!cpos.addr.empty ())
            {
                return readCell (pos, cpos.addr);
            }
        return readCell (pos, cpos.row, cpos.col);
    }

    std::string
    getSheetName (std::size_t pos) const
    {
        if (pos >= names_.size ())
            throw ExcelReader::IndexOutException ("sheets["
                                                  + std::to_string (pos) + "]");
        return names_[ pos ];
    }
};

class XLSXReadStrategy : public ReadStrategy
{
public:
    explicit XLSXReadStrategy () = default;
    ~XLSXReadStrategy () override = default;

    CellType
    readCell (const std::size_t pos, const std::size_t row,
              const std::size_t col) override
    {
        // Implement XLSX-specific logic here
        return CellType ();
    }

    CellType
    readCell (const std::size_t pos, const std::string& addr) override
    {
        // Implement XLSX-specific logic here
        return CellType ();
    }

    CellType
    readCell (const std::size_t pos, const CellPosition& cpos) override
    {
        // Implement XLSX-specific logic here
        return CellType ();
    }
};

class CSVReadStrategy : public ReadStrategy
{
public:
    explicit CSVReadStrategy () = default;
    ~CSVReadStrategy () override = default;

    CellType
    readCell (const std::size_t pos, const std::size_t row,
              const std::size_t col) override
    {
        // Implement CSV-specific logic here
        return CellType ();
    }

    CellType
    readCell (const std::size_t pos, const std::string& addr) override
    {
        // Implement CSV-specific logic here
        return CellType ();
    }

    CellType
    readCell (const std::size_t pos, const CellPosition& cpos) override
    {
        // Implement CSV-specific logic here
        return CellType ();
    }
};