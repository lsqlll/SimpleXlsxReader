#pragma once

#include "Exceptions.h"
#include "ResourceManager.h"
#include "strategy.h"

#include <filesystem>
extern "C"
{
#include "xls.h"
}

namespace fs = std::filesystem;

class TableReader
{
public:
    virtual bool open () = 0;

    virtual std::size_t getSheetsCount () const = 0;
};

class XLSReader : public TableReader
{
private:
    ResourceManager<xls::xlsWorkBook> m_resourceManager;
    std::unique_ptr<XLSReadStrategy> m_strategy;
    fs::path path_;
    std::size_t sheetCounts_;

public:
    explicit XLSReader (const fs::path& p)
        : path_ (p) {};

    explicit XLSReader (const std::string& p)
        : path_ (fs::path (p))
    {
        if (!isValide (path_))
            throw ExcelReader::FileNotFoundException (path_.string ());
    };

    bool
    open () override
    {
        auto wb = std::unique_ptr<xls::xlsWorkBook> (
            xls::xls_open (path_.string ().c_str (), "UTF-8"));

        if (!wb)
            {
                this->sheetCounts_ = 0;
                return false;
            }

        this->sheetCounts_ = wb.get ()->sheets.count;
        m_strategy = std::make_unique<XLSReadStrategy> (path_, std::move (wb));
        return true;
    };

    std::string
    getSheetName (std::size_t index) const
    {
        return m_strategy->getSheetName (index);
    }

    std::size_t
    getSheetsCount () const override
    {
        return sheetCounts_;
    }
};
