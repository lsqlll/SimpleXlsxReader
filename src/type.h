#ifndef TYPE_H
#define TYPE_H

#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <vector>

#include "XlsCell.h"

using XLSheets = std::vector<xls::xlsWorkSheet *>;
using XLSRow = std::vector<XlsCell>;
using XLSheetsName = std::vector<std::string>;

#endif
