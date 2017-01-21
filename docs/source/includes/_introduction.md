
# Introduction

```c++
#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();
    ws.cell("A1").value(5);
    ws.cell("B2").value("string data");
    ws.cell("C3").formula("=RAND()");
    ws.merge_cells("C3:C4");
    ws.freeze_panes("B2");
    wb.save("xlnt.xlsx");
    return 0;
}
```

xlnt is a C++14 library for reading, writing, and modifying XLSX files as described in [ECMA 376](http://www.ecma-international.org/publications/standards/Ecma-376.htm). The API is based on [openpyxl](https://bitbucket.org/openpyxl/openpyxl), a Python library to read/write Excel 2007 xlsx/xlsm files, and ultimately on [PHPExcel](https://github.com/PHPOffice/PHPExcel), pure PHP library for reading and writing spreadsheet files upon which openpyxl was based. This project is still very much a work in progress, but the core development work is complete.

## Features

| Feature                                                             | Read | Edit | Write |
|---------------------------------------------------------------------|------|------|-------|
| Excel-style Workbook                                                | ✓    | ✓    | ✓     |
| LibreOffice-style Workbook                                          | ✓    | ✓    | ✓     |
| Numbers-style Workbook                                              | ✓    | ✓    | ✓     |
| Encrypted Workbook (Excel 2007-2010)                                | ✓    | ✓    |       |
| Encrypted Workbook (Excel 2013-2016)                                | ✓    | ✓    |       |
| Excel Binary Workbook (.xlsb)                                       |      |      |       |
| Excel Macro-Enabled Workbook (.xlsm)                                |      |      |       |
| Excel Macro-Enabled Template (.xltm)                                |      |      |       |
| Document Properties                                                 | ✓    | ✓    | ✓     |
| Numeric Cell Values                                                 | ✓    | ✓    | ✓     |
| Inline String Cell Values                                           | ✓    | ✓    | ✓     |
| Shared String Cell Values                                           | ✓    | ✓    | ✓     |
| Shared String Text Run Formatting (e.g. varied fonts within a cell) | ✓    | ✓    | ✓     |
| Hyperlink Cell Values                                               |      |      |       |
| Formula Cell Values                                                 |      |      |       |
| Formula Evaluation                                                  |      |      |       |
| Page Margins                                                        | ✓    | ✓    | ✓     |
| Page Setup                                                          |      |      |       |
| Print Area                                                          |      |      |       |
| Comments                                                            | ✓    | ✓    |       |
| Header and Footer                                                   |      |      |       |
| Custom Views                                                        |      |      |       |
| Charts                                                              |      |      |       |
| Chartsheets                                                         |      |      |       |
| Dialogsheets                                                        |      |      |       |
| Themes                                                              | ✓    |      | ✓     |
| Cell Styles                                                         | ✓    | ✓    | ✓     |
| Cell Formats                                                        | ✓    | ✓    | ✓     |
| Formatting->Alignment (e.g. right align)                            | ✓    | ✓    | ✓     |
| Formatting->Border (e.g. red cell outline)                          | ✓    | ✓    | ✓     |
| Formatting->Fill (e.g. green cell background)                       | ✓    | ✓    | ✓     |
| Formatting->Font (e.g. blue cell text)                              | ✓    | ✓    | ✓     |
| Formatting->Number Format (e.g. show 2 decimals)                    | ✓    | ✓    | ✓     |
| Formatting->Protection (e.g. hide formulas)                         | ✓    | ✓    | ✓     |
| Column Styles                                                       |      |      |       |
| Row Styles                                                          |      |      |       |
| Sheet Styles                                                        |      |      |       |
| Conditional Formatting                                              |      |      |       |
| Tables                                                              |      |      |       |
| Table Formatting                                                    |      |      |       |
| Pivot Tables                                                        |      |      |       |
| XLSX Thumbnail                                                      | ✓    |      | ✓     |
| Custom OOXML Properties                                             |      |      |       |
| Custom OOXML Parts                                                  |      |      |       |
| Drawing                                                             |      |      |       |
| Text Box                                                            |      |      |       |
| WordArt                                                             |      |      |       |
| Embedded Content (e.g. images)                                      |      |      |       |
| Excel VBA                                                           |      |      |       |

## Performance

Creates a new workbook and writes 1 million cells to a file in X seconds using a fully optimized build on Windows 10 with MSVC 2015.
