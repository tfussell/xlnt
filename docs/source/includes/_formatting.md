
# Formatting

## Format vs. Style

```c++
#include <iostream>
#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    auto cell = wb.active_sheet().cell("A1");
    return 0;
}
```

In the context of xlnt, format and style have specific distinct meanings. A style in xlnt corresponds to a named style created in the "Cell styles" dropdown in Excel. It must have a name and optionally any of: alignment, border, fill, font, number format, protection. A format in xlnt corresponds to the alignment, border, fill, font, number format, and protection settings applied to a cell via right-click->"Format Cells". A cell can have both a format and a style. The style properties will generally override the format properties.

## Number Formatting

```c++
#include <iostream>
#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    auto cell = wb.active_sheet().cell("A1");
    cell.number_format(xlnt::number_format::percentage());
    cell.value(0.513);
    std::cout << cell.to_string() << std::endl;
    return 0;
}
```

An xlnt::number_format is the format code used when displaying a value in a cell. For example, a number_format of "0.00" implies that the number 13.726 should be displayed as "13.73". Many number formats are built-in to Excel and can be access with xlnt::number_format static constructors. Other custom number formats can be created by passing a string to the [xlnt::number_format constructor](#cell-const-cell-amp).
