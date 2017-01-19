
# Memory Model

```c++
#include <iostream>
#include <xlnt/xlnt.hpp>

void set_cell(xlnt::cell cell, int value)
{
    cell.value(value);
}

xlnt::workbook create_wb()
{
    xlnt::workbook wb;
    auto ws = wb.active_sheet();
    set_cell(wb.cell("A1"), 2);
    return wb;
}

int main()
{
    auto wb = create_wb();
    std::cout << wb.value<int>() << std::endl;
    return 0;    
}
```

xlnt uses the pimpl idiom for most of its core data structures. This primary reason for choosing this technique was simplifying usage of the library. Instead of using pointers or references, classes can be passed around by value. Internally they hold a pointer to memory which is within the primary workbook implementation struct. Methods called on the wrapper object dereference the opaque pointer and manipulate its data directly.

For the user, this means that workbooks, worksheets, cells, formats, and styles can be passed and stored by value.
