#include <iostream>
#include <xlnt/xlnt.hpp>

// Make 3 sheets containing 10,000 cells with unique numeric values.
int main()
{
    // create the workbook
    xlnt::workbook workbook;
    
    // workbooks have a single sheet called "Sheet" initially so let's get rid of it
    auto to_remove = workbook.get_sheet_by_name("Sheet");
    workbook.remove_sheet(to_remove);

    // this loop will create three sheets and populate them with data
    for(int i = 0; i < 3; i++)
    {
        // the title will be "Sample2-1", "Sample2-2", or "Sample2-3"
        // if we don't specify a title, these would be "Sheet#" where
        // # is the lowest number that doesn't have a corresponding sheet yet.
        auto sheet = workbook.create_sheet("Sample2-" + std::to_string(i));

        for(int row = 1; row < 101; row++)
        {
            for(int column = 1; column < 101; column++)
            {
                // Since we can't overload subscript to accept both number,
                // create a cell_reference which "points" to the current cell.
                xlnt::cell_reference ref(column, row);
                
                // This is important!
                // The cell class is really just a wrapper around a pointer.
                // For this reason, we can store them by value and still modify
                // the data in the containing worksheet.
                auto cell = sheet[ref];
                
                // set_value has overloads for many types such as strings and ints
                cell.set_value(row * 100 + column);
            }
        }
    }

    // This will be written to the current directory.
    workbook.save("sample2.xlsx");
}
