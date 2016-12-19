#include <iostream>
#include <xlnt/xlnt.hpp>

// Read sample1.xlsx and print out a 2-dimensional
// representation of each sheet. Cells are separated by commas.
// Each new line is a new row.
int main()
{
    // Create a new workbook by reading sample1.xlsx in the current directory.
    xlnt::workbook wb;
    wb.load("sample1.xlsx");
    
    // The workbook class has begin and end methods so it can be iterated upon.
    // Each item is a sheet in the workbook.
    for(const auto sheet : wb)
    {
        // Print the title of the sheet on its own line.
        std::cout << sheet.get_title() << ": " << std::endl;

        // Iterating on a range, such as from worksheet::rows, yields cell_vectors.
        // Cell vectors don't actually contain cells to reduce overhead.
        // Instead they hold a reference to a worksheet and the current cell_reference.
        // Internally, calling worksheet::get_cell with the current cell_reference yields the next cell.
        // This allows easy and fast iteration over a row (sometimes a column) in the worksheet.
        for(auto row : sheet.rows())
        {
            for(auto cell : row)
            {
                // cell::operator<< adds a string represenation of the cell's value to the stream.
                std::cout << cell << ", ";
            }

            std::cout << std::endl;
        }
    }
}
