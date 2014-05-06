#include "../include/xlnt/reader.h"

int main()
{
    auto wb = xlnt::reader::load_workbook("empty_book.xlsx");
    auto sheet_ranges = wb.get_sheet_by_name("range names");
    std::cout << (std::string)sheet_ranges.cell("D18").value() << std::endl;
}
