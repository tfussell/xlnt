#include <xlnt/reader/workbook_reader.hpp>
#include <xlnt/workbook/workbook.hpp>

namespace xlnt {

xlnt::workbook load_workbook(const std::vector<std::uint8_t> &bytes)
{
    xlnt::workbook wb;
    wb.load(bytes);
    return wb;
}

}
