#include <xlnt/xlnt.hpp>
#include <xlntarrow.hpp>

namespace xlnt {
namespace arrow {

void xlsx2arrow(std::istream &s, ::arrow::Table &table)
{
    xlnt::streaming_workbook_reader reader;
    reader.open(s);

    reader.begin_worksheet();
    int first_row = 0;

    while (reader.has_cell())
    {
        auto cell = reader.read_cell();

        if (first_row < 1)
        {
            first_row = cell.row();
        }

        if (cell.reference().row() % 1000 == 1)
        {
            std::cout << cell.reference().to_string() << std::endl;
        }
    }

    reader.end_worksheet();
}

void arrow2xlsx(const ::arrow::Table &table, std::istream &s)
{

}

}
}
