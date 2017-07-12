#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/workbook/streaming_workbook_reader.hpp>
#include <xlnt/workbook/streaming_workbook_writer.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/utils/xlntarrow.hpp>

namespace xlnt {
namespace arrow {

void XLNT_API xlsx2arrow(std::istream &s, ::arrow::Table &table)
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

void XLNT_API arrow2xlsx(const ::arrow::Table &table, std::ostream &s)
{
  xlnt::streaming_workbook_writer writer;
  writer.open(s);

  writer.add_worksheet("Sheet1");
  writer.add_cell("A1").value("test");
}

}
}
