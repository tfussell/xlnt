#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.get_active_sheet();
    ws.get_cell("A1").set_value(5);
    ws.get_cell("B2").set_value("string data");
    ws.get_cell("C3").set_formula("=RAND()");
    ws.merge_cells("C3:C4");
    ws.freeze_panes("B2");
    wb.save("data/sample.xlsx");

    return 0;
}
