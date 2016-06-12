#include <string>
#include <unordered_map>
#include <utility>

#include <xlnt/xlnt.hpp>

int main()
{
    const std::vector<std::pair<std::string, double>> amounts =
    {
        { "Anne", 17.31 },
        { "Brent", 21.99 },
        { "Catelyn", 94.47 },
        { "Diedrich", 101.05 }
    };
    
    xlnt::workbook wb;
    auto ws = wb.get_active_sheet();
    
    ws.get_cell("A1").set_value("Name");
    ws.get_cell("B1").set_value("Amount");
    
    std::size_t row = 2;
    auto money_format = xlnt::number_format::from_builtin_id(44);
    auto &style = wb.create_style("Currency");
    style.set_builtin_id(4);
    style.set_number_format(money_format);
    
    for (const auto &amount : amounts)
    {
        ws.get_cell(xlnt::cell_reference(1, row)).set_value(amount.first);
        ws.get_cell(xlnt::cell_reference(2, row)).set_value(amount.second);
        ws.get_cell(xlnt::cell_reference(2, row)).set_style("Currency");

        row++;
    }

    std::string sum_formula = "=SUM(B2:B" + std::to_string(row - 1) + ")";
    ws.get_cell(xlnt::cell_reference(2, row)).set_style("Currency");
    ws.get_cell(xlnt::cell_reference(2, row)).set_formula(sum_formula);
    
    wb.save("create.xlsx");
    
    return 0;
}
