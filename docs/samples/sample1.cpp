#include <iostream>
#include <xlnt/xlnt.hpp>

int main()
{
	for(auto sheet : xlnt::reader::load_workbook("book.xlsx"))
	{
		std::cout << sheet.get_title() << ": " << std::endl;

		for(auto row : sheet.rows())
		{
			for(auto cell : row)
			{
				std::cout << cell.get_value().as<std::string>() << ", ";
			}

			std::cout << std::endl;
		}
	}

	xlnt::workbook workbook;

	for(int i = 0; i < 3; i++)
	{
		auto sheet = workbook.create_sheet("Sheet" + std::to_string(i));

		for(int row = 0; row < 100; row++)
		{
			for(int column = 0; column < 100; column++)
			{
				sheet[xlnt::cell_reference(column, row)].set_value(row * 100 + column);
			}
		}
	}

	workbook.save("book2.xlsx");
}
