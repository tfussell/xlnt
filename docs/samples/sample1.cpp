
int main()
{
	for(auto sheet : xlnt::load_workbook("book.xlsx"))
	{
		std::cout << sheet.get_title() << ": " << std::endl;

		for(auto row : sheet)
		{
			for(auto cell : row)
			{
				std::cout << cell << ", ";
			}

			std::cout << std::endl;
		}
	}

	xlnt::workbook workbook;

	for(int i = 0; i < 3; i++)
	{
		auto sheet = workbook.create_sheet("Sheet" + std::to_string(i));

		for(int row = 0; row < 1000; row++)
		{
			for(int column = 0; column < 1000; column++)
			{
				sheet[xlnt::cell_reference(column, row)] = row * 1000 + column;
			}
		}
	}

	workbook.save("book2.xlsx");
}
