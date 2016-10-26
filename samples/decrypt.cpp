#include <xlnt/xlnt.hpp>

int main()
{
	xlnt::workbook wb;

	const auto password = std::string("secret");
	wb.load("data/encrypted.xlsx", password);
	wb.save("data/decrypted.xlsx");

	return 0;
}
