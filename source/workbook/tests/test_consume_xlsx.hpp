#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_consume_xlsx : public CxxTest::TestSuite
{
public:
	void test_decrypt_agile()
	{
#ifdef CRYPTO_ENABLED
		xlnt::workbook wb;
		// key is { 125, 177, 16, 188, 228, 63, 60, 108, 222, 145, 79, 49, 49, 13, 157, 98, 217, 33, 52, 3, 222, 124, 200, 242, 179, 57, 61, 97, 225, 195, 141, 242 };
		wb.load(path_helper::get_data_directory("14_encrypted_excel_2016.xlsx"), "secret");
#endif
	}

	void test_decrypt_standard()
	{
#ifdef CRYPTO_ENABLED
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("16_encrypted_excel_2007.xlsx"), "password");
#endif
	}
};
