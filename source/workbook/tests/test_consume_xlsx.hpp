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
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("14_encrypted_excel_2016.xlsx"), "secret");
	}

	void test_decrypt_libre_office()
	{
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("15_encrypted_libre_office.xlsx"), "secret");
	}

	void test_decrypt_standard()
	{
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("16_encrypted_excel_2007.xlsx"), "password");
	}
    
    void test_decrypt_numbers()
	{
        TS_SKIP("");
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("17_encrypted_numbers.xlsx"), "secret");
	}
};
