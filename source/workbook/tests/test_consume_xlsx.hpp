#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_consume_xlsx : public CxxTest::TestSuite
{
public:
	void test_consume_password_protected()
	{
#ifdef CRYPTO_ENABLED
		xlnt::workbook wb;
		wb.load(path_helper::get_data_directory("14_encrypted.xlsx"), "password");
#endif
	}
};
