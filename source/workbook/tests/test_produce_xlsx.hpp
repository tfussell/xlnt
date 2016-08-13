#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/temporary_file.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_produce_xlsx : public CxxTest::TestSuite
{
public:
	bool workbook_matches_file(xlnt::workbook &wb, const xlnt::path &file)
	{
		std::vector<std::uint8_t> buffer;
		wb.save(buffer);
		wb.save(xlnt::path("a.xlsx"));
		xlnt::zip_file wb_archive(buffer);

		xlnt::zip_file file_archive(file);

		return xml_helper::archives_match(wb_archive, file_archive);
	}

	void test_minimal()
	{
		xlnt::workbook wb = xlnt::workbook::minimal();
		TS_ASSERT(workbook_matches_file(wb, path_helper::get_data_directory("10_minimal.xlsx")));
	}

	void test_empty_excel()
	{
		xlnt::workbook wb = xlnt::workbook::empty_excel();
		TS_ASSERT(workbook_matches_file(wb, path_helper::get_data_directory("8_default-excel.xlsx")));
	}
};
