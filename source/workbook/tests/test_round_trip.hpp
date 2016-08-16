#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_round_trip : public CxxTest::TestSuite
{
public:
	bool round_trip_matches_wr(const xlnt::workbook &original)
	{
		std::vector<std::uint8_t> serialized;
		original.save(serialized);

		xlnt::workbook deserialized;
		deserialized.load(serialized);

		return xml_helper::workbooks_match(original, deserialized);
	}

	bool round_trip_matches_rw(const xlnt::path &file)
	{
		xlnt::zip_file original_archive(file);

		xlnt::workbook deserialized;
		deserialized.load(file);

		std::vector<std::uint8_t> serialized;
		deserialized.save(serialized);

		xlnt::zip_file resulting_archive(serialized);

		return xml_helper::archives_match(original_archive, resulting_archive);
	}

	void test_round_trip_minimal_wr()
	{
		xlnt::workbook wb = xlnt::workbook::minimal();
		TS_ASSERT(round_trip_matches_wr(wb));
	}

	void test_round_trip_empty_excel_wr()
	{
		xlnt::workbook wb = xlnt::workbook::empty_excel();
		TS_ASSERT(round_trip_matches_wr(wb));
	}

	void test_round_trip_empty_libre_office_wr()
	{
		TS_SKIP("");
		xlnt::workbook wb = xlnt::workbook::empty_libre_office();
		TS_ASSERT(round_trip_matches_wr(wb));
	}

	void test_round_trip_empty_pages_wr()
	{
		TS_SKIP("");
		xlnt::workbook wb = xlnt::workbook::empty_numbers();
		TS_ASSERT(round_trip_matches_wr(wb));
	}

	void test_round_trip_empty_excel_rw()
	{
		auto path = path_helper::get_data_directory("9_default-excel.xlsx");
		TS_ASSERT(round_trip_matches_rw(path));
	}

	void test_round_trip_empty_libre_rw()
	{
		TS_SKIP("");
		auto path = path_helper::get_data_directory("10_default-libre-office.xlsx");
		TS_ASSERT(round_trip_matches_rw(path));
	}
};
