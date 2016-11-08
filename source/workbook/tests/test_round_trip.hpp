#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/vector_streambuf.hpp>
#include <helpers/path_helper.hpp>
#include <helpers/xml_helper.hpp>
#include <xlnt/workbook/workbook.hpp>

class test_round_trip : public CxxTest::TestSuite
{
public:
    /// <summary>
    /// Write original to a memory as an XLSX-formatted ZIP, read it from memory back to a workbook,
    /// then ensure that the resulting workbook matches the original.
    /// </summary>
	bool round_trip_matches_wrw(const xlnt::workbook &original)
	{
		std::vector<std::uint8_t> original_buffer;
		original.save(original_buffer);
        original.save("b.xlsx");

		xlnt::workbook resulting_workbook;
		resulting_workbook.load(original_buffer);
        
        std::vector<std::uint8_t> resulting_buffer;
        resulting_workbook.save(resulting_buffer);
        resulting_workbook.save("a.xlsx");

		return xml_helper::xlsx_archives_match(original_buffer, resulting_buffer);
	}

    /// <summary>
    /// Read file as an XLSX-formatted ZIP file in the filesystem to a workbook,
    /// write the workbook back to memory, then ensure that the contents of the two files are equivalent.
    /// </summary>
	bool round_trip_matches_rw(const xlnt::path &original)
	{
        std::ifstream file_stream(original.string(), std::ios::binary);
        std::vector<std::uint8_t> original_data;

        {
            xlnt::detail::vector_ostreambuf file_data_buffer(original_data);
            std::ostream file_data_stream(&file_data_buffer);
            file_data_stream << file_stream.rdbuf();
        }

		xlnt::workbook original_workbook;
		original_workbook.load(original);
        
        std::vector<std::uint8_t> buffer;
        original_workbook.save(buffer);

		return xml_helper::xlsx_archives_match(original_data, buffer);
	}

	void test_round_trip_minimal_wrw()
	{
		xlnt::workbook wb = xlnt::workbook::minimal();
		TS_ASSERT(round_trip_matches_wrw(wb));
	}

	void test_round_trip_empty_excel_wrw()
	{
		xlnt::workbook wb = xlnt::workbook::empty_excel();
		TS_ASSERT(round_trip_matches_wrw(wb));
	}

	void test_round_trip_empty_libre_office_wrw()
	{
		TS_SKIP("");
		xlnt::workbook wb = xlnt::workbook::empty_libre_office();
		TS_ASSERT(round_trip_matches_wrw(wb));
	}

	void test_round_trip_empty_pages_wrw()
	{
		TS_SKIP("");
		xlnt::workbook wb = xlnt::workbook::empty_numbers();
		TS_ASSERT(round_trip_matches_wrw(wb));
	}

	void test_round_trip_minimal_rw()
	{
		auto path = path_helper::get_data_directory("8_minimal.xlsx");
		TS_ASSERT(round_trip_matches_rw(path));
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

	void test_round_trip_all_styles_rw()
	{
		auto path = path_helper::get_data_directory("13_all_styles.xlsx");
		xlnt::workbook original_workbook;
		original_workbook.load(path);
        
        std::vector<std::uint8_t> buffer;
        original_workbook.save(buffer);
		//TS_ASSERT(round_trip_matches_rw(path));
	}
};
