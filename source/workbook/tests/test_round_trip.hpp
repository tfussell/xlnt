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
        original_workbook.save("round_trip_out.xlsx");

		return xml_helper::xlsx_archives_match(original_data, buffer);
	}

	void test_round_trip_empty_excel_rw()
	{
        const auto files = std::vector<std::string>
        {
            "2_minimal",
            "3_default",
            "4_every_style",
            "5_encrypted_agile",
            "6_encrypted_libre",
            "7_encrypted_standard",
            "8_encrypted_numbers",
            "10_comments_hyperlinks_formulae",
            "11_print_settings",
            "12_advanced_properties"
        };

        for (const auto file : files)
        {
            auto path = path_helper::data_directory(file + ".xlsx");
            TS_ASSERT(round_trip_matches_rw(path));
        }
	}
};
