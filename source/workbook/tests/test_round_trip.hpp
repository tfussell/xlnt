#pragma once

#include <fstream>
#include <iostream>
#include <cxxtest/TestSuite.h>

#include <detail/vector_streambuf.hpp>
#include <detail/crypto/xlsx_crypto.hpp>
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
	bool round_trip_matches_rw(const xlnt::path &source)
	{
		xlnt::workbook source_workbook;
        source_workbook.load(source);

        std::vector<std::uint8_t> destination;
        source_workbook.save(destination);
        source_workbook.save("temp.xlsx");

        std::ifstream source_stream(source.string(), std::ios::binary);

		return xml_helper::xlsx_archives_match(xlnt::detail::to_vector(source_stream), destination);
	}

	bool round_trip_matches_rw(const xlnt::path &source, const std::string &password)
	{
		xlnt::workbook source_workbook;
        source_workbook.load(source, password);

        std::vector<std::uint8_t> destination;
        source_workbook.save(destination);

        std::ifstream source_stream(source.string(), std::ios::binary);
        const auto source_decrypted = xlnt::detail::decrypt_xlsx(
            xlnt::detail::to_vector(source_stream), password);

		return xml_helper::xlsx_archives_match(source_decrypted, destination);
    }

	void test_round_trip_rw()
	{
        const auto files = std::vector<std::string>
        {
            "2_minimal",
            "3_default",
            "4_every_style",
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

	void test_round_trip_rw_encrypted()
	{
        const auto files = std::vector<std::string>
        {
            "5_encrypted_agile",
            "6_encrypted_libre",
            "7_encrypted_standard",
            "8_encrypted_numbers"
        };

        for (const auto file : files)
        {
            auto path = path_helper::data_directory(file + ".xlsx");
            TS_ASSERT(round_trip_matches_rw(path, file
                == "7_encrypted_standard" ? "password" : "secret"));
        }
	}
};
