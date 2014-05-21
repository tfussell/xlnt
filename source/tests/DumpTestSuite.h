#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "TemporaryFile.h"
#include "../xlnt.h"

class DumpTestSuite : public CxxTest::TestSuite
{
public:
    DumpTestSuite()
    {
        
    }

    //_COL_CONVERSION_CACHE = dict((get_column_letter(i), i) for i in range(1, 18279));

    void test_dump_sheet_title()
    {
        xlnt::workbook wb(xlnt::optimization::write);
        auto ws = wb.create_sheet("Test1");
        wb.save(temp_file.GetFilename());
        xlnt::workbook wb2;
        wb2.load(temp_file.GetFilename());
        ws = wb2.get_sheet_by_name("Test1");
        TS_ASSERT_DIFFERS(ws, nullptr);
        TS_ASSERT_EQUALS("Test1", ws.get_title());
    }

    void test_dump_sheet()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb(xlnt::optimization::write);
        auto ws = wb.create_sheet();

        std::vector<std::string> letters;

        for(int i = 0; i < 20; i++)
        {
            letters.push_back(xlnt::cell_reference::column_string_from_index(i + 1));
        }

        std::vector<std::vector<std::string>> expected_rows;

        for(int row = 0; row < 20; row++)
        {
            expected_rows.push_back(std::vector<std::string>());
            for(auto letter : letters)
            {
                expected_rows.back().push_back(letter + std::to_string(row + 1));
            }
        }

        for(int row = 0; row < 20; row++)
        {
            expected_rows.push_back(std::vector<std::string>());
            for(auto letter : letters)
            {
                expected_rows.back().push_back(letter + std::to_string(row + 1));
            }
        }

        for(int row = 0; row < 10; row++)
        {
            expected_rows.push_back(std::vector<std::string>());
            for(auto letter : letters)
            {
                expected_rows.back().push_back(letter + std::to_string(row + 1));
            }
        }

        for(auto row : expected_rows)
        {
            ws.append(row);
        }

        wb.save(test_filename);

        //xlnt::workbook wb2;
        //wb2.load(test_filename);
        //ws = wb2[0];

        auto expected_row = expected_rows.begin();
        for(auto row : ws.rows())
        {
            auto expected_cell = expected_row->begin();
            for(auto cell : row)
            {
                TS_ASSERT_EQUALS(cell, *expected_cell);
                expected_cell++;
            }
            expected_row++;
        }
    }

    void test_table_builder()
    {
        xlnt::string_table_builder sb;

        std::vector<std::pair<std::string, int>> result = {{"a", 0}, {"b", 1}, {"c", 2}, {"d", 3}};

        for(auto pair : result)
        {
            for(int i = 0; i < 5; i++)
            {
                sb.add(pair.first);
            }
        }

        auto table = sb.get_table();

        for(auto pair : result)
        {
            TS_ASSERT_EQUALS(pair.second, table[pair.first]);
        }
    }

    void test_open_too_many_files()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb(xlnt::optimization::write);

        for(int i = 0; i < 2; i++) // over 200 worksheets should raise an OSError("too many open files")
        {
            wb.create_sheet();
            wb.save(test_filename);
	    std::remove(test_filename.c_str());
        }
    }

    void test_dump_twice()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb(xlnt::optimization::write);
        auto ws = wb.create_sheet();

        std::vector<std::string> to_append = {"hello"};
        ws.append(to_append);

        wb.save(test_filename);
        std::remove(test_filename.c_str());
        wb.save(test_filename);
    }
                
    void test_append_after_save()
    {
        xlnt::workbook wb(xlnt::optimization::write);
        auto ws = wb.create_sheet();

        std::vector<std::string> to_append = {"hello"};
        ws.append(to_append);

        {
            TemporaryFile temp2;
            wb.save(temp2.GetFilename());
        }

        ws.append(to_append);
    }

private:
    TemporaryFile temp_file;
};
