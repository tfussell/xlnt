#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "helpers/temporary_file.hpp"
#include <xlnt/xlnt.hpp>

class test_dump : public CxxTest::TestSuite
{
public:
    test_dump() : wb(xlnt::optimization::write)
    {
        
    }

    void test_dump_sheet_title()
    {
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
	TS_SKIP("");
        auto test_filename = temp_file.GetFilename();
        auto ws = wb.create_sheet();

        std::vector<std::string> letters;
        for(int i = 0; i < 20; i++)
        {
            letters.push_back(xlnt::cell_reference::column_string_from_index(i + 1));
        }

        for(int row = 0; row < 20; row++)
        {
	    std::vector<std::string> current_row;
            for(auto letter : letters)
            {
                current_row.push_back(letter + std::to_string(row + 1));
            }
	    ws.append(current_row);
        }

        for(int row = 0; row < 20; row++)
        {
	    std::vector<int> current_row;
            for(auto letter : letters)
            {
 	        current_row.push_back(row + 1);
            }
	    ws.append(current_row);
        }

        for(int row = 0; row < 10; row++)
        {
	    std::vector<xlnt::date> current_row;
            for(std::size_t x = 0; x < letters.size(); x++)
            {
	        current_row.push_back(xlnt::date(2010, (x % 12) + 1, row + 1));
            }
	    ws.append(current_row);
        }

        for(int row = 0; row < 20; row++)
        {
	    std::vector<std::string> current_row;
	    for(auto letter : letters)
            {
	      current_row.push_back("=" + letter + std::to_string(row + 1));
            }
	    ws.append(current_row);
        }

        wb.save(test_filename);

        xlnt::workbook wb2;
        wb2.load(test_filename);
        ws = wb[2];

        for(auto row : ws.rows())
        {
            for(auto cell : row)
            {
	        auto row = cell.get_row();

		if(row <= 20)
		{
		    std::string expected = cell.get_reference().to_string();
		    TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::string);
		    TS_ASSERT_EQUALS(cell, expected);
		}
		else if(row <= 40)
		{
		    TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::numeric);
		    TS_ASSERT_EQUALS(cell, (int)row);
		}
		else if(row <= 50)
		{
		    xlnt::date expected(2010, (cell.get_reference().get_column_index() % 12) + 1, row + 1);
		    TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::numeric);
		    TS_ASSERT(cell.is_date());
		    TS_ASSERT_EQUALS(cell, expected);
		}
		else
		{
		    std::string expected = "=" + cell.get_reference().to_string();
		    TS_ASSERT_EQUALS(cell.get_data_type(), xlnt::cell::type::formula);
		    TS_ASSERT_EQUALS(cell, expected);
		}
            }
        }
    }

    void test_table_builder()
    {
        static const std::vector<std::pair<std::string, int>> result = {{"a", 0}, {"b", 1}, {"c", 2}, {"d", 3}};
        xlnt::string_table_builder sb;

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

    void test_dump_twice()
    {
        auto test_filename = temp_file.GetFilename();

        auto ws = wb.create_sheet();

        std::vector<std::string> to_append = {"hello"};
        ws.append(to_append);

        wb.save(test_filename);
        std::remove(test_filename.c_str());
	
        TS_ASSERT_THROWS(wb.save(test_filename), xlnt::workbook_already_saved);
    }
                
    void test_append_after_save()
    {
        auto ws = wb.create_sheet();

        std::vector<std::string> to_append = {"hello"};
        ws.append(to_append);

        wb.save(temp_file.GetFilename());
        std::remove(temp_file.GetFilename().c_str());

        TS_ASSERT_THROWS(ws.append(to_append), xlnt::workbook_already_saved);
    }

private:
    xlnt::workbook wb;
    TemporaryFile temp_file;
};
