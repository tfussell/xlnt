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
        xlnt::workbook wb;
        wb.optimized_write(true);
        auto ws = wb.create_sheet("Test1");
        wb.save(temp_file.GetFilename());
        xlnt::workbook wb2;
        wb2.load(temp_file.GetFilename());
        ws = wb2.get_sheet_by_name("Test1");
        TS_ASSERT_EQUALS("Test1", ws.get_title());
    }

    void test_dump_sheet()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb;
        wb.optimized_write(true);
        auto ws = wb.create_sheet();

        std::vector<std::string> letters;

        for(int i = 0; i < 20; i++)
        {
            letters.push_back(xlnt::cell::get_column_letter(i + 1));
        }

        std::vector<xlnt::cell> expected_rows;

        for(int row = 0; row < 20; row++)
        {
            expected_rows.append(["%s%d" % (letter, row + 1) for letter in letters]);
            for(auto row in range(20))
            {
                expected_rows.append([(row + 1) for letter in letters]);
                for(auto row in range(10))
                {
                    expected_rows.append([datetime(2010, ((x % 12) + 1), row + 1) for x in range(len(letters))]);
                    for(auto row in range(20))
                    {
                        expected_rows.append(["=%s%d" % (letter, row + 1) for letter in letters]);
                        for(auto row in expected_rows)
                        {
                            ws.append(row);
                        }

                        wb.save(test_filename);
                        wb2 = load_workbook(test_filename, True);
                    }

                    ws = wb2.worksheets[0];
                }
            }
        }


        for(auto ex_row, ws_row : zip(expected_rows[:-20], ws.iter_rows()))
        {
            for(auto ex_cell, ws_cell : zip(ex_row, ws_row))
            {
                TS_ASSERT_EQUALS(ex_cell, ws_cell.internal_value);

                os.remove(test_filename);
            }
        }
    }

    void test_table_builder()
    {
        StringTableBuilder sb;

        std::unordered_map<std::string, int> result = {{"a", 0}, {"b", 1}, {"c", 2}, {"d", 3}};

        for(auto pair : result)
        {
            for(int i = 0; i < 5; i++)
            {
                sb.add(pair.first);

                auto table = sb.get_table();

                try
                {
                    result_items = result.items();
                }

                for key, idx in result_items
                {
                    TS_ASSERT_EQUALS(idx, table[key])
                }
            }
        }
    }

    void test_open_too_many_files()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb;
        wb.optimized_write(true);

        for(int i = 0; i < 200; i++) // over 200 worksheets should raise an OSError("too many open files")
        {
            wb.create_sheet();
            wb.save(test_filename);
            unlink(test_filename.c_str());
        }
    }

    void test_create_temp_file()
    {
        f = dump_worksheet.create_temporary_file();

        if(!osp.isfile(f))
        {
            raise Exception("The file %s does not exist" % f)
        }
    }

    void test_dump_twice()
    {
        auto test_filename = temp_file.GetFilename();

        xlnt::workbook wb;
        wb.optimized_write(true);
        auto ws = wb.create_sheet();

        std::vector<std::string> to_append = {"hello"};
        ws.append(to_append);

        wb.save(test_filename);
        unlink(test_filename.c_str());
        wb.save(test_filename);
    }
                
    void test_append_after_save()
    {
        xlnt::workbook wb;
        wb.optimized_write(true);
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
