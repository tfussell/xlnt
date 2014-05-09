#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "../xlnt.h"

class DumpTestSuite : public CxxTest::TestSuite
{
public:
    DumpTestSuite()
    {

    }

    void _get_test_filename()
    {
        /*test_file = NamedTemporaryFile(mode = "w", prefix = "openpyxl.", suffix = ".xlsx", delete = False);
        test_file.close();
        return test_file.name;*/
    }

    //_COL_CONVERSION_CACHE = dict((get_column_letter(i), i) for i in range(1, 18279));

    void test_dump_sheet_title()
    {
        /*test_filename = _get_test_filename();
        wb = Workbook(optimized_write = True);
        ws = wb.create_sheet(title = "Test1");
        wb.save(test_filename);
        wb2 = load_workbook(test_filename, True);
        ws = wb2.get_sheet_by_name("Test1");
        TS_ASSERT_EQUALS("Test1", ws.title);*/
    }

    void test_dump_sheet()
    {
        /*test_filename = _get_test_filename();
        wb = Workbook(optimized_write = True);
        ws = wb.create_sheet();
        letters = [get_column_letter(x + 1) for x in range(20)];
        expected_rows = [];
        for(auto row : range(20))
        {
            expected_rows.append(["%s%d" % (letter, row + 1) for letter in letters]);
            for row in range(20)
            {
                expected_rows.append([(row + 1) for letter in letters]);
                for row in range(10)
                {
                    expected_rows.append([datetime(2010, ((x % 12) + 1), row + 1) for x in range(len(letters))]);
                    for row in range(20)
                    {
                        expected_rows.append(["=%s%d" % (letter, row + 1) for letter in letters]);
                        for row in expected_rows
                        {
                            ws.append(row);
                        }

                        wb.save(test_filename);
                        wb2 = load_workbook(test_filename, True)
                    }

                    ws = wb2.worksheets[0]
                }
            }
        }


        for ex_row, ws_row in zip(expected_rows[:-20], ws.iter_rows())
        {
            for ex_cell, ws_cell in zip(ex_row, ws_row)
            {
                TS_ASSERT_EQUALS(ex_cell, ws_cell.internal_value)

                    os.remove(test_filename)
            }
        }*/
    }

    void test_table_builder()
    {
        //sb = StringTableBuilder()

        //    result = {"a":0, "b" : 1, "c" : 2, "d" : 3}

        //    for letter in sorted(result.keys())
        //    {
        //        for x in range(5)
        //        {
        //            sb.add(letter)

        //                table = dict(sb.get_table())

        //                try
        //            {
        //                result_items = result.items()
        //            }

        //            for key, idx in result_items
        //            {
        //                TS_ASSERT_EQUALS(idx, table[key])
        //            }
        //        }
        //    }
    }

    void test_open_too_many_files()
    {
        //test_filename = _get_test_filename();
        //wb = Workbook(optimized_write = True);

        //for i in range(200) // over 200 worksheets should raise an OSError("too many open files")
        //{
        //    wb.create_sheet();
        //    wb.save(test_filename);
        //    os.remove(test_filename);
        //}
    }

    void test_create_temp_file()
    {
        //f = dump_worksheet.create_temporary_file();

        //if(!osp.isfile(f))
        //{
        //    raise Exception("The file %s does not exist" % f)
        //}
    }

    void test_dump_twice()
    {
        //test_filename = _get_test_filename();

        //wb = Workbook(optimized_write = True);
        //ws = wb.create_sheet();
        //ws.append(["hello"]);

        //wb.save(test_filename);
        //os.remove(test_filename);

        //wb.save(test_filename);
    }
                
    void test_append_after_save()
    {
        //test_filename = _get_test_filename();

        //wb = Workbook(optimized_write = True);
        //ws = wb.create_sheet();
        //ws.append(["hello"]);

        //wb.save(test_filename);
        //os.remove(test_filename);

        //ws.append(["hello"]);
    }
};
