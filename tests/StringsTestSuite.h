#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class StringsTestSuite : public CxxTest::TestSuite
{
public:
    StringsTestSuite()
    {

    }

    void test_create_string_table()
    {
        /*wb = Workbook()
        ws = wb.create_sheet()
        ws.cell("B12").value = "hello"
            ws.cell("B13").value = "world"
            ws.cell("D28").value = "hello"
            table = create_string_table(wb)
            TS_ASSERT_EQUALS({"hello": 1, "world" : 0}, table)*/
    }

    void test_read_string_table()
    {
        /*handle = open(os.path.join(DATADIR, "reader", "sharedStrings.xml"))
            try :
            content = handle.read()
            string_table = read_string_table(content)
            TS_ASSERT_EQUALS({0: "This is cell A1 in Sheet 1", 1 : "This is cell G5"}, string_table)
            finally :
            handle.close()*/
    }

    void test_empty_string()
    {
        /*handle = open(os.path.join(DATADIR, "reader", "sharedStrings-emptystring.xml"))
            try :
            content = handle.read()
            string_table = read_string_table(content)
            TS_ASSERT_EQUALS({0: "Testing empty cell", 1 : ""}, string_table)
            finally :
            handle.close()*/
    }

    void test_formatted_string_table()
    {
        /*handle = open(os.path.join(DATADIR, "reader", "shared-strings-rich.xml"))
            try :
            content = handle.read()
            string_table = read_string_table(content)
            TS_ASSERT_EQUALS({0: "Welcome", 1 : "to the best shop in town",
            2 : "     let"s play "}, string_table)
            finally :
            handle.close()*/
    }
};
