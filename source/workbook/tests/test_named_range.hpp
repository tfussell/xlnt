#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/workbook/named_range.hpp>

class test_named_range : public CxxTest::TestSuite
{
public:
    void test_split()
    {
    /*
        using string_pair = std::pair<std::string, std::string>;
        using string_pair_vector = std::vector<string_pair>;
        using expected_pair = std::pair<std::string, string_pair_vector>;
        
        std::vector<expected_pair> expected_pairs =
        {
            { "'My Sheet'!$D$8", {{ "My Sheet", "$D$8" }} },
            { "Sheet1!$A$1", {{ "Sheet1", "$A$1" }} },
            { "[1]Sheet1!$A$1", {{ "[1]Sheet1", "$A$1" }} },
            { "[1]!B2range", {{ "[1]", "" }} },
            { "Sheet1!$C$5:$C$7,Sheet1!$C$9:$C$11,Sheet1!$E$5:$E$7,Sheet1!$E$9:$E$11,Sheet1!$D$8",
                {
                    { "Sheet1", "$C$5:$C$7" },
                    { "Sheet1", "$C$9:$C$11" },
                    { "Sheet1", "$E$5:$E$7" },
                    { "Sheet1", "$E$9:$E$11" },
                    { "Sheet1", "$D$8" }
                }
            }
        };

        for(auto pair : expected_pairs)
        {
            TS_ASSERT_EQUALS(xlnt::split_named_range(pair.first), pair.second);
        }
        */
    }

    void test_split_no_quotes()
    {
        /*TS_ASSERT_EQUALS([("HYPOTHESES", "$B$3:$L$3"), ], split_named_range("HYPOTHESES!$B$3:$L$3"))*/
    }

    void test_bad_range_name()
    {
        /*assert_raises(NamedRangeException, split_named_range, "HYPOTHESES$B$3")*/
    }

    void test_range_name_worksheet_special_chars()
    {
        /*class DummyWs
        {
            title = "My Sheeet with a , and \""

                void __str__()
            {
                    return title
            }

            ws = DummyWs()
        };

        class DummyWB
        {
            void get_sheet_by_name(self, name)
            {
                if name == ws.title :
                return ws
            }
        };

        handle = open(os.path.join(DATADIR, "reader", "workbook_namedrange.xml"))
            try
        {
            content = handle.read()
                named_ranges = read_named_ranges(content, DummyWB())
                TS_ASSERT_EQUALS(1, len(named_ranges))
                ok_(isinstance(named_ranges[0], NamedRange))
                TS_ASSERT_EQUALS([(ws, "$U$16:$U$24"), (ws, "$V$28:$V$36")], named_ranges[0].destinations)
        }
        finally
        {
            handle.close()
        }*/
    }


    void test_read_named_ranges()
    {
        /*class DummyWs
        {
            title = "My Sheeet"

                void __str__()
            {
                    return title;
            }
        };

        class DummyWB
        {
            void get_sheet_by_name(self, name)
            {
                return DummyWs()
            }
        };

        handle = open(os.path.join(DATADIR, "reader", "workbook.xml"))
            try :
            content = handle.read()
            named_ranges = read_named_ranges(content, DummyWB())
            TS_ASSERT_EQUALS(["My Sheeet!$D$8"], [str(range) for range in named_ranges])
            finally :
            handle.close()*/
    }

    void test_oddly_shaped_named_ranges()
    {
       /* ranges_counts = ((4, "TEST_RANGE"),
            (3, "TRAP_1"),
            (13, "TRAP_2"))*/
    }

    //void check_ranges(ws, count, range_name)
    //{
    //    /*TS_ASSERT_EQUALS(count, len(ws.range(range_name)))

    //        wb = load_workbook(os.path.join(DATADIR, "genuine", "merge_range.xlsx"),
    //        use_iterators = False)

    //        ws = wb.worksheets[0]

    //        for count, range_name in ranges_counts
    //        {
    //        yield check_ranges, ws, count, range_name
    //        }*/
    //}


    void test_merged_cells_named_range()
    {
        /*wb = load_workbook(os.path.join(DATADIR, "genuine", "merge_range.xlsx"),
            use_iterators = False)

            ws = wb.worksheets[0]

            cell = ws.range("TRAP_3")

            TS_ASSERT_EQUALS("B15", cell.get_coordinate())

            TS_ASSERT_EQUALS(10, cell.value)*/
    }

    void setUp()
    {
        /*wb = load_workbook(os.path.join(DATADIR, "genuine", "NameWithValueBug.xlsx"))
            ws = wb.get_sheet_by_name("Sheet1")
            make_tmpdir()*/
    }

    void tearDown()
    {
       /* clean_tmpdir()*/
    }

    void test_has_ranges()
    {
        /*ranges = wb.get_named_ranges()
            TS_ASSERT_EQUALS(["MyRef", "MySheetRef", "MySheetRef", "MySheetValue", "MySheetValue", "MyValue"], [range.name for range in ranges])*/
    }

    void test_workbook_has_normal_range()
    {
        /*normal_range = wb.get_named_range("MyRef")
            TS_ASSERT_EQUALS("MyRef", normal_range.name)*/
    }

    void test_workbook_has_value_range()
    {
        /*value_range = wb.get_named_range("MyValue")
            TS_ASSERT_EQUALS("MyValue", value_range.name)
            TS_ASSERT_EQUALS("9.99", value_range.value)*/
    }

    void test_worksheet_range()
    {
        /*range = ws.range("MyRef")*/
    }

    void test_worksheet_range_error_on_value_range()
    {
        /*assert_raises(NamedRangeException, ws.range, "MyValue")*/
    }

    //void range_as_string(self, range, include_value = False)
    //{
    //    /*void scope_as_string(range)
    //    {
    //        if range.scope
    //        {
    //            return range.scope.title
    //        }
    //        else
    //            {
    //            return "Workbook"
    //            }
    //    }
    //    retval = "%s: %s" % (range.name, scope_as_string(range))
    //        if include_value :
    //    if isinstance(range, NamedRange) :
    //        retval += "=[range]"
    //    else :
    //    retval += "=" + range.value
    //    return retval*/
    //}

    void test_handles_scope()
    {
        /*ranges = wb.get_named_ranges()
            TS_ASSERT_EQUALS(["MyRef: Workbook", "MySheetRef: Sheet1", "MySheetRef: Sheet2", "MySheetValue: Sheet1", "MySheetValue: Sheet2", "MyValue: Workbook"],
            [range_as_string(range) for range in ranges])*/
    }

    void test_can_be_saved()
    {
        /*FNAME = os.path.join(TMPDIR, "foo.xlsx")
            wb.save(FNAME)

            wbcopy = load_workbook(FNAME)
            TS_ASSERT_EQUALS(["MyRef: Workbook=[range]", "MySheetRef: Sheet1=[range]", "MySheetRef: Sheet2=[range]", "MySheetValue: Sheet1=3.33", "MySheetValue: Sheet2=14.4", "MyValue: Workbook=9.99"],
            [range_as_string(range, include_value = True) for range in wbcopy.get_named_ranges()])*/
    }
};
