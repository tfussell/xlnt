#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_style : public CxxTest::TestSuite
{
public:
    void test_create_style_table()
    {
        //TS_ASSERT_EQUALS(3, len(writer.style_table));
    }

    void test_write_style_table()
    {
        //reference_file = os.path.join(DATADIR, "writer", "expected", "simple-styles.xml");
        //assert_equals_file_content(reference_file, writer.write_table());
    }

    void setUp_writer()
    {
        //workbook = Workbook()
        //    worksheet = workbook.create_sheet()
    }

    void test_no_style()
    {

        //w = StyleWriter(workbook)
        //    TS_ASSERT_EQUALS(0, len(w.style_table))
    }

    void test_nb_style()
    {
        //for i in range(1, 6)
        //{
        //    worksheet.cell(row = 1, column = i).style.font.size += i
        //        w = StyleWriter(workbook)
        //        TS_ASSERT_EQUALS(5, len(w.style_table))

        //        worksheet.cell("A10").style.borders.top = Border.BORDER_THIN
        //        w = StyleWriter(workbook)
        //        TS_ASSERT_EQUALS(6, len(w.style_table))
        //}
    }

    void test_style_unicity()
    {
        //for i in range(1, 6)
        //{
        //    worksheet.cell(row = 1, column = i).style.font.bold = True
        //        w = StyleWriter(workbook)
        //        TS_ASSERT_EQUALS(1, len(w.style_table))
        //}
    }

    void test_fonts()
    {
        //worksheet.cell("A1").style.font.size = 12
        //    worksheet.cell("A1").style.font.bold = True
        //    w = StyleWriter(workbook)
        //    w._write_fonts()
        //    TS_ASSERT_EQUALS(get_xml(w._root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /><u val="none" /></font></fonts></styleSheet>")

        //    worksheet.cell("A1").style.font.underline = Font.UNDERLINE_SINGLE
        //    w = StyleWriter(workbook)
        //    w._write_fonts()
        //    TS_ASSERT_EQUALS(get_xml(w._root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11" /><color theme="1" /><name val="Calibri" /><family val="2" /><scheme val="minor" /></font><font><sz val="12" /><color rgb="FF000000" /><name val="Calibri" /><family val="2" /><b /><u /></font></fonts></styleSheet>")
    }

    void test_fills()
    {
        //worksheet.cell("A1").style.fill.fill_type = "solid"
        //    worksheet.cell("A1").style.fill.start_color.index = Color.DARKYELLOW
        //    w = StyleWriter(workbook)
        //    w._write_fills()
        //    TS_ASSERT_EQUALS(get_xml(w._root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fills count="3"><fill><patternFill patternType="none" /></fill><fill><patternFill patternType="gray125" /></fill><fill><patternFill patternType="solid"><fgColor rgb="FF808000" /></patternFill></fill></fills></styleSheet>")
    }

    void test_borders()
    {
        //worksheet.cell("A1").style.borders.top.border_style = Border.BORDER_THIN
        //    worksheet.cell("A1").style.borders.top.color.index = Color.DARKYELLOW
        //    w = StyleWriter(workbook)
        //    w._write_borders()
        //    TS_ASSERT_EQUALS(get_xml(w._root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><borders count="2"><border><left /><right /><top /><bottom /><diagonal /></border><border><left style="none"><color rgb="FF000000" /></left><right style="none"><color rgb="FF000000" /></right><top style="thin"><color rgb="FF808000" /></top><bottom style="none"><color rgb="FF000000" /></bottom><diagonal style="none"><color rgb="FF000000" /></diagonal></border></borders></styleSheet>")
    }

    void test_write_cell_xfs_1()
    {
        //worksheet.cell("A1").style.font.size = 12
        //    w = StyleWriter(workbook)
        //    ft = w._write_fonts()
        //    nft = w._write_number_formats()
        //    w._write_cell_xfs(nft, ft, {}, {})
        //    xml = get_xml(w._root)
        //    ok_("applyFont="1"" in xml)
        //    ok_("applyFillId="1"" not in xml)
        //    ok_("applyBorder="1"" not in xml)
        //    ok_("applyAlignment="1"" not in xml)
    }

    void test_alignment()
    {
        //worksheet.cell("A1").style.alignment.horizontal = "center"
        //    worksheet.cell("A1").style.alignment.vertical = "center"
        //    w = StyleWriter(workbook)
        //    nft = w._write_number_formats()
        //    w._write_cell_xfs(nft, {}, {}, {})
        //    xml = get_xml(w._root)
        //    ok_("applyAlignment="1"" in xml)
        //    ok_("horizontal="center"" in xml)
        //    ok_("vertical="center"" in xml)
    }

    void test_alignment_rotation()
    {
        //worksheet.cell("A1").style.alignment.vertical = "center"
        //    worksheet.cell("A1").style.alignment.text_rotation = 90
        //    worksheet.cell("A2").style.alignment.vertical = "center"
        //    worksheet.cell("A2").style.alignment.text_rotation = 135
        //    w = StyleWriter(workbook)
        //    nft = w._write_number_formats()
        //    w._write_cell_xfs(nft, {}, {}, {})
        //    xml = get_xml(w._root)
        //    ok_("textRotation="90"" in xml)
        //    ok_("textRotation="135"" in xml)
    }


    void test_format_comparisions()
    {
        //format1 = NumberFormat()
        //    format2 = NumberFormat()
        //    format3 = NumberFormat()
        //    format1.format_code = "m/d/yyyy"
        //    format2.format_code = "m/d/yyyy"
        //    format3.format_code = "mm/dd/yyyy"
        //    assert not format1 < format2
        //    assert format1 < format3
        //    assert format1 == format2
        //    assert format1 != format3
    }


    void test_builtin_format()
    {
        //format = NumberFormat()
        //    format.format_code = "0.00"
        //    TS_ASSERT_EQUALS(format.builtin_format_code(2), format._format_code)
    }


    void test_read_style()
    {
        //reference_file = os.path.join(DATADIR, "reader", "simple-styles.xml")

        //    handle = open(reference_file, "r")
        //    try :
        //    content = handle.read()
        //    finally :
        //    handle.close()
        //    style_table = read_style_table(content)
        //    TS_ASSERT_EQUALS(4, len(style_table))
        //    TS_ASSERT_EQUALS(NumberFormat._BUILTIN_FORMATS[9],
        //    style_table[1].number_format.format_code)
        //    TS_ASSERT_EQUALS("yyyy-mm-dd", style_table[2].number_format.format_code)
    }


    void test_read_cell_style()
    {
        //reference_file = os.path.join(
        //    DATADIR, "reader", "empty-workbook-styles.xml")
        //    handle = open(reference_file, "r")
        //    try :
        //    content = handle.read()
        //    finally :
        //    handle.close()
        //    style_table = read_style_table(content)
        //    TS_ASSERT_EQUALS(2, len(style_table))
    }
};
