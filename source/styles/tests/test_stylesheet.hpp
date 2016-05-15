#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <helpers/helper.hpp>
#include <helpers/path_helper.hpp>

class test_number_style : public CxxTest::TestSuite
{
public:
    void test_ctor()
    {
        xlnt::workbook wb;
        xlnt::style_serializer style_serializer(wb);
        xlnt::zip_file archive;
        auto stylesheet_xml = style_serializer.write_stylesheet().to_string();
        
        const std::string expected =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" mc:Ignorable=\"x14ac\">"
        "  <fonts count=\"1\">"
        "    <font>"
        "      <sz val=\"12\"/>"
        "      <color theme=\"1\"/>"
        "      <name val=\"Calibri\"/>"
        "      <family val=\"2\"/>"
        "      <scheme val=\"minor\"/>"
        "    </font>"
        "  </fonts>"
        "  <fills count=\"1\">"
        "    <fill>"
        "      <patternFill patternType=\"none\"/>"
        "    </fill>"
        "  </fills>"
        "  <borders count=\"1\">"
        "    <border>"
        "      <left/>"
        "      <right/>"
        "      <top/>"
        "      <bottom/>"
        "      <diagonal/>"
        "    </border>"
        "  </borders>"
        "  <cellStyleXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
        "  </cellStyleXfs>"
        "  <cellXfs count=\"1\">"
        "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        "  </cellStyles>"
        "  <dxfs count=\"0\"/>"
        "  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium7\"/>"
        "  <extLst>"
        "    <ext xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\">"
        "      <x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>"
        "    </ext>"
        "  </extLst>"
        "</styleSheet>";
        
        TS_ASSERT(Helper::compare_xml(expected, stylesheet_xml));
    }


    void test_from_simple()
    {
        xlnt::xml_document doc;
        auto xml = PathHelper::read_file(PathHelper::GetDataDirectory("/reader/styles/simple-styles.xml"));
        doc.from_string(xml);
        xlnt::workbook wb;
        xlnt::style_serializer s(wb);
        TS_ASSERT(s.read_stylesheet(doc));
        TS_ASSERT_EQUALS(s.get_number_formats().size(), 1);
    }

    void test_from_complex()
    {
        xlnt::xml_document doc;
        auto xml = PathHelper::read_file(PathHelper::GetDataDirectory("/reader/styles/complex-styles.xml"));
        doc.from_string(xml);
        xlnt::workbook wb;
        xlnt::style_serializer s(wb);
        TS_ASSERT(s.read_stylesheet(doc));
        TS_ASSERT_EQUALS(s.get_borders().size(), 7);
        TS_ASSERT_EQUALS(s.get_fills().size(), 6);
        TS_ASSERT_EQUALS(s.get_fonts().size(), 8);
        TS_ASSERT_EQUALS(s.get_number_formats().size(), 1);
    }

/*
    void _test_unprotected_cell()
    {
        datadir.chdir();
        src = open ("worksheet_unprotected_style.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        styles  = stylesheet.cell_styles;
        assert len(styles) == 3;
        // default is cells are locked
        assert styles[1] == StyleArray([4,0,0,0,0,0,0,0,0]);
        assert styles[2] == StyleArray([3,0,0,0,1,0,0,0,0]);
    }

    void _test_read_cell_style()
    {
        datadir.chdir();
        src = open("empty-workbook-styles.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        styles  = stylesheet.cell_styles;
        assert len(styles) == 2;
        assert styles[1] == StyleArray([0,0,0,9,0,0,0,0,1];
    }

    void _test_read_xf_no_number_format()
    {
        datadir.chdir();
        src = open("no_number_format.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        styles = stylesheet.cell_styles;
        assert len(styles) == 3;
        assert styles[1] == StyleArray([1,0,1,0,0,0,0,0,0]);
        assert styles[2] == StyleArray([0,0,0,14,0,0,0,0,0]);
    }

    void _test_none_values()
    {
        datadir.chdir();
        src = open("none_value_styles.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        fonts = stylesheet.fonts;
        assert fonts[0].scheme is None;
        assert fonts[0].vertAlign is None;
        assert fonts[1].u is None;
    }

    void _test_alignment()
    {
        datadir.chdir();
        src = open("alignment_styles.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        styles = stylesheet.cell_styles;
        assert len(styles) == 3;
        assert styles[2] == StyleArray([0,0,0,0,0,2,0,0,0]);

        assert stylesheet.alignments == [
            Alignment(),
            Alignment(textRotation=180),
            Alignment(vertical='top', textRotation=255),
            ];
    }

    void _test_rgb_colors()
    {
        datadir.chdir();
        src = open("rgb_colors.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        assert len(stylesheet.colors.index) == 64;
        assert stylesheet.colors.index[0] == "00000000";
        assert stylesheet.colors.index[-1] == "00333333";
    }


    void _test_custom_number_formats()
    {
        datadir.chdir();
        src = open("styles_number_formats.xml", "rb");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        assert stylesheet.number_formats == [
            '_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
            "#,##0.00_ ",
            "yyyy/m/d;@",
            "0.00000_ "
        ];
    }

    void _test_assign_number_formats()
    {
        node = fromstring(
        "<styleSheet>"
        "<numFmts count=\"1\">"
        "  <numFmt numFmtId=\"43\" formatCode=\"_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * \"-\"??_ ;_ @_ \" />"
        "</numFmts>"
        "<cellXfs count=\"0\">"
        "<xf numFmtId=\"43\" fontId=\"2\" fillId=\"0\" borderId=\"0\""
        "     applyFont=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">"
        "    <alignment vertical=\"center\"/>"
        "</xf>"
        "</cellXfs>"
        "</styleSheet>");
        stylesheet = Stylesheet.from_tree(node);
        styles = stylesheet.cell_styles;

        assert styles[0] == StyleArray([2, 0, 0, 164, 0, 1, 0, 0, 0]);
    }


    void _test_named_styles()
    {
        datadir.chdir();
        src = open("complex-styles.xml");
        xml = src.read();
        node = fromstring(xml);
        stylesheet = Stylesheet.from_tree(node);

        followed = stylesheet.named_styles['Followed Hyperlink'];
        assert followed.name == "Followed Hyperlink";
        assert followed.font == stylesheet.fonts[2];
        assert followed.fill == DEFAULT_EMPTY_FILL;
        assert followed.border == Border();

        link = stylesheet.named_styles['Hyperlink'];
        assert link.name == "Hyperlink";
        assert link.font == stylesheet.fonts[1];
        assert link.fill == DEFAULT_EMPTY_FILL;
        assert link.border == Border();

        normal = stylesheet.named_styles['Normal'];
        assert normal.name == "Normal";
        assert normal.font == stylesheet.fonts[0];
        assert normal.fill == DEFAULT_EMPTY_FILL;
        assert normal.border == Border();
    }

    void _test_no_styles()
    {
        wb1 = wb2 = Workbook();
        archive = ZipFile(BytesIO(), "a");
        apply_stylesheet(archive, wb1);
        assert wb1._cell_styles == wb2._cell_styles;
        assert wb2._named_styles == wb2._named_styles;
    }

    void _test_write_worksheet()
    {
        wb = Workbook()
        node = write_stylesheet(wb);
        xml = tostring(node);
        
        const std::string expected =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "  <numFmts count=\"0\" />"
        "  <fonts count=\"1\">"
        "    <font>"
        "      <name val=\"Calibri\"></name>"
        "      <family val=\"2\"></family>"
        "      <color theme=\"1\"></color>"
        "      <sz val=\"11\"></sz>"
        "      <scheme val=\"minor\"></scheme>"
        "    </font>"
        "  </fonts>"
        "  <fills count=\"2\">"
        "    <fill>"
        "      <patternFill></patternFill>"
        "    </fill>"
        "    <fill>"
        "      <patternFill patternType=\"gray125\"></patternFill>"
        "    </fill>"
        "  </fills>"
        "  <borders count=\"1\">"
        "    <border>"
        "      <left></left>"
        "      <right></right>"
        "      <top></top>"
        "      <bottom></bottom>"
        "      <diagonal></diagonal>"
        "    </border>"
        "  </borders>"
        "  <cellStyleXfs count=\"1\">"
        "    <xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\"></xf>"
        "  </cellStyleXfs>"
        "  <cellXfs count=\"1\">"
        "    <xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" pivotButton=\"0\" quotePrefix=\"0\" xfId=\"0\"></xf>"
        "  </cellXfs>"
        "  <cellStyles count=\"1\">"
        "    <cellStyle builtinId=\"0\" hidden=\"0\" name=\"Normal\" xfId=\"0\"></cellStyle>"
        "  </cellStyles>"
        "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        "</styleSheet>";
        
        diff = compare_xml(xml, expected);
        assert diff is None, diff;
    }

    void _test_simple_styles() {
        wb = Workbook();
        wb.guess_types = True;
        ws = wb.active;
        now = datetime.date.today();
        for idx, v in enumerate(['12.34%', now, 'This is a test', '31.31415', None], 1)
        {
            ws.append([v]);
            _ = ws.cell(column=1, row=idx).style_id;
        }

        // set explicit formats
        ws['D9'].number_format = numbers.FORMAT_NUMBER_00;
        ws['D9'].protection = Protection(locked=True);
        ws['D9'].style_id;
        ws['E1'].protection = Protection(hidden=True);
        ws['E1'].style_id;

        assert len(wb._cell_styles) == 5;
        stylesheet = write_stylesheet(wb);

        datadir.chdir();
        reference_file = open('simple-styles.xml');
        expected = reference_file.read();
        xml = tostring(stylesheet);
        diff = compare_xml(xml, expected);
        assert diff is None, diff;
    }
    */
};
