#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include "pugixml.hpp"
#include <xlnt/xlnt.hpp>

class test_worksheet : public CxxTest::TestSuite
{
public:
    void test_new_worksheet()
    {
        xlnt::worksheet ws = wb_.create_sheet();
	TS_ASSERT(wb_ == ws.get_parent());
    }

    void test_new_sheet_name()
    {
	xlnt::worksheet ws = wb_.create_sheet("TestName");
        TS_ASSERT_EQUALS(ws.to_string(), "<Worksheet \"TestName\">");
    }

    void test_get_cell()
    {
	xlnt::worksheet ws(wb_);
        auto cell = ws.get_cell("A1");
        TS_ASSERT_EQUALS(cell.get_reference().to_string(), "A1");
    }

    void test_set_bad_title()
    {
        std::string title(50, 'X');
	TS_ASSERT_THROWS(wb_.create_sheet(title), xlnt::sheet_title_exception);
    }
    
    void test_increment_title()
    {
    	auto ws1 = wb_.create_sheet("Test");
    	TS_ASSERT_EQUALS(ws1.get_title(), "Test");
    	auto ws2 = wb_.create_sheet("Test");
    	TS_ASSERT_EQUALS(ws2.get_title(), "Test1");
    }

    void test_set_bad_title_character()
    {
        TS_ASSERT_THROWS(wb_.create_sheet("["), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("]"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("*"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet(":"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("?"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("/"), xlnt::sheet_title_exception);
        TS_ASSERT_THROWS(wb_.create_sheet("\\"), xlnt::sheet_title_exception);
    }

    void test_worksheet_dimension()
    {
        xlnt::worksheet ws(wb_);
        ws.get_cell("A1") = "AAA";
        TS_ASSERT_EQUALS("A1:A1", ws.calculate_dimension().to_string());
        ws.get_cell("B12") = "AAA";
        TS_ASSERT_EQUALS("A1:B12", ws.calculate_dimension().to_string());
    }

    void test_worksheet_range()
    {
        xlnt::worksheet ws(wb_);
        auto xlrange = ws.get_range("A1:C4");
        TS_ASSERT_EQUALS(4, xlrange.length());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
    }

    void test_worksheet_named_range()
    {
        xlnt::worksheet ws(wb_);
        wb_.create_named_range("test_range", ws, "C5");
        auto xlrange = ws.get_named_range("test_range");
        TS_ASSERT_EQUALS(1, xlrange.length());
        TS_ASSERT_EQUALS(1, xlrange[0].num_cells());
        TS_ASSERT_EQUALS(5, xlrange[0][0].get_row());
    }

    void test_bad_named_range()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_THROWS(ws.get_named_range("bad_range"), xlnt::named_range_exception);
    }

    void test_named_range_wrong_sheet()
    {
        xlnt::worksheet ws1(wb_);
        xlnt::worksheet ws2(wb_);
        wb_.create_named_range("wrong_sheet_range", ws1, "C5");
        TS_ASSERT_THROWS(ws2.get_named_range("wrong_sheet_range"), xlnt::named_range_exception);
    }

    void test_cell_offset()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS("C17", ws.get_cell(xlnt::cell_reference("B15").make_offset(1, 2)).get_reference().to_string());
    }

    void test_range_offset()
    {
        xlnt::worksheet ws(wb_);
        auto xlrange = ws.get_range(xlnt::range_reference("A1:C4").make_offset(3, 1));
        TS_ASSERT_EQUALS(4, xlrange.length());
        TS_ASSERT_EQUALS(3, xlrange[0].num_cells());
        TS_ASSERT_EQUALS("D2", xlrange[0][0].get_reference().to_string());
    }

    void test_cell_alternate_coordinates()
    {
        xlnt::worksheet ws(wb_);
        auto cell = ws.get_cell(xlnt::cell_reference(4, 8));
        TS_ASSERT_EQUALS("E9", cell.get_reference().to_string());
    }

    void test_cell_range_name()
    {
        xlnt::worksheet ws(wb_);
        wb_.create_named_range("test_range_single", ws, "B12");
        TS_ASSERT_THROWS(ws.get_cell("test_range_single"), xlnt::cell_coordinates_exception);
        auto c_range_name = ws.get_named_range("test_range_single");
        auto c_range_coord = ws.get_range("B12");
        auto c_cell = ws.get_cell("B12");
        TS_ASSERT_EQUALS(c_range_coord, c_range_name);
        TS_ASSERT(c_range_coord[0][0] == c_cell);
    }

    void test_garbage_collect()
    {
        xlnt::worksheet ws(wb_);

        ws.get_cell("A1").set_null();
        ws.get_cell("B2") = "0";
        ws.get_cell("C4") = 0;
        ws.get_cell("D1").set_comment(xlnt::comment("Comment", "Comment"));

        ws.garbage_collect();

        auto cell_collection = ws.get_cell_collection();
        std::set<xlnt::cell> cells(cell_collection.begin(), cell_collection.end());
        std::set<xlnt::cell> expected = {ws.get_cell("B2"), ws.get_cell("C4"), ws.get_cell("D1")};

        // Set difference
        std::set<xlnt::cell> difference;

        for(auto a : expected)
        {
            if(cells.find(a) == cells.end())
            {
                difference.insert(a);
            }
        }

        for(auto a : cells)
        {
            if(expected.find(a) == expected.end())
            {
                difference.insert(a);
            }
        }

        TS_ASSERT(difference.empty());
    }

    void test_hyperlink_relationships()
    {
        xlnt::worksheet ws(wb_);
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 0);

        ws.get_cell("A1").set_hyperlink("http://test.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 1);
        TS_ASSERT_EQUALS("rId1", ws.get_cell("A1").get_hyperlink().get_id());
        TS_ASSERT_EQUALS("rId1", ws.get_relationships()[0].get_id());
        TS_ASSERT_EQUALS("http://test.com", ws.get_relationships()[0].get_target_uri());
        TS_ASSERT_EQUALS(xlnt::target_mode::external, ws.get_relationships()[0].get_target_mode());

        ws.get_cell("A2").set_hyperlink("http://test2.com");
        TS_ASSERT_EQUALS(ws.get_relationships().size(), 2);
        TS_ASSERT_EQUALS("rId2", ws.get_cell("A2").get_hyperlink().get_id());
        TS_ASSERT_EQUALS("rId2", ws.get_relationships()[1].get_id());
        TS_ASSERT_EQUALS("http://test2.com", ws.get_relationships()[1].get_target_uri());
        TS_ASSERT_EQUALS(xlnt::target_mode::external, ws.get_relationships()[1].get_target_mode());
    }

    void test_bad_relationship_type()
    {
        xlnt::relationship rel("bad");
    }

    void test_append_list()
    {
        xlnt::worksheet ws(wb_);

        ws.append(std::vector<std::string> {"This is A1", "This is B1"});

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is B1", ws.get_cell("B1"));
    }

    void test_append_dict_letter()
    {
        xlnt::worksheet ws(wb_);

        ws.append(std::unordered_map<std::string, std::string> {{"A", "This is A1"}, {"C", "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1"));
    }

    void test_append_dict_index()
    {
        xlnt::worksheet ws(wb_);

        ws.append(std::unordered_map<int, std::string> {{0, "This is A1"}, {2, "This is C1"}});

        TS_ASSERT_EQUALS("This is A1", ws.get_cell("A1"));
        TS_ASSERT_EQUALS("This is C1", ws.get_cell("C1"));
    }

    void test_append_2d_list()
    {
        xlnt::worksheet ws(wb_);

        ws.append(std::vector<std::string> {"This is A1", "This is B1"});
        ws.append(std::vector<std::string> {"This is A2", "This is B2"});

        auto vals = ws.get_range("A1:B2");

        TS_ASSERT_EQUALS(vals[0][0], "This is A1");
        TS_ASSERT_EQUALS(vals[0][1], "This is B1");
        TS_ASSERT_EQUALS(vals[1][0], "This is A2");
        TS_ASSERT_EQUALS(vals[1][1], "This is B2");
    }

    void test_rows()
    {
        xlnt::worksheet ws(wb_);

        ws.get_cell("A1") = "first";
        ws.get_cell("C9") = "last";

        auto rows = ws.rows();

        TS_ASSERT_EQUALS(rows.length(), 9);

        TS_ASSERT_EQUALS(rows[0][0], "first");
        TS_ASSERT_EQUALS(rows[8][2], "last");
    }
    
    void test_cols()
    {
        xlnt::worksheet ws(wb_);

        ws.get_cell("A1") = "first";
        ws.get_cell("C9") = "last";

        auto cols = ws.columns();

        TS_ASSERT_EQUALS(cols.length(), 3);

        TS_ASSERT_EQUALS(cols[0][0], "first");
        TS_ASSERT_EQUALS(cols[2][8], "last");
    }

    void test_auto_filter()
    {
        xlnt::worksheet ws(wb_);

        ws.auto_filter(ws.get_range("a1:f1"));
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "A1:F1");

        ws.unset_auto_filter();
        TS_ASSERT_EQUALS(ws.has_auto_filter(), false);

        ws.auto_filter("c1:g9");
        TS_ASSERT_EQUALS(ws.get_auto_filter(), "C1:G9");
    }

    void test_freeze()
    {
        xlnt::worksheet ws(wb_);

        ws.freeze_panes(ws.get_cell("b2"));
        TS_ASSERT_EQUALS(ws.get_frozen_panes().to_string(), "B2");

        ws.unfreeze_panes();
        TS_ASSERT(!ws.has_frozen_panes());

        ws.freeze_panes("c5");
        TS_ASSERT_EQUALS(ws.get_frozen_panes().to_string(), "C5");

        ws.freeze_panes(ws.get_cell("A1"));
        TS_ASSERT(!ws.has_frozen_panes());
    }

    void test_write_empty()
    {
    	TS_SKIP("not implemented");
    	
        xlnt::worksheet ws(wb_);
        
        auto xml_string = xlnt::writer::write_worksheet(ws);

        pugi::xml_document doc;
        doc.load(xml_string.c_str());
    }
    
    void test_page_margins()
    {
        TS_SKIP("not implemented");
    	
        xlnt::worksheet ws(wb_);
        
        auto xml_string = xlnt::writer::write_worksheet(ws);

        pugi::xml_document doc;
        doc.load(xml_string.c_str());
    }
   
    void test_merge()
    {
        TS_SKIP("not implemented");
    	
        xlnt::worksheet ws(wb_);
        
        auto xml_string = xlnt::writer::write_worksheet(ws);

        pugi::xml_document doc;
        doc.load(xml_string.c_str());
    }

    void test_printer_settings()
    {
        TS_SKIP("not implemented");
    	
        xlnt::worksheet ws(wb_);

        ws.get_page_setup().set_orientation(xlnt::page_setup::orientation::landscape);
        ws.get_page_setup().set_paper_size(xlnt::page_setup::paper_size::tabloid);
        ws.get_page_setup().set_fit_to_page(true);
        ws.get_page_setup().set_fit_to_height(false);
        ws.get_page_setup().set_fit_to_width(true);

        auto xml_string = xlnt::writer::write_worksheet(ws);

        pugi::xml_document doc;
        doc.load(xml_string.c_str());

        xlnt::worksheet ws2(wb_);
        xml_string = xlnt::writer::write_worksheet(ws2);
        doc.load(xml_string.c_str());
    }
    
    void test_header_footer()
    {
    	auto ws = wb_.create_sheet();
    	ws.get_header_footer().get_left_header().set_text("Left Header Text");
        ws.get_header_footer().get_center_header().set_text("Center Header Text");
        ws.get_header_footer().get_center_header().set_font_name("Arial,Regular");
        ws.get_header_footer().get_center_header().set_font_size(6);
        ws.get_header_footer().get_center_header().set_font_color("445566");
        ws.get_header_footer().get_right_header().set_text("Right Header Text");
        ws.get_header_footer().get_right_header().set_font_name("Arial,Bold");
        ws.get_header_footer().get_right_header().set_font_size(8);
        ws.get_header_footer().get_right_header().set_font_color("112233");
        ws.get_header_footer().get_left_footer().set_text("Left Footer Text\nAnd &[Date] and &[Time]");
        ws.get_header_footer().get_left_footer().set_font_name("Times New Roman,Regular");
        ws.get_header_footer().get_left_footer().set_font_size(10);
        ws.get_header_footer().get_left_footer().set_font_color("445566");
        ws.get_header_footer().get_center_footer().set_text("Center Footer Text &[Path]&[File] on &[Tab]");
        ws.get_header_footer().get_center_footer().set_font_name("Times New Roman,Bold");
        ws.get_header_footer().get_center_footer().set_font_size(12);
        ws.get_header_footer().get_center_footer().set_font_color("778899");
        ws.get_header_footer().get_right_footer().set_text("Right Footer Text &[Page] of &[Pages]");
        ws.get_header_footer().get_right_footer().set_font_name("Times New Roman,Italic");
        ws.get_header_footer().get_right_footer().set_font_size(14);
        ws.get_header_footer().get_right_footer().set_font_color("AABBCC");
        
        std::string expected_xml_string =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "  <headerFooter>"
        "    <oddHeader>&amp;L&amp;\"Calibri,Regular\"&amp;K000000Left Header Text&amp;C&amp;\"Arial,Regular\"&amp;6&amp;K445566Center Header Text&amp;R&amp;\"Arial,Bold\"&amp;8&amp;K112233Right Header Text</oddHeader>"
        "    <oddFooter>&amp;L&amp;\"Times New Roman,Regular\"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;\"Times New Roman,Bold\"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;\"Times New Roman,Italic\"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>"
        "  </headerFooter>"
        "</worksheet>";
        
        pugi::xml_document expected_doc;
        pugi::xml_document observed_doc;
        
        expected_doc.load(expected_xml_string.c_str());
        observed_doc.load(xlnt::writer::write_worksheet(ws, {}, {}).c_str());
        
        TS_ASSERT(Helper::compare_xml(expected_doc, observed_doc));
        
        ws = wb_.create_sheet();
        
        expected_xml_string =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "  <sheetPr>"
        "    <outlinePr summaryRight=\"1\" summaryBelow=\"1\"/>"
        "  </sheetPr>"
        "  <dimension ref=\"A1:A1\"/>"
        "  <sheetViews>"
        "    <sheetView workbookViewId=\"0\">"
        "      <selection sqref=\"A1\" activeCell=\"A1\"/>"
        "    </sheetView>"
        "  </sheetViews>"
        "  <sheetFormatPr baseColWidth=\"10\" defaultRowHeight=\"15\"/>"
        "  <sheetData/>"
        "  <pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>"
        "</worksheet>";
        
        expected_doc.load(expected_xml_string.c_str());
        observed_doc.load(xlnt::writer::write_worksheet(ws, {}, {}).c_str());
        
        TS_ASSERT(Helper::compare_xml(expected_doc, observed_doc));
    }
    
    void test_positioning_point()
    {
        TS_SKIP("not implemented");
        /*
        auto ws = wb_.create_sheet();
         */
    }
    
    void test_positioning_roundtrip()
    {
        TS_SKIP("not implemented");
        /*
    	auto ws = wb_.create_sheet();
    	TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("A1").get_anchor()), xlnt::cell_reference("A1"));
    	TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("D52").get_anchor()), xlnt::cell_reference("D52"));
    	TS_ASSERT_EQUALS(ws.get_point_pos(ws.get_cell("X11").get_anchor()), xlnt::cell_reference("X11"));
         */
    }
    
    void test_page_setup()
    {
        TS_SKIP("not implemented");
        /*
    	xlnt::page_setup p;
    	TS_ASSERT(p.get_page_setup().empty());
    	p.set_scale(1);
    	TS_ASSERT_EQUALS(p.get_page_setup().at("scale"), 1);
         */
    }
    
    void test_page_options()
    {
        TS_SKIP("not implemented");
        /*
    	xlnt::page_setup p;
    	TS_ASSERT(p.get_options().empty());
    	p.set_horizontal_centered(true);
    	p.set_vertical_centered(true);
    	TS_ASSERT_EQUALS(p.get_options().at("verticalCentered"), "1");
    	TS_ASSERT_EQUALS(p.get_options().at("horizontalCentered"), "1");
         */
    }

private:
    xlnt::workbook wb_;
};
