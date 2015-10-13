#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>

class test_style : public CxxTest::TestSuite
{
public:
    test_style() : writer_(workbook_)
    {
        workbook_.set_guess_types(true);
        auto now = xlnt::datetime::now();
        auto ws = workbook_.get_active_sheet();
        ws.get_cell("A1").set_value("12.34%"); // 2
        ws.get_cell("B1").set_value(now); // 3
        ws.get_cell("C1").set_value(now);
        ws.get_cell("D1").set_value("This is a test"); // 1
        ws.get_cell("E1").set_value("31.31415"); // 3
        xlnt::style st; // 4
        st.set_number_format(xlnt::number_format(xlnt::number_format::format::number_00));
        st.set_protection(xlnt::protection(xlnt::protection::type::unprotected));
        ws.get_cell("F1").set_style(st);
        xlnt::style st2; // 5
        st.set_protection(xlnt::protection(xlnt::protection::type::unprotected));
        ws.get_cell("G1").set_style(st2);
        
    }
    
    void test_create_style_table()
    {
        TS_SKIP("");
        //TS_ASSERT_EQUALS(5, writer_.get_styles().size());
    }

    void test_write_style_table()
    {
		TS_SKIP("");

		/*
        auto path = PathHelper::GetDataDirectory("/writer/expected/simple-styles.xml");
        pugi::xml_document expected;
        expected.load_file(path.c_str());
        
        auto content = writer_.write_table();
        pugi::xml_document observed;
        observed.load(content.c_str());
        
        auto diff = Helper::compare_xml(expected, observed);
        TS_ASSERT(diff);
		*/
    }

    void test_no_style()
    {
		TS_SKIP("");

		/*
        xlnt::workbook wb;
        xlnt::style_writer w(wb);
        TS_ASSERT_EQUALS(w.get_styles().size(), 1); // there is always the empty (default) style
		*/
    }

    void test_nb_style()
    {
        TS_SKIP("");
    }

    void test_style_unicity()
    {
        TS_SKIP("");
    }

    void test_fonts()
    {
		TS_SKIP("");

		/*
        auto expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "   <fonts count=\"2\">"
        "      <font>"
        "         <sz val=\"11\" />"
        "         <color theme=\"1\" />"
        "         <name val=\"Calibri\" />"
        "         <family val=\"2\" />"
        "         <scheme val=\"minor\" />"
        "      </font>"
        "      <font>"
        "         <sz val=\"12.0\" />"
        "         <color rgb=\"00000000\" />"
        "         <name val=\"Calibri\" />"
        "         <family val=\"2\" />"
        "         <b />"
        "      </font>"
        "   </fonts>"
        "</styleSheet>";
        
        pugi::xml_document expected_doc;
        expected_doc.load(expected);
        
        std::string observed = "";
        pugi::xml_document observed_doc;
        observed_doc.load(observed.c_str());
        
        auto diff = Helper::compare_xml(expected_doc, observed_doc);
        TS_ASSERT(diff);
		*/
    }
    
    void test_fonts_with_underline()
    {
		TS_SKIP("");

		/*
        auto expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "   <fonts count=\"2\">"
        "      <font>"
        "         <sz val=\"11\" />"
        "         <color theme=\"1\" />"
        "         <name val=\"Calibri\" />"
        "         <family val=\"2\" />"
        "         <scheme val=\"minor\" />"
        "      </font>"
        "      <font>"
        "         <sz val=\"12.0\" />"
        "         <color rgb=\"00000000\" />"
        "         <name val=\"Calibri\" />"
        "         <family val=\"2\" />"
        "         <b />"
        "         <u />"
        "      </font>"
        "   </fonts>"
        "</styleSheet>";
        
        pugi::xml_document expected_doc;
        expected_doc.load(expected);
        
        std::string observed = "";
        pugi::xml_document observed_doc;
        observed_doc.load(observed.c_str());
        
        auto diff = Helper::compare_xml(expected_doc, observed_doc);
        TS_ASSERT(diff);
		*/
    }

    void test_fills()
    {
		TS_SKIP("");

		/*
        auto expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "   <fills count=\"3\">"
        "      <fill>"
        "         <patternFill patternType=\"none\" />"
        "      </fill>"
        "      <fill>"
        "         <patternFill patternType=\"gray125\" />"
        "      </fill>"
        "      <fill>"
        "         <patternFill patternType=\"solid\">"
        "            <fgColor rgb=\"0000FF00\" />"
        "         </patternFill>"
        "      </fill>"
        "   </fills>"
        "</styleSheet>";
        
        pugi::xml_document expected_doc;
        expected_doc.load(expected);
        
        std::string observed = "";
        pugi::xml_document observed_doc;
        observed_doc.load(observed.c_str());
        
        auto diff = Helper::compare_xml(expected_doc, observed_doc);
        TS_ASSERT(diff);
		*/
    }

    void test_borders()
    {
		TS_SKIP("");

		/*
        auto expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "   <borders count=\"2\">"
        "      <border>"
        "         <left />"
        "         <right />"
        "         <top />"
        "         <bottom />"
        "         <diagonal />"
        "      </border>"
        "      <border>"
        "         <left />"
        "         <right />"
        "         <top style=\"thin\">"
        "            <color rgb=\"0000FF00\" />"
        "         </top>"
        "         <bottom />"
        "         <diagonal />"
        "      </border>"
        "   </borders>"
        "</styleSheet>";
        
        pugi::xml_document expected_doc;
        expected_doc.load(expected);
        
        std::string observed = "";
        pugi::xml_document observed_doc;
        observed_doc.load(observed.c_str());
        
        auto diff = Helper::compare_xml(expected_doc, observed_doc);
        TS_ASSERT(diff);
		*/
    }
    
    void test_write_color()
    {
        TS_SKIP("");
    }

    void test_write_cell_xfs_1()
    {
        TS_SKIP("");
    }

    void test_alignment()
    {
        TS_SKIP("");
    }

    void test_alignment_rotation()
    {
        TS_SKIP("");
    }

    void test_alignment_indent()
    {
        TS_SKIP("");
    }
    
    void test_rewrite_styles()
    {
        TS_SKIP("");
    }
    
    void test_write_dxf()
    {
		TS_SKIP("");

		/*
        auto expected =
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "   <dxfs count=\"1\">"
        "      <dxf>"
        "         <font>"
        "            <color rgb=\"FFFFFFFF\" />"
        "            <b val=\"1\" />"
        "            <i val=\"1\" />"
        "            <u val=\"single\" />"
        "            <strike />"
        "         </font>"
        "         <fill>"
        "            <patternFill patternType=\"solid\">"
        "               <fgColor rgb=\"FFEE1111\" />"
        "               <bgColor rgb=\"FFEE1111\" />"
        "            </patternFill>"
        "         </fill>"
        "         <border>"
        "            <left style=\"medium\">"
        "               <color rgb=\"000000FF\" />"
        "            </left>"
        "            <right style=\"medium\">"
        "               <color rgb=\"000000FF\" />"
        "            </right>"
        "            <top style=\"medium\">"
        "               <color rgb=\"000000FF\" />"
        "            </top>"
        "            <bottom style=\"medium\">"
        "               <color rgb=\"000000FF\" />"
        "            </bottom>"
        "         </border>"
        "      </dxf>"
        "   </dxfs>"
        "</styleSheet>";
        
        pugi::xml_document expected_doc;
        expected_doc.load(expected);
        
        std::string observed = "";
        pugi::xml_document observed_doc;
        observed_doc.load(observed.c_str());
        
        auto diff = Helper::compare_xml(expected_doc, observed_doc);
        TS_ASSERT(diff);
		*/
    }
    
    void test_protection()
    {
        TS_SKIP("");
    }
    
private:
    xlnt::workbook workbook_;
    xlnt::style_writer writer_;
};
