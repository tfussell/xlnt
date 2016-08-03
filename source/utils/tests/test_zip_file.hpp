#include <cassert>
#include <fstream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.hpp>
#include "helpers/path_helper.hpp"
#include "helpers/temporary_file.hpp"

class test_zip_file : public CxxTest::TestSuite
{
public:
    test_zip_file()
    {
        existing_file = path_helper::get_data_directory("genuine/empty.xlsx");
        expected_content_types_string = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>";
        expected_atxt_string = "<?xml version=\"1.0\" ?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"2\"><si><t>This is cell A1 in Sheet 1</t></si><si><t>This is cell G5</t></si></sst>";
        expected_printdir_string =
            "  Length      Date    Time    Name\n"
            "---------  ---------- -----   ----\n"
            "     1704  01/01/1980 00:00   [Content_Types].xml\n"
            "      588  01/01/1980 00:00   _rels/.rels\n"
            "      917  01/01/1980 00:00   docProps/app.xml\n"
            "      609  01/01/1980 00:00   docProps/core.xml\n"
            "     1254  01/01/1980 00:00   xl/_rels/workbook.xml.rels\n"
            "      169  01/01/1980 00:00   xl/calcChain.xml\n"
            "      233  01/01/1980 00:00   xl/sharedStrings.xml\n"
            "     1724  01/01/1980 00:00   xl/styles.xml\n"
            "     6995  01/01/1980 00:00   xl/theme/theme1.xml\n"
            "      898  01/01/1980 00:00   xl/workbook.xml\n"
            "     1068  07/21/2016 20:27   xl/worksheets/sheet1.xml\n"
            "     4427  01/01/1980 00:00   xl/worksheets/sheet2.xml\n"
            "     1032  01/01/1980 00:00   xl/worksheets/sheet3.xml\n"
            "     1231  01/01/1980 00:00   xl/worksheets/sheet4.xml\n"
            "---------                     -------\n"
            "    22849                     14 files\n";
    }

    bool files_equal(const xlnt::path &left, const xlnt::path &right)
    {
        if(left.to_string() == right.to_string())
        {
            return true;
        }
        
		std::ifstream stream_left(left.to_string(), std::ios::binary);
		std::ifstream stream_right(right.to_string(), std::ios::binary);

        while(stream_left && stream_right)
        {
            if(stream_left.get() != stream_right.get())
            {
                return false;
            }
        }
        
        return true;
    }

    void test_load_file()
    {
		temporary_file temp_file;
        xlnt::zip_file f(existing_file);
        f.save(temp_file.get_path());
        TS_ASSERT(files_equal(existing_file, temp_file.get_path()));
    }

    void test_load_stream()
    {
		temporary_file temp;

        std::ifstream in_stream(existing_file.to_string(), std::ios::binary);
        xlnt::zip_file f(in_stream);
        std::ofstream out_stream(temp.get_path().to_string(), std::ios::binary);
		f.save(out_stream);
		out_stream.close();

        TS_ASSERT(files_equal(existing_file, temp.get_path()));
    }

    void test_load_bytes()
    {
		temporary_file temp_file;

		std::vector<std::uint8_t> source_bytes;
        std::ifstream in_stream(existing_file.to_string(), std::ios::binary);

        while(in_stream)
        {
            source_bytes.push_back(static_cast<std::uint8_t>(in_stream.get()));
        }

        xlnt::zip_file f(source_bytes);
        f.save(temp_file.get_path());
        
        xlnt::zip_file f2;
        f2.load(temp_file.get_path());
		std::vector<std::uint8_t> result_bytes;
        f2.save(result_bytes);
        
        TS_ASSERT(source_bytes == result_bytes);
    }

    void test_reset()
    {
        xlnt::zip_file f(existing_file);
        
        TS_ASSERT(!f.namelist().empty());
        
        try
        {
            f.read(xlnt::path("[Content_Types].xml"));
        }
        catch(std::exception e)
        {
            TS_ASSERT(false);
        }
        
        f.reset();
        
        TS_ASSERT(f.namelist().empty());
        
        try
        {
            f.read(xlnt::path("[Content_Types].xml"));
            TS_ASSERT(false);
        }
        catch(std::exception e)
        {
        }

        f.write_string("b", xlnt::path("a"));
        f.reset();

        TS_ASSERT(f.namelist().empty());

        f.write_string("b", xlnt::path("a"));

        TS_ASSERT_DIFFERS(f.getinfo(xlnt::path("a")).file_size, 0);
    }

    void test_getinfo()
    {
        xlnt::zip_file f(existing_file);
        auto info = f.getinfo(xlnt::path("[Content_Types].xml"));
        TS_ASSERT(info.filename.to_string() == "[Content_Types].xml");
    }

    void test_infolist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT_EQUALS(f.infolist().size(), 14);
    }

    void test_namelist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT_EQUALS(f.namelist().size(), 14);
    }

    void test_open_by_name()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open(xlnt::path("[Content_Types].xml")).rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_content_types_string);
    }

    void test_open_by_info()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open(xlnt::path("[Content_Types].xml")).rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_content_types_string);
    }


    void test_read()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.read(xlnt::path("[Content_Types].xml")) == expected_content_types_string);
        TS_ASSERT(f.read(f.getinfo(xlnt::path("[Content_Types].xml"))) == expected_content_types_string);
    }

    void test_testzip()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(!f.check_crc());
    }

    void test_write_file()
    {
		temporary_file temp_file;

        xlnt::zip_file f;
        auto text_file = path_helper::get_data_directory("reader/sharedStrings.xml");
        f.write_file(text_file);
        f.write_file(text_file, xlnt::path("sharedStrings2.xml"));
        f.save(temp_file.get_path());
        
        xlnt::zip_file f2(temp_file.get_path());

        for(auto &info : f2.infolist())
        {
            if(info.filename.to_string() == "sharedStrings2.xml")
            {
                TS_ASSERT(f2.read(info) == expected_atxt_string);
            }
            else if(info.filename.basename() == "sharedStrings.xml")
            {
                TS_ASSERT(f2.read(info) == expected_atxt_string);
            }
        }
    }

    void test_write_string()
    {
        xlnt::zip_file f;
        f.write_string("a\na", xlnt::path("a.txt"));
        xlnt::zip_info info;
        info.filename = xlnt::path("b.txt");
        info.date_time.year = 2014;
        f.write_string("b\nb", info);

		temporary_file temp_file;
        f.save(temp_file.get_path());
        
		xlnt::zip_file f2(temp_file.get_path());
        TS_ASSERT(f2.read(xlnt::path("a.txt")) == "a\na");
        TS_ASSERT(f2.read(f2.getinfo(xlnt::path("b.txt"))) == "b\nb");
    }

    void test_comment()
    {
        xlnt::zip_file f;
        f.comment = "comment";
		temporary_file temp_file;
        f.save(temp_file.get_path());
        
        xlnt::zip_file f2(temp_file.get_path());
        TS_ASSERT(f2.comment == "comment");

        xlnt::zip_file f3;
        std::vector<std::uint8_t> bytes { 1, 2, 3 };
        TS_ASSERT_THROWS(f3.load(bytes), std::runtime_error);
    }

private:
    xlnt::path existing_file;
    std::string expected_content_types_string;
    std::string expected_atxt_string;
    std::string expected_printdir_string;
};
