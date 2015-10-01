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
        existing_file = PathHelper::GetDataDirectory("/genuine/empty.xlsx");
        expected_content_types_string = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>";
        expected_atxt_string = "<?xml version=\"1.0\" ?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"2\"><si><t>This is cell A1 in Sheet 1</t></si><si><t>This is cell G5</t></si></sst>";
        expected_printdir_string = "  Length      Date    Time    Name\n---------  ---------- -----   ----\n     1704  01/01/1980 00:00   [Content_Types].xml\n      588  01/01/1980 00:00   _rels/.rels\n     1254  01/01/1980 00:00   xl/_rels/workbook.xml.rels\n      898  01/01/1980 00:00   xl/workbook.xml\n     1231  01/01/1980 00:00   xl/worksheets/sheet4.xml\n     4427  01/01/1980 00:00   xl/worksheets/sheet2.xml\n     1032  01/01/1980 00:00   xl/worksheets/sheet3.xml\n     1026  01/01/1980 00:00   xl/worksheets/sheet1.xml\n     6995  01/01/1980 00:00   xl/theme/theme1.xml\n      233  01/01/1980 00:00   xl/sharedStrings.xml\n     1724  01/01/1980 00:00   xl/styles.xml\n      169  01/01/1980 00:00   xl/calcChain.xml\n      917  01/01/1980 00:00   docProps/app.xml\n      609  01/01/1980 00:00   docProps/core.xml\n---------                     -------\n    22807                     14 files\n";
    }

    void remove_temp_file()
    {
        std::remove(temp_file.GetFilename().c_str());
    }

    void make_temp_directory()
    {
    }

    void remove_temp_directory()
    {
    }

    bool files_equal(const std::string &a, const std::string &b)
    {
        if(a == b)
        {
            return true;
        }
        
        std::ifstream stream_a(a, std::ios::binary), stream_b(a, std::ios::binary);

        while(stream_a && stream_b)
        {
            if(stream_a.get() != stream_b.get())
            {
                return false;
            }
        }
        
        return true;
    }

    void test_load_file()
    {
        remove_temp_file();
        xlnt::zip_file f(existing_file);
        f.save(temp_file.GetFilename());
        TS_ASSERT(files_equal(existing_file, temp_file.GetFilename()));
        remove_temp_file();
    }

    void test_load_stream()
    {
        remove_temp_file();
        {
            std::ifstream in_stream(existing_file, std::ios::binary);
            xlnt::zip_file f(in_stream);
            std::ofstream out_stream(temp_file.GetFilename(), std::ios::binary);
            f.save(out_stream);
        }
        TS_ASSERT(files_equal(existing_file, temp_file.GetFilename()));
        remove_temp_file();
    }

    void test_load_bytes()
    {
        remove_temp_file();
        
        xlnt::zip_file f;
        std::vector<unsigned char> source_bytes, result_bytes;
        std::ifstream in_stream(existing_file, std::ios::binary);
        while(in_stream)
        {
            source_bytes.push_back((unsigned char)in_stream.get());
        }
        f.load(source_bytes);
        f.save(temp_file.GetFilename());
        
        xlnt::zip_file f2;
        f2.load(temp_file.GetFilename());
        result_bytes = std::vector<unsigned char>();
        f2.save(result_bytes);
        
        TS_ASSERT(source_bytes == result_bytes);
        
        remove_temp_file();
    }

    void test_reset()
    {
        xlnt::zip_file f(existing_file);
        
        TS_ASSERT(!f.namelist().empty());
        
        try
        {
            f.read("[Content_Types].xml");
        }
        catch(std::exception e)
        {
            TS_ASSERT(false);
        }
        
        f.reset();
        
        TS_ASSERT(f.namelist().empty());
        
        try
        {
            f.read("[Content_Types].xml");
            TS_ASSERT(false);
        }
        catch(std::exception e)
        {
        }
    }

    void test_getinfo()
    {
        xlnt::zip_file f(existing_file);
        auto info = f.getinfo("[Content_Types].xml");
        TS_ASSERT(info.filename == "[Content_Types].xml");
    }

    void test_infolist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.infolist().size() == 14);
    }

    void test_namelist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.namelist().size() == 14);
    }

    void test_open_by_name()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open("[Content_Types].xml").rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_content_types_string);
    }

    void test_open_by_info()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open("[Content_Types].xml").rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_content_types_string);
    }

    void test_extract_current_directory()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_extract_path()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_extractall_current_directory()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_extractall_path()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_extractall_members_name()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_extractall_members_info()
    {
        xlnt::zip_file f(existing_file);
    }

    void test_printdir()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        f.printdir(ss);
        auto printed = ss.str();
        TS_ASSERT(printed == expected_printdir_string);
    }

    void test_read()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.read("[Content_Types].xml") == expected_content_types_string);
        TS_ASSERT(f.read(f.getinfo("[Content_Types].xml")) == expected_content_types_string);
    }

    void test_testzip()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.testzip().first);
    }

    void test_write()
    {
        remove_temp_file();
        
        xlnt::zip_file f;
        auto text_file = PathHelper::GetDataDirectory("/reader/sharedStrings.xml");
        f.write(text_file);
        f.write(text_file, "sharedStrings2.xml");
        f.save(temp_file.GetFilename());
        
        xlnt::zip_file f2(temp_file.GetFilename());

        for(auto &info : f2.infolist())
        {
            if(info.filename == "sharedStrings2.xml")
            {
                TS_ASSERT(f2.read(info) == expected_atxt_string);
            }
            else if(info.filename.substr(info.filename.size() - 17) == "sharedStrings.xml")
            {
                TS_ASSERT(f2.read(info) == expected_atxt_string);
            }
        }
        
        remove_temp_file();
    }

    void test_writestr()
    {
        remove_temp_file();
        
        xlnt::zip_file f;
        f.writestr("a.txt", "a\na");
        xlnt::zip_info info;
        info.filename = "b.txt";
        info.date_time.year = 2014;
        f.writestr(info, "b\nb");
        f.save(temp_file.GetFilename());
        
        xlnt::zip_file f2(temp_file.GetFilename());
        TS_ASSERT(f2.read("a.txt") == "a\na");
        TS_ASSERT(f2.read(f2.getinfo("b.txt")) == "b\nb");
        
        remove_temp_file();
    }

    void test_comment()
    {
        remove_temp_file();
        
        xlnt::zip_file f;
        f.comment = "comment";
        f.save(temp_file.GetFilename());
        
        xlnt::zip_file f2(temp_file.GetFilename());
        TS_ASSERT(f2.comment == "comment");
        
        remove_temp_file();
    }

private:
    TemporaryFile temp_file;
    std::string existing_file;
    std::string expected_content_types_string;
    std::string expected_atxt_string;
    std::string expected_printdir_string;
};