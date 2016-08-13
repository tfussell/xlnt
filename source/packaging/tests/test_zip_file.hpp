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
        existing_file = path_helper::get_data_directory("4_not-package.xlsx");
        expected_string = "not-empty";
    }

    bool files_equal(const xlnt::path &left, const xlnt::path &right)
    {
        if(left.string() == right.string())
        {
            return true;
        }
        
		std::ifstream stream_left(left.string(), std::ios::binary);
		std::ifstream stream_right(right.string(), std::ios::binary);

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

        std::ifstream in_stream(existing_file.string(), std::ios::binary);
        xlnt::zip_file f(in_stream);
        std::ofstream out_stream(temp.get_path().string(), std::ios::binary);
		f.save(out_stream);
		out_stream.close();

        TS_ASSERT(files_equal(existing_file, temp.get_path()));
    }

    void test_load_bytes()
    {
		temporary_file temp_file;

		std::vector<std::uint8_t> source_bytes;
        std::ifstream in_stream(existing_file.string(), std::ios::binary);

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
            f.read(xlnt::path("text.txt"));
        }
        catch(std::exception e)
        {
            TS_ASSERT(false);
        }
        
        f.reset();
        
        TS_ASSERT(f.namelist().empty());
        
        try
        {
            f.read(xlnt::path("doesnt-exist.txt"));
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
        auto info = f.getinfo(xlnt::path("text.txt"));
        TS_ASSERT(info.filename.string() == "text.txt");
    }

    void test_infolist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT_EQUALS(f.infolist().size(), 1);
    }

    void test_namelist()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT_EQUALS(f.namelist().size(), 1);
    }

    void test_open_by_name()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open(xlnt::path("text.txt")).rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_string);
    }

    void test_open_by_info()
    {
        xlnt::zip_file f(existing_file);
        std::stringstream ss;
        ss << f.open(xlnt::path("text.txt")).rdbuf();
        std::string result = ss.str();
        TS_ASSERT(result == expected_string);
    }


    void test_read()
    {
        xlnt::zip_file f(existing_file);
        TS_ASSERT(f.read(xlnt::path("text.txt")) == expected_string);
        TS_ASSERT(f.read(f.getinfo(xlnt::path("text.txt"))) == expected_string);
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
        auto text_file = path_helper::get_data_directory("2_text.xlsx");
        f.write_file(text_file);
        f.write_file(text_file, xlnt::path("a.txt"));
        f.save(temp_file.get_path());
        
        xlnt::zip_file f2(temp_file.get_path());

        for(auto &info : f2.infolist())
        {
			TS_ASSERT(f2.read(info) == expected_string);
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
    std::string expected_string;
};
