#pragma once

#include <array>
#include <fstream>
#include <string>
#include <sstream>

#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/path.hpp>

#define STRING_LITERAL2(a) #a
#define LSTRING_LITERAL2(a) L#a
#define U8STRING_LITERAL2(a) u8#a
#define STRING_LITERAL(a) STRING_LITERAL2(a)
#define LSTRING_LITERAL(a) STRING_LITERAL2(a)
#define U8STRING_LITERAL(a) STRING_LITERAL2(a)

#ifndef XLNT_BENCHMARK_DATA_DIR
#define XLNT_BENCHMARK_DATA_DIR ""
#endif

#ifndef XLNT_SAMPLE_DATA_DIR
#define XLNT_SAMPLE_DATA_DIR ""
#endif

class path_helper
{
public:
    static xlnt::path test_data_directory(const std::string &append = "")
    {
        static const std::string data_dir = STRING_LITERAL(XLNT_TEST_DATA_DIR);
        return xlnt::path(data_dir);
    }

    static xlnt::path test_file(const std::string &filename)
    {
        return test_data_directory().append(xlnt::path(filename));
    }

    static xlnt::path benchmark_data_directory(const std::string &append = "")
    {
        static const std::string data_dir = STRING_LITERAL(XLNT_BENCHMARK_DATA_DIR);
        return xlnt::path(data_dir);
    }

    static xlnt::path benchmark_file(const std::string &filename)
    {
        return benchmark_data_directory().append(xlnt::path(filename));
    }

    static xlnt::path sample_data_directory(const std::string &append = "")
    {
        static const std::string data_dir = STRING_LITERAL(XLNT_SAMPLE_DATA_DIR);
        return xlnt::path(data_dir);
    }

    static xlnt::path sample_file(const std::string &filename)
    {
        return sample_data_directory().append(xlnt::path(filename));
    }

    static void copy_file(const xlnt::path &source, const xlnt::path &destination, bool overwrite)
    {
        if(!overwrite && destination.exists())
        {
            throw xlnt::exception("destination file already exists and overwrite==false");
        }

        std::ifstream src(source.string(), std::ios::binary);
        std::ofstream dst(destination.string(), std::ios::binary);

        dst << src.rdbuf();
    }

    static void delete_file(const xlnt::path &path)
    {
      std::remove(path.string().c_str());
    }
};
