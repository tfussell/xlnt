#pragma once

#include <array>
#include <fstream>
#include <string>
#include <sstream>

#include <detail/include_windows.hpp>
#include <xlnt/utils/path.hpp>

#ifdef __APPLE__
#include <mach-o/dyld.h>
#include <sys/stat.h>
#elif defined(_MSC_VER)
#include <Shlwapi.h>
#elif defined(__linux)
#include <unistd.h>
#include <linux/limits.h>
#include <sys/types.h>
#include <sys/stat.h>
#endif

#define STRING_LITERAL2(a) #a
#define LSTRING_LITERAL2(a) L#a
#define U8STRING_LITERAL2(a) u8#a
#define STRING_LITERAL(a) STRING_LITERAL2(a)
#define LSTRING_LITERAL(a) STRING_LITERAL2(a)
#define U8STRING_LITERAL(a) STRING_LITERAL2(a)

class path_helper
{
public:
    static xlnt::path data_directory(const std::string &append = "")
    {
        static const std::string data_dir = STRING_LITERAL(XLNT_TEST_DATA_DIR);
        return xlnt::path(data_dir).append(xlnt::path(append));
    }
    
    static void copy_file(const xlnt::path &source, const xlnt::path &destination, bool overwrite)
    {
        if(!overwrite && destination.exists())
        {
            throw std::runtime_error("destination file already exists and overwrite==false");
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
