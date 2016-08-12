#pragma once

#include <array>
#include <fstream>
#include <string>
#include <sstream>

#include <detail/include_windows.hpp>

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

class path_helper
{
public:
    static xlnt::path get_executable_directory()
    {
#ifdef __APPLE__
        std::array<char, 1024> path;
        uint32_t size = static_cast<uint32_t>(path.size());

        if (_NSGetExecutablePath(path.data(), &size) == 0)
        {
            return xlnt::path(std::string(path.begin(), std::find(path.begin(), path.end(), '\0') - 9));
        }

        throw std::runtime_error("buffer too small, " + std::to_string(path.size()) + ", should be: " + std::to_string(size));
#elif defined(_MSC_VER)
        
        std::array<TCHAR, MAX_PATH> buffer;
        DWORD result = GetModuleFileName(nullptr, buffer.data(), (DWORD)buffer.size());
        
        if(result == 0 || result == buffer.size())
        {
            throw std::runtime_error("GetModuleFileName failed or buffer was too small");
        }

        return xlnt::path(std::string(buffer.begin(), buffer.begin() + result - 13));
#else
        char arg1[20];
        char exepath[PATH_MAX + 1] = {0};

        sprintf(arg1, "/proc/%d/exe", getpid());
        auto bytes_written = readlink(arg1, exepath, 1024);

        return xlnt::path(std::string(exepath).substr(0, bytes_written - 9));
#endif
    }
    
    static xlnt::path get_data_directory(const std::string &append = "")
    {
        return xlnt::path("../../tests/data")
			.resolve(get_executable_directory())
			.append(xlnt::path(append));
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
