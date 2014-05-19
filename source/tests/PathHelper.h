#pragma once

#ifdef __APPLE__
#include <mach-o/dyld.h>
#elif defined(_WIN32)
#include <Windows.h>
#endif

class PathHelper
{
public:
    static std::string GetExecutableDirectory()
    {
#ifdef __APPLE__
        std::array<char, 1024> path;
        uint32_t size = static_cast<uint32_t>(path.size());
        if (_NSGetExecutablePath(path.data(), &size) == 0)
        {
            return std::string(path.begin(), std::find(path.begin(), path.end(), '\0') - 9);
        }
        throw std::runtime_error("buffer too small, " + std::to_string(path.size()) + ", should be: " + std::to_string(size));
#elif defined(_WIN32)
        std::array<TCHAR, MAX_PATH> path;
        DWORD result = GetModuleFileName(nullptr, buffer.data(), buffer.size());
        if(result == 0 || result == path.size())
        {
            throw std::runtime_error("GetModuleFileName failed or buffer was too small");
        }
        return std::string(path.begin(), path.begin() + result);
#else
        throw std::runtime_error("not implmented");
#endif
    }
    
    static std::string GetDataDirectory()
    {
        return GetExecutableDirectory() + "../source/tests/test_data";
    }
    
    static std::string GetTempDirectory()
    {
        return "/tmp/xlsx";
    }
};