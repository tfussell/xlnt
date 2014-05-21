#pragma once

#include <array>
#include <fstream>

#ifdef __APPLE__
#include <mach-o/dyld.h>
#include <sys/stat.h>
#elif defined(_WIN32)
#define NOMINMAX
#include <Shlwapi.h>
#include <Windows.h>
#endif

class PathHelper
{
public:
    static std::string WindowsToUniversalPath(const std::string &windows_path)
    {
        std::string fixed;
        std::stringstream ss(windows_path);
        std::string part;
        
        while(std::getline(ss, part, '\\'))
        {
            if(fixed == "")
            {
                fixed = part;
            }
            else
            {
                fixed += "/" + part;
            }
        }
        
        return fixed;
    }
    
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
        
        std::array<TCHAR, MAX_PATH> buffer;
        DWORD result = GetModuleFileName(nullptr, buffer.data(), buffer.size());
        
        if(result == 0 || result == buffer.size())
        {
            throw std::runtime_error("GetModuleFileName failed or buffer was too small");
        }
        return WindowsToUniversalPath(std::string(buffer.begin(), buffer.begin() + result - 13)) + "/";
        
#else
        
        throw std::runtime_error("not implmented");
        
#endif
        
    }
    
    static std::string GetDataDirectory()
    {
        return GetExecutableDirectory() + "../tests/test_data";
    }
    
    static void CopyFile(const std::string &source, const std::string &destination, bool overwrite)
    {
        if(!overwrite && FileExists(destination))
        {
            throw std::runtime_error("destination file already exists and overwrite==false");
        }
        
        std::ifstream src(source, std::ios::binary);
        std::ofstream dst(destination, std::ios::binary);
        
        dst << src.rdbuf();
    }
    
    static bool FileExists(const std::string &path)
    {
        
#ifdef _WIN32
        
        std::wstring path_wide(path.begin(), path.end());
        return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
        
#else
        
        struct stat fileAtt;
        
        if (stat(path.c_str(), &fileAtt) != 0)
        {
            throw std::runtime_error("stat failed");
        }
        
        return S_ISREG(fileAtt.st_mode);
        
#endif
        
    }
};