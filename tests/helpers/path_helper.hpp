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

class PathHelper
{
public:
    static std::string read_file(const std::string &filename)
    {
        std::ifstream f(filename);
        std::ostringstream ss;
        ss << f.rdbuf();
        
        return ss.str();
    }
    
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
        
#elif defined(_MSC_VER)
        
        std::array<TCHAR, MAX_PATH> buffer;
        DWORD result = GetModuleFileName(nullptr, buffer.data(), (DWORD)buffer.size());
        
        if(result == 0 || result == buffer.size())
        {
            throw std::runtime_error("GetModuleFileName failed or buffer was too small");
        }
        return WindowsToUniversalPath(std::string(buffer.begin(), buffer.begin() + result - 13)) + "/";
        
#else
	char arg1[20];
	char exepath[PATH_MAX + 1] = {0};
	
	sprintf(arg1, "/proc/%d/exe", getpid());
	readlink(arg1, exepath, 1024);
	return std::string(exepath).substr(0, std::strlen(exepath) - 9);
#endif
        
    }
    
    static std::string GetDataDirectory(const std::string &append = "")
    {
        return GetExecutableDirectory() + "../tests/data" + append;
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

    static void DeleteFile(const std::string &path)
    {
      std::remove(path.c_str());
    }
    
    static bool FileExists(const std::string &path)
    {        
#ifdef _MSC_VER
        
        std::wstring path_wide(path.begin(), path.end());
        return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
        
#else

        try
	{
	    struct stat fileAtt;
        
	    if (stat(path.c_str(), &fileAtt) == 0)
	    {
	        return S_ISREG(fileAtt.st_mode);
	    }
	}
	catch(...) {}

	return false;
#endif
        
    }
};
