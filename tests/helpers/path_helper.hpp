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
    static xlnt::string read_file(const xlnt::string &filename)
    {
        std::ifstream f(filename.data());
        std::ostringstream ss;
        ss << f.rdbuf();
        
        return ss.str().c_str();
    }
    
    static xlnt::string WindowsToUniversalPath(const xlnt::string &windows_path)
    {
        xlnt::string fixed;
        std::stringstream ss(windows_path.data());
        std::string part;
        
        while(std::getline(ss, part, '\\'))
        {
            if(fixed == "")
            {
                fixed = part.c_str();
            }
            else
            {
                fixed += "/" + xlnt::string(part.c_str());
            }
        }
        
        return fixed;
    }
    
    static xlnt::string GetExecutableDirectory()
    {
        
#ifdef __APPLE__
        
        std::array<char, 1024> path;
        uint32_t size = static_cast<uint32_t>(path.size());
        
        if (_NSGetExecutablePath(path.data(), &size) == 0)
        {
            return xlnt::string(path.begin(), std::find(path.begin(), path.end(), '\0') - 9);
        }
        
        throw std::runtime_error("buffer too small, " + std::to_string(path.size()) + ", should be: " + std::to_string(size));
        
#elif defined(_MSC_VER)
        
        std::array<TCHAR, MAX_PATH> buffer;
        DWORD result = GetModuleFileName(nullptr, buffer.data(), (DWORD)buffer.size());
        
        if(result == 0 || result == buffer.size())
        {
            throw std::runtime_error("GetModuleFileName failed or buffer was too small");
        }
        return WindowsToUniversalPath(xlnt::string(buffer.begin(), buffer.begin() + result - 13)) + "/";
        
#else
	char arg1[20];
	char exepath[PATH_MAX + 1] = {0};
	
	sprintf(arg1, "/proc/%d/exe", getpid());
	readlink(arg1, exepath, 1024);

	return xlnt::string(exepath).substr(0, std::strlen(exepath) - 9);
#endif
        
    }
    
    static xlnt::string GetDataDirectory(const xlnt::string &append = "")
    {
        return GetExecutableDirectory() + "../tests/data" + append;
    }
    
    static void CopyFile(const xlnt::string &source, const xlnt::string &destination, bool overwrite)
    {
        if(!overwrite && FileExists(destination))
        {
            throw std::runtime_error("destination file already exists and overwrite==false");
        }
        
        std::ifstream src(source.data(), std::ios::binary);
        std::ofstream dst(destination.data(), std::ios::binary);
        
        dst << src.rdbuf();
    }

    static void DeleteFile(const xlnt::string &path)
    {
      std::remove(path.data());
    }
    
    static bool FileExists(const xlnt::string &path)
    {        
#ifdef _MSC_VER
		std::string path_stdstring = path.data();
        std::wstring path_wide(path_stdstring.begin(), path_stdstring.end());
        return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
        
#else

        try
	{
	    struct stat fileAtt;
        
	    if (stat(path.data(), &fileAtt) == 0)
	    {
	        return S_ISREG(fileAtt.st_mode);
	    }
	}
	catch(...) {}

	return false;
#endif
        
    }
};
