#pragma once

#include <array>
#include <cstdio>
#include <string>

#include <detail/include_windows.hpp>
#include <helpers/path_helper.hpp>

class TemporaryFile
{
public:
    static std::string CreateTemporaryFilename()
    {
#ifdef _MSC_VER
	std::array<TCHAR, MAX_PATH> buffer;
	DWORD result = GetTempPath(static_cast<DWORD>(buffer.size()), buffer.data());
    
	if(result > MAX_PATH)
	{
	    throw std::runtime_error("buffer is too small");
	}
    
	if(result == 0)
	{
	    throw std::runtime_error("GetTempPath failed");
	}
    
	std::string directory(buffer.begin(), buffer.begin() + result);
    
    return PathHelper::WindowsToUniversalPath(directory + "xlnt.xlsx");
#else
	return "/tmp/xlnt.xlsx";
#endif
}

    TemporaryFile() : filename_(CreateTemporaryFilename())
    {
        if(PathHelper::FileExists(GetFilename()))
        {
            std::remove(filename_.c_str());
        }
    }

    ~TemporaryFile()
    {
        std::remove(filename_.c_str());
    }

    std::string GetFilename() const { return filename_; }

private:
    const std::string filename_;
};
