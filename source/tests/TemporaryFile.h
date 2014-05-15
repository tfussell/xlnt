#pragma once

#include <array>
#include <cstdio>
#include <string>

#ifdef _WIN32
#include <Windows.h>
#endif

class TemporaryFile
{
public:
    static std::string CreateTemporaryFilename()
    {
#ifdef _WIN32
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
        return directory + "xlnt.xlsx";
#else
	return "/tmp/xlsx.xlnt";
#endif
}

    TemporaryFile() : filename_(CreateTemporaryFilename())
    {

    }

    ~TemporaryFile()
    {
        remove(filename_.c_str());
    }

    std::string GetFilename() const { return filename_; }

private:
    const std::string filename_;
};
