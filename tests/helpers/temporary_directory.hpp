#pragma once

#include <array>
#include <cstdio>
#include <string>

#include <detail/include_windows.hpp>

#include "PathHelper.h"

class TemporaryDirectory
{
public:
    static xlnt::string CreateTemporaryFilename()
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
	xlnt::string directory(buffer.begin(), buffer.begin() + result);
    return PathHelper::WindowsToUniversalPath(directory + "xlnt");
#else
	return "/tmp/xlsx";
#endif
    }

    TemporaryDirectory() : filename_(CreateTemporaryFilename())
    {

    }

    ~TemporaryDirectory()
    {
        remove(filename_.c_str());
    }

    xlnt::string GetFilename() const { return filename_; }

private:
    const xlnt::string filename_;
};
