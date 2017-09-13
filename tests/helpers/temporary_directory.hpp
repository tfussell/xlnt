#pragma once

#include <array>
#include <cstdio>
#include <string>

#include <detail/include_windows.hpp>

#include "PathHelper.h"

class temporary_directory
{
public:
    static std::string create()
    {
#ifdef _WIN32
	std::array<TCHAR, MAX_PATH> buffer;
	DWORD result = GetTempPath(static_cast<DWORD>(buffer.size()), buffer.data());
	if(result > MAX_PATH)
	{
	    throw xlnt::exception("buffer is too small");
	}
	if(result == 0)
	{
	    throw xlnt::exception("GetTempPath failed");
	}
	std::string directory(buffer.begin(), buffer.begin() + result);
    return path_helper::windows_to_universal_path(directory + "xlnt");
#else
	return "/tmp/xlsx";
#endif
    }

    temporary_directory() : filename_(create())
    {

    }

    ~temporary_directory()
    {
        remove(filename_.c_str());
    }

    std::string get_filename() const { return filename_; }

private:
    const std::string filename_;
};
