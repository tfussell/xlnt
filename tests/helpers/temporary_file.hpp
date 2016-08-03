#pragma once

#include <array>
#include <cstdio>
#include <string>

#include <detail/include_windows.hpp>
#include <helpers/path_helper.hpp>

namespace {

static std::string create_temporary_filename()
{
#ifdef _MSC_VER
	std::array<TCHAR, MAX_PATH> buffer;
	DWORD result = GetTempPath(static_cast<DWORD>(buffer.size()), buffer.data());

	if (result > MAX_PATH)
	{
		throw std::runtime_error("buffer is too small");
	}

	if (result == 0)
	{
		throw std::runtime_error("GetTempPath failed");
	}

	return std::string(buffer.begin(), buffer.begin() + result) + "xlnt.xlsx";
#else
	return "/tmp/xlnt.xlsx";
#endif
}

} // namespace 

class temporary_file
{
public:
    temporary_file() : path_(create_temporary_filename())
    {
        if(path_.exists())
        {
            std::remove(path_.to_string().c_str());
        }
    }

    ~temporary_file()
    {
        std::remove(path_.to_string().c_str());
    }

    xlnt::path get_path() const { return path_; }

private:
    const xlnt::path path_;
};
