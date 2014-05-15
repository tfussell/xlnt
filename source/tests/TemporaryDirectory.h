#pragma once

#include <array>
#include <cstdio>
#include <string>

class TemporaryDirectory
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
	if(resule == 0)
	{
	    throw std::runtime_error("GetTempPath failed");
	}
	std::string directory(buffer.begin(), buffer.begin() + result);
        return directory + "/xlnt";
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

    std::string GetFilename() const { return filename_; }

private:
    const std::string filename_;
};
