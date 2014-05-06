#include <cassert>
#include <Shlwapi.h>
#include <Windows.h>

#include "File.h"

namespace xlnt {

void file::copy(const std::string &source, const std::string &destination, bool overwrite)
{
	assert(source.size() + 1 < MAX_PATH);
	assert(destination.size() + 1 < MAX_PATH);

	std::wstring source_wide(source.begin(), source.end());
	std::wstring destination_wide(destination.begin(), destination.end());

	BOOL result = CopyFile(source_wide.c_str(), destination_wide.c_str(), !overwrite);

	if(result == 0)
	{
		DWORD error = GetLastError();
		if(error == ERROR_ACCESS_DENIED)
		{
			throw std::runtime_error("Access is denied");
		}
		if(error == ERROR_ENCRYPTION_FAILED)
		{
			throw std::runtime_error("The specified file could not be encrypted");
		}
		if(error == ERROR_FILE_NOT_FOUND)
		{
			throw std::runtime_error("The source file wasn't found");
		}
		if(!overwrite)
		{
			throw std::runtime_error("The destination file already exists");
		}
		throw std::runtime_error("Unknown error");
	}
}

bool file::exists(const std::string &path)
{
	std::wstring path_wide(path.begin(), path.end());

	return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
}

} // namespace xlnt
