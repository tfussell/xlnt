#include <Windows.h>

#include "file_info.h"

namespace xlnt {

file_info::file_info(const std::string &path)
{
	fullName_ = path;
}

void file_info::delete_()
{
	std::wstring wFileName(fullName_.begin(), fullName_.end());
	LPCWSTR lpFileName = wFileName.c_str();

	BOOL result = DeleteFile(lpFileName);

	if(result == 0)
	{
		throw std::runtime_error("delete failed: " + fullName_);
	}
}

bool file_info::exists() const
{
	std::wstring wFileName(fullName_.begin(), fullName_.end());
	LPCWSTR lpFileName = wFileName.c_str();

	DWORD dwAttrib = GetFileAttributes(lpFileName);

	return dwAttrib != INVALID_FILE_ATTRIBUTES;
}

} // namespace xlnt
