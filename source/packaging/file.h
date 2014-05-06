#pragma once

#include <string>

namespace xlnt {

/// <summary>
/// Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects.
/// </summary>
class file
{
public:
	/// <summary>
	/// Copies an existing file to a new file. Overwriting a file of the same name is allowed.
	/// </summary>
	static void copy(const std::string &source, const std::string &destination, bool overwrite = false);

	/// <summary>
	/// Determines whether the specified file exists.
	/// </summary>
	static bool exists(const std::string &path);
};

} // namespace xlnt
