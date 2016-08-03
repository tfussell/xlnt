// Copyright (c) 2014-2016 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#include <fstream>
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

#include <xlnt/utils/path.hpp>

namespace {

#ifdef _MSC_VER

char system_separator()
{
	return '\\';
}

bool is_root(const std::string &part)
{
	return part.size() == 2 && part.back() == ':' 
		&& part.front() >= 'A' && part.front() <= 'Z';
}

bool file_exists(const std::string &path_string)
{
	std::wstring path_wide(path_string.begin(), path_string.end());
	return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
}

std::string get_working_directory()
{
	TCHAR buffer[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, buffer);
	std::basic_string<TCHAR> working_directory(buffer);
	return std::string(working_directory.begin(), working_directory.end());
}

#else

char system_separator()
{
	return '/';
}

bool is_root(const std::string &part)
{
	return part.empty();
}

bool file_exists(const std::string &path_string)
{
	try
	{
		struct stat fileAtt;

		if (stat(path.c_str(), &fileAtt) == 0)
		{
			return S_ISREG(fileAtt.st_mode);
		}
	}
	catch (...) {}

	return false;
}

std::string get_working_directory()
{
	std::size_t buffer_size = 100;
	std::vector<char> buffer(buffer_size, '\0');

	while (getcwd(&buffer[0], buffer_size) == nullptr)
	{
		buffer_size *= 2;
		buffer.resize(buffer_size, '\0');
	}

	return std::string(&buffer[0]);
}

#endif

std::vector<std::string> split_path(const std::string &path, char delim)
{
	std::vector<std::string> split;
	std::string::size_type previous_index = 0;
	auto separator_index = path.find(delim);

	while (separator_index != std::string::npos)
	{
		auto part = path.substr(previous_index, separator_index - previous_index);
		split.push_back(part);

		previous_index = separator_index + 1;
		separator_index = path.find(delim, previous_index);
	}	

	// Don't add trailing slash
	if (previous_index < path.size())
	{
		split.push_back(path.substr(previous_index));
	}

	return split;
}

std::vector<std::string> split_path_universal(const std::string &path)
{
	auto initial = split_path(path, system_separator());

	if (initial.size() == 1 && system_separator() == '\\')
	{
		auto alternative = split_path(path, '/');
		
		if (alternative.size() > 1)
		{
			return alternative;
		}
	}

	return initial;
}

bool directory_exists(const std::string path)
{
	struct stat info;
	return stat(path.c_str(), &info) == 0 && (info.st_mode & S_IFDIR);
}

} // namespace

namespace xlnt {

char path::separator()
{
	return system_separator();
}

path::path()
{

}

path::path(const std::string &path_string) : parts_(split_path_universal(path_string))
{
}

path::path(const std::string &path_string, char sep) : parts_(split_path(path_string, sep))
{
}

// general attributes 

bool path::is_relative() const
{
	return !is_absolute();
}

bool path::is_absolute() const
{
	return !parts_.empty() && ::is_root(parts_.front());
}

bool path::is_root() const
{
	return parts_.size() == 1 && ::is_root(parts_.front());
}

path path::parent() const
{
	path result;
	result.parts_ = parts_;

	if (result.parts_.size() > 1)
	{
		result.parts_.pop_back();
	}

	return result;
}

std::string path::basename()
{
	return parts_.empty() ? "" : parts_.back();
}

// conversion

std::string path::to_string(char sep) const
{
	if (parts_.empty()) return "";

	std::string result;

	for (const auto &part : parts_)
	{
		result.append(part);
		result.push_back(sep);
	}

	result.pop_back();

	return result;
}

path path::make_absolute(const path &base_path) const
{
	if (is_absolute())
	{
		return *this;
	}

	path copy = base_path;

	for (const auto &part : parts_)
	{
		if (part == ".")
		{
			continue;
		}
		else if (part == ".." && copy.parts_.size() > 1)
		{
			copy.parts_.pop_back();
		}
		else
		{
			copy.parts_.push_back(part);
		}
	}

	return copy;
}

// filesystem attributes

bool path::exists() const
{
	return is_file() || is_directory();
}

bool path::is_directory() const
{
	return directory_exists(to_string());
}

bool path::is_file() const
{
	return file_exists(to_string());
}

// filesystem

std::string path::read_contents() const
{
	std::ifstream f(to_string());
	std::ostringstream ss;
	ss << f.rdbuf();

	return ss.str();
}

// mutators

path &path::append(const std::string &to_append)
{
	parts_.push_back(to_append);
	return *this;
}

path path::append(const std::string &to_append) const
{
	path copy(*this);
	copy.append(to_append);

	return copy;
}

path &path::append(const path &to_append)
{
	parts_.insert(parts_.end(), to_append.begin(), to_append.end());
	return *this;
}

path path::append(const path &to_append) const
{
	path copy(*this);
	copy.append(to_append);

	return copy;
}

// iterators

path::iterator path::begin()
{
	return parts_.begin();
}

path::iterator path::end()
{
	return parts_.end();
}

path::const_iterator path::begin() const
{
	return cbegin();
}

path::const_iterator path::end() const
{
	return cend();
}

path::const_iterator path::cbegin() const
{
	return parts_.cbegin();
}

path::const_iterator path::cend() const
{
	return parts_.cend();
}

path::reverse_iterator path::rbegin()
{
	return parts_.rbegin();
}

path::reverse_iterator path::rend()
{
	return parts_.rend();
}

path::const_reverse_iterator path::rbegin() const
{
	return crbegin();
}

path::const_reverse_iterator path::rend() const
{
	return crend();
}

path::const_reverse_iterator path::crbegin() const
{
	return parts_.crbegin();
}

path::const_reverse_iterator path::crend() const
{
	return parts_.crend();
}

std::string path::to_hash_string() const
{
	return to_string('/');
}

} // namespace xlnt
