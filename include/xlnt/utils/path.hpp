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
#pragma once

#include <string>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/hashable.hpp>

namespace xlnt {

/// <summary>
/// Encapsulates a path that points to location in a filesystem.
/// </summary>
class XLNT_CLASS path : public hashable
{
public:
	/// <summary>
	/// The parts of this path are held in a container of this type.
	/// </summary>
	using container = std::vector<std::string>;

	/// <summary>
	/// Expose the container's iterator as the iterator for this class.
	/// </summary>
	using iterator = container::iterator;
	/// <summary>
	/// Expose the container's const_iterator as the const_iterator for this class.
	/// </summary>
	using const_iterator = container::const_iterator;

	/// <summary>
	/// Expose the container's reverse_iterator as the reverse_iterator for this class.
	/// </summary>
	using reverse_iterator = container::reverse_iterator;

	/// <summary>
	/// Expose the container's const_reverse_iterator as the const_reverse_iterator for this class.
	/// </summary>
	using const_reverse_iterator = container::const_reverse_iterator;

	/// <summary>
	/// The system-specific path separator character (e.g. '/' or '\').
	/// </summary>
	static char separator();

	/// <summary>
	/// Construct an empty path.
	/// </summary>
	path();

	/// <summary>
	/// Counstruct a path from a string representing the path.
	/// </summary>
	explicit path(const std::string &path_string);

	/// <summary>
	/// Construct a path from a string with an explicit directory seprator.
	/// </summary>
	path(const std::string &path_string, char sep);

	// general attributes 

	/// <summary>
	/// Return true iff this path doesn't begin with / (or a drive letter on Windows).
	/// </summary>
	bool is_relative() const;

	/// <summary>
	/// Return true iff path::is_relative() is false.
	/// </summary>
	bool is_absolute() const;

	/// <summary>
	/// Return true iff this path is the root of the host's filesystem.
	/// </summary>
	bool is_root() const;

	/// <summary>
	/// Return a new path that points to the directory containing the current path
	/// Return the path unchanged if this path is the absolute or relative root.
	/// </summary>
	path parent() const;

	/// <summary>
	/// Return the last component of this path.
	/// </summary>
	std::string basename();

	// conversion

	/// <summary>
	/// Create a string representing this path separated by the provided
	/// separator or the system-default separator if not provided.
	/// </summary>
	std::string to_string(char sep = separator()) const;

	/// <summary>
	/// If this path is relative, append each component of this path
	/// to base_path and return the resulting absolute path. Otherwise,
	/// the the current path will be returned and base_path will be ignored.
	/// </summary>
	path make_absolute(const path &base_path) const;

	// filesystem attributes

	/// <summary>
	/// Return true iff the file or directory pointed to by this path
	/// exists on the filesystem.
	/// </summary>
	bool exists() const;

	/// <summary>
	/// Return true if the file or directory pointed to by this path
	/// is a directory.
	/// </summary>
	bool is_directory() const;

	/// <summary>
	/// Return true if the file or directory pointed to by this path
	/// is a regular file.
	/// </summary>
	bool is_file() const;

	// filesystem

	/// <summary>
	/// Open the file pointed to by this path and return a string containing
	/// the files contents.
	/// </summary>
	std::string read_contents() const;

	// mutators

	/// <summary>
	/// Append the provided part to this path and return a reference to this same path
	/// so calls can be chained.
	/// </summary>
	path &append(const std::string &to_append);

	/// <summary>
	/// Append the provided part to this path and return the result.
	/// </summary>
	path append(const std::string &to_append) const;

	/// <summary>
	/// Append the provided part to this path and return a reference to this same path
	/// so calls can be chained.
	/// </summary>
	path &append(const path &to_append);

	/// <summary>
	/// Append the provided part to this path and return the result.
	/// </summary>
	path append(const path &to_append) const;

	// iterators

	/// <summary>
	/// An iterator to the first element in this path.
	/// </summary>
	iterator begin();

	/// <summary>
	/// An iterator to one past the last element in this path.
	/// </summary>
	iterator end();

	/// <summary>
	/// An iterator to the first element in this path.
	/// </summary>
	const_iterator begin() const;

	/// <summary>
	/// A const iterator to one past the last element in this path.
	/// </summary>
	const_iterator end() const;

	/// <summary>
	/// An iterator to the first element in this path.
	/// </summary>
	const_iterator cbegin() const;

	/// <summary>
	/// A const iterator to one past the last element in this path.
	/// </summary>
	const_iterator cend() const;

	/// <summary>
	/// A reverse iterator to the last element in this path.
	/// </summary>
	reverse_iterator rbegin();

	/// <summary>
	/// A reverse iterator to one before the first element in this path.
	/// </summary>
	reverse_iterator rend();

	/// <summary>
	/// A const reverse iterator to the last element in this path.
	/// </summary>
	const_reverse_iterator rbegin() const;

	/// <summary>
	/// A const reverse iterator to one before the first element in this path.
	/// </summary>
	const_reverse_iterator rend() const;

	/// <summary>
	/// A const reverse iterator to the last element in this path.
	/// </summary>
	const_reverse_iterator crbegin() const;

	/// <summary>
	/// A const reverse iterator to one before the first element in this path.
	/// </summary>
	const_reverse_iterator crend() const;

protected:
	std::string to_hash_string() const override;

private:
	container parts_;
};

} // namespace xlnt
