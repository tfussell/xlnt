// Copyright (c) 2014-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <string>
#include <utility>
#include <vector>

#include <xlnt/xlnt_config.hpp>

namespace xlnt {

/// <summary>
/// Encapsulates a path that points to location in a filesystem.
/// </summary>
class XLNT_API path
{
public:
    /// <summary>
    /// The system-specific path separator character (e.g. '/' or '\').
    /// </summary>
    static char system_separator();

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
    /// Return true iff this path is the root directory.
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
    std::string filename() const;

    /// <summary>
    /// Return the part of the path following the last dot in the filename.
    /// </summary>
    std::string extension() const;

    /// <summary>
    /// Return a pair of strings resulting from splitting the filename on the last dot.
    /// </summary>
    std::pair<std::string, std::string> split_extension() const;

    // conversion

    /// <summary>
    /// Create a string representing this path separated by the provided
    /// separator or the system-default separator if not provided.
    /// </summary>
    std::vector<std::string> split() const;

    /// <summary>
    /// Create a string representing this path separated by the provided
    /// separator or the system-default separator if not provided.
    /// </summary>
    const std::string &string() const;

#ifdef _MSC_VER
    /// <summary>
    /// Create a wstring representing this path separated by the provided
    /// separator or the system-default separator if not provided.
    /// </summary>
    std::wstring wstring() const;
#endif

    /// <summary>
    /// If this path is relative, append each component of this path
    /// to base_path and return the resulting absolute path. Otherwise,
    /// the the current path will be returned and base_path will be ignored.
    /// </summary>
    path resolve(const path &base_path) const;

    /// <summary>
    /// The inverse of path::resolve. Creates a relative path from an absolute
    /// path by removing the common root between base_path and this path.
    /// If the current path is already relative, return it unchanged.
    /// </summary>
    path relative_to(const path &base_path) const;

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
    /// Append the provided part to this path and return the result.
    /// </summary>
    path append(const std::string &to_append) const;

    /// <summary>
    /// Append the provided part to this path and return the result.
    /// </summary>
    path append(const path &to_append) const;

    /// <summary>
    /// Returns true if left path is equal to right path.
    /// </summary>
    bool operator==(const path &other) const;

    /// <summary>
    /// Returns true if left path is equal to right path.
    /// </summary>
    bool operator!=(const path &other) const;

private:
    /// <summary>
    /// Returns the character that separates directories in the path.
    /// On POSIX style filesystems, this is always '/'.
    /// On Windows, this is the character that separates the drive letter from
    /// the rest of the path for absolute paths with a drive letter, '/' if the path
    /// is absolute and starts with '/', and '/' or '\' for relative paths
    /// depending on which splits the path into more directory components.
    /// </summary>
    char guess_separator() const;

    /// <summary>
    /// A string that represents this path.
    /// </summary>
    std::string internal_;
};

} // namespace xlnt

namespace std {

/// <summary>
/// Template specialization to allow xlnt:path to be used as a key in a std container.
/// </summary>
template <>
struct hash<xlnt::path>
{
    /// <summary>
    /// Returns a hashed represenation of the given path.
    /// </summary>
    size_t operator()(const xlnt::path &p) const
    {
        static hash<string> hasher;
        return hasher(p.string());
    }
};

} // namespace std
