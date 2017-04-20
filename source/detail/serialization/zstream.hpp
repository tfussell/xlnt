/*
PARTIO SOFTWARE
Copyright 2010 Disney Enterprises, Inc. All rights reserved

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

* Redistributions of source code must retain the above copyright
notice, this list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright
notice, this list of conditions and the following disclaimer in
the documentation and/or other materials provided with the
distribution.

* The names "Disney", "Walt Disney Pictures", "Walt Disney Animation
Studios" or the names of its contributors may NOT be used to
endorse or promote products derived from this software without
specific prior written permission from Walt Disney Pictures.

Disclaimer: THIS SOFTWARE IS PROVIDED BY WALT DISNEY PICTURES AND
CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE, NONINFRINGEMENT AND TITLE ARE DISCLAIMED.
IN NO EVENT SHALL WALT DISNEY PICTURES, THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND BASED ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
*/

#pragma once

#include <iostream>
#include <memory>
#include <unordered_map>
#include <vector>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/path.hpp>

//TODO: don't export these classes (some tests are using them for now)

namespace xlnt {
namespace detail {

/// <summary>
/// A structure representing the header that occurs before each compressed file in a ZIP
/// archive and again at the end of the file with more information.
/// </summary>
struct XLNT_API zheader
{
    std::uint16_t version = 20;
    std::uint16_t flags = 0;
    std::uint16_t compression_type = 8;
    std::uint16_t stamp_date = 0;
    std::uint16_t stamp_time = 0;
    std::uint32_t crc = 0;
    std::uint32_t compressed_size = 0;
    std::uint32_t uncompressed_size = 0;
    std::string filename;
    std::string comment;
    std::vector<std::uint8_t> extra;
    std::uint32_t header_offset = 0;
};

/// <summary>
/// Writes a series of uncompressed binary file data as ostreams into another ostream
/// according to the ZIP format.
/// </summary>
class XLNT_API ozstream
{
public:
    /// <summary>
    /// Construct a new zip_file_writer which writes a ZIP archive to the given stream.
    /// </summary>
    ozstream(std::ostream &stream);

    /// <summary>
    /// Destructor.
    /// </summary>
    virtual ~ozstream();

    /// <summary>
    /// Returns a pointer to a streambuf which compresses the data it receives.
    /// </summary>
    std::unique_ptr<std::streambuf> open(const path &file);

private:
    std::vector<zheader> file_headers_;
    std::ostream &destination_stream_;
};

/// <summary>
/// Reads an archive containing a number of files from an istream and allows them
/// to be decompressed into an istream.
/// </summary>
class XLNT_API izstream
{
public:
    /// <summary>
    /// Construct a new zip_file_reader which reads a ZIP archive from the given stream.
    /// </summary>
    izstream(std::istream &stream);

    /// <summary>
    /// Destructor.
    /// </summary>
    virtual ~izstream();

    /// <summary>
    ///
    /// </summary>
    std::unique_ptr<std::streambuf> open(const path &file) const;

    /// <summary>
    ///
    /// </summary>
    std::string read(const path &file) const;

    /// <summary>
    ///
    /// </summary>
    std::vector<path> files() const;

    /// <summary>
    ///
    /// </summary>
    bool has_file(const path &filename) const;

private:
    /// <summary>
    ///
    /// </summary>
    bool read_central_header();

    /// <summary>
    ///
    /// </summary>
    std::unordered_map<std::string, zheader> file_headers_;

    /// <summary>
    ///
    /// </summary>
    std::istream &source_stream_;
};

} // namespace detail
} // namespace xlnt
