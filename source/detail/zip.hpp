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

namespace xlnt {
namespace detail {

struct zip_file_header
{
    std::uint16_t version = 20;
    std::uint16_t flags = 0;
    std::uint16_t compression_type = 8;
    std::uint16_t stamp_date, stamp_time = 0;
    std::uint32_t crc = 0;
    std::uint32_t compressed_size = 0;
    std::uint32_t uncompressed_size = 0;
    std::string filename;
    std::string comment;
    std::vector<std::uint8_t> extra;
    std::uint32_t header_offset = 0;

    zip_file_header();

    bool read(std::istream &istream, const bool global);
    void Write(std::ostream &ostream, const bool global) const;
};

class zip_file_istream : public std::istream
{
public:
    zip_file_istream(std::unique_ptr<std::streambuf> &&buf);
    virtual ~zip_file_istream();
    
private:
    std::unique_ptr<std::streambuf> buf;
};

class zip_file_ostream : public std::ostream
{
public:
    zip_file_ostream(std::unique_ptr<std::streambuf> &&buf);
    virtual ~zip_file_ostream();

private:
    std::unique_ptr<std::streambuf> buf;
};

class ZipFileWriter
{
public:
    ZipFileWriter(std::ostream &filename);
    virtual ~ZipFileWriter();
    std::ostream &open(const std::string &filename);
    void close();
private:
    std::vector<zip_file_header> file_headers_;
    std::ostream &tarstream_;
    std::unique_ptr<std::ostream> write_stream_;
};

class ZipFileReader
{
public:
    ZipFileReader(std::istream &stream);
    virtual ~ZipFileReader();
    std::istream &open(const std::string &filename);
    std::vector<std::string> files() const;
    bool has_file(const std::string &filename) const;

private:
    bool read_central_header();

    std::unordered_map<std::string, zip_file_header> file_headers_;
    std::istream &source_stream_;
    std::unique_ptr<std::istream> read_stream_;
};

} // namespace detail
} // namespace xlnt
