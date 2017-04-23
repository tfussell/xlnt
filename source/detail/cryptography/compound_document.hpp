/*  POLE - Portable C++ library to access OLE Storage
 Copyright (C) 2002-2007 Ariya Hidayat (ariya@kde.org).
 
 Redistribution and use in source and binary forms, with or without
 modification, are permitted provided that the following conditions
 are met:
 
 1. Redistributions of source code must retain the above copyright
 notice, this list of conditions and the following disclaimer.
 2. Redistributions in binary form must reproduce the above copyright
 notice, this list of conditions and the following disclaimer in the
 documentation and/or other materials provided with the distribution.
 
 THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR
 IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
 OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
 IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT,
 INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
 NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
 DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
 THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
 THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

#pragma once

#include <memory>
#include <string>

#include <detail/bytes.hpp>

namespace xlnt {
namespace detail {

class compound_document_reader_impl;
class compound_document_writer_impl;

class compound_document_reader
{
public:
    compound_document_reader(const std::vector<std::uint8_t> &data);
    ~compound_document_reader();

    std::vector<std::uint8_t> read_stream(const std::u16string &filename) const;

private:
    std::unique_ptr<compound_document_reader_impl> d_;
};

class compound_document_writer
{
public:
    compound_document_writer(std::vector<std::uint8_t> &data);
    ~compound_document_writer();

    void write_stream(const std::u16string &filename, const std::vector<std::uint8_t> &data);

private:
    std::unique_ptr<compound_document_writer_impl> d_;
};

} // namespace detail
} // namespace xlnt
