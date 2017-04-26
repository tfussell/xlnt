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

#include <array>
#include <memory>
#include <string>
#include <unordered_map>

#include <detail/binary.hpp>

namespace xlnt {
namespace detail {

using directory_id = std::int32_t;
using sector_id = std::int32_t;
using sector_chain = std::vector<sector_id>;

struct compound_document_header
{
    enum class byte_order_type : uint16_t
    {
        big_endian = 0xFFFE,
        little_endian = 0xFEFF
    };

    std::uint64_t file_id = 0xE11AB1A1E011CFD0;
    std::array<std::uint8_t, 16> ignore1 = { { 0 } };
    std::uint16_t revision = 0x003E;
    std::uint16_t version = 0x0003;
    byte_order_type byte_order = byte_order_type::little_endian;
    std::uint16_t sector_size_power = 9;
    std::uint16_t short_sector_size_power = 6;
    std::array<std::uint8_t, 10> ignore2 = { { 0 } };
    std::uint32_t num_msat_sectors = 0;
    sector_id directory_start = -1;
    std::array<std::uint8_t, 4> ignore3 = { { 0 } };
    std::uint32_t threshold = 4096;
    sector_id ssat_start = -2;
    std::uint32_t num_short_sectors = 0;
    sector_id extra_msat_start = -2;
    std::uint32_t num_extra_msat_sectors = 0;
    std::array<sector_id, 109> msat = { 0 };
};

struct compound_document_entry
{
    void name(const std::u16string &new_name)
    {
        name_length = std::min(static_cast<std::uint16_t>(new_name.size()), std::uint16_t(31));
        std::copy(new_name.begin(), new_name.begin() + name_length, name_array.begin());
        name_array[name_length] = 0;
        name_length = (name_length + 1) * 2;
    }

    std::u16string name() const
    {
        return std::u16string(name_array.begin(),
            name_array.begin() + (name_length - 1) / 2);
    }

    enum class entry_type : std::uint8_t
    {
        Empty = 0,
        UserStorage = 1,
        UserStream = 2,
        LockBytes = 3,
        Property = 4,
        RootStorage = 5
    };

    enum class entry_color : std::uint8_t
    {
        Red = 0,
        Black = 1
    };

    std::array<char16_t, 32> name_array = { { 0 } };
    std::uint16_t name_length = 2;
    entry_type type = entry_type::Empty;
    entry_color color = entry_color::Black;
    directory_id prev = -1;
    directory_id next = -1;
    directory_id child = -1;
    std::array<std::uint8_t, 36> ignore;
    sector_id start = -2;
    std::uint32_t size = 0;
    std::uint32_t ignore2;
};

class red_black_tree;

class compound_document
{
public:
    compound_document(std::vector<std::uint8_t> &data);
    compound_document(const std::vector<std::uint8_t> &data);
    ~compound_document();

    std::vector<std::uint8_t> read_stream(const std::u16string &filename);
    void write_stream(const std::u16string &filename, const std::vector<std::uint8_t> &data);

private:
    std::size_t sector_size();
    std::size_t short_sector_size();
    std::size_t sector_data_start();

    bool contains_entry(const std::u16string &path);
    compound_document_entry &find_entry(const std::u16string &path);

    std::vector<byte> read(sector_id start);
    std::vector<byte> read_short(sector_id start);

    sector_chain follow_chain(sector_id start);

    void read_msat();
    void read_sat();
    void read_ssat();
    void read_header();
    void read_directory_tree();

    void write(const std::vector<byte> &data, sector_id start);
    void write_short(const std::vector<byte> &data, sector_id start);

    void write_header();
    void write_directory_tree();

    void print_directory();

    sector_id allocate_sectors(std::size_t sectors);
    void reallocate_sectors(sector_id start, std::size_t sectors);
    sector_id allocate_short_sectors(std::size_t sectors);

    compound_document_entry &insert_entry(const std::u16string &path, 
        compound_document_entry::entry_type type);

    std::unique_ptr<binary_reader> reader_;
    std::unique_ptr<binary_writer> writer_;

    compound_document_header header_;

    sector_chain msat_;
    sector_chain sat_;
    sector_chain ssat_;

    std::vector<compound_document_entry> entries_;

    std::unique_ptr<red_black_tree> rb_tree_;
};

} // namespace detail
} // namespace xlnt
