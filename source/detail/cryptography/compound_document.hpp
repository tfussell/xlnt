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
#include <detail/unicode.hpp>

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
    std::array<sector_id, 109> msat = {{0}};
};

struct compound_document_entry
{
    void name(const std::string &new_name)
    {
        auto u16_name = utf8_to_utf16(new_name);
        name_length = std::min(static_cast<std::uint16_t>(u16_name.size()), std::uint16_t(31));
        std::copy(u16_name.begin(), u16_name.begin() + name_length, name_array.begin());
        name_array[name_length] = 0;
        name_length = (name_length + 1) * 2;
    }

    std::string name() const
    {
        return utf16_to_utf8(std::u16string(name_array.begin(),
            name_array.begin() + (name_length - 1) / 2));
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

class compound_document_istreambuf;
class compound_document_ostreambuf;

class compound_document
{
public:
    compound_document(std::istream &in);
    compound_document(std::ostream &out);
    ~compound_document();

    void close();

    std::istream &open_read_stream(const std::string &filename);
    std::ostream &open_write_stream(const std::string &filename);

private:
    friend class compound_document_istreambuf;
    friend class compound_document_ostreambuf;

    template<typename T>
    void read_sector(sector_id id, binary_writer<T> &writer);
    template<typename T>
    void read_sector_chain(sector_id id, binary_writer<T> &writer);
    template<typename T>
    void read_sector_chain(sector_id start, binary_writer<T> &writer, sector_id offset, std::size_t count);
    template<typename T>
    void read_short_sector(sector_id id, binary_writer<T> &writer);
    template<typename T>
    void read_short_sector_chain(sector_id start, binary_writer<T> &writer);
    template<typename T>
    void read_short_sector_chain(sector_id start, binary_writer<T> &writer, sector_id offset, std::size_t count);

    sector_chain follow_chain(sector_id start, const sector_chain &table);

    template<typename T>
    void write_sector(binary_reader<T> &reader, sector_id id);
    template<typename T>
    void write_short_sector(binary_reader<T> &reader, sector_id id);

    void read_header();
    void read_msat();
    void read_sat();
    void read_ssat();
    void read_entry(directory_id id);
    void read_directory();

    void write_header();
    void write_msat();
    void write_sat();
    void write_ssat();
    void write_entry(directory_id id);
    void write_directory();

    std::size_t sector_size();
    std::size_t short_sector_size();
    std::size_t sector_data_start();

    void print_directory();

    sector_id allocate_msat_sector();
    sector_id allocate_sat_sector();
    sector_id allocate_ssat_sector();

    sector_id allocate_sector();
    sector_chain allocate_sectors(std::size_t sectors);
    sector_id allocate_short_sector();
    sector_chain allocate_short_sectors(std::size_t sectors);

    bool contains_entry(const std::string &path,
        compound_document_entry::entry_type type);
    directory_id find_entry(const std::string &path,
        compound_document_entry::entry_type type);
    directory_id next_empty_entry();
    directory_id insert_entry(const std::string &path,
        compound_document_entry::entry_type type);

    // Red black tree helper functions
    void tree_insert(directory_id new_id, directory_id storage_id);
    void tree_insert_fixup(directory_id x);
    std::string tree_path(directory_id id);
    void tree_rotate_left(directory_id x);
    void tree_rotate_right(directory_id y);
    directory_id &tree_left(directory_id id);
    directory_id &tree_right(directory_id id);
    directory_id &tree_parent(directory_id id);
    directory_id &tree_root(directory_id id);
    directory_id &tree_child(directory_id id);
    std::string tree_key(directory_id id);
    compound_document_entry::entry_color &tree_color(directory_id id);

    compound_document_header header_;
    sector_chain msat_;
    sector_chain sat_;
    sector_chain ssat_;
    std::vector<compound_document_entry> entries_;

    std::unordered_map<directory_id, directory_id> parent_storage_;
    std::unordered_map<directory_id, directory_id> parent_;

    std::istream *in_;
    std::ostream *out_;

    std::unique_ptr<compound_document_istreambuf> stream_in_buffer_;
    std::istream stream_in_;
    std::unique_ptr<compound_document_ostreambuf> stream_out_buffer_;
    std::ostream stream_out_;
};

} // namespace detail
} // namespace xlnt
