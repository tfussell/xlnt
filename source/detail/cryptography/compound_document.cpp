// Copyright (C) 2016-2017 Thomas Fussell
// Copyright (C) 2002-2007 Ariya Hidayat (ariya@kde.org).
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions
// are met:
//
// 1. Redistributions of source code must retain the above copyright
// notice, this list of conditions and the following disclaimer.
// 2. Redistributions in binary form must reproduce the above copyright
// notice, this list of conditions and the following disclaimer in the
// documentation and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR
// IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
// OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
// IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT,
// INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
// NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
// THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
// THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 
#include <array>
#include <algorithm>
#include <cstring>
#include <fstream>
#include <iostream>
#include <list>
#include <string>
#include <unordered_set>
#include <vector>

#include <detail/binary.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using namespace xlnt::detail;

const sector_id FreeSector = -1;

// TODO: I don't think this is Unicode-aware...
std::u16string to_lower(std::u16string string)
{
    for (auto &c : string)
    {
        if (c <= 255)
        {
            c = static_cast<std::uint16_t>(::tolower(c));
        }
    }

    return string;
}

bool header_is_valid(const compound_document_header &h)
{
    if (h.threshold != 4096)
    {
        return false;
    }

    if (h.num_msat_sectors == 0
        || (h.num_msat_sectors > 109 && h.num_msat_sectors > (h.num_extra_msat_sectors * 127) + 109)
        || ((h.num_msat_sectors < 109) && (h.num_extra_msat_sectors != 0)))
    {
        return false;
    }

    if (h.short_sector_size_power > h.sector_size_power
        || h.sector_size_power <= 6
        || h.sector_size_power >= 31)
    {
        return false;
    }

    return true;
}

compound_document_entry &look_up_entry(binary_writer &writer, sector_id directory_start, directory_id index)
{
    return *reinterpret_cast<compound_document_entry *>(writer.data().data());
}

const compound_document_entry &look_up_entry(binary_reader &reader, sector_id directory_start, directory_id index)
{
    return *reinterpret_cast<const compound_document_entry *>(reader.data().data());
}

const directory_id End = -1;

sector_id next_free_sector(const sector_chain &allocation_table)
{
    auto index = sector_id(0);

    for (auto sector : allocation_table)
    {
        if (sector == FreeSector)
        {
            return index;
        }

        ++index;
    }

    return FreeSector;
}

} // namespace

namespace xlnt {
namespace detail {

compound_document::compound_document(std::vector<std::uint8_t> &data)
    : writer_(new binary_writer(data)),
      reader_(new binary_reader(data))
{
    data.resize(sizeof(compound_document_header));

    auto new_header = compound_document_header();
    new_header.msat.fill(FreeSector);
    header() = new_header;

    auto &root = insert_entry(u"Root Entry");
    root.type = compound_document_entry::entry_type::RootStorage;
    root.color = compound_document_entry::entry_color::Black;
}

compound_document::compound_document(const std::vector<std::uint8_t> &data)
    : reader_(new binary_reader(data))
{
    if (!header_is_valid(header()))
    {
        throw xlnt::exception("bad ole");
    }

    read_msat();
    read_sat();
    read_ssat();
}

compound_document::~compound_document()
{
}

std::vector<std::uint8_t> compound_document::read_stream(const std::u16string &name)
{
    if (!contains_entry(name))
    {
        throw xlnt::exception("document doesn't contain stream with the given name");
    }

    auto &entry = find_entry(name);
    auto stream_data = std::vector<byte>();

    if (entry.size < header().threshold)
    {
        stream_data = read_short(entry.start);
    }
    else
    {
        stream_data = read(entry.start);
    }

    stream_data.resize(entry.size);

    return stream_data;
}

void compound_document::write_stream(const std::u16string &name, const std::vector<std::uint8_t> &data)
{
    const auto sector_size = header().sector_size();
    const auto num_sectors = data.size() / sector_size + (data.size() % sector_size ? 1 : 0);
    auto &entry = contains_entry(name) ? find_entry(name) : insert_entry(name);
    
    entry.type = compound_document_entry::entry_type::UserStream;
    entry.start = allocate_sector(num_sectors);

    write(data, entry.start);
}

void compound_document::write(const std::vector<byte> &data, sector_id start)
{
    const auto sector_size = header().sector_size();
    const auto header_size = sizeof(compound_document_header);
    const auto num_sectors = data.size() / sector_size + (data.size() % sector_size ? 1 : 0);
    auto current_sector = start;

    for (auto i = std::size_t(0); i < num_sectors; ++i)
    {
        auto stream_position = i * sector_size;
        auto stream_size = std::min(sector_size, data.size() - stream_position);
        writer_->offset(header_size + sector_size * current_sector);
        writer_->append(data, stream_position, stream_size);
        current_sector = sat_[current_sector];
    }
}

void compound_document::write_short(const std::vector<byte> &data, sector_id start)
{
    const auto sector_size = header().sector_size();
    const auto num_sectors = data.size() / sector_size + (data.size() % sector_size ? 1 : 0);
    auto current_sector = start;

    for (auto i = std::size_t(0); i < num_sectors; ++i)
    {
        auto position = sector_size * current_sector;
        auto current_sector_size = data.size() % sector_size;
        writer_->append(data, position, current_sector_size);
        current_sector = ssat_[current_sector];
    }
}

compound_document_header &compound_document::header()
{
    reader_->offset(0);
    return *const_cast<compound_document_header *>(
        &reader_->read_reference<compound_document_header>());
}

std::vector<byte> compound_document::read(sector_id current)
{
    const auto sector_size = header().sector_size();
    const auto header_size = sizeof(compound_document_header);
    auto sector_data = std::vector<byte>();
    auto sector_data_writer = binary_writer(sector_data);

    while (current >= 0)
    {
        reader_->offset(header_size + sector_size * current);
        auto sector_data = reader_->read_vector<byte>(sector_size);
        sector_data_writer.append(sector_data);
        current = sat_[current];
    }

    return sector_data;
}

std::vector<byte> compound_document::read_short(sector_id current_short)
{
    const auto short_sector_size = header().short_sector_size();
    const auto sector_size = header().sector_size();
    const auto header_size = sizeof(compound_document_header);
    auto sector_data = std::vector<byte>();
    auto sector_data_writer = binary_writer(sector_data);
    auto first_short_sector = find_entry(u"Root Entry").start;

    while (current_short >= 0)
    {
        auto short_position = static_cast<std::size_t>(short_sector_size * current_short);
        auto current_sector_index = 0;
        auto current_sector = first_short_sector;

        while (current_sector_index < short_position / sector_size)
        {
            current_sector = sat_[current_sector];
            ++current_sector_index;
        }

        auto offset = short_position % sector_size;
        reader_->offset(header_size + sector_size * current_sector + offset);
        auto sector_data = reader_->read_vector<byte>(short_sector_size);
        sector_data_writer.append(sector_data);

        current_short = ssat_[current_short];
    }

    return sector_data;
}

void compound_document::read_msat()
{
    msat_ = sector_chain(
        header().msat.begin(),
        header().msat.begin()
        + std::min(header().msat.size(),
            static_cast<std::size_t>(header().num_msat_sectors)));

    if (header().num_msat_sectors > std::uint32_t(109))
    {
        auto current_sector = header().sector_table_start;

        for (auto r = std::uint32_t(0); r < header().num_extra_msat_sectors; ++r)
        {
            auto current_sector_data = read({ current_sector });
            auto current_sector_reader = binary_reader(current_sector_data);
            auto current_sector_sectors = current_sector_reader.to_vector<sector_id>();

            current_sector = current_sector_sectors.back();
            current_sector_sectors.pop_back();

            std::copy(
                current_sector_sectors.begin(),
                current_sector_sectors.end(),
                std::back_inserter(msat_));
        }
    }
}

void compound_document::read_sat()
{
    sat_.clear();

    const auto sector_size = header().sector_size();
    const auto sector_data_start = sizeof(compound_document_header);

    for (auto msat_sector : msat_)
    {
        reader_->offset(sector_data_start + sector_size * msat_sector);
        auto sat_sectors = reader_->read_vector<sector_id>(sector_size / sizeof(sector_id));

        std::copy(
            sat_sectors.begin(),
            sat_sectors.end(),
            std::back_inserter(sat_));
    }
}

void compound_document::read_ssat()
{
    ssat_.clear();

    const auto sector_size = header().sector_size();
    const auto sector_data_start = sizeof(compound_document_header);
    auto current = header().short_table_start;

    while (current >= 0)
    {
        reader_->offset(sector_data_start + sector_size * current);
        auto ssat_sectors = reader_->read_vector<sector_id>(sector_size / sizeof(sector_id));

        std::copy(
            ssat_sectors.begin(),
            ssat_sectors.end(),
            std::back_inserter(ssat_));

        current = sat_[current];
    }
}

sector_id compound_document::allocate_sector(std::size_t count)
{
    const auto sector_data_start = sizeof(compound_document_header);
    const auto sector_size = header().sector_size();
    const auto sectors_per_sector = sector_size / sizeof(sector_id);
    auto num_free = std::count(sat_.begin(), sat_.end(), FreeSector);

    if (num_free < count)
    {
        // increase size of allocation table, plus extra sector for sat table
        auto new_size = sat_.size() + count - num_free + 1;
        new_size = (new_size / sectors_per_sector 
            + (new_size % sectors_per_sector != 0 ? 1 : 0)) * sectors_per_sector;
        sat_.resize(new_size, FreeSector);

        // allocate new sat sector
        auto new_sat_sector = allocate_sector(1);
        msat_.push_back(new_sat_sector);
        ++header().num_msat_sectors;

        if (header().sector_table_start == -1)
        {
            header().sector_table_start = new_sat_sector;
        }

        writer_->offset(sector_data_start + new_sat_sector * sector_size);
        writer_->append(sat_, sat_.size() - sectors_per_sector, sectors_per_sector);

        if (msat_.size() > std::size_t(109))
        {
            // allocate extra msat sector
            ++header().num_extra_msat_sectors;
            auto new_msat_sector = allocate_sector(1);
            writer_->offset(sector_data_start + new_msat_sector * sector_size);
            writer_->write(new_msat_sector);
        }
        else
        {
            header().msat.at(msat_.size() - 1) = new_sat_sector;
        }
    }

    auto allocated = 0;
    const auto start_iter = std::find(sat_.begin(), sat_.end(), FreeSector);
    const auto start = sector_id(start_iter - sat_.begin());
    auto current = start;

    while (allocated < count)
    {
        const auto next_iter = std::find(sat_.begin() + current + 1, sat_.end(), FreeSector);
        const auto next = sector_id(next_iter - sat_.begin());
        sat_[current] = (allocated == count - 1) ? -2 : next;

        if (sector_data_start + (current + 1) * sector_size > writer_->size())
        {
            writer_->extend(sector_size);
        }

        current = next;
        ++allocated;
    }

    return start;
}

compound_document_entry &compound_document::insert_entry(const std::u16string &name)
{
    const auto entry_size = sizeof(compound_document_entry);
    const auto sector_size = header().sector_size();
    const auto sector_data_start = sizeof(compound_document_header);
    const auto entries_per_sector = static_cast<directory_id>(sector_size / entry_size);

    auto current_sector_id = header().directory_start;

    if (current_sector_id == -1)
    {
        current_sector_id = allocate_sector();
        header().directory_start = current_sector_id;

        auto sector_entries = std::vector<compound_document_entry>(entries_per_sector);

        writer_->offset(sector_data_start + current_sector_id * sector_size);
        writer_->append(sector_entries);
    }

    auto sector_empty_entry_index = directory_id(0);
    auto any_sector_empty_entry = false;

    while (current_sector_id >= 0 && !any_sector_empty_entry)
    {
        reader_->offset(sector_data_start + current_sector_id * sector_size);

        for (auto i = directory_id(0); i < entries_per_sector; ++i)
        {
            const auto &current_entry = reader_->read_reference<compound_document_entry>();

            if (current_entry.type == compound_document_entry::entry_type::Empty)
            {
                sector_empty_entry_index = i;
                any_sector_empty_entry = true;
                break;
            }

            if (current_entry.name() == name)
            {
                throw xlnt::exception("already exists");
            }
        }

        if (!any_sector_empty_entry)
        {
            current_sector_id = sat_[current_sector_id];
        }
    }

    if (!any_sector_empty_entry)
    {
        // add sector
    }

    reader_->offset(sector_data_start
        + current_sector_id * sector_size
        + sector_empty_entry_index * entry_size);
    auto &new_entry = const_cast<compound_document_entry &>(
        reader_->read_reference<compound_document_entry>());

    new_entry.name(name);
    new_entry.type = compound_document_entry::entry_type::UserStream;

    return new_entry;
}

compound_document_entry &compound_document::find_entry(const std::u16string &name)
{
    const auto entry_size = sizeof(compound_document_entry);
    const auto sector_size = header().sector_size();
    const auto sector_data_start = sizeof(compound_document_header);
    auto current_sector_id = header().directory_start;

    while (current_sector_id >= 0)
    {
        reader_->offset(sector_data_start + current_sector_id * sector_size);

        for (auto i = std::size_t(0); i < sector_size / entry_size; ++i)
        {
            const auto &current_entry = reader_->read_reference<compound_document_entry>();

            if (current_entry.name() == name)
            {
                return const_cast<compound_document_entry &>(current_entry);
            }
        }

        current_sector_id = sat_[current_sector_id];
    }

    throw xlnt::exception("not found");
}

compound_document_entry &compound_document::find_entry(directory_id id)
{
    return look_up_entry(*writer_, 0, 0);
}

bool compound_document::contains_entry(const std::u16string &name)
{
    const auto entry_size = sizeof(compound_document_entry);
    const auto sector_size = header().sector_size();
    const auto sector_data_start = sizeof(compound_document_header);
    auto current_sector_id = header().directory_start;

    while (current_sector_id >= 0)
    {
        reader_->offset(sector_data_start + current_sector_id * sector_size);

        for (auto i = std::size_t(0); i < sector_size / entry_size; ++i)
        {
            const auto &current_entry = reader_->read_reference<compound_document_entry>();

            if (current_entry.type != compound_document_entry::entry_type::Empty
                && current_entry.name() == name)
            {
                return true;
            }
        }

        current_sector_id = sat_[current_sector_id];
    }

    return false;
}

std::vector<directory_id> compound_document::find_siblings(directory_id entry)
{
    return {};
}

std::vector<directory_id> compound_document::children(directory_id entry)
{
    return {};
}

} // namespace detail
} // namespace xlnt
