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
#include <string>
#include <vector>

#include <detail/binary.hpp>
#include <detail/unicode.hpp>
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

const directory_id End = -1;

void visit(compound_document_entry *entry)
{
}

void traverse_storages(std::vector<compound_document_entry> &entries, compound_document_entry *storage);

void traverse_storage(std::vector<compound_document_entry> &entries, compound_document_entry *entry)
{
    if (entry->prev >= 0)
    {
        traverse_storage(entries, &entries[entry->prev]);
    }

    visit(entry);

    if (entry->type == compound_document_entry::entry_type::UserStorage
        || entry->type == compound_document_entry::entry_type::RootStorage)
    {
        traverse_storages(entries, entry);
    }

    if (entry->next >= 0)
    {
        traverse_storage(entries, &entries[entry->next]);
    }
}

void traverse_storages(std::vector<compound_document_entry> &entries, compound_document_entry *storage)
{
    visit(storage);

    if (storage->child >= 0)
    {
        traverse_storage(entries, &entries[storage->child]);
    }
}

void update_red_black_tree(std::vector<compound_document_entry> &entries)
{
    using entry_type = compound_document_entry::entry_type;

    if (entries.empty()) return;

    traverse_storages(entries, &entries[0]);

    return;
}

std::vector<std::u16string> split_path(const std::u16string &path_string)
{
    auto sep = path_string.find(u'/');
    auto prev = std::size_t(0);
    auto split = std::vector<std::u16string>();

    while (sep != std::u16string::npos)
    {
        split.push_back(path_string.substr(prev, sep - prev));
        prev = sep;
        sep = path_string.find(u'/');
    }

    return split;
}



} // namespace

namespace xlnt {
namespace detail {

compound_document::compound_document(std::vector<std::uint8_t> &data)
    : writer_(new binary_writer(data)),
      reader_(new binary_reader(data))
{
    header_ = compound_document_header();
    header_.directory_start = allocate_sectors(1);

    write_header();

    insert_entry(u"Root Entry", compound_document_entry::entry_type::RootStorage);
}

compound_document::compound_document(const std::vector<std::uint8_t> &data)
    : reader_(new binary_reader(data))
{
    read_header();

    if (!header_is_valid(header_))
    {
        throw xlnt::exception("bad ole");
    }

    read_msat();
    read_sat();
    read_ssat();
    read_directory_tree();
}

compound_document::~compound_document()
{
}

std::size_t compound_document::sector_size()
{
    return static_cast<std::size_t>(1) << header_.sector_size_power;
}

std::size_t compound_document::short_sector_size()
{
    return static_cast<std::size_t>(1) << header_.short_sector_size_power;
}


std::vector<std::uint8_t> compound_document::read_stream(const std::u16string &name)
{
    const auto &entry = find_entry(name);

    auto stream_data = (entry.size < header_.threshold)
        ? read_short(entry.start)
        : read(entry.start);
    stream_data.resize(entry.size);

    return stream_data;
}

void compound_document::write_stream(const std::u16string &name, const std::vector<std::uint8_t> &data)
{
    const auto num_sectors = data.size() / sector_size() + (data.size() % sector_size() ? 1 : 0);

    auto &entry = contains_entry(name)
        ? find_entry(name)
        : insert_entry(name, compound_document_entry::entry_type::UserStream);
    entry.start = allocate_sectors(num_sectors);

    write(data, entry.start);
    write_directory_tree();
}

void compound_document::write(const std::vector<byte> &data, sector_id start)
{
    const auto header_size = sizeof(compound_document_header);
    const auto num_sectors = data.size() / sector_size() + (data.size() % sector_size() ? 1 : 0);
    auto current_sector = start;

    for (auto i = std::size_t(0); i < num_sectors; ++i)
    {
        auto stream_position = i * sector_size();
        auto stream_size = std::min(sector_size(), data.size() - stream_position);
        writer_->offset(header_size + sector_size() * current_sector);
        writer_->append(data, stream_position, stream_size);
        current_sector = sat_[current_sector];
    }
}

void compound_document::write_short(const std::vector<byte> &data, sector_id start)
{
    const auto num_sectors = data.size() / sector_size() + (data.size() % sector_size() ? 1 : 0);
    auto current_sector = start;

    for (auto i = std::size_t(0); i < num_sectors; ++i)
    {
        auto position = sector_size() * current_sector;
        auto current_sector_size = data.size() % sector_size();
        writer_->append(data, position, current_sector_size);
        current_sector = ssat_[current_sector];
    }
}

std::vector<byte> compound_document::read(sector_id current)
{
    const auto header_size = sizeof(compound_document_header);
    auto sector_data = std::vector<byte>();
    auto sector_data_writer = binary_writer(sector_data);

    while (current >= 0)
    {
        reader_->offset(header_size + sector_size() * current);
        auto sector_data = reader_->read_vector<byte>(sector_size());
        sector_data_writer.append(sector_data);
        current = sat_[current];
    }

    return sector_data;
}

std::vector<byte> compound_document::read_short(sector_id current_short)
{
    const auto header_size = sizeof(compound_document_header);
    auto sector_data = std::vector<byte>();
    auto sector_data_writer = binary_writer(sector_data);
    auto first_short_sector = find_entry(u"Root Entry").start;

    while (current_short >= 0)
    {
        auto short_position = static_cast<std::size_t>(short_sector_size() * current_short);
        auto current_sector_index = 0;
        auto current_sector = first_short_sector;

        while (current_sector_index < short_position / sector_size())
        {
            current_sector = sat_[current_sector];
            ++current_sector_index;
        }

        auto offset = short_position % sector_size();
        reader_->offset(header_size + sector_size() * current_sector + offset);
        auto sector_data = reader_->read_vector<byte>(short_sector_size());
        sector_data_writer.append(sector_data);

        current_short = ssat_[current_short];
    }

    return sector_data;
}

void compound_document::read_msat()
{
    msat_ = sector_chain(
        header_.msat.begin(),
        header_.msat.begin()
        + std::min(header_.msat.size(),
            static_cast<std::size_t>(header_.num_msat_sectors)));

    if (header_.num_msat_sectors > std::uint32_t(109))
    {
        auto current_sector = header_.extra_msat_start;

        for (auto r = std::uint32_t(0); r < header_.num_extra_msat_sectors; ++r)
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

    for (auto msat_sector : msat_)
    {
        reader_->offset(sector_data_start() + sector_size() * msat_sector);
        auto sat_sectors = reader_->read_vector<sector_id>(sector_size() / sizeof(sector_id));

        std::copy(
            sat_sectors.begin(),
            sat_sectors.end(),
            std::back_inserter(sat_));
    }
}

void compound_document::read_ssat()
{
    ssat_.clear();
    auto current = header_.short_table_start;

    while (current >= 0)
    {
        reader_->offset(sector_data_start() + sector_size() * current);
        auto ssat_sectors = reader_->read_vector<sector_id>(sector_size() / sizeof(sector_id));

        std::copy(
            ssat_sectors.begin(),
            ssat_sectors.end(),
            std::back_inserter(ssat_));

        current = sat_[current];
    }
}

sector_id compound_document::allocate_sectors(std::size_t count)
{
    const auto sector_data_start = sizeof(compound_document_header);
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    auto num_free = std::count(sat_.begin(), sat_.end(), FreeSector);

    if (num_free < count)
    {
        // increase size of allocation table, plus extra sector for sat table
        auto new_size = sat_.size() + count - num_free + 1;
        new_size = (new_size / sectors_per_sector 
            + (new_size % sectors_per_sector != 0 ? 1 : 0)) * sectors_per_sector;
        sat_.resize(new_size, FreeSector);

        // allocate new sat sector
        auto new_sat_sector = allocate_sectors(1);
        msat_.push_back(new_sat_sector);
        ++header_.num_msat_sectors;

        writer_->offset(sector_data_start + new_sat_sector * sector_size());
        writer_->append(sat_, sat_.size() - sectors_per_sector, sectors_per_sector);

        if (msat_.size() > std::size_t(109))
        {
            // allocate extra msat sector
            ++header_.num_extra_msat_sectors;
            auto new_msat_sector = allocate_sectors(1);
            writer_->offset(sector_data_start + new_msat_sector * sector_size());
            writer_->write(new_msat_sector);
        }
        else
        {
            header_.msat.at(msat_.size() - 1) = new_sat_sector;
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

        if (sector_data_start + (current + 1) * sector_size() > writer_->size())
        {
            writer_->extend(sector_size());
        }

        current = next;
        ++allocated;
    }

    return start;
}

compound_document_entry &compound_document::insert_entry(
    const std::u16string &name,
    compound_document_entry::entry_type type)
{
    auto entry_id = directory_id(entries_.size());
    entries_.push_back(compound_document_entry());
    auto &entry = entries_.back();

    entry.name(name);
    entry.type = type;

    if (entry_id != 0)
    {
        auto &root_entry = entries_.at(0);

        if (root_entry.child == -1)
        {
            root_entry.child = entry_id;
        }
        else
        {
            auto parent = 0;
            auto current = root_entry.child;

            while (current >= 0)
            {
                if (name > entries_[current].name())
                {
                    parent = current;
                    current = entries_[current].next;
                }
                else
                {
                    parent = current;
                    current = entries_[current].prev;
                }
            }

            if (name > entries_[parent].name())
            {
                entries_[parent].next = entry_id;
            }
            else
            {
                entries_[parent].prev = entry_id;
            }
        }
    }

    update_red_black_tree(entries_);
    write_directory_tree();

    return entry;
}

void compound_document::write_directory_tree()
{
    const auto entries_per_sector = static_cast<directory_id>(sector_size() 
        / sizeof(compound_document_entry));
    const auto required_sectors = entries_.size() / entries_per_sector
        + (entries_.size() % entries_per_sector) ? 1 : 0;
    auto current_sector_id = header_.directory_start;
    auto entry_index = directory_id(0);

    for (auto &e : entries_)
    {
        writer_->offset(sector_data_start() + current_sector_id * sector_size());
        writer_->write(e);

        ++entry_index;

        if (entry_index % entries_per_sector == 0)
        {
            current_sector_id = sat_[current_sector_id];
        }        
    }
}

std::size_t compound_document::sector_data_start()
{
    return sizeof(compound_document_header);
}

bool compound_document::contains_entry(const std::u16string &path)
{
    for (auto &e : entries_)
    {
        if (e.name() == path)
        {
            return true;
        }
    }

    return false;
}

void compound_document::write_header()
{
    writer_->offset(0);
    writer_->write(header_);
}

void compound_document::read_header()
{
    reader_->offset(0);
    header_ = reader_->read<compound_document_header>();
}

compound_document_entry &compound_document::find_entry(const std::u16string &name)
{
    for (auto &e : entries_)
    {
        if (e.name() == name)
        {
            return e;
        }
    }

    throw xlnt::exception("not found");
}

void compound_document::read_directory_tree()
{
    entries_.clear();

    auto current_sector_id = header_.directory_start;
    auto current_id = directory_id(0);

    while (current_sector_id >= 0)
    {
        reader_->offset(sector_data_start() + current_sector_id * sector_size());

        for (auto i = std::size_t(0); i < sector_size() / sizeof(compound_document_entry); ++i)
        {
            entries_.push_back(reader_->read<compound_document_entry>());
        }

        current_sector_id = sat_[current_sector_id];
    }
}

} // namespace detail
} // namespace xlnt
