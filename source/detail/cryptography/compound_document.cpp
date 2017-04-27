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
#include <iostream>
#include <string>
#include <vector>

#include <detail/binary.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using namespace xlnt::detail;

int compare_keys(const std::u16string &left, const std::u16string &right)
{
    auto to_lower = [](std::u16string s)
    {
        static const auto locale = std::locale();
        std::use_facet<std::ctype<char16_t>>(locale).tolower(&s[0], &s[0] + s.size());
        return s;
    };

    return to_lower(left).compare(to_lower(right));
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

std::u16string join_path(const std::vector<std::u16string> &path)
{
    auto joined = std::u16string();

    for (auto part : path)
    {
        joined.append(part);
        joined.push_back(u'/');
    }

    return joined;
}

const sector_id FreeSector = -1;
const sector_id EndOfChain = -2;
const sector_id SATSector = -3;
const sector_id MSATSector = -4;

const directory_id End = -1;

} // namespace

namespace xlnt {
namespace detail {

compound_document::compound_document(std::vector<std::uint8_t> &data)
    : writer_(new binary_writer<byte>(data)),
      reader_(new binary_reader<byte>(data))
{
    header().msat.fill(FreeSector);
    header().directory_start = allocate_sector();

    insert_entry(u"Root Entry", compound_document_entry::entry_type::RootStorage);
}

compound_document::compound_document(const std::vector<std::uint8_t> &data)
    : reader_(new binary_reader<byte>(data))
{
    tree_initialize_parent_maps();
}

compound_document::~compound_document()
{
}

std::size_t compound_document::sector_size()
{
    return static_cast<std::size_t>(1) << header().sector_size_power;
}

std::size_t compound_document::short_sector_size()
{
    return static_cast<std::size_t>(1) << header().short_sector_size_power;
}

std::vector<byte> compound_document::read_stream(const std::string &name)
{
    const auto entry_id = find_entry(utf8_to_utf16(name));
    const auto entry = entry_cache_.at(entry_id);

    auto stream_data = std::vector<byte>();
    auto stream_data_writer = binary_writer<byte>(stream_data);

    if (entry->size < header().threshold)
    {
        for (auto sector : follow_ssat_chain(entry->start))
        {
            read_short_sector<byte>(sector, stream_data_writer);
        }
    }
    else
    {
        for (auto sector : follow_sat_chain(entry->start))
        {
            read_sector<byte>(sector, stream_data_writer);
        }
    }
    stream_data.resize(entry->size);

    return stream_data;
}

void compound_document::write_stream(const std::string &name, const std::vector<std::uint8_t> &data)
{
    auto entry_id = contains_entry(utf8_to_utf16(name))
        ? find_entry(utf8_to_utf16(name))
        : insert_entry(utf8_to_utf16(name), compound_document_entry::entry_type::UserStream);
    auto entry = entry_cache_.at(entry_id);
    entry->size = static_cast<std::uint32_t>(data.size());

    auto stream_data_reader = binary_reader<byte>(data);

    if (entry->size < header().threshold)
    {
        const auto num_sectors = data.size() / short_sector_size()
            + (data.size() % short_sector_size() ? 1 : 0);

        auto chain = allocate_short_sectors(num_sectors);
        entry->start = chain.front();

        for (auto sector : follow_ssat_chain(entry->start))
        {
            write_short_sector(stream_data_reader, sector);
        }
    }
    else
    {
        const auto num_sectors = data.size() / short_sector_size()
            + (data.size() % short_sector_size() ? 1 : 0);

        auto chain = allocate_sectors(num_sectors);
        entry->start = chain.front();

        for (auto sector : follow_sat_chain(entry->start))
        {
            write_sector(stream_data_reader, sector);
        }
    }
}

template<typename T>
void compound_document::write_sector(binary_reader<T> &reader, sector_id id)
{
    writer_->offset(sector_data_start() + sector_size() * id);
    writer_->append(reader, std::min(sector_size(), reader.bytes()) / sizeof(T));
}

template<typename T>
void compound_document::write_short_sector(binary_reader<T> &reader, sector_id id)
{
    const auto header_size = sizeof(compound_document_header);
    auto first_short_sector = entry_cache_.at(find_entry(u"Root Entry"))->start;

    auto short_position = static_cast<std::size_t>(short_sector_size() * id);
    auto current_sector_index = 0;
    auto current_sector = first_short_sector;

    while (current_sector_index < short_position / sector_size())
    {
        current_sector = sat(current_sector);
        ++current_sector_index;
    }

    auto offset = short_position % sector_size();
    writer_->offset(header_size + sector_size() * current_sector + offset);
    writer_->append(reader, short_sector_size());
}

template<typename T>
void compound_document::read_sector(sector_id current, binary_writer<T> &writer)
{
    reader_->offset(sector_data_start() + sector_size() * current);
    writer.append<byte>(*reader_, sector_size());
}

template<typename T>
void compound_document::read_short_sector(sector_id current_short, binary_writer<T> &writer)
{
    const auto header_size = sizeof(compound_document_header);
    auto sector_data = std::vector<byte>();
    auto sector_data_writer = binary_writer<byte>(sector_data);
    auto first_short_sector = entry_cache_.at(find_entry(u"Root Entry"))->start;

    auto short_position = static_cast<std::size_t>(short_sector_size() * current_short);
    auto current_sector_index = 0;
    auto current_sector = first_short_sector;

    while (current_sector_index < short_position / sector_size())
    {
        current_sector = sat(current_sector);
        ++current_sector_index;
    }

    auto offset = short_position % sector_size();
    reader_->offset(header_size + sector_size() * current_sector + offset);
    sector_data_writer.append(*reader_, short_sector_size());
}

sector_id compound_document::allocate_sector()
{
    auto msat_index = std::size_t(0);
    auto sat_index = sector_id(0);

    while (true)
    {
        if (header().msat[msat_index] == FreeSector)
        {
            // allocate msat sector
            header().msat[msat_index] = sat_index;

            auto sector_data = std::vector<sector_id>(sector_size() / sizeof(sector_id), FreeSector);
            sector_data.front() = SATSector;
            auto sector_data_reader = binary_reader<sector_id>(sector_data);
            write_sector(sector_data_reader, sat_index);
        }

        // load the current sector
        // could read one at a time, but this is easier for debugging for now
        sat_index = header().msat[msat_index];
        auto sat_sector_data = std::vector<sector_id>();
        auto sat_sector_data_writer = binary_writer<sector_id>(sat_sector_data);
        read_sector<sector_id>(sat_index, sat_sector_data_writer);

        auto next_free_iter = std::find(sat_sector_data.begin(), sat_sector_data.end(), FreeSector);

        if (next_free_iter != sat_sector_data.end()) // found a free sector
        {
            auto next_free = static_cast<sector_id>(next_free_iter - sat_sector_data.begin());

            // write empty sector
            auto sector_data = std::vector<sector_id>(sector_size() / sizeof(sector_id), 0);
            auto sector_data_reader = binary_reader<sector_id>(sector_data);
            write_sector(sector_data_reader, next_free);

            // update sat
            writer_->offset(sector_data_start() + sat_index * sector_size() + next_free * sizeof(sector_id));
            writer_->write(EndOfChain);

            return next_free;
        }

        ++msat_index;
    }
}

sector_chain compound_document::allocate_sectors(std::size_t count)
{
    auto chain = sector_chain();

    for (auto i = std::size_t(0); i < count; ++i)
    {
        chain.push_back(allocate_sector());
    }

    return chain;
}

sector_chain compound_document::follow_sat_chain(sector_id start)
{
    auto chain = sector_chain();
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    const auto msat_index = start / sectors_per_sector;
    auto sector = start;

    while (sector >= 0)
    {
        chain.push_back(sector);
        auto sat_index = header().msat[msat_index];
        reader_->offset(sector_data_start() + sat_index * sector_size() + (sector * sizeof(sector_id) % sector_size()));
        sector = reader_->read<sector_id>();
    }

    return chain;
}

sector_chain compound_document::follow_ssat_chain(sector_id start)
{
    auto chain = sector_chain();
    return chain;
}

sector_id compound_document::msat(sector_id id)
{
    if (id < sector_id(109))
    {
        return header().msat.at(id);
    }

    auto current = msat(108);

    while (current < id)
    {
        reader_->offset(sector_data_start() + current * sector_size() - sizeof(sector_id));
        current = reader_->read<sector_id>();
    }

    return current;
}

sector_id compound_document::sat(sector_id id)
{
    auto sectors_per_sector = sector_size() / sizeof(sector_id);
    auto msat_sector_index = msat(sector_id(id / sectors_per_sector));

    reader_->offset(sector_data_start() + msat_sector_index * sector_size() + id % sectors_per_sector);

    return reader_->read<sector_id>();
}

sector_chain compound_document::allocate_short_sectors(std::size_t count)
{
    return { EndOfChain };
}

directory_id compound_document::next_empty_entry()
{
    auto entry_id = directory_id(0);

    for (auto entry : entry_cache_)
    {
        if (entry.second->type == compound_document_entry::entry_type::Empty)
        {
            return entry.first;
        }

        entry_id = std::max(entry.first, entry_id);
    }

    const auto entries_per_sector = sector_size() 
        / sizeof(compound_document_entry);
    auto new_sector = allocate_sector();
    // TODO: connect chains here
    writer_->offset(sector_data_start() + new_sector * sector_size());
    reader_->offset(sector_data_start() + new_sector * sector_size());

    for (auto i = std::size_t(0); i < entries_per_sector; ++i)
    {
        auto empty_entry = compound_document_entry();
        empty_entry.type = compound_document_entry::entry_type::Empty;

        writer_->write(empty_entry);

        entry_cache_[entry_id + i] = const_cast<compound_document_entry *>(
            reader_->read_pointer<compound_document_entry>());
    }

    return entry_id;
}

directory_id compound_document::insert_entry(
    const std::u16string &name,
    compound_document_entry::entry_type type)
{
    auto entry_id = next_empty_entry();
    auto entry = entry_cache_.at(entry_id);

    entry->name(name);
    entry->type = type;

    // TODO: parse path from name and use correct parent storage instead of 0
    tree_insert(entry_id, 0);

    return entry_id;
}

std::size_t compound_document::sector_data_start()
{
    return sizeof(compound_document_header);
}

bool compound_document::contains_entry(const std::u16string &path)
{
    return find_entry(path) >= 0;
}

compound_document_header &compound_document::header()
{
    reader_->offset(0);

    if (reader_->bytes() < sizeof(compound_document_header))
    {
        if (writer_ != nullptr)
        {
            writer_->resize(sizeof(compound_document_header), 0);
            writer_->write(compound_document_header());
        }
        else
        {
            throw xlnt::exception("missing header");
        }
    }

    return const_cast<compound_document_header &>(
        reader_->read_reference<compound_document_header>());
}

directory_id compound_document::find_entry(const std::u16string &name)
{
    for (auto entry : entry_cache_)
    {
        if (entry.second->type == compound_document_entry::entry_type::Empty) continue;
        if (tree_path(entry.first) == name) return entry.first;
    }

    return End;
}

void compound_document::print_directory()
{
    for (auto entry : entry_cache_)
    {
        if (entry.second->type != compound_document_entry::entry_type::UserStream) continue;
        std::cout << utf16_to_utf8(tree_path(entry.first)) << std::endl;
    }
}

void compound_document::tree_initialize_parent_maps()
{
    const auto entries_per_sector = static_cast<directory_id>(sector_size()
        / sizeof(compound_document_entry));
    auto entry_id = directory_id(0);

    for (auto sector : follow_sat_chain(header().directory_start))
    {
        reader_->offset(sector_data_start() + sector * sector_size());

        for (auto i = std::size_t(0); i < entries_per_sector; ++i)
        {
            entry_cache_[entry_id++] = const_cast<compound_document_entry *>(
                &reader_->read_reference<compound_document_entry>());
        }
    }

    auto stack = std::vector<directory_id>();
    auto storage_siblings = std::vector<directory_id>();
    auto stream_siblings = std::vector<directory_id>();

    auto directory_stack = std::vector<directory_id>();
    directory_stack.push_back(directory_id(0));

    while (!directory_stack.empty())
    {
        auto current_storage_id = directory_stack.back();
        directory_stack.pop_back();

        if (tree_child(current_storage_id) < 0) continue;

        auto storage_stack = std::vector<directory_id>();
        auto storage_root_id = tree_child(current_storage_id);
        parent_[storage_root_id] = End;
        storage_stack.push_back(storage_root_id);

        while (!storage_stack.empty())
        {
            auto current_entry_id = storage_stack.back();
            auto current_entry = entry_cache_.at(current_entry_id);
            storage_stack.pop_back();

            parent_storage_[current_entry_id] = current_storage_id;

            if (current_entry->type == compound_document_entry::entry_type::UserStorage)
            {
                directory_stack.push_back(current_entry_id);
            }

            if (current_entry->prev >= 0)
            {
                storage_stack.push_back(current_entry->prev);
                tree_parent(tree_left(current_entry_id)) = current_entry_id;
            }

            if (current_entry->next >= 0)
            {
                storage_stack.push_back(tree_right(current_entry_id));
                tree_parent(tree_right(current_entry_id)) = current_entry_id;
            }
        }
    }
}

void compound_document::tree_insert(directory_id new_id, directory_id storage_id)
{
    using entry_color = compound_document_entry::entry_color;

    parent_storage_[new_id] = storage_id;

    tree_left(new_id) = End;
    tree_right(new_id) = End;

    if (tree_root(new_id) == End)
    {
        if (new_id != 0)
        {
            tree_root(new_id) = new_id;
        }

        tree_color(new_id) = entry_color::Black;
        tree_parent(new_id) = End;

        return;
    }

    // normal tree insert
    // (will probably unbalance the tree, fix after)
    auto x = tree_root(new_id);
    auto y = End;

    while (x >= 0)
    {
        y = x;

        if (compare_keys(tree_key(x), tree_key(new_id)) > 0)
        {
            x = tree_right(x);
        }
        else
        {
            x = tree_left(x);
        }
    }

    tree_parent(new_id) = y;

    if (compare_keys(tree_key(y), tree_key(new_id)) > 0)
    {
        tree_right(y) = new_id;
    }
    else
    {
        tree_left(y) = new_id;
    }

    tree_insert_fixup(new_id);
}

std::u16string compound_document::tree_path(directory_id id)
{
    auto storage_id = parent_storage_[id];
    auto result = std::u16string();

    while (storage_id > 0)
    {
        storage_id = parent_storage_[storage_id];
        result = tree_key(storage_id) + u"/" + result;
    }

    return result;
}

void compound_document::tree_rotate_left(directory_id x)
{
    auto y = tree_right(x);

    // turn y's left subtree into x's right subtree
    tree_right(x) = tree_left(y);

    if (tree_left(y) != End)
    {
        tree_parent(tree_left(y)) = x;
    }

    // link x's parent to y
    tree_parent(y) = tree_parent(x);

    if (tree_parent(x) == End)
    {
        tree_root(x) = y;
    }
    else if (x == tree_left(tree_parent(x)))
    {
        tree_left(tree_parent(x)) = y;
    }
    else
    {
        tree_right(tree_parent(x)) = y;
    }

    // put x on y's left
    tree_left(y) = x;
    tree_parent(x) = y;
}

void compound_document::tree_rotate_right(directory_id y)
{
    auto x = tree_left(y);

    // turn x's right subtree into y's left subtree
    tree_left(y) = tree_right(x);

    if (tree_right(x) != End)
    {
        tree_parent(tree_right(x)) = y;
    }

    // link y's parent to x
    tree_parent(x) = tree_parent(y);

    if (tree_parent(y) == End)
    {
        tree_root(y) = x;
    }
    else if (y == tree_left(tree_parent(y)))
    {
        tree_left(tree_parent(y)) = x;
    }
    else
    {
        tree_right(tree_parent(y)) = x;
    }

    // put y on x's right
    tree_right(x) = y;
    tree_parent(y) = x;
}

void compound_document::tree_insert_fixup(directory_id x)
{
    using entry_color = compound_document_entry::entry_color;

    tree_color(x) = entry_color::Red;

    while (x != tree_root(x) && tree_color(tree_parent(x)) == entry_color::Red)
    {
        if (tree_parent(x) == tree_left(tree_parent(tree_parent(x))))
        {
            auto y = tree_right(tree_parent(tree_parent(x)));

            if (tree_color(y) == entry_color::Red)
            {
                // case 1
                tree_color(tree_parent(x)) = entry_color::Black;
                tree_color(y) = entry_color::Black;
                tree_color(tree_parent(tree_parent(x))) = entry_color::Red;
                x = tree_parent(tree_parent(x));
            }
            else
            {
                if (x == tree_right(tree_parent(x)))
                {
                    // case 2
                    x = tree_parent(x);
                    tree_rotate_left(x);
                }

                // case 3
                tree_color(tree_parent(x)) = entry_color::Black;
                tree_color(tree_parent(tree_parent(x))) = entry_color::Red;
                tree_rotate_right(tree_parent(tree_parent(x)));
            }
        }
        else // same as above with left and right switched
        {
            auto y = tree_left(tree_parent(tree_parent(x)));

            if (tree_color(y) == entry_color::Red)
            {
                //case 1
                tree_color(tree_parent(x)) = entry_color::Black;
                tree_color(y) = entry_color::Black;
                tree_color(tree_parent(tree_parent(x))) = entry_color::Red;
                x = tree_parent(tree_parent(x));
            }
            else
            {
                if (x == tree_left(tree_parent(x)))
                {
                    // case 2
                    x = tree_parent(x);
                    tree_rotate_right(x);
                }

                // case 3
                tree_color(tree_parent(x)) = entry_color::Black;
                tree_color(tree_parent(tree_parent(x))) = entry_color::Red;
                tree_rotate_left(tree_parent(tree_parent(x)));
            }
        }
    }

    tree_color(tree_root(x)) = entry_color::Black;
}

directory_id &compound_document::tree_left(directory_id id)
{
    return entry_cache_[id]->prev;
}

directory_id &compound_document::tree_right(directory_id id)
{
    return entry_cache_[id]->next;
}

directory_id &compound_document::tree_parent(directory_id id)
{
    return parent_[id];
}

directory_id &compound_document::tree_root(directory_id id)
{
    return tree_child(parent_storage_[id]);
}

directory_id &compound_document::tree_child(directory_id id)
{
    return entry_cache_[id]->child;
}

std::u16string compound_document::tree_key(directory_id id)
{
    return entry_cache_[id]->name();
}

compound_document_entry::entry_color &compound_document::tree_color(directory_id id)
{
    return entry_cache_[id]->color;
}

} // namespace detail
} // namespace xlnt
