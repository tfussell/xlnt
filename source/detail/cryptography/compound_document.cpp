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
#include <locale>
#include <string>
#include <vector>

#include <detail/binary.hpp>
#include <detail/unicode.hpp>
#include <detail/cryptography/compound_document.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

using namespace xlnt::detail;

int compare_keys(const std::string &left, const std::string &right)
{
    auto to_lower = [](std::string s)
    {
        static const auto locale = std::locale();
        std::use_facet<std::ctype<char>>(locale).tolower(&s[0], &s[0] + s.size());
        return s;
    };

    return to_lower(left).compare(to_lower(right));
}

std::string join_path(const std::vector<std::string> &path)
{
    auto joined = std::string();

    for (auto part : path)
    {
        joined.append(part);
        joined.push_back('/');
    }

    return joined;
}

const sector_id FreeSector = -1;
const sector_id EndOfChain = -2;
const sector_id SATSector = -3;
//const sector_id MSATSector = -4;

const directory_id End = -1;

} // namespace

namespace xlnt {
namespace detail {

compound_document::compound_document(std::vector<std::uint8_t> &data)
    : writer_(new binary_writer<byte>(data)),
      reader_(new binary_reader<byte>(data))
{
    header_.msat.fill(FreeSector);
    writer_->offset(0);
    writer_->write(header_);

    header_.directory_start = allocate_sector();
    writer_->offset(0);
    writer_->write(header_);

    insert_entry("Root Entry", compound_document_entry::entry_type::RootStorage);
}

compound_document::compound_document(const std::vector<std::uint8_t> &data)
    : writer_(nullptr),
      reader_(new binary_reader<byte>(data))
{
    header_ = reader_->read<compound_document_header>();

    // read msat
    auto current = header_.extra_msat_start;

    for (auto i = std::size_t(0); i < header_.num_msat_sectors; ++i)
    {
        if (i < 109)
        {
            msat_.push_back(header_.msat[i]);
        }
        else
        {
            auto extra_msat_sector = std::vector<sector_id>();
            auto extra_msat_sector_writer = binary_writer<sector_id>(extra_msat_sector);

            read_sector(current, extra_msat_sector_writer);
            
            std::copy(extra_msat_sector.begin(),
                extra_msat_sector.end() - 1,
                std::back_inserter(msat_));
            current = extra_msat_sector.back();
        }
    }
    
    for (auto sat_sector_id : msat_)
    {
        if (sat_sector_id < 0) continue;
        
        auto sat_sector = std::vector<sector_id>();
        auto sat_sector_writer = binary_writer<sector_id>(sat_sector);
        
        read_sector(sat_sector_id, sat_sector_writer);
        
        std::copy(sat_sector.begin(),
            sat_sector.end(),
            std::back_inserter(sat_));
    }

    for (auto ssat_sector_id : follow_chain(header_.ssat_start, sat_))
    {
        auto ssat_sector = std::vector<sector_id>();
        auto ssat_sector_writer = binary_writer<sector_id>(ssat_sector);
        
        read_sector(ssat_sector_id, ssat_sector_writer);
        
        std::copy(ssat_sector.begin(),
            ssat_sector.end(),
            std::back_inserter(ssat_));
    }

    tree_initialize_parent_maps();
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

std::vector<byte> compound_document::read_stream(const std::string &name)
{
    const auto entry_id = find_entry(name, compound_document_entry::entry_type::UserStream);
    const auto entry = entry_cache_.at(entry_id);

    auto stream_data = std::vector<byte>();
    auto stream_data_writer = binary_writer<byte>(stream_data);

    if (entry->size < header_.threshold)
    {
        for (auto sector : follow_chain(entry->start, ssat_))
        {
            read_short_sector<byte>(sector, stream_data_writer);
        }
    }
    else
    {
        for (auto sector : follow_chain(entry->start, sat_))
        {
            read_sector<byte>(sector, stream_data_writer);
        }
    }
    stream_data.resize(entry->size);

    return stream_data;
}

void compound_document::write_stream(const std::string &name, const std::vector<std::uint8_t> &data)
{
    auto entry_id = contains_entry(name, compound_document_entry::entry_type::UserStream)
        ? find_entry(name, compound_document_entry::entry_type::UserStream)
        : insert_entry(name, compound_document_entry::entry_type::UserStream);
    auto entry = entry_cache_.at(entry_id);
    entry->size = static_cast<std::uint32_t>(data.size());

    auto stream_data_reader = binary_reader<byte>(data);

    if (entry->size < header_.threshold)
    {
        const auto num_sectors = data.size() / short_sector_size()
            + (data.size() % short_sector_size() ? 1 : 0);

        auto chain = allocate_short_sectors(num_sectors);
        entry->start = chain.front();

        for (auto sector : follow_chain(entry->start, ssat_))
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

        for (auto sector : follow_chain(entry->start, sat_))
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

}

template<typename T>
void compound_document::read_sector(sector_id id, binary_writer<T> &writer)
{
    reader_->offset(sector_data_start() + sector_size() * id);
    writer.append(*reader_, sector_size());
}

template<typename T>
void compound_document::read_short_sector(sector_id id, binary_writer<T> &writer)
{
    const auto container_chain = follow_chain(entry_cache_[0]->start, sat_);
    auto container = std::vector<byte>();
    auto container_writer = binary_writer<byte>(container);
    
    for (auto sector : container_chain)
    {
        read_sector(sector, container_writer);
    }

    auto container_reader = binary_reader<byte>(container);
    container_reader.offset(id * short_sector_size());

    writer.append(container_reader, short_sector_size());
}

sector_id compound_document::allocate_sector()
{
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    auto next_free_iter = std::find(sat_.begin(), sat_.end(), FreeSector);
    
    if (next_free_iter == sat_.end())
    {
        auto next_msat_index = header_.num_msat_sectors;
        auto new_sat_sector_id = sector_id(sat_.size());
        msat_.push_back(new_sat_sector_id);
        header_.msat[msat_.size() - 1] = new_sat_sector_id;
        sat_.resize(sat_.size() + sectors_per_sector, FreeSector);
        sat_[new_sat_sector_id] = SATSector;
        auto sat_reader = binary_reader<sector_id>(sat_);
        sat_reader.offset(next_msat_index * sectors_per_sector);
        write_sector(sat_reader, new_sat_sector_id);
        next_free_iter = std::find(sat_.begin(), sat_.end(), FreeSector);
    }
    
    auto next_free = next_free_iter - sat_.begin();
    sat_[next_free] = EndOfChain;

    auto next_free_msat_index = next_free / sectors_per_sector;;
    auto sat_index = msat_[next_free_msat_index];
    writer_->offset(sector_data_start()
        + (sat_index * sector_size())
        + (next_free % sectors_per_sector) * sizeof(sector_id));
    writer_->write(EndOfChain);
    
    auto empty_sector = std::vector<byte>(sector_size());
    auto empty_sector_reader = binary_reader<byte>(empty_sector);
    write_sector(empty_sector_reader, next_free);
    
    return next_free;
}

sector_chain compound_document::allocate_sectors(std::size_t count)
{
    if (count == std::size_t(0)) return {};

    auto chain = sector_chain();
    auto current = allocate_sector();

    for (auto i = std::size_t(1); i < count; ++i)
    {
        chain.push_back(current);
        auto next = allocate_sector();
        sat_[current] = next;
        current = next;
    }

    return chain;
}

sector_chain compound_document::follow_chain(sector_id start, const sector_chain &table)
{
    auto chain = sector_chain();
    auto current = start;

    while (current >= 0)
    {
        chain.push_back(current);
        current = table[current];
    }

    return chain;
}

sector_chain compound_document::allocate_short_sectors(std::size_t count)
{
    if (count == std::size_t(0)) return {};

    auto chain = sector_chain();
    auto current = allocate_short_sector();

    for (auto i = std::size_t(1); i < count; ++i)
    {
        chain.push_back(current);
        auto next = allocate_short_sector();
        ssat_[current] = next;
        current = next;
    }
    
    chain.push_back(current);

    return chain;
}

sector_id compound_document::allocate_short_sector()
{
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    auto next_free_iter = std::find(ssat_.begin(), ssat_.end(), FreeSector);
    
    if (next_free_iter == ssat_.end())
    {
        auto new_ssat_sector_id = allocate_sector();
        
        ++header_.num_short_sectors;
        
        if (header_.ssat_start < 0)
        {
            header_.ssat_start = new_ssat_sector_id;
        }
        
        writer_->offset(0);
        writer_->write(header_);
        
        auto old_size = ssat_.size();
        ssat_.resize(old_size + sectors_per_sector, FreeSector);

        auto ssat_reader = binary_reader<sector_id>(ssat_);
        ssat_reader.offset(old_size / sectors_per_sector);
        write_sector(ssat_reader, new_ssat_sector_id);

        next_free_iter = std::find(ssat_.begin(), ssat_.end(), FreeSector);
    }
    
    auto next_free = next_free_iter - ssat_.begin();
    ssat_[next_free] = EndOfChain;
    
    auto sat_chain = follow_chain(header_.ssat_start, sat_);
    auto next_free_sat_chain_index = next_free / sectors_per_sector;
    auto sat_index = sat_chain[next_free_sat_chain_index];
    writer_->offset(sector_data_start()
        + (sat_index * sectors_per_sector) * sizeof(sector_id)
        + (next_free % sectors_per_sector) * sizeof(sector_id));
    writer_->write(EndOfChain);
    
    const auto short_sectors_per_sector = sector_size() / short_sector_size();
    const auto required_container_sectors = std::size_t(next_free / short_sectors_per_sector + 1);
    
    if (required_container_sectors > 0)
    {
        if (entry_cache_[0]->start < 0)
        {
            entry_cache_[0]->start = allocate_sector();
        }
        
        auto container_chain = follow_chain(entry_cache_[0]->start, sat_);
        
        if (required_container_sectors > container_chain.size())
        {
            sat_[container_chain.back()] = allocate_sector();
        }
    }
    
    return next_free;
}

directory_id compound_document::next_empty_entry()
{
    auto entry_id = directory_id(0);

    for (; entry_id < directory_id(entry_cache_.size()); ++entry_id)
    {
        auto entry = entry_cache_[entry_id];

        if (entry->type == compound_document_entry::entry_type::Empty)
        {
            return entry_id;
        }
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
    const std::string &name,
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

bool compound_document::contains_entry(const std::string &path,
        compound_document_entry::entry_type type)
{
    return find_entry(path, type) >= 0;
}

directory_id compound_document::find_entry(const std::string &name,
    compound_document_entry::entry_type type)
{
    if (type == compound_document_entry::entry_type::RootStorage
        && (name == "/" || name == "/Root Entry")) return 0;

    for (auto entry : entry_cache_)
    {
        if (entry.second->type == compound_document_entry::entry_type::Empty) continue;
        if (entry.second->type != type) continue;
        if (tree_path(entry.first) == name) return entry.first;
    }

    return End;
}

void compound_document::print_directory()
{
    for (auto entry : entry_cache_)
    {
        if (entry.second->type != compound_document_entry::entry_type::UserStream) continue;
        std::cout << tree_path(entry.first) << std::endl;
    }
}

void compound_document::tree_initialize_parent_maps()
{
    const auto entries_per_sector = sector_size() / sizeof(compound_document_entry);
    auto entry_id = directory_id(0);

    for (auto sector : follow_chain(header_.directory_start, sat_))
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

        if (compare_keys(tree_key(new_id), tree_key(x)) > 0)
        {
            x = tree_right(x);
        }
        else
        {
            x = tree_left(x);
        }
    }

    tree_parent(new_id) = y;

    if (compare_keys(tree_key(new_id), tree_key(y)) > 0)
    {
        tree_right(y) = new_id;
    }
    else
    {
        tree_left(y) = new_id;
    }

    tree_insert_fixup(new_id);
}

std::string compound_document::tree_path(directory_id id)
{
    auto storage_id = parent_storage_[id];
    auto result = std::vector<std::string>();

    while (storage_id > 0)
    {
        storage_id = parent_storage_[storage_id];
        result.push_back(entry_cache_[storage_id]->name());
    }

    return "/" + join_path(result) + entry_cache_[id]->name();
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

            if (y >= 0 && tree_color(y) == entry_color::Red)
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

            if (y >= 0 && tree_color(y) == entry_color::Red)
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

std::string compound_document::tree_key(directory_id id)
{
    return entry_cache_[id]->name();
}

compound_document_entry::entry_color &compound_document::tree_color(directory_id id)
{
    return entry_cache_[id]->color;
}

} // namespace detail
} // namespace xlnt
