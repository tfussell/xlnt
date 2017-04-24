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

using xlnt::detail::byte;
using xlnt::detail::binary_reader;
using xlnt::detail::binary_writer;

using directory_id = std::int32_t;
using sector_id = std::int32_t;
using sector_chain = std::vector<sector_id>;

const sector_id FreeSector = -1;

sector_chain follow_sector_chain(const sector_chain &table, sector_id start)
{
    auto chain = sector_chain();
    auto added = std::unordered_set<sector_id>();
    auto last_sector = static_cast<sector_id>(table.size());

    if (start >= last_sector)
    {
        return chain;
    }

    auto current = start;

    while (current < last_sector && current >= 0)
    {
        if (added.find(current) != added.end())
        {
            break;
        }

        chain.push_back(current);
        added.insert(current); //TODO: why would there be a repeat?

        current = table[current];
    }

    return chain;
}

struct header
{
    enum class byte_order_type : uint16_t
    {
        big_endian = 0xFFFE,
        little_endian = 0xFEFF
    };

    std::uint64_t file_id = 0xe11ab1a1e011cfd0;
    std::array<std::uint8_t, 16> ignore1 = {{0}};
    std::uint16_t revision = 0x003E;
    std::uint16_t version = 0x0003;
    byte_order_type byte_order = byte_order_type::little_endian;
    std::uint16_t sector_size_power = 9;
    std::uint16_t short_sector_size_power = 6;
    std::array<std::uint8_t, 10> ignore2 = {{0}};
    std::uint32_t num_sectors = 0;
    sector_id directory_start = 0;
    std::array<std::uint8_t, 4> ignore3 = {{0}};
    std::uint32_t threshold = 4096;
    sector_id short_table_start = 0;
    std::uint32_t num_short_sectors = 0;
    sector_id sector_table_start = 0;
    std::uint32_t num_master_alloc_table_sectors = 0;
    std::array<sector_id, 109> master_sector_alloc_table = {{FreeSector}};
};

bool header_is_valid(const header &h)
{
    if (h.threshold != 4096)
    {
        return false;
    }

    if (h.num_sectors == 0
        || (h.num_sectors > 109 && h.num_sectors > (h.num_master_alloc_table_sectors * 127) + 109)
        || ((h.num_sectors < 109) && (h.num_master_alloc_table_sectors != 0)))
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

struct directory_entry
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

    std::array<char16_t, 32> name_array = {{0}};
    std::uint16_t name_length = 0;
    entry_type type;
    entry_color color;
    directory_id prev = -1;
    directory_id next = -1;
    directory_id child = -1;
    std::array<std::uint8_t, 36> ignore;
    sector_id first = 0;
    std::uint32_t size = 0;
    std::uint32_t ignore2;
};

class directory_tree
{
public:
    static const directory_id End;

    directory_tree()
    {
        clear();
    }

    void clear()
    {
        entries = { create_root_entry() };
    }

    std::size_t entry_count() const
    {
        return entries.size();
    }

    directory_entry &entry(directory_id index)
    {
        return entries[static_cast<std::size_t>(index)];
    }

    const directory_entry &entry(directory_id index) const
    {
        return entries[static_cast<std::size_t>(index)];
    }

    const directory_entry &entry(const std::u16string &name) const
    {
        return entry(find_entry(name).first);
    }

    directory_entry &entry(const std::u16string &name, bool create)
    {
        auto find_result = find_entry(name);
        auto index = find_result.first;
        auto found = find_result.second;

        if (!found)
        {
            // not found among children
            if (!create)
            {
                throw xlnt::exception("not found");
            }

            // create a new entry
            auto parent = index;
            entries.push_back(directory_entry());
            index = static_cast<directory_id>(entry_count() - 1);
            auto &e = entry(index);
            e.name(name);
            e.type = directory_entry::entry_type::UserStream;
            e.size = 0;
            e.first = 0;
            e.child = End;
            e.prev = End;
            e.next = entry(parent).child;
            entry(parent).child = index;
        }

        return entry(index);
    }
/*
    directory_id parent(directory_id index)
    {
        // brute-force, basically we iterate for each entries, find its children
        // and check if one of the children is 'index'
        for (auto j = directory_id(0); j < static_cast<directory_id>(entry_count()); j++)
        {
            auto chi = children(j);

            for (std::size_t i = 0; i < chi.size(); i++)
            {
                if (chi[i] == index)
                {
                    return j;
                }
            }
        }

        return -1;
    }
*/
/*
    std::u16string path(directory_id index)
    {
        // don't use root name ("Root Entry"), just give "/"
        if (index == 0) return u"/";

        auto current_entry = entry(index);

        auto result = std::u16string(entry(index).name.data());
        result.insert(0, u"/");

        auto current_parent = parent(index);

        while (current_parent > 0)
        {
            current_entry = entry(current_parent);

            result.insert(0, std::u16string(current_entry.name.data()));
            result.insert(0, u"/");

            --current_parent;
            index = current_parent;

            if (current_parent <= 0) break;
        }

        return result;
    }
*/
    std::vector<directory_id> children(directory_id index) const
    {
        auto result = std::vector<directory_id>();
        auto &e = entry(index);

        if (e.child >= 0 && e.child < static_cast<directory_id>(entry_count()))
        {
            find_siblings(result, e.child);
        }

        return result;
    }


    void load(const std::vector<byte> &data)
    {
        auto reader = binary_reader(data);
        entries = reader.as_vector_of<directory_entry>();
        
        auto is_empty = [](const directory_entry &entry)
        {
            return entry.type == directory_entry::entry_type::Empty;
        };

        entries.erase(std::remove_if(entries.begin(), entries.end(), is_empty));
    }

    directory_entry create_root_entry() const
    {
        directory_entry root;

        root.name(u"Root Entry");
        root.type = directory_entry::entry_type::RootStorage;
        root.color = directory_entry::entry_color::Black;
        root.size = 0;

        return root;
    }

private:
    // helper function: recursively find siblings of index
    void find_siblings(std::vector<directory_id> &result, directory_id index) const
    {
        auto e = entry(index);

        // prevent infinite loop
        for (std::size_t i = 0; i < result.size(); i++)
        {
            if (result[i] == index) return;
        }

        // add myself
        result.push_back(index);

        // visit previous sibling, don't go infinitely
        auto prev = e.prev;

        if ((prev > 0) && (prev < static_cast<directory_id>(entry_count())))
        {
            for (std::size_t i = 0; i < result.size(); i++)
            {
                if (result[i] == prev)
                {
                    prev = 0;
                }
            }

            if (prev)
            {
                find_siblings(result, prev);
            }
        }

        // visit next sibling, don't go infinitely
        auto next = e.next;

        if ((next > 0) && (next < static_cast<directory_id>(entry_count())))
        {
            for (std::size_t i = 0; i < result.size(); i++)
            {
                if (result[i] == next) next = 0;
            }

            if (next)
            {
                find_siblings(result, next);
            }
        }
    }

    std::pair<directory_id, bool> find_entry(const std::u16string &name) const
    {
        // quick check for "/" (that's root)
        if (name == u"/Root Entry")
        {
            return { 0, true };
        }

        // split the names, e.g  "/ObjectPool/_1020961869" will become:
        // "ObjectPool" and "_1020961869"
        auto names = std::vector<std::u16string>();
        auto start = std::size_t(0);
        auto end = std::size_t(0);

        if (name[0] == u'/') start++;

        while (start < name.length())
        {
            end = name.find_first_of('/', start);
            if (end == std::string::npos) end = name.length();
            names.push_back(name.substr(start, end - start));
            start = end + 1;
        }

        // start from root
        auto index = directory_id(0);

        for (auto &name : names)
        {
            // find among the children of index
            auto chi = children(index);
            std::ptrdiff_t child = 0;

            for (std::size_t i = 0; i < chi.size(); i++)
            {
                auto ce = entry(chi[i]);

                if (ce.name() == name)
                {
                    child = static_cast<std::ptrdiff_t>(chi[i]);
                }
            }

            // traverse to the child
            if (child > 0)
            {
                index = static_cast<directory_id>(child);
            }
            else
            {
                return { index, false };
            }
        }

        return { index, true };
    }

    std::vector<directory_entry> entries;
};

const directory_id directory_tree::End = -1;

} // namespace

namespace xlnt {
namespace detail {

class compound_document_reader_impl
{
public:
    compound_document_reader_impl(const std::vector<byte> &bytes)
        : sectors_(bytes.data() + sizeof(header)),
          sectors_size_(bytes.size())
    {
        auto reader = binary_reader(bytes);

        header_ = reader.read<header>();
        
        if(!header_is_valid(header_))
        {
            throw xlnt::exception("bad ole");
        }

        // Master allocation table
        const auto sector_table_sectors = load_master_sector_allocation_table();
        const auto sector_table_bytes = read(sector_table_sectors);
        auto sector_table_reader = binary_reader(sector_table_bytes);
        sector_table_ = sector_table_reader.as_vector_of<sector_id>();

        // Short sector allocation table
        const auto short_table_chain = follow_sector_chain(sector_table_, header_.short_table_start);
        const auto short_table_bytes = read(short_table_chain);
        auto short_sector_table_reader = binary_reader(short_table_bytes);
        short_sector_table_ = short_sector_table_reader.as_vector_of<sector_id>();

        // Directory
        const auto directory_chain = follow_sector_chain(sector_table_, header_.directory_start);
        const auto directory_sectors = read(directory_chain);
        directory_.load(directory_sectors);

        // Short stream container
        auto first_short_sector = directory_.entry(u"/Root Entry", false).first;
        short_container_stream_ = follow_sector_chain(sector_table_, first_short_sector);
    }

    std::vector<byte> read(const sector_chain &sectors) const
    {
        const auto sector_size = 1 << header_.sector_size_power;
        auto result = std::vector<byte>();
        auto writer = binary_writer(result);

        for (auto sector : sectors)
        {
            auto position = static_cast<std::size_t>(sector_size * sector);
            writer.append(sectors_, sectors_size_, position, sector_size);
        }
        
        return result;
    }

    std::vector<byte> read_short(const sector_chain &sectors) const
    {
        const auto short_sector_size = 1 << header_.short_sector_size_power;
        const auto sector_size = 1 << header_.sector_size_power;
        auto result = std::vector<byte>();
        auto writer = binary_writer(result);

        for (auto sector : sectors)
        {
            auto position = static_cast<std::size_t>(short_sector_size * sector);
            auto master_allocation_table_index = position / sector_size;

            auto sector_data = read({ short_container_stream_[master_allocation_table_index] });

            auto offset = position % sector_size;
            writer.append(sector_data, offset, short_sector_size);
        }

        return result;
    }

    sector_chain load_master_sector_allocation_table() const
    {
        auto sectors = sector_chain(
            header_.master_sector_alloc_table.begin(),
            header_.master_sector_alloc_table.begin()
                + std::min(header_.master_sector_alloc_table.size(),
                    static_cast<std::size_t>(header_.num_sectors)));

        if (header_.num_sectors > std::uint32_t(109))
        {
            auto current_sector = header_.sector_table_start;

            for (auto r = std::uint32_t(0); r < header_.num_master_alloc_table_sectors; ++r)
            {
                auto current_sector_data = read({ current_sector });
                auto current_sector_reader = binary_reader(current_sector_data);
                auto current_sector_sectors = current_sector_reader.as_vector_of<sector_id>();
                
                current_sector = current_sector_sectors.back();
                current_sector_sectors.pop_back();
                
                sectors.insert(
                    current_sector_sectors.begin(),
                    current_sector_sectors.end(),
                    sectors.end());
            }
        }

        return sectors;
    }

    std::vector<byte> read_stream(const std::u16string &name) const
    {
        const auto entry = directory_.entry(name);

        const auto entry_sectors = entry.size < header_.threshold
            ? follow_sector_chain(short_sector_table_, entry.first)
            : follow_sector_chain(sector_table_, entry.first);
        auto result = entry.size < header_.threshold
            ? read_short(entry_sectors)
            : read(entry_sectors);
        result.resize(entry.size);

        return result;
    }

private:
    const byte *sectors_;
    const std::size_t sectors_size_;
    directory_tree directory_;
    header header_;
    std::vector<sector_id> sector_table_;
    std::vector<sector_id> short_sector_table_;
    std::vector<sector_id> short_container_stream_;
};

class compound_document_writer_impl
{
public:
    compound_document_writer_impl(std::vector<byte> &bytes)
        : writer_(bytes),
          sector_table_(128, FreeSector),
          short_sector_table_(128, FreeSector)
    {
    }

    void write_sectors(const std::vector<byte> &data, directory_entry &/*entry*/)
    {
        const auto sector_size = 1 << header_.sector_size_power;
        const auto num_sectors = data.size() / sector_size;

        for (auto i = std::size_t(0); i < num_sectors; ++i)
        {
            auto position = sector_size * i;
            auto current_sector_size = data.size() % sector_size;
            writer_.append(data, position, current_sector_size);
        }
    }

    void write_short_sectors(const std::vector<byte> &data, directory_entry &/*entry*/)
    {
        const auto sector_size = 1 << header_.sector_size_power;
        const auto num_sectors = data.size() / sector_size;

        for (auto i = std::size_t(0); i < num_sectors; ++i)
        {
            auto position = sector_size * i;
            auto current_sector_size = data.size() % sector_size;
            writer_.append(data, position, current_sector_size);
        }
    }

    void write_stream(const std::u16string &name, const std::vector<byte> &data)
    {
        auto &entry = directory_.entry(name, true);
        
        if (entry.size < header_.threshold)
        {
            write_short_sectors(data, entry);
        }
        else
        {
            write_sectors(data, entry);
        }
    }

private:
    binary_writer writer_;
    directory_tree directory_;
    header header_;
    std::vector<sector_id> sector_table_;
    std::vector<sector_id> short_sector_table_;
    std::vector<sector_id> short_container_stream_;
};

compound_document_reader::compound_document_reader(const std::vector<std::uint8_t> &data)
    : d_(new compound_document_reader_impl(data))
{
}

compound_document_reader::~compound_document_reader()
{
}

std::vector<std::uint8_t> compound_document_reader::read_stream(const std::u16string &name) const
{
    return d_->read_stream(name);
}

compound_document_writer::compound_document_writer(std::vector<std::uint8_t> &data)
    : d_(new compound_document_writer_impl(data))
{
}

compound_document_writer::~compound_document_writer()
{
}

void compound_document_writer::write_stream(const std::u16string &name, const std::vector<std::uint8_t> &data)
{
    d_->write_stream(name, data);
}


} // namespace detail
} // namespace xlnt
