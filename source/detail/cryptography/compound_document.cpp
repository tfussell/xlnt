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

const sector_id FreeSector = -1;
const sector_id EndOfChain = -2;
const sector_id SATSector = -3;
const sector_id MSATSector = -4;

const directory_id End = -1;

} // namespace

namespace xlnt {
namespace detail {

class red_black_tree
{
public:
    using entry_color = compound_document_entry::entry_color;

    red_black_tree(std::vector<compound_document_entry> &entries)
        : entries_(entries)
    {
        initialize_tree();
    }

    void initialize_tree()
    {
        if (entries_.empty()) return;

        auto stack = std::vector<directory_id>();
        auto storage_siblings = std::vector<directory_id>();
        auto stream_siblings = std::vector<directory_id>();

        auto directory_stack = std::vector<directory_id>();
        directory_stack.push_back(directory_id(0));

        while (!directory_stack.empty())
        {
            auto current_storage_id = directory_stack.back();
            directory_stack.pop_back();

            if (entries_[current_storage_id].child < 0) continue;

            auto storage_stack = std::vector<directory_id>();
            auto storage_root_id = entries_[current_storage_id].child;
            parent_[storage_root_id] = End;
            storage_stack.push_back(storage_root_id);

            while (!storage_stack.empty())
            {
                auto current_entry_id = storage_stack.back();
                auto &current_entry = entries_[current_entry_id];
                storage_stack.pop_back();

                parent_storage_[current_entry_id] = current_storage_id;

                if (current_entry.type == compound_document_entry::entry_type::UserStorage)
                {
                    directory_stack.push_back(current_entry_id);
                }

                if (current_entry.prev >= 0)
                {
                    storage_stack.push_back(current_entry.prev);
                    parent_[current_entry.prev] = current_entry_id;
                }

                if (current_entry.next >= 0)
                {
                    storage_stack.push_back(entries_[current_entry_id].next);
                    parent_[entries_[current_entry_id].next] = current_entry_id;
                }
            }
        }
    }

    void insert(directory_id new_id, directory_id storage_id)
    {
        parent_storage_[new_id] = storage_id;

        left(new_id) = End;
        right(new_id) = End;

        if (root(new_id) == End)
        {
            if (new_id != 0)
            {
                root(new_id) = new_id;
            }

            color(new_id, entry_color::Black);
            parent(new_id) = End;

            return;
        }

        // normal tree insert
        // (will probably unbalance the tree, fix after)
        auto x = root(new_id);
        auto y = End;

        while (x >= 0)
        {
            y = x;

            if (compare_keys(key(x), key(new_id)) > 0)
            {
                x = right(x);
            }
            else
            {
                x = left(x);
            }
        }

        parent(new_id) = y;

        if (compare_keys(key(y), key(new_id)) > 0)
        {
            right(y) = new_id;
        }
        else
        {
            left(y) = new_id;
        }

        insert_fixup(new_id);
    }

    std::vector<directory_id> path(directory_id id)
    {
        auto storage_id = parent_storage_[id];
        auto storages = std::vector<directory_id>();
        storages.push_back(storage_id);
        
        while (storage_id > 0)
        {
            storage_id = parent_storage_[storage_id];
            storages.push_back(storage_id);
        }

        return storages;
    }

private:
    void rotate_left(directory_id x)
    {
        auto y = right(x);

        // turn y's left subtree into x's right subtree
        right(x) = left(y);

        if (left(y) != End)
        {
            parent(left(y)) = x;
        }

        // link x's parent to y
        parent(y) = parent(x);

        if (parent(x) == End)
        {
            root(x) = y;
        }
        else if (x == left(parent(x)))
        {
            left(parent(x)) = y;
        }
        else
        {
            right(parent(x)) = y;
        }

        // put x on y's left
        left(y) = x;
        parent(x) = y;
    }

    void rotate_right(directory_id y)
    {
        auto x = left(y);

        // turn x's right subtree into y's left subtree
        left(y) = right(x);

        if (right(x) != End)
        {
            parent(right(x)) = y;
        }

        // link y's parent to x
        parent(x) = parent(y);

        if (parent(y) == End)
        {
            root(y) = x;
        }
        else if (y == left(parent(y)))
        {
            left(parent(y)) = x;
        }
        else
        {
            right(parent(y)) = x;
        }

        // put y on x's right
        right(x) = y;
        parent(y) = x;
    }

    void insert_fixup(directory_id x)
    {
        color(x, entry_color::Red);

        while (x != root(x) && color(parent(x)) == entry_color::Red)
        {
            if (parent(x) == left(parent(parent(x))))
            {
                auto y = right(parent(parent(x)));

                if (color(y) == entry_color::Red)
                {
                    // case 1
                    color(parent(x), entry_color::Black);
                    color(y, entry_color::Black);
                    color(parent(parent(x)), entry_color::Red);
                    x = parent(parent(x));
                }
                else
                {
                    if (x == right(parent(x)))
                    {
                        // case 2
                        x = parent(x);
                        rotate_left(x);
                    }

                    // case 3
                    color(parent(x), entry_color::Black);
                    color(parent(parent(x)), entry_color::Red);
                    rotate_right(parent(parent(x)));
                }
            }
            else // same as above with left and right switched
            {
                auto y = left(parent(parent(x)));

                if (color(y) == entry_color::Red)
                {
                    //case 1
                    color(parent(x), entry_color::Black);
                    color(y, entry_color::Black);
                    color(parent(parent(x)), entry_color::Red);
                    x = parent(parent(x));
                }
                else
                {
                    if (x == left(parent(x)))
                    {
                        // case 2
                        x = parent(x);
                        rotate_right(x);
                    }

                    // case 3
                    color(parent(x), entry_color::Black);
                    color(parent(parent(x)), entry_color::Red);
                    rotate_left(parent(parent(x)));
                }
            }
        }

        color(root(x), entry_color::Black);
    }

    void color(directory_id id, entry_color c)
    {
        entries_.at(id).color = c;
    }

    entry_color color(directory_id id)
    {
        return entries_.at(id).color;
    }

    directory_id &parent(directory_id id)
    {
        if (parent_.find(id) == parent_.end())
        {
            parent_[id] = End;
        }

        return parent_[id];
    }

    directory_id &parent_storage(directory_id id)
    {
        if (parent_storage_.find(id) == parent_storage_.end())
        {
            throw xlnt::exception("not found");
        }

        return parent_storage_[id];
    }

    directory_id &root(directory_id id)
    {
        return entries_[parent_storage(id)].child;
    }

    directory_id &left(directory_id id)
    {
        return entries_.at(id).prev;
    }

    directory_id &right(directory_id id)
    {
        return entries_.at(id).next;
    }

    std::u16string key(directory_id id)
    {
        return entries_.at(id).name();
    }

private:
    std::vector<compound_document_entry> &entries_;
    std::unordered_map<directory_id, directory_id> parent_storage_;
    std::unordered_map<directory_id, directory_id> parent_;
};


compound_document::compound_document(std::vector<std::uint8_t> &data)
    : writer_(new binary_writer(data)),
      reader_(new binary_reader(data)),
      rb_tree_(new red_black_tree(entries_))
{
    header_ = compound_document_header();
    header_.msat.fill(FreeSector);
    msat_.resize(109, FreeSector);
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
    auto &entry = contains_entry(name)
        ? find_entry(name)
        : insert_entry(name, compound_document_entry::entry_type::UserStream);
    entry.size = data.size();

    if (entry.size < header_.threshold)
    {
        const auto num_sectors = data.size() / short_sector_size() + (data.size() % short_sector_size() ? 1 : 0);
        entry.start = allocate_short_sectors(num_sectors);
        write_short(data, entry.start);
    }
    else
    {
        const auto num_sectors = data.size() / sector_size() + (data.size() % sector_size() ? 1 : 0);
        entry.start = allocate_sectors(num_sectors);
        write(data, entry.start);
    }

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

void compound_document::write_short(const std::vector<byte> &data, sector_id current_short)
{
    const auto header_size = sizeof(compound_document_header);
    auto sector_data_reader = binary_reader(data);
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
        writer_->offset(header_size + sector_size() * current_sector + offset);
        writer_->append(data, offset, short_sector_size());

        current_short = ssat_[current_short];
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
        auto current_sector_data = reader_->read_vector<byte>(sector_size());
        sector_data_writer.append(current_sector_data);
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
        auto current_sector_data = reader_->read_vector<byte>(short_sector_size());
        sector_data_writer.append(current_sector_data);

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
    auto current = header_.ssat_start;

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
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    auto num_free = static_cast<std::size_t>(std::count(sat_.begin(), sat_.end(), FreeSector));

    while (num_free < count)
    {
        const auto sat_index = sat_.size() / sectors_per_sector;
        const auto msat_index = std::find(msat_.begin(), msat_.end(), FreeSector) - msat_.begin();

        msat_[msat_index] = sat_index;
        header_.msat[msat_index] = sat_index;

        ++header_.num_msat_sectors;

        write_header();

        auto previous_size = sat_.size();
        sat_.resize(sat_.size() + sectors_per_sector, FreeSector);
        sat_[sat_index] = SATSector;

        writer_->offset(sector_data_start() + msat_index * sectors_per_sector);
        writer_->append(sat_, previous_size, sectors_per_sector);

        num_free = static_cast<std::size_t>(std::count(sat_.begin(), sat_.end(), FreeSector));
    }

    auto allocated = 0;
    const auto start_iter = std::find(sat_.begin(), sat_.end(), FreeSector);
    const auto start = sector_id(start_iter - sat_.begin());
    auto current = start;

    while (allocated < count)
    {
        const auto next_iter = std::find(sat_.begin() + current + 1, sat_.end(), FreeSector);
        const auto next = sector_id(next_iter - sat_.begin());

        sat_[current] = (allocated == count - 1) ? EndOfChain : next;

        auto msat_index = current / sectors_per_sector;
        writer_->offset(sector_data_start() + msat_index * sector_size() + current * sizeof(sector_id));
        writer_->write(sat_[current]);

        if (sector_data_start() + (current + 1) * sector_size() > writer_->size())
        {
            writer_->extend(sector_size());
        }

        current = next;
        ++allocated;
    }

    return start;
}

void compound_document::reallocate_sectors(sector_id start, std::size_t count)
{
    auto current_sector = start;

    for (auto i = std::size_t(0); i < count; ++i)
    {
        if (current_sector == EndOfChain)
        {
            sat_[current_sector] = allocate_sectors(count - i);
        }

        current_sector = sat_[current_sector];
    }
}

sector_chain compound_document::follow_chain(sector_id start)
{
    auto chain = sector_chain();
    chain.push_back(start);

    while (start != EndOfChain)
    {
        start = sat_[start];
    }

    return chain;
}

sector_id compound_document::allocate_short_sectors(std::size_t count)
{
    const auto sectors_per_sector = sector_size() / sizeof(sector_id);
    const auto num_free = std::count(ssat_.begin(), ssat_.end(), FreeSector);
    auto next_sat = EndOfChain;

    if (num_free < count)
    {
        const auto num_to_allocate = (count - num_free) / sectors_per_sector + (count - num_free ? 1 : 0);
        next_sat = allocate_sectors(num_to_allocate);

        if (header_.num_short_sectors == 0)
        {
            header_.ssat_start = next_sat;
        }

        ssat_.resize(ssat_.size() + sectors_per_sector * num_to_allocate, FreeSector);
    }

    header_.num_short_sectors += count;
    write_header();

    const auto start_iter = std::find(ssat_.begin(), ssat_.end(), FreeSector);
    const auto start = sector_id(start_iter - ssat_.begin());
    auto current = start;
    auto allocated = std::size_t(0);
    auto current_sat = header_.ssat_start;

    while (allocated < count)
    {
        const auto next_iter = std::find(ssat_.begin() + current + 1, ssat_.end(), FreeSector);
        const auto next = sector_id(next_iter - ssat_.begin());

        ssat_[current] = (allocated == count - 1) ? EndOfChain : next;

        current = next;
        ++allocated;
    }

    if (find_entry(u"Root Entry").start < 0)
    {
        find_entry(u"Root Entry").start = start;
    }

    return start;

}

compound_document_entry &compound_document::insert_entry(
    const std::u16string &name,
    compound_document_entry::entry_type type)
{
    auto any_empty_entry = false;
    auto entry_id = directory_id(0);

    for (auto &entry : entries_)
    {
        if (entry.type == compound_document_entry::entry_type::Empty)
        {
            any_empty_entry = true;
            break;
        }

        ++entry_id;
    }

    if (!any_empty_entry)
    {
        const auto entries_per_sector = static_cast<directory_id>(sector_size()
            / sizeof(compound_document_entry));

        for (auto i = std::size_t(0); i < entries_per_sector; ++i)
        {
            auto empty_entry = compound_document_entry();
            empty_entry.type = compound_document_entry::entry_type::Empty;
            entries_.push_back(empty_entry);
        }
    }

    auto &entry = entries_[entry_id];

    entry.name(name);
    entry.type = type;

    rb_tree_->insert(entry_id, 0);
    write_directory_tree();

    return entry;
}

void compound_document::write_directory_tree()
{
    const auto entries_per_sector = static_cast<directory_id>(sector_size() 
        / sizeof(compound_document_entry));
    const auto required_sectors = entries_.size() / entries_per_sector;
    auto current_sector_id = header_.directory_start;
    auto entry_index = directory_id(0);

    reallocate_sectors(header_.directory_start, required_sectors);
    writer_->offset(sector_data_start() + current_sector_id * sector_size());

    for (auto &e : entries_)
    {
        writer_->write(e);

        ++entry_index;

        if (entry_index % entries_per_sector == 0)
        {
            current_sector_id = sat_[current_sector_id];
            writer_->offset(sector_data_start() + current_sector_id * sector_size());
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

void compound_document::print_directory()
{
    red_black_tree rb_tree(entries_);

    for (auto entry_id = directory_id(0); entry_id < directory_id(entries_.size()); ++entry_id)
    {
        if (entries_[entry_id].type != compound_document_entry::entry_type::UserStream) continue;

        auto path = std::string("/");

        for (auto part : rb_tree.path(entry_id))
        {
            if (part > 0)
            {
                path.append(utf16_to_utf8(entries_[part].name()) + "/");
            }
        }

        std::cout << path << utf16_to_utf8(entries_[entry_id].name()) << std::endl;
    }
}

} // namespace detail
} // namespace xlnt
