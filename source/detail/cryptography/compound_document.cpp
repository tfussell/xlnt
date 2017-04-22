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
#include <vector>

#include <detail/cryptography/compound_document.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

struct header
{
    std::array<std::uint8_t, 8> id; // signature, or magic identifier
    std::uint16_t big_shift;        // bbat->blockSize = 1 << b_shift
    std::uint16_t small_shift;      // sbat->blockSize = 1 << s_shift
    std::uint32_t num_big_blocks;   // blocks allocated for big bat
    std::uint32_t directory_start;  // starting block for directory info
    std::uint32_t threshold;        // switch from small to big file (usually 4K)
    std::uint32_t small_start;      // starting block index to store small bat
    std::uint32_t num_small_blocks;  // blocks allocated for small bat
    std::uint32_t meta_start;       // starting block to store meta bat
    std::uint32_t num_meta_blocks;  // blocks allocated for meta bat
    std::array<std::uint32_t, 109> bb_blocks;
};

class allocation_table
{
public:
    static const std::uint32_t Eof;
    static const std::uint32_t Avail;
    static const std::uint32_t Bat;
    static const std::uint32_t MetaBat;
    std::size_t blockSize;
    allocation_table();
    void clear();
    std::size_t count();
    void resize(std::size_t newsize);
    //void preserve(std::size_t n);
    //void set(std::size_t index, std::uint32_t val);
    //std::size_t unused();
    //void setChain(std::vector<std::uint32_t>);
    std::vector<std::uint32_t> follow(std::uint32_t start);
    void load(const std::vector<std::uint8_t> &data);
    //void save(std::vector<std::uint8_t> &data);
    //std::size_t size();
private:
    std::vector<std::uint32_t> data_;
};

struct directory_entry
{
public:
    bool valid;            // false if invalid (should be skipped)
    std::string name;      // the name, not in unicode anymore
    bool dir;              // true if directory
    std::uint32_t size;    // size (not valid if directory)
    std::uint32_t start;   // starting block
    std::uint32_t prev;         // previous sibling
    std::uint32_t next;         // next sibling
    std::uint32_t child;        // first child
};

class directory_tree
{
public:
    static const std::uint32_t End;
    directory_tree();
    void clear();
    std::size_t entryCount();
    directory_entry* entry(std::size_t index);
    directory_entry* entry(const std::string& name, bool create = false);
    std::ptrdiff_t indexOf(directory_entry* e);
    std::ptrdiff_t parent(std::size_t index);
    std::string fullName(std::size_t index);
    std::vector<std::size_t> children(std::size_t index);
    void load(const std::vector<std::uint8_t> &data);
    //void save(std::vector<std::uint8_t> &data);
    //std::size_t size();
private:
    std::vector<directory_entry> entries;
};

const std::uint32_t allocation_table::Avail = 0xffffffff;
const std::uint32_t allocation_table::Eof = 0xfffffffe;
const std::uint32_t allocation_table::Bat = 0xfffffffd;
const std::uint32_t allocation_table::MetaBat = 0xfffffffc;

allocation_table::allocation_table()
    : blockSize(4096)
{
    // initial size
    resize(128);
}

std::size_t allocation_table::count()
{
    return data_.size();
}

void allocation_table::resize(std::size_t newsize)
{
    data_.resize(newsize, Avail);
}

// make sure there're still free blocks
/*
void allocation_table::preserve(std::size_t n)
{
    std::vector<std::size_t> pre;
    for (std::size_t i = 0; i < n; i++)
        pre.push_back(unused());
}
*/

/*
void allocation_table::set(std::size_t index, std::uint32_t value)
{
    if (index >= count()) resize(index + 1);
    data_[index] = value;
}
*/

/*
void allocation_table::setChain(std::vector<std::uint32_t> chain)
{
    if (chain.size())
    {
        for (std::size_t i = 0; i < chain.size() - 1; i++)
            set(chain[i], chain[i + 1]);
        set(chain[chain.size() - 1], allocation_table::Eof);
    }
}
*/

// TODO: optimize this with better search
static bool already_exist(const std::vector<std::uint32_t> &chain, std::uint32_t item)
{
    for (std::size_t i = 0; i < chain.size(); i++)
        if (chain[i] == item) return true;

    return false;
}

// follow
std::vector<std::uint32_t> allocation_table::follow(std::uint32_t start)
{
    auto chain = std::vector<std::uint32_t>();
    if (start >= count()) return chain;

    auto p = start;

    while (p < count())
    {
        if (p == static_cast<std::size_t>(Eof)) break;
        if (p == static_cast<std::size_t>(Bat)) break;
        if (p == static_cast<std::size_t>(MetaBat)) break;
        if (already_exist(chain, p)) break;
        chain.push_back(p);
        if (data_[p] >= count()) break;
        p = data_[p];
    }

    return chain;
}

/*
std::size_t allocation_table::unused()
{
    // find first available block
    for (std::size_t i = 0; i < data_.size(); i++)
        if (data_[i] == Avail) return i;

    // completely full, so enlarge the table
    std::size_t block = data_.size();
    resize(data_.size() + 10);
    return block;
}
*/

void allocation_table::load(const std::vector<std::uint8_t> &data)
{
    resize(data.size() / 4);
    std::copy(
        reinterpret_cast<const std::uint32_t *>(&data[0]),
        reinterpret_cast<const std::uint32_t *>(&data[0] + data.size()),
        data_.begin());
}

// return space required to save this dirtree
/*
std::size_t allocation_table::size()
{
    return count() * 4;
}
*/

/*
void allocation_table::save(std::vector<std::uint8_t> &data)
{
    auto offset = std::size_t(0);

    for (std::size_t i = 0; i < count(); i++)
    {
        xlnt::detail::write_int(data_[i], data, offset);
    }
}
*/

const std::uint32_t directory_tree::End = 0xffffffff;

directory_tree::directory_tree()
    : entries()
{
    clear();
}

void directory_tree::clear()
{
    // leave only root entry
    entries.resize(1);
    entries[0].valid = true;
    entries[0].name = "Root Entry";
    entries[0].dir = true;
    entries[0].size = 0;
    entries[0].start = End;
    entries[0].prev = End;
    entries[0].next = End;
    entries[0].child = End;
}

std::size_t directory_tree::entryCount()
{
    return entries.size();
}

directory_entry *directory_tree::entry(std::size_t index)
{
    if (index >= entryCount()) return nullptr;
    return &entries[index];
}

/*
std::ptrdiff_t directory_tree::indexOf(directory_entry *e)
{
    for (std::size_t i = 0; i < entryCount(); i++)
        if (entry(i) == e) return static_cast<std::ptrdiff_t>(i);

    return -1;
}
*/

/*
std::ptrdiff_t directory_tree::parent(std::size_t index)
{
    // brute-force, basically we iterate for each entries, find its children
    // and check if one of the children is 'index'
    for (std::size_t j = 0; j < entryCount(); j++)
    {
        std::vector<std::size_t> chi = children(j);
        for (std::size_t i = 0; i < chi.size(); i++)
            if (chi[i] == index) return static_cast<std::ptrdiff_t>(j);
    }

    return -1;
}
*/

/*
std::string directory_tree::fullName(std::size_t index)
{
    // don't use root name ("Root Entry"), just give "/"
    if (index == 0) return "/";

    std::string result = entry(index)->name;
    result.insert(0, "/");
    auto p = parent(index);
    directory_entry *_entry = nullptr;
    while (p > 0)
    {
        _entry = entry(static_cast<std::size_t>(p));
        if (_entry->dir && _entry->valid)
        {
            result.insert(0, _entry->name);
            result.insert(0, "/");
        }
        --p;
        index = static_cast<std::size_t>(p);
        if (p <= 0) break;
    }
    return result;
}
*/

// given a fullname (e.g "/ObjectPool/_1020961869"), find the entry
// if not found and create is false, return 0
// if create is true, a new entry is returned
directory_entry *directory_tree::entry(const std::string &name, bool create)
{
    if (!name.length()) return nullptr;

    // quick check for "/" (that's root)
    if (name == "/") return entry(0);

    // split the names, e.g  "/ObjectPool/_1020961869" will become:
    // "ObjectPool" and "_1020961869"
    std::list<std::string> names;
    std::string::size_type start = 0, end = 0;
    if (name[0] == '/') start++;
    while (start < name.length())
    {
        end = name.find_first_of('/', start);
        if (end == std::string::npos) end = name.length();
        names.push_back(name.substr(start, end - start));
        start = end + 1;
    }

    // start from root
    std::size_t index = 0;

    // trace one by one
    std::list<std::string>::iterator it;

    for (it = names.begin(); it != names.end(); ++it)
    {
        // find among the children of index
        std::vector<std::size_t> chi = children(index);
        std::ptrdiff_t child = 0;
        for (std::size_t i = 0; i < chi.size(); i++)
        {
            directory_entry *ce = entry(chi[i]);
            if (ce)
                if (ce->valid && (ce->name.length() > 1))
                    if (ce->name == *it) child = static_cast<std::ptrdiff_t>(chi[i]);
        }

        // traverse to the child
        if (child > 0)
            index = static_cast<std::size_t>(child);
        else
        {
            // not found among children
            if (!create) return nullptr;

            // create a new entry
            std::size_t parent = index;
            entries.push_back(directory_entry());
            index = entryCount() - 1;
            directory_entry *e = entry(index);
            e->valid = true;
            e->name = *it;
            e->dir = false;
            e->size = 0;
            e->start = 0;
            e->child = End;
            e->prev = End;
            e->next = entry(parent)->child;
            entry(parent)->child = static_cast<std::uint32_t>(index);
        }
    }

    return entry(index);
}

// helper function: recursively find siblings of index
void dirtree_find_siblings(directory_tree *dirtree, std::vector<std::size_t> &result, std::size_t index)
{
    auto e = dirtree->entry(index);
    if (!e) return;
    if (!e->valid) return;

    // prevent infinite loop
    for (std::size_t i = 0; i < result.size(); i++)
        if (result[i] == index) return;

    // add myself
    result.push_back(index);

    // visit previous sibling, don't go infinitely
    std::size_t prev = e->prev;
    if ((prev > 0) && (prev < dirtree->entryCount()))
    {
        for (std::size_t i = 0; i < result.size(); i++)
            if (result[i] == prev) prev = 0;
        if (prev) dirtree_find_siblings(dirtree, result, prev);
    }

    // visit next sibling, don't go infinitely
    std::size_t next = e->next;
    if ((next > 0) && (next < dirtree->entryCount()))
    {
        for (std::size_t i = 0; i < result.size(); i++)
            if (result[i] == next) next = 0;
        if (next) dirtree_find_siblings(dirtree, result, next);
    }
}

std::vector<std::size_t> directory_tree::children(std::size_t index)
{
    std::vector<std::size_t> result;

    directory_entry *e = entry(index);
    if (e)
        if (e->valid && e->child < entryCount()) dirtree_find_siblings(this, result, e->child);

    return result;
}

void directory_tree::load(const std::vector<std::uint8_t> &data)
{
    entries.clear();

    for (std::size_t i = 0; i < data.size() / 128; i++)
    {
        std::size_t p = i * 128;
        auto offset = p + 0x40;

        // parse name of this entry, which stored as Unicode 16-bit
        std::string name;
        auto name_len = static_cast<std::size_t>(xlnt::detail::read_int<std::uint16_t>(data, offset));
        if (name_len > 64) name_len = 64;
        for (std::size_t j = 0; (data[j + p]) && (j < name_len); j += 2)
        {
            name.append(1, static_cast<char>(data[j + p]));
        }

        // first char isn't printable ? remove it...
        if (data[p] < 32)
        {
            name.erase(0, 1);
        }

        // 2 = file (aka stream), 1 = directory (aka storage), 5 = root
        std::size_t type = data[0x42 + p];

        directory_entry e;
        e.valid = true;
        e.name = name;
        offset = 0x74 + p;
        e.start = xlnt::detail::read_int<std::uint32_t>(data, offset);
        offset = 0x78 + p;
        e.size = xlnt::detail::read_int<std::uint32_t>(data, offset);
        offset = 0x44 + p;
        e.prev = xlnt::detail::read_int<std::uint32_t>(data, offset);
        offset = 0x48 + p;
        e.next = xlnt::detail::read_int<std::uint32_t>(data, offset);
        offset = 0x4C + p;
        e.child = xlnt::detail::read_int<std::uint32_t>(data, offset);
        e.dir = (type != 2);

        // sanity checks
        if ((type != 2) && (type != 1) && (type != 5)) e.valid = false;
        if (name_len < 1) e.valid = false;

        entries.push_back(e);
    }
}

// return space required to save this dirtree
/*
std::size_t directory_tree::size()
{
    return entryCount() * 128;
}
*/

/*
void directory_tree::save(std::vector<std::uint8_t> &data)
{
    std::fill(data.begin(), data.begin() + size(), std::uint8_t(0));

    // root is fixed as "Root Entry"
    directory_entry *root = entry(0);
    std::string entry_name = "Root Entry";

    for (std::size_t j = 0; j < entry_name.length(); j++)
    {
        data[j * 2] = static_cast<std::uint8_t>(entry_name[j]);
    }

    auto offset = std::size_t(0x40);
    xlnt::detail::write_int(static_cast<std::uint16_t>(entry_name.length() * 2 + 2), data, offset);
    data[0x42] = 5;
    data[0x43] = 1;
    xlnt::detail::write_int(static_cast<std::uint32_t>(0xffffffff), data, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(0xffffffff), data, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(root->child), data, offset);

    offset = 0x74;
    xlnt::detail::write_int(0xffffffff, data, offset);

    for (std::size_t i = 1; i < entryCount(); i++)
    {
        directory_entry *e = entry(i);
        if (!e) continue;

        if (e->dir)
        {
            e->start = 0xffffffff;
            e->size = 0;
        }

        // max length for name is 32 chars
        entry_name = e->name;
        if (entry_name.length() > 32) entry_name.erase(32, entry_name.length());

        // write name as Unicode 16-bit
        for (std::size_t j = 0; j < entry_name.length(); j++)
        {
            data[i * 128 + j * 2] = static_cast<std::uint8_t>(entry_name[j]);
        }

        offset = 0x40 + i * 128;
        xlnt::detail::write_int(static_cast<std::uint16_t>(entry_name.length() * 2 + 2), data, offset);
        xlnt::detail::write_int(static_cast<std::uint16_t>(0), data, offset);
        xlnt::detail::write_int(e->prev, data, offset);
        xlnt::detail::write_int(e->next, data, offset);
        xlnt::detail::write_int(e->child, data, offset);

        offset = 0x74 + i * 128;
        xlnt::detail::write_int(e->start, data, offset);
        xlnt::detail::write_int(e->size, data, offset);

        data[i * 128 + 0x42] = e->dir ? 1 : 2;
        data[i * 128 + 0x43] = 1; // always black
    }
}
*/

static const std::array<std::uint8_t, 8> pole_magic = {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1};

header initialize_header()
{
    header h;

    h.big_shift = 9;
    h.small_shift = 6;
    h.num_big_blocks = 0;
    h.directory_start = 0;
    h.threshold = 4096;
    h.small_start = 0;
    h.num_small_blocks = 0;
    h.meta_start = 0;
    h.num_meta_blocks = 0;
    std::copy(pole_magic.begin(), pole_magic.end(), h.id.begin());
    h.bb_blocks.fill(allocation_table::Avail);

    return h;
}

bool header_is_valid(const header &h)
{
    if (h.threshold != 4096) return false;
    if (h.num_big_blocks == 0) return false;
    if ((h.num_big_blocks > 109) && (h.num_big_blocks > (h.num_meta_blocks * 127) + 109)) return false;
    if ((h.num_big_blocks < 109) && (h.num_meta_blocks != 0)) return false;
    if (h.small_shift > h.big_shift) return false;
    if (h.big_shift <= 6) return false;
    if (h.big_shift >= 31) return false;

    return true;
}

header load_header(const std::vector<std::uint8_t> &data)
{
    header h;

    auto in = &data[0];
    auto offset = std::size_t(0x1e);

    h.big_shift = xlnt::detail::read_int<std::uint16_t>(data, offset);
    h.small_shift = xlnt::detail::read_int<std::uint16_t>(data, offset);

    offset = 0x2c;
    h.num_big_blocks = xlnt::detail::read_int<std::uint32_t>(data, offset);
    h.directory_start = xlnt::detail::read_int<std::uint32_t>(data, offset);

    offset = 0x38;
    h.threshold = xlnt::detail::read_int<std::uint32_t>(data, offset);
    h.small_start = xlnt::detail::read_int<std::uint32_t>(data, offset);
    h.num_small_blocks = xlnt::detail::read_int<std::uint32_t>(data, offset);
    h.meta_start = xlnt::detail::read_int<std::uint32_t>(data, offset);
    h.num_meta_blocks = xlnt::detail::read_int<std::uint32_t>(data, offset);

    for (std::size_t i = 0; i < 8; i++)
    {
        h.id[i] = in[i];
    }

    offset = 0x4c;

    for (std::size_t i = 0; i < 109; i++)
    {
        h.bb_blocks[i] = xlnt::detail::read_int<std::uint32_t>(data, offset);
    }

    return h;
}

/*
void save_header(const header &h, std::vector<std::uint8_t> &out)
{
    std::fill(out.begin(), out.begin() + 0x4c, std::uint8_t(0));
    std::copy(h.id.begin(), h.id.end(), out.begin());

    auto offset = std::size_t(0);

    xlnt::detail::write_int(static_cast<std::uint32_t>(0), out, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(0), out, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(0), out, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(0), out, offset);
    xlnt::detail::write_int(static_cast<std::uint16_t>(0x003e), out, offset);
    xlnt::detail::write_int(static_cast<std::uint16_t>(3), out, offset);
    xlnt::detail::write_int(static_cast<std::uint16_t>(0xfffe), out, offset);

    xlnt::detail::write_int(h.big_shift, out, offset);
    xlnt::detail::write_int(h.small_shift, out, offset);

    offset = 0x2c;

    xlnt::detail::write_int(h.num_big_blocks, out, offset);
    xlnt::detail::write_int(h.directory_start, out, offset);
    xlnt::detail::write_int(static_cast<std::uint32_t>(0), out, offset);
    xlnt::detail::write_int(h.threshold, out, offset);
    xlnt::detail::write_int(h.small_start, out, offset);
    xlnt::detail::write_int(h.num_small_blocks, out, offset);
    xlnt::detail::write_int(h.meta_start, out, offset);
    xlnt::detail::write_int(h.num_meta_blocks, out, offset);

    for (std::size_t i = 0; i < 109; i++)
    {
        xlnt::detail::write_int(h.bb_blocks[i], out, offset);
    }
}
*/

} // namespace

namespace xlnt {
namespace detail {

struct compound_document_impl
{
    byte_vector buffer_;
    directory_tree directory_;
    header header_;
    allocation_table small_block_table_;
    allocation_table big_block_table_;
    std::vector<std::uint32_t> small_blocks_;
};

std::vector<std::uint8_t> load_big_blocks(const std::vector<std::uint32_t> &blocks, compound_document_impl &d)
{
    std::vector<std::uint8_t> result;
    auto bytes_loaded = std::size_t(0);
    const auto block_size = d.big_block_table_.blockSize;

    for (auto block : blocks)
    {
        auto position = block_size * (block + 1);
        auto block_length = std::size_t(512);
        auto current_size = result.size();
        result.resize(result.size() + block_length);

        std::copy(
            d.buffer_.begin() + position,
            d.buffer_.begin() + position + block_length,
            result.begin() + current_size);

        bytes_loaded += block_length;
    }

    return result;
}

std::vector<std::uint8_t> load_small_blocks(const std::vector<std::uint32_t> &blocks, compound_document_impl &d)
{
    std::vector<std::uint8_t> result;
    auto bytes_loaded = std::size_t(0);
    const auto small_block_size = d.small_block_table_.blockSize;
    const auto big_block_size = d.big_block_table_.blockSize;

    for (auto block : blocks)
    {
        auto position = block * small_block_size;
        auto bbindex = position / big_block_size;

        if (bbindex >= d.small_blocks_.size()) break;

        auto block_data = load_big_blocks({ d.small_blocks_[bbindex] }, d);

        auto offset = position % big_block_size;
        auto current_size = result.size();
        result.resize(result.size() + small_block_size);

        std::copy(
            block_data.begin() + offset,
            block_data.begin() + offset + small_block_size,
            result.begin() + current_size);

        bytes_loaded += small_block_size;
    }

    return result;
}

compound_document::compound_document()
    : d_(new compound_document_impl())
{
    d_->header_ = initialize_header();

    d_->big_block_table_.blockSize = static_cast<std::size_t>(1) << d_->header_.big_shift;
    d_->small_block_table_.blockSize = static_cast<std::size_t>(1) << d_->header_.small_shift;
}

compound_document::~compound_document()
{
}

void compound_document::load(const std::vector<std::uint8_t> &data)
{
    d_->buffer_ = data;

    // 1. load header
    {
        if (data.size() < 512)
        {
            throw xlnt::exception("not ole");
        }

        d_->header_ = load_header(data);

        if (d_->header_.id != pole_magic)
        {
            throw xlnt::exception("not ole");
        }

        if (!header_is_valid(d_->header_) || d_->header_.threshold != 4096)
        {
            throw xlnt::exception("bad ole");
        }
    }

    // important block size
    auto big_block_size = static_cast<std::size_t>(1) << d_->header_.big_shift;
    d_->big_block_table_.blockSize = big_block_size;
    auto small_block_size = static_cast<std::size_t>(1) << d_->header_.small_shift;
    d_->small_block_table_.blockSize = small_block_size;

    // find blocks allocated to store big bat
    // the first 109 blocks are in header, the rest in meta bat
    auto num_header_blocks = std::min(std::uint32_t(109), d_->header_.num_big_blocks);
    auto blocks = std::vector<std::uint32_t>(
        d_->header_.bb_blocks.begin(), 
        d_->header_.bb_blocks.begin() + num_header_blocks);
    auto buffer = byte_vector();

    if ((d_->header_.num_big_blocks > 109) && (d_->header_.num_meta_blocks > 0))
    {
        buffer.resize(big_block_size);

        std::size_t k = 109;

        for (std::size_t r = 0; r < d_->header_.num_meta_blocks; r++)
        {
            load_big_blocks(blocks, *d_);

            for (std::size_t s = 0; s < big_block_size - 4; s += 4)
            {
                if (k >= d_->header_.num_big_blocks) break;
                auto offset = s;
                blocks[k++] = xlnt::detail::read_int<std::uint32_t>(buffer, offset);
            }

            auto offset = big_block_size - 4;
            xlnt::detail::read_int<std::uint32_t>(buffer, offset);
        }
    }

    // load big block table
    if (!blocks.empty())
    {
        d_->big_block_table_.load(load_big_blocks(blocks, *d_));
    }

    // load small bat
    blocks = d_->big_block_table_.follow(d_->header_.small_start);

    if (!blocks.empty())
    {
        d_->small_block_table_.load(load_big_blocks(blocks, *d_));
    }

    // load directory tree
    blocks = d_->big_block_table_.follow(d_->header_.directory_start);
    auto directory_data = load_big_blocks(blocks, *d_);
    d_->directory_.load(directory_data);

    auto offset = std::size_t(0x74);
    auto sb_start = xlnt::detail::read_int<std::uint32_t>(directory_data, offset);

    // fetch block chain as data for small-files
    d_->small_blocks_ = d_->big_block_table_.follow(sb_start); // small files
}

std::vector<std::uint8_t> compound_document::save() const
{
    return d_->buffer_;
}

bool compound_document::has_stream(const std::string &filename) const
{
    return d_->directory_.entry(filename, false) != nullptr;
}

void compound_document::add_stream(
    const std::string &name,
    const std::vector<std::uint8_t> &/*data*/)
{
    d_->directory_.entry(name, !has_stream(name));
}

std::vector<std::uint8_t> compound_document::stream(const std::string &name) const
{
    if (!has_stream(name))
    {
        throw xlnt::exception("document doesn't contain stream with the given name");
    }

    auto entry = d_->directory_.entry(name);
    byte_vector result;

    if (entry->size < d_->header_.threshold)
    {
        result = load_small_blocks(d_->small_block_table_.follow(entry->start), *d_);
        result.resize(entry->size);
    }
    else
    {
        result = load_big_blocks(d_->big_block_table_.follow(entry->start), *d_);
        result.resize(entry->size);
    }

    return result;
}

} // namespace detail
} // namespace xlnt
