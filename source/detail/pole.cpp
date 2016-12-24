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

#include <cassert>
#include <cstring>
#include <fstream>
#include <iostream>
#include <list>
#include <string>
#include <vector>

#include <detail/pole.hpp>

// enable to activate debugging output
// #define POLE_DEBUG

namespace {

// helper function: recursively find siblings of index
void dirtree_find_siblings(POLE::DirTree *dirtree, std::vector<std::size_t> &result, std::size_t index)
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
}

namespace POLE {

} // namespace POLE

using namespace POLE;

static inline std::uint16_t readU16(const std::uint8_t *ptr)
{
    return static_cast<std::uint16_t>(ptr[0] + (ptr[1] << 8));
}

static inline std::uint32_t readU32(const std::uint8_t *ptr)
{
    return static_cast<std::uint32_t>(ptr[0] + (ptr[1] << 8) + (ptr[2] << 16) + (ptr[3] << 24));
}

static inline void writeU16(std::uint8_t *ptr, std::uint16_t data)
{
    ptr[0] = static_cast<std::uint8_t>(data & 0xff);
    ptr[1] = static_cast<std::uint8_t>((data >> 8) & 0xff);
}

static inline void writeU32(std::uint8_t *ptr, std::uint32_t data)
{
    ptr[0] = static_cast<std::uint8_t>(data & 0xff);
    ptr[1] = static_cast<std::uint8_t>((data >> 8) & 0xff);
    ptr[2] = static_cast<std::uint8_t>((data >> 16) & 0xff);
    ptr[3] = static_cast<std::uint8_t>((data >> 24) & 0xff);
}

static const std::uint8_t pole_magic[] = {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1};

// =========== Header ==========

Header::Header()
    : b_shift(9),
      s_shift(6),
      num_bat(0),
      dirent_start(0),
      threshold(4096),
      sbat_start(0),
      num_sbat(0),
      mbat_start(0),
      num_mbat(0)
{
    for (std::size_t i = 0; i < 8; i++)
        id[i] = pole_magic[i];
    for (std::size_t i = 0; i < 109; i++)
        bb_blocks[i] = AllocTable::Avail;
}

bool Header::valid()
{
    if (threshold != 4096) return false;
    if (num_bat == 0) return false;
    if ((num_bat > 109) && (num_bat > (num_mbat * 127) + 109)) return false;
    if ((num_bat < 109) && (num_mbat != 0)) return false;
    if (s_shift > b_shift) return false;
    if (b_shift <= 6) return false;
    if (b_shift >= 31) return false;

    return true;
}

void Header::load(const std::uint8_t *buffer)
{
    b_shift = readU16(buffer + 0x1e);
    s_shift = readU16(buffer + 0x20);
    num_bat = readU32(buffer + 0x2c);
    dirent_start = readU32(buffer + 0x30);
    threshold = readU32(buffer + 0x38);
    sbat_start = readU32(buffer + 0x3c);
    num_sbat = readU32(buffer + 0x40);
    mbat_start = readU32(buffer + 0x44);
    num_mbat = readU32(buffer + 0x48);

    for (std::size_t i = 0; i < 8; i++)
        id[i] = buffer[i];
    for (std::size_t i = 0; i < 109; i++)
        bb_blocks[i] = readU32(buffer + 0x4C + i * 4);
}

void Header::save(std::uint8_t *buffer)
{
    memset(buffer, 0, 0x4c);
    memcpy(buffer, pole_magic, 8); // ole signature
    writeU32(buffer + 8, 0); // unknown
    writeU32(buffer + 12, 0); // unknown
    writeU32(buffer + 16, 0); // unknown
    writeU16(buffer + 24, 0x003e); // revision ?
    writeU16(buffer + 26, 3); // version ?
    writeU16(buffer + 28, 0xfffe); // unknown
    writeU16(buffer + 0x1e, b_shift);
    writeU16(buffer + 0x20, s_shift);
    writeU32(buffer + 0x2c, num_bat);
    writeU32(buffer + 0x30, dirent_start);
    writeU32(buffer + 0x38, threshold);
    writeU32(buffer + 0x3c, sbat_start);
    writeU32(buffer + 0x40, num_sbat);
    writeU32(buffer + 0x44, mbat_start);
    writeU32(buffer + 0x48, num_mbat);

    for (std::size_t i = 0; i < 109; i++)
        writeU32(buffer + 0x4C + i * 4, bb_blocks[i]);
}

void Header::debug()
{
    std::cout << std::endl;
    std::cout << "b_shift " << b_shift << std::endl;
    std::cout << "s_shift " << s_shift << std::endl;
    std::cout << "num_bat " << num_bat << std::endl;
    std::cout << "dirent_start " << dirent_start << std::endl;
    std::cout << "threshold " << threshold << std::endl;
    std::cout << "sbat_start " << sbat_start << std::endl;
    std::cout << "num_sbat " << num_sbat << std::endl;
    std::cout << "mbat_start " << mbat_start << std::endl;
    std::cout << "num_mbat " << num_mbat << std::endl;

    std::size_t s = (num_bat <= 109) ? num_bat : 109;
    std::cout << "bat blocks: ";
    for (std::size_t i = 0; i < s; i++)
        std::cout << bb_blocks[i] << " ";
    std::cout << std::endl;
}

// =========== AllocTable ==========

const std::uint32_t AllocTable::Avail = 0xffffffff;
const std::uint32_t AllocTable::Eof = 0xfffffffe;
const std::uint32_t AllocTable::Bat = 0xfffffffd;
const std::uint32_t AllocTable::MetaBat = 0xfffffffc;

AllocTable::AllocTable()
    : blockSize(4096), data()
{
    // initial size
    resize(128);
}

std::size_t AllocTable::count()
{
    return data.size();
}

void AllocTable::resize(std::size_t newsize)
{
    std::size_t oldsize = data.size();
    data.resize(newsize);
    if (newsize > oldsize)
        for (std::size_t i = oldsize; i < newsize; i++)
            data[i] = Avail;
}

// make sure there're still free blocks
void AllocTable::preserve(std::size_t n)
{
    std::vector<std::size_t> pre;
    for (std::size_t i = 0; i < n; i++)
        pre.push_back(unused());
}

std::size_t AllocTable::operator[](std::size_t index)
{
    std::size_t result;
    result = data[index];
    return result;
}

void AllocTable::set(std::size_t index, std::uint32_t value)
{
    if (index >= count()) resize(index + 1);
    data[index] = value;
}

void AllocTable::setChain(std::vector<std::uint32_t> chain)
{
    if (chain.size())
    {
        for (std::size_t i = 0; i < chain.size() - 1; i++)
            set(chain[i], chain[i + 1]);
        set(chain[chain.size() - 1], AllocTable::Eof);
    }
}

// TODO: optimize this with better search
static bool already_exist(const std::vector<std::size_t> &chain, std::size_t item)
{
    for (std::size_t i = 0; i < chain.size(); i++)
        if (chain[i] == item) return true;

    return false;
}

// follow
std::vector<std::size_t> AllocTable::follow(std::size_t start)
{
    std::vector<std::size_t> chain;

    if (start >= count()) return chain;

    std::size_t p = start;
    while (p < count())
    {
        if (p == static_cast<std::size_t>(Eof)) break;
        if (p == static_cast<std::size_t>(Bat)) break;
        if (p == static_cast<std::size_t>(MetaBat)) break;
        if (already_exist(chain, p)) break;
        chain.push_back(p);
        if (data[p] >= count()) break;
        p = data[p];
    }

    return chain;
}

std::size_t AllocTable::unused()
{
    // find first available block
    for (std::size_t i = 0; i < data.size(); i++)
        if (data[i] == Avail) return i;

    // completely full, so enlarge the table
    std::size_t block = data.size();
    resize(data.size() + 10);
    return block;
}

void AllocTable::load(const std::uint8_t *buffer, std::size_t len)
{
    resize(len / 4);
    for (std::size_t i = 0; i < count(); i++)
        set(i, readU32(buffer + i * 4));
}

// return space required to save this dirtree
std::size_t AllocTable::size()
{
    return count() * 4;
}

void AllocTable::save(std::uint8_t *buffer)
{
    for (std::size_t i = 0; i < count(); i++)
        writeU32(buffer + i * 4, data[i]);
}

void AllocTable::debug()
{
    std::cout << "block size " << data.size() << std::endl;
    for (std::size_t i = 0; i < data.size(); i++)
    {
        if (data[i] == Avail) continue;
        std::cout << i << ": ";
        if (data[i] == Eof)
            std::cout << "[eof]";
        else if (data[i] == Bat)
            std::cout << "[bat]";
        else if (data[i] == MetaBat)
            std::cout << "[metabat]";
        else
            std::cout << data[i];
        std::cout << std::endl;
    }
}

// =========== DirTree ==========

const std::uint32_t DirTree::End = 0xffffffff;

DirTree::DirTree()
    : entries()
{
    clear();
}

void DirTree::clear()
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

std::size_t DirTree::entryCount()
{
    return entries.size();
}

DirEntry *DirTree::entry(std::size_t index)
{
    if (index >= entryCount()) return nullptr;
    return &entries[index];
}

std::ptrdiff_t DirTree::indexOf(DirEntry *e)
{
    for (std::size_t i = 0; i < entryCount(); i++)
        if (entry(i) == e) return static_cast<std::ptrdiff_t>(i);

    return -1;
}

std::ptrdiff_t DirTree::parent(std::size_t index)
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

std::string DirTree::fullName(std::size_t index)
{
    // don't use root name ("Root Entry"), just give "/"
    if (index == 0) return "/";

    std::string result = entry(index)->name;
    result.insert(0, "/");
    auto p = parent(index);
    DirEntry *_entry = 0;
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

// given a fullname (e.g "/ObjectPool/_1020961869"), find the entry
// if not found and create is false, return 0
// if create is true, a new entry is returned
DirEntry *DirTree::entry(const std::string &name, bool create)
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
            DirEntry *ce = entry(chi[i]);
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
            entries.push_back(DirEntry());
            index = entryCount() - 1;
            DirEntry *e = entry(index);
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

std::vector<std::size_t> DirTree::children(std::size_t index)
{
    std::vector<std::size_t> result;

    DirEntry *e = entry(index);
    if (e)
        if (e->valid && e->child < entryCount()) dirtree_find_siblings(this, result, e->child);

    return result;
}

void DirTree::load(std::uint8_t *buffer, std::size_t size)
{
    entries.clear();

    for (std::size_t i = 0; i < size / 128; i++)
    {
        std::size_t p = i * 128;

        // parse name of this entry, which stored as Unicode 16-bit
        std::string name;
        auto name_len = static_cast<std::size_t>(readU16(buffer + 0x40 + p));
        if (name_len > 64) name_len = 64;
        for (std::size_t j = 0; (buffer[j + p]) && (j < name_len); j += 2)
            name.append(1, static_cast<char>(buffer[j + p]));

        // first char isn't printable ? remove it...
        if (buffer[p] < 32)
        {
            name.erase(0, 1);
        }

        // 2 = file (aka stream), 1 = directory (aka storage), 5 = root
        std::size_t type = buffer[0x42 + p];

        DirEntry e;
        e.valid = true;
        e.name = name;
        e.start = readU32(buffer + 0x74 + p);
        e.size = readU32(buffer + 0x78 + p);
        e.prev = readU32(buffer + 0x44 + p);
        e.next = readU32(buffer + 0x48 + p);
        e.child = readU32(buffer + 0x4C + p);
        e.dir = (type != 2);

        // sanity checks
        if ((type != 2) && (type != 1) && (type != 5)) e.valid = false;
        if (name_len < 1) e.valid = false;

        entries.push_back(e);
    }
}

// return space required to save this dirtree
std::size_t DirTree::size()
{
    return entryCount() * 128;
}

void DirTree::save(std::uint8_t *buffer)
{
    memset(buffer, 0, size());

    // root is fixed as "Root Entry"
    DirEntry *root = entry(0);
    std::string entry_name = "Root Entry";
    for (std::size_t j = 0; j < entry_name.length(); j++)
        buffer[j * 2] = static_cast<std::uint8_t>(entry_name[j]);
    writeU16(buffer + 0x40, static_cast<std::uint16_t>(entry_name.length() * 2 + 2));
    writeU32(buffer + 0x74, 0xffffffff);
    writeU32(buffer + 0x78, 0);
    writeU32(buffer + 0x44, 0xffffffff);
    writeU32(buffer + 0x48, 0xffffffff);
    writeU32(buffer + 0x4c, root->child);
    buffer[0x42] = 5;
    buffer[0x43] = 1;

    for (std::size_t i = 1; i < entryCount(); i++)
    {
        DirEntry *e = entry(i);
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
            buffer[i * 128 + j * 2] = static_cast<std::uint8_t>(entry_name[j]);

        writeU16(buffer + i * 128 + 0x40, static_cast<std::uint16_t>(entry_name.length() * 2 + 2));
        writeU32(buffer + i * 128 + 0x74, e->start);
        writeU32(buffer + i * 128 + 0x78, e->size);
        writeU32(buffer + i * 128 + 0x44, e->prev);
        writeU32(buffer + i * 128 + 0x48, e->next);
        writeU32(buffer + i * 128 + 0x4c, e->child);
        buffer[i * 128 + 0x42] = e->dir ? 1 : 2;
        buffer[i * 128 + 0x43] = 1; // always black
    }
}

void DirTree::debug()
{
    for (std::size_t i = 0; i < entryCount(); i++)
    {
        DirEntry *e = entry(i);
        if (!e) continue;
        std::cout << i << ": ";
        if (!e->valid) std::cout << "INVALID ";
        std::cout << e->name << " ";
        if (e->dir)
            std::cout << "(Dir) ";
        else
            std::cout << "(File) ";
        std::cout << e->size << " ";
        std::cout << "s:" << e->start << " ";
        std::cout << "(";
        if (e->child == End)
            std::cout << "-";
        else
            std::cout << e->child;
        std::cout << " ";
        if (e->prev == End)
            std::cout << "-";
        else
            std::cout << e->prev;
        std::cout << ":";
        if (e->next == End)
            std::cout << "-";
        else
            std::cout << e->next;
        std::cout << ")";
        std::cout << std::endl;
    }
}

// =========== StorageIO ==========

StorageIO::StorageIO(Storage *st, char *bytes, std::size_t length)
    : storage(st),
      filedata(reinterpret_cast<std::uint8_t *>(bytes)),
      dataLength(length),
      result(Storage::Ok),
      opened(false),
      filesize(0),
      header(new Header()),
      dirtree(new DirTree()),
      bbat(new AllocTable()),
      sbat(new AllocTable()),
      sb_blocks(),
      streams()
{
    bbat->blockSize = static_cast<std::size_t>(1) << header->b_shift;
    sbat->blockSize = static_cast<std::size_t>(1) << header->s_shift;
}

StorageIO::~StorageIO()
{
    if (opened) close();
    delete sbat;
    delete bbat;
    delete dirtree;
    delete header;
}

bool StorageIO::open()
{
    // already opened ? close first
    if (opened) close();

    load();

    return result == Storage::Ok;
}

void StorageIO::load()
{
    std::uint8_t *buffer = 0;
    std::size_t buflen = 0;
    std::vector<std::size_t> blocks;

    // open the file, check for error
    result = Storage::OpenFailed;
    // FSTREAM file.open( filename.c_str(), std::ios::binary | std::ios::in );
    // FSTREAM if( !file.good() ) return;

    // find size of input file
    // FSTREAM file.seekg( 0, std::ios::end );
    // FSTREAM filesize = file.tellg();
    filesize = dataLength;

    // load header
    buffer = new std::uint8_t[512];
    // FSTREAM file.seekg( 0 );
    // FSTREAM file.read( (char*)buffer, 512 );
    memcpy(buffer, filedata, 512);
    header->load(buffer);
    delete[] buffer;

    // check OLE magic id
    result = Storage::NotOLE;
    for (std::size_t i = 0; i < 8; i++)
        if (header->id[i] != pole_magic[i]) return;

    // sanity checks
    result = Storage::BadOLE;
    if (!header->valid()) return;
    if (header->threshold != 4096) return;

    // important block size
    bbat->blockSize = static_cast<std::size_t>(1) << header->b_shift;
    sbat->blockSize = static_cast<std::size_t>(1) << header->s_shift;

    // find blocks allocated to store big bat
    // the first 109 blocks are in header, the rest in meta bat
    blocks.clear();
    blocks.resize(header->num_bat);
    for (std::size_t i = 0; i < 109; i++)
        if (i >= header->num_bat)
            break;
        else
            blocks[i] = header->bb_blocks[i];
    if ((header->num_bat > 109) && (header->num_mbat > 0))
    {
        std::uint8_t *buffer2 = new std::uint8_t[bbat->blockSize];
        memset(buffer2, 0, bbat->blockSize);
        std::size_t k = 109;
        std::size_t mblock = header->mbat_start;
        for (std::size_t r = 0; r < header->num_mbat; r++)
        {
            loadBigBlock(mblock, buffer2, bbat->blockSize);
            for (std::size_t s = 0; s < bbat->blockSize - 4; s += 4)
            {
                if (k >= header->num_bat)
                    break;
                else
                    blocks[k++] = readU32(buffer2 + s);
            }
            mblock = readU32(buffer2 + bbat->blockSize - 4);
        }
        delete[] buffer2;
    }

    // load big bat
    buflen = blocks.size() * bbat->blockSize;
    if (buflen > 0)
    {
        buffer = new std::uint8_t[buflen];
        memset(buffer, 0, buflen);
        loadBigBlocks(blocks, buffer, buflen);
        bbat->load(buffer, buflen);
        delete[] buffer;
    }

    // load small bat
    blocks.clear();
    blocks = bbat->follow(header->sbat_start);
    buflen = blocks.size() * bbat->blockSize;
    if (buflen > 0)
    {
        buffer = new std::uint8_t[buflen];
        memset(buffer, 0, buflen);
        loadBigBlocks(blocks, buffer, buflen);
        sbat->load(buffer, buflen);
        delete[] buffer;
    }

    // load directory tree
    blocks.clear();
    blocks = bbat->follow(header->dirent_start);
    buflen = blocks.size() * bbat->blockSize;
    buffer = new std::uint8_t[buflen];
    memset(buffer, 0, buflen);
    loadBigBlocks(blocks, buffer, buflen);
    dirtree->load(buffer, buflen);
    std::size_t sb_start = readU32(buffer + 0x74);
    delete[] buffer;

    // fetch block chain as data for small-files
    sb_blocks = bbat->follow(sb_start); // small files

// for troubleshooting, just enable this block
#if 0
    header->debug();
    sbat->debug();
    bbat->debug();
    dirtree->debug();
#endif

    // so far so good
    result = Storage::Ok;
    opened = true;
}

void StorageIO::create()
{
    // std::cout << "Creating " << filename << std::endl;

    /*FSTREAM file.open( filename.c_str(), std::ios::out|std::ios::binary );
    if( !file.good() )
    {
        std::cerr << "Can't create " << filename << std::endl;
        result = Storage::OpenFailed;
        return;
    }*/

    // so far so good
    opened = true;
    result = Storage::Ok;
}

void StorageIO::flush()
{
    /* Note on Microsoft implementation:
     - directory entries are stored in the last block(s)
     - BATs are as second to the last
     - Meta BATs are third to the last
     */
}

void StorageIO::close()
{
    if (!opened) return;

    // FSTREAM file.close();
    opened = false;

    std::list<Stream *>::iterator it;
    for (it = streams.begin(); it != streams.end(); ++it)
        delete *it;
}

StreamIO *StorageIO::streamIO(const std::string &name)
{
    // sanity check
    if (!name.length()) return nullptr;

    // search in the entries
    DirEntry *entry = dirtree->entry(name);
    // if( entry) std::cout << "FOUND\n";
    if (!entry) return nullptr;
    // if( !entry->dir ) std::cout << "  NOT DIR\n";
    if (entry->dir) return nullptr;

    StreamIO *stream = new StreamIO(this, entry);
    stream->fullName = name;

    return stream;
}

std::size_t StorageIO::loadBigBlocks(std::vector<std::size_t> blocks, std::uint8_t *data, std::size_t maxlen)
{
    // sentinel
    if (!data) return 0;
    if (blocks.size() < 1) return 0;
    if (maxlen == 0) return 0;

    // read block one by one, seems fast enough
    std::size_t bytes = 0;
    for (std::size_t i = 0; (i < blocks.size()) && (bytes < maxlen); i++)
    {
        std::size_t block = blocks[i];
        std::size_t pos = bbat->blockSize * (block + 1);
        std::size_t p = (bbat->blockSize < maxlen - bytes) ? bbat->blockSize : maxlen - bytes;
        if (pos + p > filesize) p = filesize - pos;
        // FSTREAM file.seekg( pos );
        // FSTREAM file.read( (char*)data + bytes, p );
        memcpy(reinterpret_cast<char *>(data) + bytes, filedata + pos, p);
        bytes += p;
    }

    return bytes;
}

std::size_t StorageIO::loadBigBlock(std::size_t block, std::uint8_t *data, std::size_t maxlen)
{
    // sentinel
    if (!data) return 0;

    // wraps call for loadBigBlocks
    std::vector<std::size_t> blocks;
    blocks.resize(1);
    blocks[0] = block;

    return loadBigBlocks(blocks, data, maxlen);
}

// return number of bytes which has been read
std::size_t StorageIO::loadSmallBlocks(std::vector<std::size_t> blocks, std::uint8_t *data, std::size_t maxlen)
{
    // sentinel
    if (!data) return 0;
    if (blocks.size() < 1) return 0;
    if (maxlen == 0) return 0;

    // our own local buffer
    std::uint8_t *buf = new std::uint8_t[bbat->blockSize];

    // read small block one by one
    std::size_t bytes = 0;
    for (std::size_t i = 0; (i < blocks.size()) && (bytes < maxlen); i++)
    {
        std::size_t block = blocks[i];

        // find where the small-block exactly is
        std::size_t pos = block * sbat->blockSize;
        std::size_t bbindex = pos / bbat->blockSize;
        if (bbindex >= sb_blocks.size()) break;

        loadBigBlock(sb_blocks[bbindex], buf, bbat->blockSize);

        // copy the data
        std::size_t offset = pos % bbat->blockSize;
        std::size_t p = (maxlen - bytes < bbat->blockSize - offset) ? maxlen - bytes : bbat->blockSize - offset;
        p = (sbat->blockSize < p) ? sbat->blockSize : p;
        memcpy(data + bytes, buf + offset, p);
        bytes += p;
    }

    delete[] buf;

    return bytes;
}

std::size_t StorageIO::loadSmallBlock(std::size_t block, std::uint8_t *data, std::size_t maxlen)
{
    // sentinel
    if (!data) return 0;

    // wraps call for loadSmallBlocks
    std::vector<std::size_t> blocks;
    blocks.resize(1);
    blocks.assign(1, block);

    return loadSmallBlocks(blocks, data, maxlen);
}

// =========== StreamIO ==========

StreamIO::StreamIO(StorageIO *s, DirEntry *e)
    : io(s),
      entry(e),
      fullName(),
      eof(false),
      fail(false),
      blocks(),
      m_pos(0),
      cache_data(0),
      cache_size(4096), // optimal ?
      cache_pos(0)
{
    if (entry->size >= io->header->threshold)
        blocks = io->bbat->follow(entry->start);
    else
        blocks = io->sbat->follow(entry->start);

    // prepare cache
    cache_data = new std::uint8_t[cache_size];
    updateCache();
}

// FIXME tell parent we're gone
StreamIO::~StreamIO()
{
    delete[] cache_data;
}

void StreamIO::seek(std::size_t pos)
{
    m_pos = pos;
}

std::size_t StreamIO::tell()
{
    return m_pos;
}

int StreamIO::getch()
{
    // past end-of-file ?
    if (m_pos > entry->size) return -1;

    // need to update cache ?
    if (!cache_size || (m_pos < cache_pos) || (m_pos >= cache_pos + cache_size)) updateCache();

    // something bad if we don't get good cache
    if (!cache_size) return -1;

    int data = cache_data[m_pos - cache_pos];
    m_pos++;

    return data;
}

std::size_t StreamIO::read(std::size_t pos, std::uint8_t *data, std::size_t maxlen)
{
    // sanity checks
    if (!data) return 0;
    if (maxlen == 0) return 0;

    std::size_t totalbytes = 0;

    if (entry->size < io->header->threshold)
    {
        // small file
        std::size_t index = pos / io->sbat->blockSize;

        if (index >= blocks.size()) return 0;

        std::uint8_t *buf = new std::uint8_t[io->sbat->blockSize];
        std::size_t offset = pos % io->sbat->blockSize;
        while (totalbytes < maxlen)
        {
            if (index >= blocks.size()) break;
            io->loadSmallBlock(blocks[index], buf, io->bbat->blockSize);
            std::size_t count = io->sbat->blockSize - offset;
            if (count > maxlen - totalbytes) count = maxlen - totalbytes;
            memcpy(data + totalbytes, buf + offset, count);
            totalbytes += count;
            offset = 0;
            index++;
        }
        delete[] buf;
    }
    else
    {
        // big file
        std::size_t index = pos / io->bbat->blockSize;

        if (index >= blocks.size()) return 0;

        std::uint8_t *buf = new std::uint8_t[io->bbat->blockSize];
        std::size_t offset = pos % io->bbat->blockSize;
        while (totalbytes < maxlen)
        {
            if (index >= blocks.size()) break;
            io->loadBigBlock(blocks[index], buf, io->bbat->blockSize);
            std::size_t count = io->bbat->blockSize - offset;
            if (count > maxlen - totalbytes) count = maxlen - totalbytes;
            memcpy(data + totalbytes, buf + offset, count);
            totalbytes += count;
            index++;
            offset = 0;
        }
        delete[] buf;
    }

    return totalbytes;
}

std::size_t StreamIO::read(std::uint8_t *data, std::size_t maxlen)
{
    std::size_t bytes = read(tell(), data, maxlen);
    m_pos += bytes;
    return bytes;
}

void StreamIO::updateCache()
{
    // sanity check
    if (!cache_data) return;

    cache_pos = m_pos - (m_pos % cache_size);
    std::size_t bytes = cache_size;
    if (cache_pos + bytes > entry->size) bytes = entry->size - cache_pos;
    cache_size = read(cache_pos, cache_data, bytes);
}

// =========== Storage ==========

Storage::Storage(char *bytes, std::size_t length)
    : io(new StorageIO(this, bytes, length))
{
}

Storage::~Storage()
{
    delete io;
}

int Storage::result()
{
    return io->result;
}

bool Storage::open()
{
    return io->open();
}

void Storage::close()
{
    io->close();
}

std::list<std::string> Storage::entries(const std::string &path)
{
    std::list<std::string> result;
    DirTree *dt = io->dirtree;
    DirEntry *e = dt->entry(path, false);
    if (e && e->dir)
    {
        auto parent = dt->indexOf(e);
        std::vector<std::size_t> children = dt->children(static_cast<std::size_t>(parent));
        for (std::size_t i = 0; i < children.size(); i++)
            result.push_back(dt->entry(children[i])->name);
    }

    return result;
}

bool Storage::isDirectory(const std::string &name)
{
    DirEntry *e = io->dirtree->entry(name, false);
    return e ? e->dir : false;
}

DirTree *Storage::dirTree()
{
    return io->dirtree;
}

StorageIO *Storage::storageIO()
{
    return io;
}

std::list<DirEntry *> Storage::dirEntries(const std::string &path)
{
    std::list<DirEntry *> result;
    DirTree *dt = io->dirtree;
    DirEntry *e = dt->entry(path, false);
    if (e && e->dir)
    {
        auto parent = dt->indexOf(e);
        std::vector<std::size_t> children = dt->children(static_cast<std::size_t>(parent));
        for (std::size_t i = 0; i < children.size(); i++)
            result.push_back(dt->entry(children[i]));
    }

    return result;
}

// =========== Stream ==========

Stream::Stream(Storage *storage, const std::string &name)
    : io(storage->io->streamIO(name))
{
}

// FIXME tell parent we're gone
Stream::~Stream()
{
    delete io;
}

std::string Stream::fullName()
{
    return io ? io->fullName : std::string();
}

std::size_t Stream::tell()
{
    return io ? io->tell() : 0;
}

void Stream::seek(std::size_t newpos)
{
    if (io) io->seek(newpos);
}

std::size_t Stream::size()
{
    return io ? io->entry->size : 0;
}

int Stream::getch()
{
    return io ? io->getch() : 0;
}

std::size_t Stream::read(std::uint8_t *data, std::size_t maxlen)
{
    return io ? io->read(data, maxlen) : 0;
}

bool Stream::eof()
{
    return io ? io->eof : false;
}

bool Stream::fail()
{
    return io ? io->fail : true;
}
