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

#include <string>
#include <list>
#include <fstream>
#include <iostream>
#include <list>
#include <string>
#include <vector>

namespace POLE
{
    class StorageIO;
    class Stream;
    class StreamIO;
    class DirTree;
    class DirEntry;

    class Storage
    {
        friend class Stream;
        friend class StreamOut;
        
    public:
        
        // for Storage::result()
        enum { Ok, OpenFailed, NotOLE, BadOLE, UnknownError };
        
        /**
         * Constructs a storage with name filename.
         **/
        Storage( char* bytes, std::size_t length );
        
        /**
         * Destroys the storage.
         **/
        ~Storage();
        
        /**
         * Opens the storage. Returns true if no error occurs.
         **/
        bool open();
        
        /**
         * Closes the storage.
         **/
        void close();
        
        /**
         * Returns the error code of last operation.
         **/
        int result();
        
        /**
         * Finds all stream and directories in given path.
         **/
        std::list<std::string> entries( const std::string& path = "/" );
        
        /**
         * Returns true if specified entry name is a directory.
         */
        bool isDirectory( const std::string& name );
        
        DirTree* dirTree();
        
        StorageIO* storageIO();
        
        std::list<DirEntry *> dirEntries( const std::string& path = "/" );
        
        /**
         * Finds and returns a stream with the specified name.
         * If reuse is true, this function returns the already created stream
         * (if any). Otherwise it will create the stream.
         *
         * When errors occur, this function returns NULL.
         *
         * You do not need to delete the created stream, it will be handled
         * automatically.
         **/
        Stream* stream( const std::string& name, bool reuse = true );
        //Stream* stream( const std::string& name, int mode = Stream::ReadOnly, bool reuse = true );
        
    private:
        StorageIO* io;
        
        // no copy or assign
        Storage( const Storage& );
        Storage& operator=( const Storage& );
        
    };
    
    class Stream
    {
        friend class Storage;
        friend class StorageIO;
        
    public:
        
        /**
         * Creates a new stream.
         */
        // name must be absolute, e.g "/Workbook"
        Stream( Storage* storage, const std::string& name );
        
        /**
         * Destroys the stream.
         */
        ~Stream();
        
        /**
         * Returns the full stream name.
         */
        std::string fullName();
        
        /**
         * Returns the stream size.
         **/
        std::size_t size();
        
        /**
         * Returns the current read/write position.
         **/
        std::size_t tell();
        
        /**
         * Sets the read/write position.
         **/
        void seek( std::size_t pos );
        
        /**
         * Reads a byte.
         **/
        int getch();
        
        /**
         * Reads a block of data.
         **/
        std::size_t read( std::uint8_t* data, std::size_t maxlen );
        
        /**
         * Returns true if the read/write position is past the file.
         **/
        bool eof();
        
        /**
         * Returns true whenever error occurs.
         **/
        bool fail();
        
    private:
        StreamIO* io;
        
        // no copy or assign
        Stream( const Stream& );
        Stream& operator=( const Stream& );
    };
    
    class Header
    {
    public:
        std::uint8_t id[8];       // signature, or magic identifier
        std::uint16_t b_shift;          // bbat->blockSize = 1 << b_shift
        std::uint16_t s_shift;          // sbat->blockSize = 1 << s_shift
        std::uint32_t num_bat;          // blocks allocated for big bat
        std::uint32_t dirent_start;     // starting block for directory info
        std::uint32_t threshold;        // switch from small to big file (usually 4K)
        std::uint32_t sbat_start;       // starting block index to store small bat
        std::uint32_t num_sbat;         // blocks allocated for small bat
        std::uint32_t mbat_start;       // starting block to store meta bat
        std::uint32_t num_mbat;         // blocks allocated for meta bat
        std::uint32_t bb_blocks[109];
        
        Header();
        bool valid();
        void load( const std::uint8_t* buffer );
        void save( std::uint8_t* buffer );
        void debug();
    };
    
    class AllocTable
    {
    public:
        static const std::uint32_t Eof;
        static const std::uint32_t Avail;
        static const std::uint32_t Bat;
        static const std::uint32_t MetaBat;
        std::size_t blockSize;
        AllocTable();
        void clear();
        std::size_t count();
        void resize( std::size_t newsize );
        void preserve( std::size_t n );
        void set( std::size_t index, std::uint32_t val );
        std::size_t unused();
        void setChain( std::vector<std::uint32_t> );
        std::vector<std::size_t> follow( std::size_t start );
        std::size_t operator[](std::size_t index );
        void load( const std::uint8_t* buffer, std::size_t len );
        void save( std::uint8_t* buffer );
        std::size_t size();
        void debug();
    private:
        std::vector<std::uint32_t> data;
        AllocTable( const AllocTable& );
        AllocTable& operator=( const AllocTable& );
    };
    
    class DirEntry
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
        
        DirEntry(): valid(false), name(), dir(false), size(0), start(0),
        prev(0), next(0), child(0) {}
    };
    
    class DirTree
    {
    public:
        static const std::uint32_t End;
        DirTree();
        void clear();
        std::size_t entryCount();
        DirEntry* entry( std::size_t index );
        DirEntry* entry( const std::string& name, bool create=false );
        std::ptrdiff_t indexOf( DirEntry* e );
        std::ptrdiff_t parent( std::size_t index );
        std::string fullName( std::size_t index );
        std::vector<std::size_t> children( std::size_t index );
        void load( std::uint8_t* buffer, std::size_t len );
        void save( std::uint8_t* buffer );
        std::size_t size();
        void debug();
    private:
        std::vector<DirEntry> entries;
        DirTree( const DirTree& );
        DirTree& operator=( const DirTree& );
    };
    
    class StorageIO
    {
    public:
        Storage* storage;         // owner
        std::uint8_t *filedata;
        std::size_t dataLength;
        int result;               // result of operation
        bool opened;              // true if file is opened
        std::size_t filesize;   // size of the file
        
        Header* header;           // storage header
        DirTree* dirtree;         // directory tree
        AllocTable* bbat;         // allocation table for big blocks
        AllocTable* sbat;         // allocation table for small blocks
        
        std::vector<std::size_t> sb_blocks; // blocks for "small" files
        
        std::list<Stream*> streams;
        
        StorageIO( Storage* storage, char* bytes, std::size_t length );
        ~StorageIO();
        
        bool open();
        void close();
        void flush();
        void load();
        void create();
        
        std::size_t loadBigBlocks( std::vector<std::size_t> blocks, std::uint8_t* buffer, std::size_t maxlen );
        
        std::size_t loadBigBlock( std::size_t block, std::uint8_t* buffer, std::size_t maxlen );
        
        std::size_t loadSmallBlocks( std::vector<std::size_t> blocks, std::uint8_t* buffer, std::size_t maxlen );
        
        std::size_t loadSmallBlock( std::size_t block, std::uint8_t* buffer, std::size_t maxlen );
        
        StreamIO* streamIO( const std::string& name );
        
    private:
        // no copy or assign
        StorageIO( const StorageIO& );
        StorageIO& operator=( const StorageIO& );
        
    };
    
    class StreamIO
    {
    public:
        StorageIO* io;
        DirEntry* entry;
        std::string fullName;
        bool eof;
        bool fail;
        
        StreamIO( StorageIO* io, DirEntry* entry );
        ~StreamIO();
        std::size_t size();
        void seek( std::size_t pos );
        std::size_t tell();
        int getch();
        std::size_t read( std::uint8_t* data, std::size_t maxlen );
        std::size_t read( std::size_t pos, std::uint8_t* data, std::size_t maxlen );
        
        
    private:
        std::vector<std::size_t> blocks;
        
        // no copy or assign
        StreamIO( const StreamIO& );
        StreamIO& operator=( const StreamIO& );
        
        // pointer for read
        std::size_t m_pos;
        
        // simple cache system to speed-up getch()
        std::uint8_t* cache_data;
        std::size_t cache_size;
        std::size_t cache_pos;
        void updateCache();
    };
    
} // namespace POLE
