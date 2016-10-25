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

#ifndef POLE_H
#define POLE_H

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
        Storage( char* bytes, unsigned long length );
        
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
        unsigned long size();
        
        /**
         * Returns the current read/write position.
         **/
        unsigned long tell();
        
        /**
         * Sets the read/write position.
         **/
        void seek( unsigned long pos );
        
        /**
         * Reads a byte.
         **/
        int getch();
        
        /**
         * Reads a block of data.
         **/
        unsigned long read( unsigned char* data, unsigned long maxlen );
        
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
        unsigned char id[8];       // signature, or magic identifier
        unsigned b_shift;          // bbat->blockSize = 1 << b_shift
        unsigned s_shift;          // sbat->blockSize = 1 << s_shift
        unsigned num_bat;          // blocks allocated for big bat
        unsigned dirent_start;     // starting block for directory info
        unsigned threshold;        // switch from small to big file (usually 4K)
        unsigned sbat_start;       // starting block index to store small bat
        unsigned num_sbat;         // blocks allocated for small bat
        unsigned mbat_start;       // starting block to store meta bat
        unsigned num_mbat;         // blocks allocated for meta bat
        unsigned long bb_blocks[109];
        
        Header();
        bool valid();
        void load( const unsigned char* buffer );
        void save( unsigned char* buffer );
        void debug();
    };
    
    class AllocTable
    {
    public:
        static const unsigned Eof;
        static const unsigned Avail;
        static const unsigned Bat;
        static const unsigned MetaBat;
        unsigned blockSize;
        AllocTable();
        void clear();
        unsigned long count();
        void resize( unsigned long newsize );
        void preserve( unsigned long n );
        void set( unsigned long index, unsigned long val );
        unsigned unused();
        void setChain( std::vector<unsigned long> );
        std::vector<unsigned long> follow( unsigned long start );
        unsigned long operator[](unsigned long index );
        void load( const unsigned char* buffer, unsigned len );
        void save( unsigned char* buffer );
        unsigned size();
        void debug();
    private:
        std::vector<unsigned long> data;
        AllocTable( const AllocTable& );
        AllocTable& operator=( const AllocTable& );
    };
    
    class DirEntry
    {
    public:
        bool valid;            // false if invalid (should be skipped)
        std::string name;      // the name, not in unicode anymore
        bool dir;              // true if directory
        unsigned long size;    // size (not valid if directory)
        unsigned long start;   // starting block
        unsigned prev;         // previous sibling
        unsigned next;         // next sibling
        unsigned child;        // first child
        
        DirEntry(): valid(false), name(), dir(false), size(0), start(0),
        prev(0), next(0), child(0) {}
    };
    
    class DirTree
    {
    public:
        static const unsigned End;
        DirTree();
        void clear();
        unsigned entryCount();
        DirEntry* entry( unsigned index );
        DirEntry* entry( const std::string& name, bool create=false );
        int indexOf( DirEntry* e );
        int parent( unsigned index );
        std::string fullName( unsigned index );
        std::vector<unsigned> children( unsigned index );
        void load( unsigned char* buffer, unsigned len );
        void save( unsigned char* buffer );
        unsigned size();
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
        unsigned char *filedata;
        unsigned long dataLength;
        int result;               // result of operation
        bool opened;              // true if file is opened
        unsigned long filesize;   // size of the file
        
        Header* header;           // storage header
        DirTree* dirtree;         // directory tree
        AllocTable* bbat;         // allocation table for big blocks
        AllocTable* sbat;         // allocation table for small blocks
        
        std::vector<unsigned long> sb_blocks; // blocks for "small" files
        
        std::list<Stream*> streams;
        
        StorageIO( Storage* storage, char* bytes, unsigned long length );
        ~StorageIO();
        
        bool open();
        void close();
        void flush();
        void load();
        void create();
        
        unsigned long loadBigBlocks( std::vector<unsigned long> blocks, unsigned char* buffer, unsigned long maxlen );
        
        unsigned long loadBigBlock( unsigned long block, unsigned char* buffer, unsigned long maxlen );
        
        unsigned long loadSmallBlocks( std::vector<unsigned long> blocks, unsigned char* buffer, unsigned long maxlen );
        
        unsigned long loadSmallBlock( unsigned long block, unsigned char* buffer, unsigned long maxlen );
        
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
        unsigned long size();
        void seek( unsigned long pos );
        unsigned long tell();
        int getch();
        unsigned long read( unsigned char* data, unsigned long maxlen );
        unsigned long read( unsigned long pos, unsigned char* data, unsigned long maxlen );
        
        
    private:
        std::vector<unsigned long> blocks;
        
        // no copy or assign
        StreamIO( const StreamIO& );
        StreamIO& operator=( const StreamIO& );
        
        // pointer for read
        unsigned long m_pos;
        
        // simple cache system to speed-up getch()
        unsigned char* cache_data;
        unsigned long cache_size;
        unsigned long cache_pos;
        void updateCache();
    };
    
} // namespace POLE


#endif // POLE_H