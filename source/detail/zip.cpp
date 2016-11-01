/*
PARTIO SOFTWARE
Copyright 2010 Disney Enterprises, Inc. All rights reserved

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

* Redistributions of source code must retain the above copyright
notice, this list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright
notice, this list of conditions and the following disclaimer in
the documentation and/or other materials provided with the
distribution.

* The names "Disney", "Walt Disney Pictures", "Walt Disney Animation
Studios" or the names of its contributors may NOT be used to
endorse or promote products derived from this software without
specific prior written permission from Walt Disney Pictures.

Disclaimer: THIS SOFTWARE IS PROVIDED BY WALT DISNEY PICTURES AND
CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE, NONINFRINGEMENT AND TITLE ARE DISCLAIMED.
IN NO EVENT SHALL WALT DISNEY PICTURES, THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND BASED ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
*/

#include <algorithm>
#include <cassert>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <iterator> // for std::back_inserter
#include <stdexcept>
#include <cstring>
#include <string>

extern "C" {
#include <zlib.h>
}

#include <detail/zip.hpp>

namespace xlnt {
namespace detail {

template<class T>
inline T read_int(std::istream &stream)
{
    T value;
    stream.read(reinterpret_cast<char *>(&value), sizeof(T));
    return value;
}

template<class T>
inline void write_int(std::ostream &stream, T value)
{
    stream.write(reinterpret_cast<char *>(&value), sizeof(T));
}

struct zip_file_header
{
    std::uint16_t version = 20;
    std::uint16_t flags = 0;
    std::uint16_t compression_type = 8;
    std::uint16_t stamp_date,stamp_time = 0;
    std::uint32_t crc = 0;
    std::uint32_t compressed_size = 0;
    std::uint32_t uncompressed_size = 0;
    std::string filename;
    std::string comment;
    std::vector<std::uint8_t> extra;
    std::uint32_t header_offset = 0; // local header offset

    zip_file_header()
    {
    }

    bool read(std::istream& istream,const bool global)
    {
        auto sig = read_int<std::uint32_t>(istream);

        // read and check for local/global magic
        if(global)
        {
            if(sig!=0x02014b50)
            {
                std::cerr<<"Did not find global header signature"<<std::endl;
                return false;
            }

            version = read_int<std::uint16_t>(istream);
        }
        else if(sig!=0x04034b50)
        {
            std::cerr<<"Did not find local header signature"<<std::endl;
            return false;
        }

        // Read rest of header
        version = read_int<std::uint16_t>(istream);
        flags = read_int<std::uint16_t>(istream);
        compression_type = read_int<std::uint16_t>(istream);
        stamp_date = read_int<std::uint16_t>(istream);
        stamp_time = read_int<std::uint16_t>(istream);
        crc = read_int<std::uint32_t>(istream);
        compressed_size = read_int<std::uint32_t>(istream);
        uncompressed_size = read_int<std::uint32_t>(istream);

        auto filename_length = read_int<std::uint16_t>(istream);
        auto extra_length = read_int<std::uint16_t>(istream);

        std::uint16_t comment_length = 0;

        if(global)
        {
            comment_length = read_int<std::uint16_t>(istream);
            /*std::uint16_t disk_number_start = */read_int<std::uint16_t>(istream);
            /*std::uint16_t int_file_attrib = */read_int<std::uint16_t>(istream);
            /*std::uint32_t ext_file_attrib = */read_int<std::uint32_t>(istream);
            header_offset = read_int<std::uint32_t>(istream);
        }

        filename.resize(filename_length, '\0');
        istream.read(&filename[0], filename_length);
        
        extra.resize(extra_length, 0);
        istream.read(reinterpret_cast<char *>(extra.data()), extra_length);

        if (global)
        {
            comment.resize(comment_length, '\0');
            istream.read(&comment[0], comment_length);
        }

        return true;
    }

    void Write(std::ostream& ostream,const bool global) const
    {if(global){
        write_int(ostream,(unsigned int)0x02014b50); // header sig
        write_int(ostream,(unsigned short)00);} // version made by
    else write_int(ostream,(unsigned int)0x04034b50);
    write_int(ostream,version);
    write_int(ostream,flags);
    write_int(ostream,compression_type);
    write_int(ostream,stamp_date);
    write_int(ostream,stamp_time);
    write_int(ostream,crc);
    write_int(ostream,compressed_size);
    write_int(ostream,uncompressed_size);
    write_int(ostream,(unsigned short)filename.length());
    write_int(ostream,(unsigned short)0); // extra lengthx
    if(global){
        write_int(ostream,(unsigned short)0); // filecomment
        write_int(ostream,(unsigned short)0); // disk# start
        write_int(ostream,(unsigned short)0); // internal file
        write_int(ostream,(unsigned int)0); // ext final
        write_int(ostream,(unsigned int)header_offset);} // rel offset
    for(unsigned int i=0;i<filename.length();i++) write_int(ostream,filename.c_str()[i]);}
};

class ZipStreambufDecompress:public std::streambuf
{
    static const unsigned int buffer_size=512;
    std::istream& istream;

    z_stream strm;
    unsigned char in[buffer_size],out[buffer_size];
    zip_file_header header;
    int total_read,total_uncompressed;
    //bool own_istream;
    bool valid;
    bool compressed_data;

    static const unsigned short DEFLATE=8;
    static const unsigned short UNCOMPRESSED=0;
public:
    ZipStreambufDecompress(std::istream& stream,zip_file_header central_header)
        :istream(stream),header(central_header),total_read(0),total_uncompressed(0),valid(true)
    {
        strm.zalloc=Z_NULL;strm.zfree=Z_NULL;strm.opaque=Z_NULL;strm.avail_in=0;strm.next_in=Z_NULL;
        setg((char*)in,(char*)in,(char*)in);
        setp(0,0);
        // skip the header
        valid=header.read(istream,false);
        if(header.compression_type==DEFLATE) compressed_data=true;
        else if(header.compression_type==UNCOMPRESSED) compressed_data=false;
        else{
            compressed_data=false;std::cerr<<"ZIP: got unrecognized compressed data (Supported deflate/uncompressed)"<<std::endl;
            valid=false;}
        // initialize the inflate
        if(compressed_data && valid){
            int result=inflateInit2(&strm,-MAX_WBITS);
            if(result!=Z_OK){std::cerr<<"gzip: inflateInit2 did not return Z_OK"<<std::endl;valid=false;}}
		header = central_header;
    }

    virtual ~ZipStreambufDecompress()
    {if(compressed_data && valid) inflateEnd(&strm);}

    int process()
    {if(!valid) return -1;
    if(compressed_data){
        strm.avail_out=buffer_size-4;
        strm.next_out=(Bytef*)(out+4);
        while(strm.avail_out!=0){
            if(strm.avail_in==0){ // buffer empty, read some more from file
                istream.read((char*)in,std::min((unsigned int)buffer_size,header.compressed_size-total_read));
                strm.avail_in=istream.gcount();
                total_read+=strm.avail_in;
                strm.next_in=(Bytef*)in;}
            int ret=inflate(&strm,Z_NO_FLUSH); // decompress
            switch(ret){
                case Z_STREAM_ERROR: 
                    std::cerr<<"libz error Z_STREAM_ERROR"<<std::endl;
                    valid=false;return -1;
                case Z_NEED_DICT:
                case Z_DATA_ERROR:
                case Z_MEM_ERROR: 
                    std::cerr<<"gzip error "<<strm.msg<<std::endl;
                    valid=false;return -1;}
            if(ret==Z_STREAM_END) break;}
        int unzip_count=buffer_size-strm.avail_out-4;
        total_uncompressed+=unzip_count;
        return unzip_count;}
    else{ // uncompressed, so just read
        istream.read((char*)(out+4),std::min(buffer_size-4,header.uncompressed_size-total_read));
        int count=istream.gcount();
        total_read+=count;
        return count;}
    return 1;}

    virtual int underflow()
    {if(gptr() && (gptr()<egptr())) return traits_type::to_int_type(*gptr()); // if we already have data just use it
    int put_back_count=gptr()-eback();
    if(put_back_count>4) put_back_count=4;
    std::memmove(out+(4-put_back_count),gptr()-put_back_count,put_back_count);
    int num=process();
    setg((char*)(out+4-put_back_count),(char*)(out+4),(char*)(out+4+num));
    if(num<=0) return EOF;
    return traits_type::to_int_type(*gptr());}

    virtual int overflow(int c=EOF)
    {assert(false);return EOF;}
};

class ZipStreambufCompress:public std::streambuf
{
    static const int buffer_size=512;
    std::ostream& ostream; // owned when header==0 (when not part of zip file)

    z_stream strm;
    unsigned char in[buffer_size],out[buffer_size];

    zip_file_header* header;
    unsigned int uncompressed_size;
    unsigned int crc;

    bool valid;

public:
    ZipStreambufCompress(zip_file_header* central_header,std::ostream& stream)
        :ostream(stream),header(central_header),valid(true)
    {
        strm.zalloc=Z_NULL;strm.zfree=Z_NULL;strm.opaque=Z_NULL;
        int ret=deflateInit2(&strm,Z_DEFAULT_COMPRESSION,Z_DEFLATED,-MAX_WBITS,8,Z_DEFAULT_STRATEGY); 
        if(ret != Z_OK){std::cerr<<"libz: failed to deflateInit"<<std::endl;valid=false;return;}
        setg(0,0,0);
        setp((char*)in,(char*)(in+buffer_size-4)); // we want to be 4 aligned
        // Write appropriate header
        if(header){header->header_offset=stream.tellp();header->Write(ostream,false);}
        uncompressed_size=crc=0;
    }

    virtual ~ZipStreambufCompress()
    {if(valid){
        process(true);
        deflateEnd(&strm);
        if(header){
            std::ios::streampos final_position=ostream.tellp();
            header->uncompressed_size=uncompressed_size;
            header->crc=crc;
            ostream.seekp(header->header_offset);
            header->Write(ostream,false);
            ostream.seekp(final_position);}
        else{write_int(ostream,crc);write_int(ostream,uncompressed_size);}}
    if(!header) delete &ostream;}

protected:
    int process(bool flush)
    {if(!valid) return -1;
    strm.next_in=(Bytef*)pbase();
    strm.avail_in=pptr()-pbase();
    while(strm.avail_in!=0 || flush){
        strm.avail_out=buffer_size;
        strm.next_out=(Bytef*)out;
        int ret=deflate(&strm,flush?Z_FINISH:Z_NO_FLUSH);
        if(!(ret!=Z_BUF_ERROR && ret!=Z_STREAM_ERROR)){
            valid=false;
            std::cerr<<"gzip: gzip error "<<strm.msg<<std::endl;;
            return -1;}
        int generated_output=strm.next_out-(Bytef*)out;
        ostream.write((char*)out,generated_output);
        if(header) header->compressed_size+=generated_output;
        if(ret==Z_STREAM_END) break;}
    // update counts, crc's and buffers
    int consumed_input=pptr()-pbase();
    uncompressed_size+=consumed_input;
    crc=crc32(crc,(Bytef*)in,consumed_input);
    setp(pbase(),pbase()+buffer_size-4);return 1;}

    virtual int sync()
    {if(pptr() && pptr()>pbase()) return process(false);return 0;}

    virtual int underflow()
    {std::runtime_error("Attempt to read write only ostream");return 0;}

    virtual int overflow(int c=EOF)
    {if(c!=EOF){*pptr()=c;pbump(1);}
    if(process(false)==EOF) return EOF;
    return c;}
};

class ZIP_FILE_ISTREAM:public std::istream
{
    ZipStreambufDecompress buf;

public:
    ZIP_FILE_ISTREAM(std::istream& istream,zip_file_header header)
        :std::istream(&buf),buf(istream,header)
    {
    }

    virtual ~ZIP_FILE_ISTREAM()
    {
    }
};

class ZIP_FILE_OSTREAM:public std::ostream
{
    ZipStreambufCompress buf;

public:
    ZIP_FILE_OSTREAM(zip_file_header* header,std::ostream& ostream)
        :std::ostream(&buf),buf(header,ostream)
    {
    }

    virtual ~ZIP_FILE_OSTREAM()
    {
    }
};

ZipFileWriter::ZipFileWriter(std::ostream& stream) : target_stream_(stream)
{
    if(!target_stream_) throw std::runtime_error("ZIP: Invalid file handle");
}

ZipFileWriter::~ZipFileWriter()
{
    // Write all file headers
    std::ios::streampos final_position=target_stream_.tellp();
    for(unsigned int i=0;i<file_headers_.size();i++)
    {
        file_headers_[i].Write(target_stream_,true);
    }
    std::ios::streampos central_end=target_stream_.tellp();
    // Write end of central
    write_int(target_stream_,(unsigned int)0x06054b50); // end of central
    write_int(target_stream_,(unsigned short)0); // this disk number
    write_int(target_stream_,(unsigned short)0); // this disk number
    write_int(target_stream_,(unsigned short)file_headers_.size()); // one entry in center in this disk
    write_int(target_stream_,(unsigned short)file_headers_.size()); // one entry in center
    write_int(target_stream_,(unsigned int)(central_end - final_position)); // size of header
    write_int(target_stream_,(unsigned int)final_position); // offset to header
    write_int(target_stream_,(unsigned short)0); // zip comment
}

std::ostream* ZipFileWriter::open(const std::string& filename)
{
    zip_file_header header;
    header.filename = filename;
    file_headers_.push_back(header);
    return new ZIP_FILE_OSTREAM(&file_headers_.back(),target_stream_);
}

ZipFileReader::ZipFileReader(std::istream &stream) : source_stream_(stream)
{
    if(!stream)
    {
        throw std::runtime_error("ZIP: Invalid file handle");
    }

    read_central_header();
}

ZipFileReader::~ZipFileReader()
{
}

bool ZipFileReader::read_central_header()
{
    // Find the header
    // NOTE: this assumes the zip file header is the last thing written to file...
    source_stream_.seekg(0,std::ios_base::end);
    std::ios::streampos end_position=source_stream_.tellg();

    auto max_comment_size = std::uint32_t(0xffff); // max size of header
    auto read_size_before_comment = std::uint32_t(22);

    std::ios::streamoff read_start=max_comment_size+read_size_before_comment;

    if(read_start>end_position) read_start=end_position;
    source_stream_.seekg(end_position-read_start);
    std::vector<char> buf(static_cast<std::size_t>(read_start), '\0');
    if(read_start<=0){std::cerr<<"ZIP: Invalid read buffer size"<<std::endl;return false;}
    source_stream_.read(buf.data(), read_start);

    auto found_header = false;
    std::size_t header_index = 0;

    for(std::size_t i = 0; i < read_start - 3; ++i)
    {
        if(buf[i] == 0x50 && buf[i + 1] == 0x4b && buf[i + 2] == 0x05 && buf[i + 3] == 0x06)
        {
            found_header = true;
            header_index = i;
            break;
        }
    }

    if(!found_header)
    {
        std::cerr<<"ZIP: Failed to find zip header"<<std::endl;
        return false;
    }

    // seek to end of central header and read
    source_stream_.seekg(end_position - (read_start - static_cast<std::ptrdiff_t>(header_index)));

    /*auto word = */read_int<std::uint32_t>(source_stream_);
    auto disk_number1 = read_int<std::uint16_t>(source_stream_);
    auto disk_number2 = read_int<std::uint16_t>(source_stream_);

    if(disk_number1 != disk_number2 || disk_number1 != 0)
    {
        std::cerr<<"ZIP: multiple disk zip files are not supported"<<std::endl;
        return false;
    }

    auto num_files = read_int<std::uint16_t>(source_stream_); // one entry in center in this disk
    auto num_files_this_disk = read_int<std::uint16_t>(source_stream_); // one entry in center

    if(num_files != num_files_this_disk)
    {
        std::cerr<<"ZIP: multi disk zip files are not supported"<<std::endl;
        return false;
    }

    /*auto size_of_header = */read_int<std::uint32_t>(source_stream_); // size of header
    auto header_offset = read_int<std::uint32_t>(source_stream_); // offset to header

    // go to header and read all file headers
    source_stream_.seekg(header_offset);

    for (std::uint16_t i = 0; i < num_files; ++i)
    {
        zip_file_header header;

        if (header.read(source_stream_, true))
        {
            file_headers_[header.filename] = header;
        }
    }

    return true;
}

std::istream &ZipFileReader::open(const std::string &filename)
{
    if (!has_file(filename))
    {
        throw "not found";
    }

    auto header = file_headers_.at(filename);
    source_stream_.seekg(header.header_offset);
    read_stream_.reset(new ZIP_FILE_ISTREAM(source_stream_, header));

    return *read_stream_;
}

std::vector<std::string> ZipFileReader::files() const
{
    std::vector<std::string> filenames;
    std::transform(file_headers_.begin(), file_headers_.end(), std::back_inserter(filenames),
        [](const std::pair<std::string, zip_file_header> &h) { return h.first; });
    return filenames;
}

bool ZipFileReader::has_file(const std::string &filename) const
{
	return file_headers_.count(filename) != 0;
}

} // namespace detail
} // namespace xlnt
