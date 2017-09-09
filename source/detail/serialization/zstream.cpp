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
#include <array>
#include <cassert>
#include <cstring>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <iterator> // for std::back_inserter
#include <stdexcept>
#include <string>

#include <xlnt/utils/exceptions.hpp>
#include <detail/serialization/miniz.hpp>
#include <detail/serialization/vector_streambuf.hpp>
#include <detail/serialization/zstream.hpp>

namespace {

template <class T>
T read_int(std::istream &stream)
{
    T value;
    stream.read(reinterpret_cast<char *>(&value), sizeof(T));

    return value;
}

template <class T>
void write_int(std::ostream &stream, T value)
{
    stream.write(reinterpret_cast<char *>(&value), sizeof(T));
}

xlnt::detail::zheader read_header(std::istream &istream, const bool global)
{
    xlnt::detail::zheader header;

    auto sig = read_int<std::uint32_t>(istream);

    // read and check for local/global magic
    if (global)
    {
        if (sig != 0x02014b50)
        {
            throw xlnt::exception("missing global header signature");
        }

        header.version = read_int<std::uint16_t>(istream);
    }
    else if (sig != 0x04034b50)
    {
        throw xlnt::exception("missing local header signature");
    }

    // Read rest of header
    header.version = read_int<std::uint16_t>(istream);
    header.flags = read_int<std::uint16_t>(istream);
    header.compression_type = read_int<std::uint16_t>(istream);
    header.stamp_date = read_int<std::uint16_t>(istream);
    header.stamp_time = read_int<std::uint16_t>(istream);
    header.crc = read_int<std::uint32_t>(istream);
    header.compressed_size = read_int<std::uint32_t>(istream);
    header.uncompressed_size = read_int<std::uint32_t>(istream);

    auto filename_length = read_int<std::uint16_t>(istream);
    auto extra_length = read_int<std::uint16_t>(istream);

    std::uint16_t comment_length = 0;

    if (global)
    {
        comment_length = read_int<std::uint16_t>(istream);
        /*std::uint16_t disk_number_start = */ read_int<std::uint16_t>(istream);
        /*std::uint16_t int_file_attrib = */ read_int<std::uint16_t>(istream);
        /*std::uint32_t ext_file_attrib = */ read_int<std::uint32_t>(istream);
        header.header_offset = read_int<std::uint32_t>(istream);
    }

    header.filename.resize(filename_length, '\0');
    istream.read(&header.filename[0], filename_length);

    header.extra.resize(extra_length, 0);
    istream.read(reinterpret_cast<char *>(header.extra.data()), extra_length);

    if (global)
    {
        header.comment.resize(comment_length, '\0');
        istream.read(&header.comment[0], comment_length);
    }

    return header;
}

void write_header(const xlnt::detail::zheader &header, std::ostream &ostream, const bool global)
{
    if (global)
    {
        write_int(ostream, static_cast<std::uint32_t>(0x02014b50)); // header sig
        write_int(ostream, static_cast<std::uint16_t>(20)); // version made by
    }
    else
    {
        write_int(ostream, static_cast<std::uint32_t>(0x04034b50));
    }

    write_int(ostream, header.version);
    write_int(ostream, header.flags);
    write_int(ostream, header.compression_type);
    write_int(ostream, header.stamp_date);
    write_int(ostream, header.stamp_time);
    write_int(ostream, header.crc);
    write_int(ostream, header.compressed_size);
    write_int(ostream, header.uncompressed_size);
    write_int(ostream, static_cast<std::uint16_t>(header.filename.length()));
    write_int(ostream, static_cast<std::uint16_t>(0)); // extra lengthx

    if (global)
    {
        write_int(ostream, static_cast<std::uint16_t>(0)); // filecomment
        write_int(ostream, static_cast<std::uint16_t>(0)); // disk# start
        write_int(ostream, static_cast<std::uint16_t>(0)); // internal file
        write_int(ostream, static_cast<std::uint32_t>(0)); // ext final
        write_int(ostream, static_cast<std::uint32_t>(header.header_offset)); // rel offset
    }

    for (auto c : header.filename)
    {
        write_int(ostream, c);
    }
}

} // namespace

namespace xlnt {
namespace detail {

static const std::size_t buffer_size = 512;

class zip_streambuf_decompress : public std::streambuf
{
    std::istream &istream;

    z_stream strm;
    std::array<char, buffer_size> in;
    std::array<char, buffer_size> out;
    zheader header;
    std::size_t total_read;
    std::size_t total_uncompressed;
    bool valid;
    bool compressed_data;

    static const unsigned short DEFLATE = 8;
    static const unsigned short UNCOMPRESSED = 0;

public:
    zip_streambuf_decompress(std::istream &stream, zheader central_header)
        : istream(stream), header(central_header), total_read(0), total_uncompressed(0), valid(true)
    {
        in.fill(0);
        out.fill(0);

        strm.zalloc = nullptr;
        strm.zfree = nullptr;
        strm.opaque = nullptr;
        strm.avail_in = 0;
        strm.next_in = nullptr;

        setg(in.data(), in.data(), in.data());
        setp(nullptr, nullptr);

        // skip the header
        read_header(istream, false);

        if (header.compression_type == DEFLATE)
        {
            compressed_data = true;
        }
        else if (header.compression_type == UNCOMPRESSED)
        {
            compressed_data = false;
        }
        else
        {
            compressed_data = false;
            throw xlnt::exception("unsupported compression type, should be DEFLATE or uncompressed");
        }

        // initialize the inflate
        if (compressed_data && valid)
        {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wold-style-cast"
            int result = inflateInit2(&strm, -MAX_WBITS);
#pragma clang diagnostic pop

            if (result != Z_OK)
            {
                throw xlnt::exception("couldn't inflate ZIP, possibly corrupted");
            }
        }

        header = central_header;
    }

    virtual ~zip_streambuf_decompress()
    {
        if (compressed_data && valid)
        {
            inflateEnd(&strm);
        }
    }

    int process()
    {
        if (!valid) return -1;

        if (compressed_data)
        {
            strm.avail_out = buffer_size - 4;
            strm.next_out = reinterpret_cast<Bytef *>(out.data() + 4);

            while (strm.avail_out != 0)
            {
                if (strm.avail_in == 0)
                {
                    // buffer empty, read some more from file
                    istream.read(in.data(),
                        static_cast<std::streamsize>(std::min(buffer_size, header.compressed_size - total_read)));
                    strm.avail_in = static_cast<unsigned int>(istream.gcount());
                    total_read += strm.avail_in;
                    strm.next_in = reinterpret_cast<Bytef *>(in.data());
                }

                const auto ret = inflate(&strm, Z_NO_FLUSH); // decompress

                if (ret == Z_STREAM_ERROR || ret == Z_NEED_DICT || ret == Z_DATA_ERROR || ret == Z_MEM_ERROR)
                {
                    throw xlnt::exception("couldn't inflate ZIP, possibly corrupted");
                }

                if (ret == Z_STREAM_END) break;
            }

            auto unzip_count = buffer_size - strm.avail_out - 4;
            total_uncompressed += unzip_count;
            return static_cast<int>(unzip_count);
        }

        // uncompressed, so just read
        istream.read(out.data() + 4,
            static_cast<std::streamsize>(std::min(buffer_size - 4, header.uncompressed_size - total_read)));
        auto count = istream.gcount();
        total_read += static_cast<std::size_t>(count);
        return static_cast<int>(count);
    }

    virtual int underflow()
    {
        if (gptr() && (gptr() < egptr()))
            return traits_type::to_int_type(*gptr()); // if we already have data just use it
        auto put_back_count = gptr() - eback();
        if (put_back_count > 4) put_back_count = 4;
        std::memmove(
            out.data() + (4 - put_back_count), gptr() - put_back_count, static_cast<std::size_t>(put_back_count));
        int num = process();
        setg(out.data() + 4 - put_back_count, out.data() + 4, out.data() + 4 + num);
        if (num <= 0) return EOF;
        return traits_type::to_int_type(*gptr());
    }

    virtual int overflow(int c = EOF);
};

int zip_streambuf_decompress::overflow(int)
{
    throw xlnt::exception("writing to read-only buffer");
}

class zip_streambuf_compress : public std::streambuf
{
    std::ostream &ostream; // owned when header==0 (when not part of zip file)

    z_stream strm;
    std::array<char, buffer_size> in;
    std::array<char, buffer_size> out;

    zheader *header;
    std::uint32_t uncompressed_size;
    std::uint32_t crc;

    bool valid;

public:
    zip_streambuf_compress(zheader *central_header, std::ostream &stream)
        : ostream(stream), header(central_header), valid(true)
    {
        strm.zalloc = nullptr;
        strm.zfree = nullptr;
        strm.opaque = nullptr;

#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wold-style-cast"
        int ret = deflateInit2(&strm, Z_DEFAULT_COMPRESSION, Z_DEFLATED, -MAX_WBITS, 8, Z_DEFAULT_STRATEGY);
#pragma clang diagnostic pop

        if (ret != Z_OK)
        {
            std::cerr << "libz: failed to deflateInit" << std::endl;
            valid = false;
            return;
        }

        setg(nullptr, nullptr, nullptr);
        setp(in.data(), in.data() + buffer_size - 4); // we want to be 4 aligned

        // Write appropriate header
        if (header)
        {
            header->header_offset = static_cast<std::uint32_t>(stream.tellp());
            write_header(*header, ostream, false);
        }

        uncompressed_size = crc = 0;
    }

    virtual ~zip_streambuf_compress()
    {
        if (valid)
        {
            process(true);
            deflateEnd(&strm);
            if (header)
            {
                std::ios::streampos final_position = ostream.tellp();
                header->uncompressed_size = uncompressed_size;
                header->crc = crc;
                ostream.seekp(header->header_offset);
                write_header(*header, ostream, false);
                ostream.seekp(final_position);
            }
            else
            {
                write_int(ostream, crc);
                write_int(ostream, uncompressed_size);
            }
        }
        if (!header) delete &ostream;
    }

protected:
    int process(bool flush)
    {
        if (!valid) return -1;

        strm.next_in = reinterpret_cast<Bytef *>(pbase());
        strm.avail_in = static_cast<unsigned int>(pptr() - pbase());

        while (strm.avail_in != 0 || flush)
        {
            strm.avail_out = buffer_size;
            strm.next_out = reinterpret_cast<Bytef *>(out.data());

            int ret = deflate(&strm, flush ? Z_FINISH : Z_NO_FLUSH);

            if (!(ret != Z_BUF_ERROR && ret != Z_STREAM_ERROR))
            {
                valid = false;
                std::cerr << "gzip: gzip error " << strm.msg << std::endl;
                ;
                return -1;
            }

            auto generated_output = static_cast<int>(strm.next_out - reinterpret_cast<std::uint8_t *>(out.data()));
            ostream.write(out.data(), generated_output);
            if (header) header->compressed_size += static_cast<std::uint32_t>(generated_output);
            if (ret == Z_STREAM_END) break;
        }

        // update counts, crc's and buffers
        auto consumed_input = static_cast<std::uint32_t>(pptr() - pbase());
        uncompressed_size += consumed_input;
        crc = static_cast<std::uint32_t>(crc32(crc, reinterpret_cast<Bytef *>(in.data()), consumed_input));
        setp(pbase(), pbase() + buffer_size - 4);

        return 1;
    }

    virtual int sync()
    {
        if (pptr() && pptr() > pbase()) return process(false);
        return 0;
    }

    virtual int underflow()
    {
        throw xlnt::exception("Attempt to read write only ostream");
    }

    virtual int overflow(int c = EOF);
};

int zip_streambuf_compress::overflow(int c)
{
    if (c != EOF)
    {
        *pptr() = static_cast<char>(c);
        pbump(1);
    }
    if (process(false) == EOF) return EOF;
    return c;
}

ozstream::ozstream(std::ostream &stream)
    : destination_stream_(stream)
{
    if (!destination_stream_)
    {
        throw xlnt::exception("bad zip stream");
    }
}

ozstream::~ozstream()
{
    // Write all file headers
    std::ios::streampos final_position = destination_stream_.tellp();

    for (const auto &header : file_headers_)
    {
        write_header(header, destination_stream_, true);
    }

    std::ios::streampos central_end = destination_stream_.tellp();

    // Write end of central
    write_int(destination_stream_, static_cast<std::uint32_t>(0x06054b50)); // end of central
    write_int(destination_stream_, static_cast<std::uint16_t>(0)); // this disk number
    write_int(destination_stream_, static_cast<std::uint16_t>(0)); // this disk number
    write_int(destination_stream_, static_cast<std::uint16_t>(file_headers_.size())); // one entry in center in this disk
    write_int(destination_stream_, static_cast<std::uint16_t>(file_headers_.size())); // one entry in center
    write_int(destination_stream_, static_cast<std::uint32_t>(central_end - final_position)); // size of header
    write_int(destination_stream_, static_cast<std::uint32_t>(final_position)); // offset to header
    write_int(destination_stream_, static_cast<std::uint16_t>(0)); // zip comment
}

std::unique_ptr<std::streambuf> ozstream::open(const path &filename)
{
    zheader header;
    header.filename = filename.string();
    file_headers_.push_back(header);
    auto buffer = new zip_streambuf_compress(&file_headers_.back(), destination_stream_);

    return std::unique_ptr<zip_streambuf_compress>(buffer);
}

izstream::izstream(std::istream &stream)
    : source_stream_(stream)
{
    if (!stream)
    {
        throw xlnt::exception("Invalid file handle");
    }

    read_central_header();
}

izstream::~izstream()
{
}

bool izstream::read_central_header()
{
    // Find the header
    // NOTE: this assumes the zip file header is the last thing written to file...
    source_stream_.seekg(0, std::ios_base::end);
    std::ios::streampos end_position = source_stream_.tellg();

    auto max_comment_size = std::uint32_t(0xffff); // max size of header
    auto read_size_before_comment = std::uint32_t(22);

    std::ios::streamoff read_start = max_comment_size + read_size_before_comment;

    if (read_start > end_position)
    {
        read_start = end_position;
    }

    source_stream_.seekg(end_position - read_start);
    std::vector<std::uint8_t> buf(static_cast<std::size_t>(read_start), '\0');

    if (read_start <= 0)
    {
        throw xlnt::exception("file is empty");
    }

    source_stream_.read(reinterpret_cast<char *>(buf.data()), read_start);

    if (buf[0] == 0xd0 && buf[1] == 0xcf && buf[2] == 0x11 && buf[3] == 0xe0
        && buf[4] == 0xa1 && buf[5] == 0xb1 && buf[6] == 0x1a && buf[7] == 0xe1)
    {
        throw xlnt::exception("encrypted xlsx, password required");
    }

    auto found_header = false;
    std::size_t header_index = 0;

    for (std::size_t i = 0; i < static_cast<std::size_t>(read_start - 3); ++i)
    {
        if (buf[i] == 0x50 && buf[i + 1] == 0x4b && buf[i + 2] == 0x05 && buf[i + 3] == 0x06)
        {
            found_header = true;
            header_index = i;
            break;
        }
    }

    if (!found_header)
    {
        throw xlnt::exception("failed to find zip header");
    }

    // seek to end of central header and read
    source_stream_.seekg(end_position - (read_start - static_cast<std::ptrdiff_t>(header_index)));

    /*auto word = */ read_int<std::uint32_t>(source_stream_);
    auto disk_number1 = read_int<std::uint16_t>(source_stream_);
    auto disk_number2 = read_int<std::uint16_t>(source_stream_);

    if (disk_number1 != disk_number2 || disk_number1 != 0)
    {
        throw xlnt::exception("multiple disk zip files are not supported");
    }

    auto num_files = read_int<std::uint16_t>(source_stream_); // one entry in center in this disk
    auto num_files_this_disk = read_int<std::uint16_t>(source_stream_); // one entry in center

    if (num_files != num_files_this_disk)
    {
        throw xlnt::exception("multi disk zip files are not supported");
    }

    /*auto size_of_header = */ read_int<std::uint32_t>(source_stream_); // size of header
    auto header_offset = read_int<std::uint32_t>(source_stream_); // offset to header

    // go to header and read all file headers
    source_stream_.seekg(header_offset);

    for (std::uint16_t i = 0; i < num_files; ++i)
    {
        auto header = read_header(source_stream_, true);
        file_headers_[header.filename] = header;
    }

    return true;
}

std::unique_ptr<std::streambuf> izstream::open(const path &filename) const
{
    if (!has_file(filename))
    {
        throw xlnt::exception("file not found");
    }

    auto header = file_headers_.at(filename.string());
    source_stream_.seekg(header.header_offset);
    auto buffer = new zip_streambuf_decompress(source_stream_, header);

    return std::unique_ptr<zip_streambuf_decompress>(buffer);
}

std::string izstream::read(const path &filename) const
{
    auto buffer = open(filename);
    std::istream stream(buffer.get());
    auto bytes = to_vector(stream);

    return std::string(bytes.begin(), bytes.end());
}

std::vector<path> izstream::files() const
{
    std::vector<path> filenames;
    std::transform(file_headers_.begin(), file_headers_.end(), std::back_inserter(filenames),
        [](const std::pair<std::string, zheader> &h) { return path(h.first); });

    return filenames;
}

bool izstream::has_file(const path &filename) const
{
    return file_headers_.count(filename.string()) != 0;
}

} // namespace detail
} // namespace xlnt
