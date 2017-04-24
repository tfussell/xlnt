// Copyright (c) 2014-2017 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include <cstdint>
#include <cstring>
#include <vector>

#include <xlnt/utils/exceptions.hpp>

namespace xlnt {
namespace detail {

using byte = std::uint8_t;

class binary_reader
{
public:
    binary_reader() = delete;

    binary_reader(const std::vector<byte> &bytes)
        : bytes_(&bytes)
    {
    }

    binary_reader(const binary_reader &other) = default;

    binary_reader &operator=(const binary_reader &other)
    {
        offset_ = other.offset_;
        bytes_ = other.bytes_;

        return *this;
    }

    ~binary_reader()
    {
    }

    const std::vector<std::uint8_t> &data() const
    {
        return *bytes_;
    }

    void offset(std::size_t offset)
    {
        offset_ = offset;
    }

    std::size_t offset() const
    {
        return offset_;
    }

    void reset()
    {
        offset_ = 0;
    }

    template<typename T>
    T read()
    {
        return read_reference<T>();
    }

    template<typename T>
    const T &read_reference()
    {
        const auto &result = *reinterpret_cast<const T *>(bytes_->data() + offset_);
        offset_ += sizeof(T);

        return result;
    }

    template<typename T>
    std::vector<T> to_vector() const
    {
        auto result = std::vector<T>(size() / sizeof(T), T());
        std::memcpy(result.data(), bytes_->data(), size());

        return result;
    }

    template<typename T>
    std::vector<T> read_vector(std::size_t count)
    {
        auto result = std::vector<T>(count, T());
        std::memcpy(result.data(), bytes_->data() + offset_, count * sizeof(T));
        offset_ += count * sizeof(T);

        return result;
    }

    std::size_t size() const
    {
        return bytes_->size();
    }

private:
    std::size_t offset_ = 0;
    const std::vector<std::uint8_t> *bytes_;
};

class binary_writer
{
public:
    binary_writer(std::vector<byte> &bytes)
        : bytes_(&bytes)
    {
    }

    binary_writer(const binary_writer &other)
    {
        *this = other;
    }

    ~binary_writer()
    {
    }

    binary_writer &operator=(const binary_writer &other)
    {
        bytes_ = other.bytes_;
        offset_ = other.offset_;

        return *this;
    }

    std::vector<byte> &data()
    {
        return *bytes_;
    }
    
    template<typename T>
    void assign(const std::vector<T> &ints)
    {
        resize(ints.size() * sizeof(T));
        std::memcpy(bytes_->data(), ints.data(), bytes_->size());
    }

    template<typename T>
    void assign(const std::basic_string<T> &string)
    {
        resize(string.size() * sizeof(T));
        std::memcpy(bytes_->data(), string.data(), bytes_->size());
    }

    void offset(std::size_t new_offset)
    {
        offset_ = new_offset;
    }

    std::size_t offset() const
    {
        return offset_;
    }

    void reset()
    {
        offset_ = 0;
        bytes_->clear();
    }

    template<typename T>
    void write(T value)
    {
        const auto num_bytes = sizeof(T);

        if (offset() + num_bytes > size())
        {
            extend(offset() + num_bytes - size());
        }

        std::memcpy(bytes_->data() + offset(), &value, num_bytes);
        offset_ += num_bytes;
    }

    std::size_t size() const
    {
        return bytes_->size();
    }

    void resize(std::size_t new_size, byte fill = 0)
    {
        bytes_->resize(new_size, fill);
    }

    void extend(std::size_t amount, byte fill = 0)
    {
        bytes_->resize(size() + amount, fill);
    }

    std::vector<byte>::iterator iterator()
    {
        return bytes_->begin() + static_cast<std::ptrdiff_t>(offset());
    }

    /*
    void append(const byte *data, const std::size_t data_size, std::size_t offset, std::size_t count)
    {
        if (offset + count > data_size)
        {
            throw xlnt::exception("out of bounds read");
        }

        const auto end_index = size();
        extend(count);
        std::memcpy(bytes_->data() + end_index, data + offset, count);
    }
    */

    template<typename T>
    void append(const std::vector<T> &data)
    {
        append(data, 0, data.size());
    }

    template<typename T>
    void append(const std::vector<T> &data, std::size_t offset, std::size_t count)
    {
        const auto byte_count = count * sizeof(T);

        if (offset_ + byte_count > size())
        {
            extend(offset_ + byte_count - size());
        }

        std::memcpy(bytes_->data() + offset_, data.data() + offset, byte_count);
        offset_ += byte_count;
    }

private:
    std::vector<byte> *bytes_;
    std::size_t offset_ = 0;
};

template<typename T>
std::vector<byte> string_to_bytes(const std::basic_string<T> &string)
{
    std::vector<byte> bytes;
    binary_writer writer(bytes);
    writer.assign(string);
    
    return bytes;
}

} // namespace detail
} // namespace xlnt
