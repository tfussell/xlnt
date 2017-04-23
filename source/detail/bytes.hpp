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

namespace xlnt {
namespace detail {

using byte = std::uint8_t;

class byte_reader
{
public:
    byte_reader() = delete;

    byte_reader(const std::vector<byte> &bytes)
        : bytes_(&bytes)
    {
    }

    byte_reader &operator=(const byte_reader &other)
    {
        offset_ = other.offset_;
        bytes_ = other.bytes_;

        return *this;
    }

    ~byte_reader()
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
        T result;
        std::memcpy(&result, bytes_->data() + offset_, sizeof(T));
        offset_ += sizeof(T);

        return result;
    }

    template<typename T>
    std::vector<T> as_vector_of() const
    {
        auto result = std::vector<T>(size() / sizeof(T), 0);
        std::memcpy(result.data(), bytes_->data(), size());

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

class byte_vector
{
public:
    template<typename T>
    static byte_vector from(const std::vector<T> &ints)
    {
        byte_vector result;

        result.resize(ints.size() * sizeof(T));
        std::memcpy(result.bytes_.data(), ints.data(), result.bytes_.size());

        return result;
    }

    template<typename T>
    static byte_vector from(const std::basic_string<T> &string)
    {
        byte_vector result;

        result.resize(string.size() * sizeof(T));
        std::memcpy(result.bytes_.data(), string.data(), result.bytes_.size());

        return result;
    }

    byte_vector()
        : reader_(bytes_)
    {
    }

    byte_vector(std::vector<byte> &bytes)
        : bytes_(bytes),
          reader_(bytes_)
    {
    }

    template<typename T>
    byte_vector(const std::vector<T> &ints)
        : byte_vector()
    {
        bytes_ = from(ints).data();
    }

    byte_vector(const byte_vector &other)
        : byte_vector()
    {
        *this = other;
    }

    ~byte_vector()
    {
    }

    byte_vector &operator=(const byte_vector &other)
    {
        bytes_ = other.bytes_;
        reader_ = byte_reader(bytes_);

        return *this;
    }

    const std::vector<byte> &data() const
    {
        return bytes_;
    }

    std::vector<byte> data()
    {
        return bytes_;
    }

    void data(std::vector<byte> &bytes)
    {
        bytes_ = bytes;
    }

    void offset(std::size_t offset)
    {
        reader_.offset(offset);
    }

    std::size_t offset() const
    {
        return reader_.offset();
    }

    void reset()
    {
        reader_.reset();
        bytes_.clear();
    }

    template<typename T>
    T read()
    {
        return reader_.read<T>();
    }

    template<typename T>
    std::vector<T> as_vector_of() const
    {
        return reader_.as_vector_of<T>();
    }

    template<typename T>
    void write(T value)
    {
        const auto num_bytes = sizeof(T);

        if (offset() + num_bytes > size())
        {
            extend(offset() + num_bytes - size());
        }

        std::memcpy(bytes_.data() + offset(), &value, num_bytes);
        reader_.offset(reader_.offset() + num_bytes);
    }

    std::size_t size() const
    {
        return bytes_.size();
    }

    void resize(std::size_t new_size, byte fill = 0)
    {
        bytes_.resize(new_size, fill);
    }

    void extend(std::size_t amount, byte fill = 0)
    {
        bytes_.resize(size() + amount, fill);
    }

    std::vector<byte>::iterator iterator()
    {
        return bytes_.begin() + offset();
    }

    void append(const std::vector<std::uint8_t> &data, std::size_t offset, std::size_t count)
    {
        auto end_index = size();
        extend(count);
        std::memcpy(bytes_.data() + end_index, data.data() + offset, count);
    }

    void append(const std::vector<std::uint8_t> &data)
    {
        append(data, 0, data.size());
    }

private:
    std::vector<byte> bytes_;
    byte_reader reader_;
};

} // namespace detail
} // namespace xlnt
