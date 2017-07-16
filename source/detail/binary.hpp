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
#include <iostream>
#include <vector>

#include <xlnt/utils/exceptions.hpp>

namespace xlnt {
namespace detail {

using byte = std::uint8_t;

template<typename T>
class binary_reader
{
public:
    binary_reader() = delete;

    binary_reader(const std::vector<T> &vector)
        : vector_(&vector),
          data_(nullptr),
          size_(0)
    {
    }

    binary_reader(const T *source_data, std::size_t size)
        : vector_(nullptr),
          data_(source_data),
          size_(size)
    {
    }

    binary_reader(const binary_reader &other) = default;

    binary_reader &operator=(const binary_reader &other)
    {
        vector_ = other.vector_;
        offset_ = other.offset_;
        data_ = other.data_;

        return *this;
    }

    ~binary_reader()
    {
    }

    const T *data() const
    {
        return vector_ == nullptr ? data_ : vector_->data();
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

    template<typename U>
    U read()
    {
        return read_reference<U>();
    }

    template<typename U>
    const U *read_pointer()
    {
        const auto result = reinterpret_cast<const U *>(data() + offset_);
        offset_ += sizeof(U) / sizeof(T);

        return result;
    }

    template<typename U>
    const U &read_reference()
    {
        return *read_pointer<U>();
    }

    template<typename U>
    std::vector<U> as_vector() const
    {
        auto result = std::vector<T>(bytes() / sizeof(U), U());
        std::memcpy(result.data(), data(), bytes());

        return result;
    }

    template<typename U>
    std::vector<U> read_vector(std::size_t count)
    {
        auto result = std::vector<U>(count, U());
        std::memcpy(result.data(), data() + offset_, count * sizeof(U));
        offset_ += count * sizeof(T) / sizeof(U);

        return result;
    }

    std::size_t count() const
    {
        return vector_ != nullptr ? vector_->size() : size_;
    }

    std::size_t bytes() const
    {
        return count() * sizeof(T);
    }

private:
    std::size_t offset_ = 0;
    const std::vector<T> *vector_;
    const T *data_;
    const std::size_t size_;
};

template<typename T>
class binary_writer
{
public:
    binary_writer(std::vector<T> &bytes)
        : data_(&bytes)
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
        data_ = other.data_;
        offset_ = other.offset_;

        return *this;
    }

    std::vector<T> &data()
    {
        return *data_;
    }
    
    // Make the bytes of the data pointed to by this writer equivalent to those in the given vector
    // sizeof(U) should be a multiple of sizeof(T)
    template<typename U>
    void assign(const std::vector<U> &ints)
    {
        resize(ints.size() * sizeof(U));
        std::memcpy(data_->data(), ints.data(), bytes());
    }

    // Make the bytes of the data pointed to by this writer equivalent to those in the given string
    // sizeof(U) should be a multiple of sizeof(T)
    template<typename U>
    void assign(const std::basic_string<U> &string)
    {
        resize(string.size() * sizeof(U));
        std::memcpy(data_->data(), string.data(), bytes());
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
        data_->clear();
    }

    template<typename U>
    void write(U value)
    {
        const auto num_bytes = sizeof(U);
        const auto remaining_bytes = bytes() - offset() * sizeof(T);

        if (remaining_bytes < num_bytes)
        {
            extend((num_bytes - remaining_bytes) / sizeof(T));
        }

        std::memcpy(data_->data() + offset(), &value, num_bytes);
        offset_ += num_bytes / sizeof(T);
    }

    std::size_t count() const
    {
        return data_->size();
    }

    std::size_t bytes() const
    {
        return count() * sizeof(T);
    }

    void resize(std::size_t new_size, byte fill = 0)
    {
        data_->resize(new_size, fill);
    }

    void extend(std::size_t amount, byte fill = 0)
    {
        data_->resize(count() + amount, fill);
    }

    std::vector<byte>::iterator iterator()
    {
        return data_->begin() + static_cast<std::ptrdiff_t>(offset());
    }

    template<typename U>
    void append(const std::vector<U> &data)
    {
        binary_reader<U> reader(data);
        append(reader, data.size() * sizeof(U));
    }

    template<typename U>
    void append(binary_reader<U> &reader, std::size_t reader_element_count)
    {
        const auto num_bytes = sizeof(U) * reader_element_count;
        const auto remaining_bytes = bytes() - offset() * sizeof(T);

        if (remaining_bytes < num_bytes)
        {
            extend((num_bytes - remaining_bytes) / sizeof(T));
        }

        if ((reader.offset() + reader_element_count) * sizeof(U) > reader.bytes())
        {
            throw xlnt::exception("reading past end");
        }

        std::memcpy(data_->data() + offset_, reader.data() + reader.offset(), reader_element_count * sizeof(U));
        offset_ += reader_element_count * sizeof(U) / sizeof(T);
    }

private:
    std::vector<T> *data_;
    std::size_t offset_ = 0;
};

template<typename T>
std::vector<byte> string_to_bytes(const std::basic_string<T> &string)
{
    std::vector<byte> bytes;
    binary_writer<byte> writer(bytes);
    writer.assign(string);
    
    return bytes;
}

template<typename T>
T read(std::istream &in)
{
    T result;
    in.read(reinterpret_cast<char *>(&result), sizeof(T));

    return result;
}

template<typename T>
std::vector<T> read_vector(std::istream &in, std::size_t count)
{
    std::vector<T> result(count, T());
    in.read(reinterpret_cast<char *>(&result[0]),
        static_cast<std::streamsize>(sizeof(T) * count));

    return result;
}

template<typename T>
std::basic_string<T> read_string(std::istream &in, std::size_t count)
{
    std::basic_string<T> result(count, T());
    in.read(reinterpret_cast<char *>(&result[0]),
        static_cast<std::streamsize>(sizeof(T) * count));

    return result;
}

} // namespace detail
} // namespace xlnt
