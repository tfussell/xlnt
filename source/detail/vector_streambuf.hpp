// Copyright (c) 2016 Thomas Fussell
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

#include <iostream>
#include <vector>

namespace xlnt {
namespace detail {

/// <summary>
/// Allows a std::vector to be read through a std::istream.
/// </summary>
class vector_istreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;

public:
    vector_istreambuf(const std::vector<std::uint8_t> &data)
        : data_(data),
          position_(0)
    {
    }

    vector_istreambuf(const vector_istreambuf &) = delete;
    vector_istreambuf &operator=(const vector_istreambuf &) = delete;

private:
    int_type underflow()
    {
        return (position_ == static_cast<std::streampos>(data_.size()))
            ? traits_type::eof() : traits_type::to_int_type(data_[position_]);
    }

    int_type uflow()
    {
        if (position_ == static_cast<std::streampos>(data_.size()))
        {
            return traits_type::eof();
        }

        auto previous = position_;
        position_ += 1;

        return traits_type::to_int_type(data_[previous]);
    }

    int_type pbackfail(int_type ch)
    {
        if (position_ == std::streampos(0) || (ch != traits_type::eof() && ch != data_[position_ - std::streampos(1)]))
        {
            return traits_type::eof();
        }

        auto old_pos = position_;
        position_ -= 1;

        return traits_type::to_int_type(data_[position_]);
    }

    std::streamsize showmanyc()
    {
        return position_ < data_.size() ? data_.size() - position_ : -1;
    }

    std::streampos seekoff(std::streamoff off,
        std::ios_base::seekdir way,
        std::ios_base::openmode which = std::ios_base::in | std::ios_base::out)
    {
        if (way == std::ios_base::beg)
        {
            position_ = 0;
        }
        else if (way == std::ios_base::end)
        {
            position_ = data_.size();
        }

        position_ += off;

        if (position_ < 0) return -1;
        if (position_ > data_.size()) return -1;

        return position_;
    }

    std::streampos seekpos(std::streampos sp,
        std::ios_base::openmode which = std::ios_base::in | std::ios_base::out)
    {
        return position_ = sp;
    }

private:
    const std::vector<std::uint8_t> &data_;
    std::streampos position_;
};

/// <summary>
/// Allows a std::vector to be written through a std::ostream.
/// </summary>
class vector_ostreambuf : public std::streambuf
{
    using int_type = std::streambuf::int_type;

public:
    vector_ostreambuf(std::vector<std::uint8_t> &data)
        : data_(data),
          position_(0)
    {
    }

    vector_ostreambuf(const vector_ostreambuf &) = delete;
    vector_ostreambuf &operator=(const vector_ostreambuf &) = delete;

private:
    int_type overflow(int_type c = traits_type::eof())
    {
        if (c != traits_type::eof())
        {
            data_.push_back(c);
            position_ = data_.size() - 1;
        }

        return traits_type::to_int_type(data_[position_]);
    }

    std::streamsize xsputn(const char *s, std::streamsize n)
    {
        if (data_.empty())
        {
            data_.resize(n);
        }
        else
        {
            auto position_size = data_.size();
            auto required_size = static_cast<std::size_t>(position_ + n);
            data_.resize(std::max(position_size, required_size));
        }

        std::copy(s, s + n, data_.begin() + position_);
        position_ += n;

        return n;
    }

    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way,
        std::ios_base::openmode which = std::ios_base::in | std::ios_base::out)
    {
        if (way == std::ios_base::beg)
        {
            position_ = off;
        }
        else if (way == std::ios_base::cur)
        {
            position_ += off;
        }
        else if (way == std::ios_base::end)
        {
            position_ = data_.size();
        }

        return (position_ < 0 || position_ > data_.size())
          ? std::streampos(-1) : position_;
    }

    std::streampos seekpos(std::streampos sp,
        std::ios_base::openmode which = std::ios_base::in | std::ios_base::out)
    {
        position_ = sp;
        return (position_ < 0 || position_ > data_.size())
          ? std::streampos(-1) : position_;
    }

private:
    std::vector<std::uint8_t> &data_;
    std::streampos position_;
};

} // namespace detail
} // namespace xlnt
