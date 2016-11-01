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
        if (position_ == data_.size())
        {
            return traits_type::eof();
        }

        return traits_type::to_int_type(static_cast<char>(data_[position_]));
    }

    int_type uflow()
    {
        if (position_ == data_.size())
        {
            return traits_type::eof();
        }

        return traits_type::to_int_type(static_cast<char>(data_[position_++]));
    }

    std::streamsize showmanyc()
    {
        if (position_ == data_.size())
        {
            return static_cast<std::streamsize>(-1);
        }

        return static_cast<std::streamsize>(data_.size() - position_);
    }

    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
    {
        if (way == std::ios_base::beg)
        {
            position_ = 0;
        }
        else if (way == std::ios_base::end)
        {
            position_ = data_.size();
        }

        if (off < 0)
        {
            if (static_cast<std::size_t>(-off) > position_)
            {
                position_ = 0;
                return static_cast<std::ptrdiff_t>(-1);
            }
            else
            {
                position_ -= static_cast<std::size_t>(-off);
            }
        }
        else if (off > 0)
        {
            if (static_cast<std::size_t>(off) + position_ > data_.size())
            {
                position_ = data_.size();
                return static_cast<std::ptrdiff_t>(-1);
            }
            else
            {
                position_ += static_cast<std::size_t>(off);
            }
        }

        return static_cast<std::ptrdiff_t>(position_);
    }

    std::streampos seekpos(std::streampos sp, std::ios_base::openmode)
    {
        if (sp < 0)
        {
            position_ = 0;
        }
        else if (static_cast<std::size_t>(sp) > data_.size())
        {
            position_ = data_.size();
        }
        else
        {
            position_ = static_cast<std::size_t>(sp);
        }
        
        return static_cast<std::ptrdiff_t>(position_);
    }

private:
    const std::vector<std::uint8_t> &data_;
    std::size_t position_;
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

        return traits_type::to_int_type(static_cast<char>(data_[position_]));
    }

    std::streamsize xsputn(const char *s, std::streamsize n)
    {
        if (data_.empty())
        {
            data_.resize(static_cast<std::size_t>(n));
        }
        else
        {
            auto position_size = data_.size();
            auto required_size = static_cast<std::size_t>(position_ + static_cast<std::size_t>(n));
            data_.resize(std::max(position_size, required_size));
        }

        std::copy(s, s + n, data_.begin() + static_cast<std::ptrdiff_t>(position_));
        position_ += static_cast<std::size_t>(n);

        return n;
    }

    std::streampos seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
    {
        if (way == std::ios_base::beg)
        {
            position_ = 0;
        }
        else if (way == std::ios_base::end)
        {
            position_ = data_.size();
        }

        if (off < 0)
        {
            if (static_cast<std::size_t>(-off) > position_)
            {
                position_ = 0;
                return static_cast<std::ptrdiff_t>(-1);
            }
            else
            {
                position_ -= static_cast<std::size_t>(-off);
            }
        }
        else if (off > 0)
        {
            if (static_cast<std::size_t>(off) + position_ > data_.size())
            {
                position_ = data_.size();
                return static_cast<std::ptrdiff_t>(-1);
            }
            else
            {
                position_ += static_cast<std::size_t>(off);
            }
        }

        return static_cast<std::ptrdiff_t>(position_);
    }

    std::streampos seekpos(std::streampos sp, std::ios_base::openmode)
    {
        if (sp < 0)
        {
            position_ = 0;
        }
        else if (static_cast<std::size_t>(sp) > data_.size())
        {
            position_ = data_.size();
        }
        else
        {
            position_ = static_cast<std::size_t>(sp);
        }
        
        return static_cast<std::ptrdiff_t>(position_);
    }

private:
    std::vector<std::uint8_t> &data_;
    std::size_t position_;
};

} // namespace detail
} // namespace xlnt
