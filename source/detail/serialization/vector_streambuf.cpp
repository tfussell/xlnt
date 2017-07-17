// Copyright (c) 2016-2017 Thomas Fussell
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

#include <detail/serialization/vector_streambuf.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {
namespace detail {

vector_istreambuf::vector_istreambuf(const std::vector<std::uint8_t> &data)
    : data_(data),
        position_(0)
{
}

vector_istreambuf::int_type vector_istreambuf::underflow()
{
    if (position_ == data_.size())
    {
        return traits_type::eof();
    }

    return traits_type::to_int_type(static_cast<char>(data_[position_]));
}

vector_istreambuf::int_type vector_istreambuf::uflow()
{
    if (position_ == data_.size())
    {
        return traits_type::eof();
    }

    return traits_type::to_int_type(static_cast<char>(data_[position_++]));
}

std::streamsize vector_istreambuf::showmanyc()
{
    if (position_ == data_.size())
    {
        return static_cast<std::streamsize>(-1);
    }

    return static_cast<std::streamsize>(data_.size() - position_);
}

std::streampos vector_istreambuf::seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
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

std::streampos vector_istreambuf::seekpos(std::streampos sp, std::ios_base::openmode)
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

vector_ostreambuf::vector_ostreambuf(std::vector<std::uint8_t> &data)
    : data_(data),
        position_(0)
{
}

vector_ostreambuf::int_type vector_ostreambuf::overflow(int_type c)
{
    if (c != traits_type::eof())
    {
        data_.push_back(static_cast<std::uint8_t>(c));
        position_ = data_.size() - 1;
    }

    return traits_type::to_int_type(static_cast<char>(data_[position_]));
}

std::streamsize vector_ostreambuf::xsputn(const char *s, std::streamsize n)
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

std::streampos vector_ostreambuf::seekoff(std::streamoff off, std::ios_base::seekdir way, std::ios_base::openmode)
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

std::streampos vector_ostreambuf::seekpos(std::streampos sp, std::ios_base::openmode)
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

XLNT_API std::vector<std::uint8_t> to_vector(std::istream &in_stream)
{
    if (!in_stream)
    {
        throw xlnt::exception("bad stream");
    }

    return std::vector<std::uint8_t>(
        std::istreambuf_iterator<char>(in_stream),
        std::istreambuf_iterator<char>());
}

XLNT_API void to_stream(const std::vector<std::uint8_t> &bytes, std::ostream &out_stream)
{
    if (!out_stream)
    {
        throw xlnt::exception("bad stream");
    }

    out_stream.write(reinterpret_cast<const char *>(bytes.data()), static_cast<std::ptrdiff_t>(bytes.size()));
}

XLNT_API std::ostream &operator<<(std::ostream &out_stream, const std::vector<std::uint8_t> &bytes)
{
    to_stream(bytes, out_stream);
    return out_stream;
}

} // namespace detail
} // namespace xlnt
