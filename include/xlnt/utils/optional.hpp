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

#pragma once

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

/// <summary>
/// Many settings in xlnt are allowed to not have a value set. This class
/// encapsulates a value which may or may not be set. Memory is allocated
/// within the optional class.
/// </summary>
template <typename T>
class XLNT_API optional
{
public:
    /// <summary>
    /// Default contructor. is_set() will be false initially.
    /// </summary>
    optional() : has_value_(false), value_(T())
    {
    }

    /// <summary>
    /// Constructs this optional with a value.
    /// </summary>
    explicit optional(const T &value) : has_value_(true), value_(value)
    {
    }

    /// <summary>
    /// Returns true if this object currently has a value set. This should
    /// be called before accessing the value with optional::get().
    /// </summary>
    bool is_set() const
    {
        return has_value_;
    }

    /// <summary>
    /// Sets the value to value.
    /// </summary>
    void set(const T &value)
    {
        has_value_ = true;
        value_ = value;
    }

    /// <summary>
    /// Gets the value. If no value has been initialized in this object,
    /// an xlnt::invalid_attribute exception will be thrown.
    /// </summary>
    T &get()
    {
        if (!has_value_)
        {
            throw invalid_attribute();
        }

        return value_;
    }

    /// <summary>
    /// Gets the value. If no value has been initialized in this object,
    /// an xlnt::invalid_attribute exception will be thrown.
    /// </summary>
    const T &get() const
    {
        if (!has_value_)
        {
            throw invalid_attribute();
        }

        return value_;
    }

    /// <summary>
    /// Resets the internal value using its default constructor. After this is
    /// called, is_set() will return false until a new value is provided.
    /// </summary>
    void clear()
    {
        has_value_ = false;
        value_ = T();
    }

    /// <summary>
    /// Assignment operator. Equivalent to setting the value using optional::set.
    /// </summary>
    optional &operator=(const T &rhs)
    {
        has_value_ = true;
        value_ = rhs;

        return *this;
    }

    /// <summary>
    /// Returns true if neither this nor other have a value
    /// or both have a value and those values are equal according to
    /// their equality operator.
    /// </summary>
    bool operator==(const optional<T> &other) const
    {
        return has_value_ == other.has_value_
            && (!has_value_ || (has_value_ && value_ == other.value_));
    }

private:
    bool has_value_;
    T value_;
};

} // namespace xlnt
