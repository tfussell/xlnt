// Copyright (c) 2016-2021 Thomas Fussell
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
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#pragma once

#include "xlnt/utils/exceptions.hpp"
#include "xlnt/utils/numeric.hpp"
#include "xlnt/xlnt_config.hpp"
#include <type_traits>

namespace xlnt {

/// <summary>
/// Many settings in xlnt are allowed to not have a value set. This class
/// encapsulates a value which may or may not be set. Memory is allocated
/// within the optional class.
/// </summary>
template <typename T>
class optional
{
#if ((defined(_MSC_VER) && _MSC_VER <= 1900) || (defined(__GNUC__) && __GNUC__ < 5))
// Disable enhanced type checking on Visual Studio <= 2015 and GCC <5
#define XLNT_NOEXCEPT_VALUE_COMPAT(...) (false)
#else
#define XLNT_NOEXCEPT_VALUE_COMPAT(...) (__VA_ARGS__)
    using ctor_copy_T_noexcept = typename std::conditional<std::is_nothrow_copy_constructible<T>{}, std::true_type, std::false_type>::type;
    using ctor_move_T_noexcept = typename std::conditional<std::is_nothrow_move_constructible<T>{}, std::true_type, std::false_type>::type;
    using copy_ctor_noexcept = ctor_copy_T_noexcept;
    using move_ctor_noexcept = ctor_move_T_noexcept;
    using set_copy_noexcept_t = typename std::conditional<std::is_nothrow_copy_constructible<T>{} && std::is_nothrow_assignable<T, T>{}, std::true_type, std::false_type>::type;
    using set_move_noexcept_t = typename std::conditional<std::is_nothrow_move_constructible<T>{} && std::is_nothrow_move_assignable<T>{}, std::true_type, std::false_type>::type;
    using clear_noexcept_t = typename std::conditional<std::is_nothrow_destructible<T>{}, std::true_type, std::false_type>::type;
#endif
    /// <summary>
    /// Default equality operation, just uses operator==
    /// </summary>
    template <typename U = T, typename std::enable_if<!std::is_floating_point<U>::value>::type * = nullptr>
    constexpr bool compare_equal(const U &lhs, const U &rhs) const
    {
        return lhs == rhs;
    }

    /// <summary>
    /// equality operation for floating point numbers. Provides "fuzzy" equality
    /// </summary>
    template <typename U = T, typename std::enable_if<std::is_floating_point<U>::value>::type * = nullptr>
    constexpr bool compare_equal(const U &lhs, const U &rhs) const
    {
        return detail::float_equals(lhs, rhs);
    }

public:
    /// <summary>
    /// Default contructor. is_set() will be false initially.
    /// </summary>
    optional() noexcept
        : has_value_(false)
    {
    }

    /// <summary>
    /// Constructs this optional with a value.
    /// noexcept if T copy ctor is noexcept
    /// </summary>
    optional(const T &value) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(ctor_copy_T_noexcept{}))
        : has_value_(true)
    {
        new (&storage_) T(value);
    }

    /// <summary>
    /// Constructs this optional with a value.
    /// noexcept if T move ctor is noexcept
    /// </summary>
    optional(T &&value) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(ctor_move_T_noexcept{}))
        : has_value_(true)
    {
        new (&storage_) T(std::move(value));
    }

    /// <summary>
    /// Copy constructs this optional from other
    /// noexcept if T copy ctor is noexcept
    /// </summary>
    optional(const optional &other) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(copy_ctor_noexcept{}))
        : has_value_(other.has_value_)
    {
        if (has_value_)
        {
            new (&storage_) T(other.value_ref());
        }
    }

    /// <summary>
    /// Move constructs this optional from other. Clears the value from other if set
    /// noexcept if T move ctor is noexcept
    /// </summary>
    optional(optional &&other) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(move_ctor_noexcept{}))
        : has_value_(other.has_value_)
    {
        if (has_value_)
        {
            new (&storage_) T(std::move(other.value_ref()));
            other.clear();
        }
    }

    /// <summary>
    /// Copy assignment of this optional from other
    /// noexcept if set and clear are noexcept for T&
    /// </summary>
    optional &operator=(const optional &other) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_copy_noexcept_t{} && clear_noexcept_t{}))
    {
        if (other.has_value_)
        {
            set(other.value_ref());
        }
        else
        {
            clear();
        }
        return *this;
    }

    /// <summary>
    /// Move assignment of this optional from other
    /// noexcept if set and clear are noexcept for T&&
    /// </summary>
    optional &operator=(optional &&other) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_move_noexcept_t{} && clear_noexcept_t{}))
    {
        if (other.has_value_)
        {
            set(std::move(other.value_ref()));
            other.clear();
        }
        else
        {
            clear();
        }
        return *this;
    }

    /// <summary>
    /// Destructor cleans up the T instance if set
    /// </summary>
    ~optional() noexcept // note:: unconditional because msvc freaks out otherwise
    {
        clear();
    }

    /// <summary>
    /// Returns true if this object currently has a value set. This should
    /// be called before accessing the value with optional::get().
    /// </summary>
    bool is_set() const noexcept
    {
        return has_value_;
    }

    /// <summary>
    /// Copies the value into the stored value
    /// </summary>
    void set(const T &value) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_copy_noexcept_t{}))
    {
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wmaybe-uninitialized"
#endif
        if (has_value_)
        {
            value_ref() = value;
        }
        else
        {
            new (&storage_) T(value);
            has_value_ = true;
        }
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic pop
#endif
    }

    /// <summary>
    /// Moves the value into the stored value
    /// </summary>
    void set(T &&value) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_move_noexcept_t{}))
    {
        // note seperate overload for two reasons (as opposed to perfect forwarding)
        // 1. have to deal with implicit conversions internally with perfect forwarding
        // 2. have to deal with the noexcept specfiers for all the different variations
        // overload is just far and away the simpler solution
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wmaybe-uninitialized"
#endif
        if (has_value_)
        {
            value_ref() = std::move(value);
        }
        else
        {
            new (&storage_) T(std::move(value));
            has_value_ = true;
        }
#if defined(__GNUC__) && !defined(__clang__)
#pragma GCC diagnostic pop
#endif
    }

    /// <summary>
    /// Assignment operator overload. Equivalent to setting the value using optional::set.
    /// </summary>
    optional &operator=(const T &rhs) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_copy_noexcept_t{}))
    {
        set(rhs);
        return *this;
    }

    /// <summary>
    /// Assignment operator overload. Equivalent to setting the value using optional::set.
    /// </summary>
    optional &operator=(T &&rhs) noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(set_move_noexcept_t{}))
    {
        set(std::move(rhs));
        return *this;
    }

    /// <summary>
    /// After this is called, is_set() will return false until a new value is provided.
    /// </summary>
    void clear() noexcept(XLNT_NOEXCEPT_VALUE_COMPAT(clear_noexcept_t{}))
    {
        if (has_value_)
        {
            reinterpret_cast<T *>(&storage_)->~T();
        }
        has_value_ = false;
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

        return value_ref();
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

        return value_ref();
    }

    /// <summary>
    /// Returns true if neither this nor other have a value
    /// or both have a value and those values are equal according to
    /// their equality operator.
    /// </summary>
    bool operator==(const optional<T> &other) const noexcept
    {
        if (has_value_ != other.has_value_)
        {
            return false;
        }
        if (!has_value_)
        {
            return true;
        }
        // equality is overloaded to provide fuzzy equality when T is a fp number
        return compare_equal(value_ref(), other.value_ref());
    }

    /// <summary>
    /// Returns false if neither this nor other have a value
    /// or both have a value and those values are equal according to
    /// their equality operator.
    /// </summary>
    bool operator!=(const optional<T> &other) const noexcept
    {
        return !operator==(other);
    }

private:
    // helpers for getting a T out of storage
    T &value_ref() noexcept
    {
        return *reinterpret_cast<T *>(&storage_);
    }

    const T &value_ref() const noexcept
    {
        return *reinterpret_cast<const T *>(&storage_);
    }

    bool has_value_;
    typename std::aligned_storage<sizeof(T), alignof(T)>::type storage_;
};

#ifdef XLNT_NOEXCEPT_VALUE_COMPAT
#undef XLNT_NOEXCEPT_VALUE_COMPAT
#endif

} // namespace xlnt
