#pragma once

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

/// <summary>
/// Many settings in xlnt are allowed to not have a value set. This class
/// encapsulates a value which may or may not be set. Memory is allocated
/// within the optional class.
/// </summary>
template<typename T>
class XLNT_API optional
{
public:
    /// <summary>
    ///
    /// </summary>
	optional() : has_value_(false), value_(T())
	{
	}

    /// <summary>
    ///
    /// </summary>
	optional(const T &value) : has_value_(true), value_(value)
	{
	}

    /// <summary>
    ///
    /// </summary>
	operator bool() const
	{
		return is_set();
	}

    /// <summary>
    ///
    /// </summary>
	bool is_set() const
	{
		return has_value_;
	}

    /// <summary>
    ///
    /// </summary>
	void set(const T &value)
	{
		has_value_ = true;
		value_ = value;
	}

    /// <summary>
    ///
    /// </summary>
	T &operator*()
	{
		return get();
	}

    /// <summary>
    ///
    /// </summary>
	const T &operator*() const
	{
		return get();
	}

    /// <summary>
    ///
    /// </summary>
	T *operator->()
	{
		return &get();
	}

    /// <summary>
    ///
    /// </summary>
	const T *operator->() const
	{
		return &get();
	}

    /// <summary>
    ///
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
    ///
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
    ///
    /// </summary>
	void clear()
	{
		has_value_ = false;
		value_ = T();
	}

    /// <summary>
    ///
    /// </summary>
    bool operator==(const optional<T> &other) const
    {
        return has_value_ == other.has_value_
            && (!has_value_
            || (has_value_ && value_ == other.value_));
    }

private:
    /// <summary>
    ///
    /// </summary>
	bool has_value_;

    /// <summary>
    ///
    /// </summary>
	T value_;
};

} // namespace xlnt
