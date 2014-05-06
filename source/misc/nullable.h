#pragma once

#include <stdexcept>


namespace xlnt {

/// <summary>
/// Represents a value type that may or may not have an assigned value.
/// </summary>
template<typename T>
struct nullable
{
public:
	/// <summary>
	/// Initializes a new instance of the nullable&lt;T&gt; structure with a "null" value.
	/// </summary>
	nullable() : has_value_(false) {}

	/// <summary>
	/// Initializes a new instance of the nullable&lt;T&gt; structure to the specified value.
	/// </summary>
	nullable(const T &value) : has_value_(true), value_(value) {}

	/// <summary>
	/// gets a value indicating whether the current nullable&lt;T&gt; object has a valid value of its underlying type.
	/// </summary>
	bool has_value() const { return has_value_; }

	/// <summary>
	/// gets the value of the current nullable&lt;T&gt; object if it has been assigned a valid underlying value.
	/// </summary>
	T get_value() const 
	{
		if(!has_value())
		{
			throw std::runtime_error("invalid operation");
		} 

		return value_;
	}

	/// <summary>
	/// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
	/// </summary>
	bool equals(const nullable<T> &other) const
	{
		if(!has_value())
		{
			return !other.has_value();
		}

		if(other.has_value())
		{
			return get_value() == other.get_value();
		}

		return false;
	}

	/// <summary>
	/// Retrieves the value of the current nullable&lt;T&gt; object, or the object's default value.
	/// </summary>
	T get_value_or_default()
	{
		if(has_value())
		{
			return get_value();
		}

		return T();
	}

	/// <summary>
	/// Retrieves the value of the current nullable&lt;T&gt; object or the specified default value.
	/// </summary>
	T get_value_or_default(const T &default_)
	{
		if(has_value())
		{
			return get_value();
		}

		return default_;
	}

	/// <summary>
	/// Returns the value of a specified nullable&lt;T&gt; value.
	/// </summary>
	operator T() const
	{
		return get_value();
	}

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool operator==(nullptr_t) const
    {
        return !has_value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
    /// </summary>
    bool operator!=(nullptr_t) const
    {
        return has_value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
    /// </summary>
    bool operator==(T other) const
    {
        return has_value_ && other == value_;
    }

    /// <summary>
    /// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
    /// </summary>
    bool operator!=(T other) const
    {
        return has_value_ && other != value_;
    }

	/// <summary>
	/// Indicates whether the current nullable&lt;T&gt; object is equal to a specified object.
	/// </summary>
	bool operator==(const nullable<T> &other) const
	{
		return equals(other);
	}

	/// <summary>
	/// Indicates whether the current nullable&lt;T&gt; object is not equal to a specified object.
	/// </summary>
	bool operator!=(const nullable<T> &other) const
	{
		return !(*this == other);
	}

private:
	bool has_value_;
	T value_;
};

} // namespace xlnt
