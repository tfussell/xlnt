#pragma once

#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

template<typename T>
class optional
{
public:
	optional() : has_value_(false), value_(T())
	{
	}

	optional(const T &value) : has_value_(true), value_(value)
	{
	}

	operator bool() const
	{
		return is_set();
	}

	bool is_set() const
	{
		return has_value_;
	}

	void set(const T &value)
	{
		has_value_ = true;
		value_ = value;
	}

	T &operator*()
	{
		return get();
	}

	const T &operator*() const
	{
		return get();
	}

	T *operator->()
	{
		return &get();
	}

	const T *operator->() const
	{
		return &get();
	}

	T &get()
	{
		if (!has_value_)
		{
			throw invalid_attribute();
		}

		return value_;
	}

	const T &get() const
	{
		if (!has_value_)
		{
			throw invalid_attribute();
		}

		return value_;
	}

	void clear()
	{
		has_value_ = false;
		value_ = T();
	}

private:
	bool has_value_;
	T value_;
};

} // namespace xlnt
