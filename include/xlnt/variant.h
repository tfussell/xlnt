/*
  Copyright (c) 2012-2014 Thomas Fussell

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:
  
  The above copyright notice and this permission notice shall be included in
  all copies or substantial portions of the Software.
  
  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
  THE SOFTWARE.
*/

/** @file variant.h
 *  @brief A dynamic container holding one of the seven JSON primitives.
 *
 *  Declaration of variant class which represents a dynamic type. It can be assigned any
 *  one of the following value types as defined in RFC 4627: false, null, true,
 *  object, array, number, string.
 *
 *  @author Thomas A. Fussell (tfussell)
 */
#pragma once

#include <exception>
#include <string>
#include <unordered_map>
#include <vector>

namespace xlnt {

/**
 * A dynamic type exception may be raised when an operation is attempted
 * on a dynamic type which does not support it. For example, trying to
 * index an integer type.
 */
class dynamic_type_exception : public std::runtime_error
{
public:
	dynamic_type_exception()  : std::runtime_error("Error: wrong type")
	{
	}
};

/**
 * Must be forward declared so that object and array can store pointers to this
 * heretofore incomplete type.
 */
class variant;

/**
 * A variant class can be dynamically assigned to one of the following: null, true,
 * false, number, string.
 *
 * Note: This class uses a long double to store its dynamic number values. This
 * is consistent with JavaScript's number type, but it may limit the precision
 * on very large or very small numbers and introduce rounding errors.
 */
class variant
{
private:
	/** 
	 * An enumeration of possible dynamic types that the containing class may assume.
	 */
	enum Type
	{
		T_Null, /**< null */
		T_True, /**< true */
		T_False, /**< false */
		T_Number, /**< number */
		T_String, /**< string */
	};

public:
	#pragma region Constructors
	/**
	 * Constructs a new empty variant of null type.
	 */
	variant() : type_(T_Null) { }

	/**
	 * Constructs a new variant class of type T_True or T_False if the given paramter, b, is true or false respectively.
	 */
	variant(bool b) : type_(b ? T_True : T_False) { }

	/**
	 * Constructs a new variant class of type T_Number with a dynamic numeric value based on the given parameter, i.
	 */
	variant(int i) : type_(T_Number), number_(i) { }

	/**
	 * Constructs a new variant class of type T_Number with a dynamic numeric value based on the given parameter, i.
	 */
	variant(long long i) : type_(T_Number), number_(static_cast<long double>(i)) { }

	/**
	 * Constructs a new variant class of type T_Number with a dynamic numeric value based on the given parameter, d.
	 */
	variant(double d) : type_(T_Number), number_(d) { }

	/**
	 * Constructs a new variant class of type T_Number with a dynamic numeric value based on the given parameter, d.
	 */
	variant(long double d) : type_(T_Number), number_(d) { }

	/**
	 * Constructs a new variant class of type T_String with a dynamic string value based on the given parameter, s.
	 */
	variant(const char *s) : type_(T_String), string_(s) { }

	/**
	 * Constructs a new variant class of type T_String with a dynamic string value based on the given parameter, s.
	 */
	variant(const std::string &s) : type_(T_String), string_(s) { }

	/**
	 * Constructs a new variant as a deep-copy of an existing object.
	 */
	variant(const variant &v) { *this = v; }

	/**
	 * If this is a container type, destroys contents and clears the container.
	 */
	~variant();
	#pragma endregion

	#pragma region Operators
	/**
	 * Assigns the fields of this object to those of the given parameter, rhs, by creating a deep-copy of containers.
	 */
	variant &operator=(const variant &rhs);

	/**
	 * Returns true if the types of this object and the given parameter, rhs, are equal and checks the corresponding
	 * field for equality. For container types, this recursively checks for element-by-element equality in the same 
	 * way.
	 */
	bool operator==(const variant &rhs) const;

	/**
	 * Returns the negation of the result of the equality operator.
	 */
	bool operator!=(const variant &rhs) const { return !(*this == rhs); }
	#pragma endregion

	#pragma region Casts
	template <class T> T cast() const;

	template <>
	bool variant::cast() const
	{
		if(type_ == T_True)
		{
			return true;
		}
		else if(type_ == T_False)
		{
			return false;
		}

		throw dynamic_type_exception();
	}

	template <>
	int variant::cast() const
	{
		if(type_ == T_Number)
		{
			return static_cast<int>(number_);
		}

		throw dynamic_type_exception();
	}

	template <>
	double variant::cast() const
	{
		if(type_ == T_Number)
		{
			return static_cast<double>(number_);
		}

		throw dynamic_type_exception();
	}

	template <>
	long long variant::cast() const
	{
		if(type_ == T_Number)
		{
			return static_cast<long long>(number_);
		}

		throw dynamic_type_exception();
	}

	template <>
	long double variant::cast() const
	{
		if(type_ == T_Number)
		{
			return number_;
		}

		throw dynamic_type_exception();
	}

	template <>
	std::string variant::cast() const
	{
		if(type_ == T_String)
		{
			return string_;
		}

		throw dynamic_type_exception();
	}

	/**
	 * See: bool ToBool() const
	 */
	operator bool() const { return cast<bool>(); }

	/**
	 * See: int ToInt() const
	 */
	operator int() const { return cast<int>(); }

	/**
	 * See: double ToDouble() const
	 */
	operator double() const { return cast<double>(); }

	/**
	 * See: long long ToLongLong() const
	 */
	operator long long() const { return cast<long long>(); }

	/**
	 * See: long double ToLongDouble() const
	 */
	operator long double() const { return cast<long double>(); }

	/**
	 * See: std::string ToString() const
	 */
	operator std::string() const { return cast<std::string>(); }
	#pragma endregion

	#pragma region Type checks

	/**
	 * Returns true if the type of this object is T_Null.
	 */
	bool isnull() const { return type_ == T_Null; }

	/**
	 * Returns true if the type of this object is T_True.
	 */
	bool istrue() const { return type_ == T_True; }

	/**
	 * Returns true if the type of this object is T_False.
	 */
	bool isfalse() const { return type_ == T_False; }

	/**
	 * Returns true if the type of this object is T_True or T_False.
	 */
	bool isboolean() const { return istrue() || isfalse(); }

	/**
	 * Returns true if the type of this object is T_Number.
	 */
	bool isnumber() const { return type_ == T_Number; }

	/**
	 * Returns true if the type of this object is T_Number and number has no decimal part.
	 */
	bool isint() const { return isnumber() && !isfloat(); }

	/**
	 * Returns true if the type of this object is T_Number and number has a decimal part.
	 */
	bool isfloat() const { return isnumber() && ceil(number_) != number_; }

	/**
	 * Returns true if the type of this object is T_String.
	 */
	bool isstring() const { return type_ == T_String; }

	#pragma endregion

private:
	/**
	* Constructs a new variant of the given type with an optional parameter indicating whether this
	* variant preserves order (only relevant if it is of type T_Object).
	*/
	variant(Type type) : type_(type) { }

	/**
	* Attempts to convert the current value of this object to a bool.
	* This will raise an error if the internal type is not T_True or T_False.
	*/
	bool asbool() const;

	/*
	* Attempts to convert the current value of this object to an int.
	* This will raise an error if the internal type is not T_Number.
	* This conversion may overflow.
	*/
	int asint() const;

	/*
	* Attempts to convert the current value of this object to a double.
	* This will raise an error if the internal type is not T_Number.
	* This conversion may overflow.
	*/
	double asdouble() const;

	/*
	* Attempts to convert the current value of this object to a long long int.
	* This will raise an error if the internal type is not T_Number.
	*/
	long long aslonglong() const;

	/*
	* Attempts to convert the current value of this object to a long double.
	* This will raise an error if the internal type is not T_Number.
	*/
	long double aslongdouble() const;

	/*
	* Attempts to convert the current value of this object to a string.
	* This will raise an error if the internal type is not T_String.
	*/
	std::string asstring() const;

	Type type_; /**< Dynamic type of the contained value. */

	long double number_; /**< Holder for numeric value if this has type T_Number */
	std::string string_; /**< Holder for string value if this has type T_String */
};

extern const variant $; /**< Special variant which represents null JSON literal */

} // namespace xlnt
