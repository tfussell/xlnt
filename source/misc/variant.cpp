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
#include <string>
#include <unordered_map>
#include <vector>

#include "variant.h"

namespace xlnt {

const variant $ = variant();

variant::~variant()
{
}

variant &variant::operator=(const variant &rhs)
{
	type_ = rhs.type_;

	switch(type_)
	{
	case T_Null: 
	case T_True:
	case T_False:
		{
			break;
		}
	case T_Number:
		{
			number_ = rhs.number_;
			break;
		}
	case T_String:
		{
			string_ = rhs.string_;
			break;
		}
	}

	return *this;
}

bool variant::operator==(const variant &rhs) const
{
	if(rhs.type_ != type_)
	{
		return false;
	}

	switch(type_)
	{
	case T_Null:
	case T_True:
	case T_False:
		return true;
	case T_Number:
		return std::abs(number_ - rhs.number_) < 0.001;
	case T_String:
		return string_.compare(rhs.string_) == 0;
	}

	throw dynamic_type_exception();
}

} // namespace xlnt
