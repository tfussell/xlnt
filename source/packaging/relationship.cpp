// Copyright (c) 2014-2017 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
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

#include <xlnt/packaging/relationship.hpp>

namespace xlnt {

relationship::relationship()
{
}

relationship::relationship(
    const std::string &id, relationship_type t, const uri &source, const uri &target, xlnt::target_mode mode)
    : id_(id), type_(t), source_(source), target_(target), mode_(mode)
{
}

std::string relationship::id() const
{
    return id_;
}

target_mode relationship::target_mode() const
{
    return mode_;
}

uri relationship::source() const
{
    return source_;
}

uri relationship::target() const
{
    return target_;
}

relationship_type relationship::type() const
{
    return type_;
}

bool relationship::operator==(const relationship &rhs) const
{
    return type_ == rhs.type_ && id_ == rhs.id_ && source_ == rhs.source_ && target_ == rhs.target_
        && mode_ == rhs.mode_;
}

} // namespace xlnt
