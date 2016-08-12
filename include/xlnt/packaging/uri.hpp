// Copyright (c) 2014-2016 Thomas Fussell
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

#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/utils/path.hpp>

namespace xlnt {

/// <summary>
///
/// </summary>
class XLNT_CLASS uri
{
public:
    uri();
    uri(const uri &base, const uri &relative);
    uri(const uri &base, const path &relative);
    uri(const std::string &uri_string);
    
    bool is_relative() const;
    bool is_absolute() const;

    std::string get_scheme() const;
    std::string get_authority() const;
    bool has_authentication() const;
    std::string get_authentication() const;
    std::string get_username() const;
    std::string get_password() const;
    std::string get_host() const;
    bool has_port() const;
    std::size_t get_port() const;
    path get_path() const;
    bool has_query() const;
    std::string get_query() const;
    bool has_fragment() const;
    std::string get_fragment() const;
    
    std::string to_string() const;
    
    uri make_absolute(const uri &base);
    uri make_reference(const uri &base);
    
    friend XLNT_FUNCTION bool operator==(const uri &left, const uri &right);
    
private:
    bool absolute_;
    std::string scheme_;
    bool has_authentication_;
    std::string username_;
    std::string password_;
    std::string host_;
    bool has_port_;
    std::size_t port_;
    bool has_query_;
    std::string query_;
    bool has_fragment_;
    std::string fragment_;
    path path_;
};

} // namespace xlnt
