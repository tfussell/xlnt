// Copyright (c) 2014-2017 Thomas Fussell
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
class XLNT_API uri
{
public:
    /// <summary>
    ///
    /// </summary>
    uri();

    /// <summary>
    ///
    /// </summary>
    uri(const uri &base, const uri &relative);

    /// <summary>
    ///
    /// </summary>
    uri(const uri &base, const path &relative);

    /// <summary>
    ///
    /// </summary>
    uri(const std::string &uri_string);

    /// <summary>
    ///
    /// </summary>
    bool is_relative() const;

    /// <summary>
    ///
    /// </summary>
    bool is_absolute() const;

    /// <summary>
    ///
    /// </summary>
    std::string scheme() const;

    /// <summary>
    ///
    /// </summary>
    std::string authority() const;

    /// <summary>
    ///
    /// </summary>
    bool has_authentication() const;

    /// <summary>
    ///
    /// </summary>
    std::string authentication() const;

    /// <summary>
    ///
    /// </summary>
    std::string username() const;

    /// <summary>
    ///
    /// </summary>
    std::string password() const;

    /// <summary>
    ///
    /// </summary>
    std::string host() const;

    /// <summary>
    ///
    /// </summary>
    bool has_port() const;

    /// <summary>
    ///
    /// </summary>
    std::size_t port() const;

    /// <summary>
    ///
    /// </summary>
    class path path() const;

    /// <summary>
    ///
    /// </summary>
    bool has_query() const;

    /// <summary>
    ///
    /// </summary>
    std::string query() const;

    /// <summary>
    ///
    /// </summary>
    bool has_fragment() const;

    /// <summary>
    ///
    /// </summary>
    std::string fragment() const;

    /// <summary>
    ///
    /// </summary>
    std::string to_string() const;

    /// <summary>
    ///
    /// </summary>
    uri make_absolute(const uri &base);

    /// <summary>
    ///
    /// </summary>
    uri make_reference(const uri &base);

    /// <summary>
    ///
    /// </summary>
    friend XLNT_API bool operator==(const uri &left, const uri &right);

private:
    /// <summary>
    ///
    /// </summary>
    bool absolute_ = false;

    /// <summary>
    ///
    /// </summary>
    std::string scheme_;

    /// <summary>
    ///
    /// </summary>
    bool has_authentication_ = false;

    /// <summary>
    ///
    /// </summary>
    std::string username_;

    /// <summary>
    ///
    /// </summary>
    std::string password_;

    /// <summary>
    ///
    /// </summary>
    std::string host_;

    /// <summary>
    ///
    /// </summary>
    bool has_port_ = false;

    /// <summary>
    ///
    /// </summary>
    std::size_t port_ = 0;

    /// <summary>
    ///
    /// </summary>
    bool has_query_ = false;

    /// <summary>
    ///
    /// </summary>
    std::string query_;

    /// <summary>
    ///
    /// </summary>
    bool has_fragment_ = false;

    /// <summary>
    ///
    /// </summary>
    std::string fragment_;

    /// <summary>
    ///
    /// </summary>
    class path path_;
};

} // namespace xlnt
