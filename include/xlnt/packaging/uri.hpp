// Copyright (c) 2014-2018 Thomas Fussell
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
/// Encapsulates a uniform resource identifier (URI) as described
/// by RFC 3986.
/// </summary>
class XLNT_API uri
{
public:
    /// <summary>
    /// Constructs an empty URI.
    /// </summary>
    uri();

    /// <summary>
    /// Constructs a URI by combining base with relative.
    /// </summary>
    uri(const uri &base, const uri &relative);

    /// <summary>
    /// Constructs a URI by combining base with relative path.
    /// </summary>
    uri(const uri &base, const path &relative);

    /// <summary>
    /// Constructs a URI by parsing the given uri_string.
    /// </summary>
    uri(const std::string &uri_string);

    /// <summary>
    /// Returns true if this URI is relative.
    /// </summary>
    bool is_relative() const;

    /// <summary>
    /// Returns true if this URI is not relative (i.e. absolute).
    /// </summary>
    bool is_absolute() const;

    /// <summary>
    /// Returns the scheme of this URI.
    /// E.g. the scheme of http://user:pass@example.com is "http"
    /// </summary>
    std::string scheme() const;

    /// <summary>
    /// Returns the authority of this URI.
    /// E.g. the authority of http://user:pass@example.com:80/document is "user:pass@example.com:80"
    /// </summary>
    std::string authority() const;

    /// <summary>
    /// Returns true if an authentication section is specified for this URI.
    /// </summary>
    bool has_authentication() const;

    /// <summary>
    /// Returns the authentication of this URI.
    /// E.g. the authentication of http://user:pass@example.com is "user:pass"
    /// </summary>
    std::string authentication() const;

    /// <summary>
    /// Returns the username of this URI.
    /// E.g. the username of http://user:pass@example.com is "user"
    /// </summary>
    std::string username() const;

    /// <summary>
    /// Returns the password of this URI.
    /// E.g. the password of http://user:pass@example.com is "pass"
    /// </summary>
    std::string password() const;

    /// <summary>
    /// Returns the host of this URI.
    /// E.g. the host of http://example.com:80/document is "example.com"
    /// </summary>
    std::string host() const;

    /// <summary>
    /// Returns true if a non-default port is specified for this URI.
    /// </summary>
    bool has_port() const;

    /// <summary>
    /// Returns the port of this URI.
    /// E.g. the port of https://example.com:443/document is "443"
    /// </summary>
    std::size_t port() const;

    /// <summary>
    /// Returns the path of this URI.
    /// E.g. the path of http://example.com/document is "/document"
    /// </summary>
    const class path& path() const;

    /// <summary>
    /// Returns true if this URI has a non-null query string section.
    /// </summary>
    bool has_query() const;

    /// <summary>
    /// Returns the query string of this URI.
    /// E.g. the query of http://example.com/document?v=1&x=3#abc is "v=1&x=3"
    /// </summary>
    std::string query() const;

    /// <summary>
    /// Returns true if this URI has a non-empty fragment section.
    /// </summary>
    bool has_fragment() const;

    /// <summary>
    /// Returns the fragment section of this URI.
    /// E.g. the fragment of http://example.com/document#abc is "abc"
    /// </summary>
    std::string fragment() const;

    /// <summary>
    /// Returns a string representation of this URI.
    /// </summary>
    std::string to_string() const;

    /// <summary>
    /// If this URI is relative, an absolute URI will be returned by appending
    /// the path to the given absolute base URI.
    /// </summary>
    uri make_absolute(const uri &base);

    /// <summary>
    /// If this URI is absolute, a relative URI will be returned by removing the
    /// common base path from the given absolute base URI.
    /// </summary>
    uri make_reference(const uri &base);

    /// <summary>
    /// Returns true if this URI is equivalent to other.
    /// </summary>
    bool operator==(const uri &other) const;

private:
    /// <summary>
    /// True if this URI is absolute.
    /// </summary>
    bool absolute_ = false;

    /// <summary>
    /// The scheme, like "http"
    /// </summary>
    std::string scheme_;

    /// <summary>
    /// True if this URI has an authentication section.
    /// </summary>
    bool has_authentication_ = false;

    /// <summary>
    /// The username
    /// </summary>
    std::string username_;

    /// <summary>
    /// The password
    /// </summary>
    std::string password_;

    /// <summary>
    /// The host
    /// </summary>
    std::string host_;

    /// <summary>
    /// True if this URI has a non-default port specified
    /// </summary>
    bool has_port_ = false;

    /// <summary>
    /// The numeric port
    /// </summary>
    std::size_t port_ = 0;

    /// <summary>
    /// True if this URI has a query section
    /// </summary>
    bool has_query_ = false;

    /// <summary>
    /// The query section
    /// </summary>
    std::string query_;

    /// <summary>
    /// True if this URI has a fragment section
    /// </summary>
    bool has_fragment_ = false;

    /// <summary>
    /// The fragment section
    /// </summary>
    std::string fragment_;

    /// <summary>
    /// The path section
    /// </summary>
    class path path_;
};

} // namespace xlnt
