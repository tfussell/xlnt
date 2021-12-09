// Copyright (c) 2014-2021 Thomas Fussell
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

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/uri.hpp>

#include <string>
#include <vector>

namespace xml {
class parser;
class serializer;
} // namespace xml

namespace xlnt {

/// <summary>
/// A list of xml extensions that may or may not be understood by the parser
/// preservation is required for round-tripping even if extension is not understood
/// [serialised: extLst]
/// </summary>
class XLNT_API ext_list
{
public:
    struct ext
    {
        ext(xml::parser &parser, const std::string &ns);
        ext(const uri &ID, const std::string &serialised);
        void serialise(xml::serializer &serialiser, const std::string &ns);

        uri extension_ID_;
        std::string serialised_value_;
    };
    ext_list() = default; // default ctor required by xlnt::optional
    explicit ext_list(xml::parser &parser, const std::string &ns);
    void serialize(xml::serializer &serialiser, const std::string &ns);

    void add_extension(const uri &ID, const std::string &element);

    bool has_extension(const uri &extension_uri) const;

    const ext &extension(const uri &extension_uri) const;

    const std::vector<ext> &extensions() const;

    bool operator==(const ext_list &rhs) const;

private:
    std::vector<ext> extensions_;
};

inline bool operator==(const ext_list::ext &lhs, const ext_list::ext &rhs)
{
    return lhs.extension_ID_ == rhs.extension_ID_
        && lhs.serialised_value_ == rhs.serialised_value_;
}
} // namespace xlnt
