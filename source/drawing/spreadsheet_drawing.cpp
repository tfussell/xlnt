// Copyright (c) 2018
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

#include <xlnt/drawing/spreadsheet_drawing.hpp>
#include <detail/constants.hpp>

#include <detail/external/include_libstudxml.hpp>

namespace {
// copy elements to the serializer provided and extract the embed ids
// from blip elements
std::vector<std::string> copy_and_extract(xml::parser &p, xml::serializer &s)
{
    std::vector<std::string> embed_ids;
    int nest_level = 0;
    while (nest_level > 0 || (p.peek() != xml::parser::event_type::end_element && p.peek() != xml::parser::event_type::eof))
    {
        switch (p.next())
        {
        case xml::parser::start_element: {
            ++nest_level;
            auto attribs = p.attribute_map();
            auto current_ns = p.namespace_();
            s.start_element(p.qname());
            s.namespace_decl(current_ns, p.prefix());
            if (p.qname().name() == "blip")
            {
                embed_ids.push_back(attribs.at(xml::qname(xlnt::constants::ns("r"), "embed")).value);
            }
            p.peek();
            auto new_ns = p.namespace_();
            if (new_ns != current_ns)
            {
                auto pref = p.prefix();
                s.namespace_decl(new_ns, pref);
            }
            for (auto &ele : attribs)
            {
                s.attribute(ele.first, ele.second.value);
            }
            break;
        }
        case xml::parser::end_element: {
            --nest_level;
            s.end_element();
            break;
        }
        case xml::parser::start_namespace_decl: {
            s.namespace_decl(p.namespace_(), p.prefix());
            break;
        }
        case xml::parser::end_namespace_decl: { // nothing required here
            break;
        }
        case xml::parser::characters: {
            s.characters(p.value());
            break;
        }
        case xml::parser::eof:
            return embed_ids;
        case xml::parser::start_attribute:
        case xml::parser::end_attribute:
        default:
            break;
        }
    }
    return embed_ids;
}
} // namespace

namespace xlnt {
namespace drawing {

spreadsheet_drawing::spreadsheet_drawing(xml::parser &parser)
{
    std::ostringstream serialization_stream;
    xml::serializer s(serialization_stream, "", 0);
    embed_ids_ = copy_and_extract(parser, s);
    serialized_value_ = serialization_stream.str();
}

// void spreadsheet_drawing::serialize(xml::serializer &serializer, const std::string& ns)
void spreadsheet_drawing::serialize(xml::serializer &serializer)
{
    std::istringstream ser(serialized_value_);
    xml::parser p(ser, "", xml::parser::receive_default);
    copy_and_extract(p, serializer);
}

std::vector<std::string> spreadsheet_drawing::get_embed_ids()
{
    return embed_ids_;
}

} // namespace drawing
} // namespace xlnt
