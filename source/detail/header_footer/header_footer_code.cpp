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

#include <detail/header_footer/header_footer_code.hpp>
#include <xlnt/utils/serialisation_utils.hpp>

namespace xlnt {
namespace detail {

std::array<xlnt::optional<xlnt::rich_text>, 3> decode_header_footer(const std::string &hf_string)
{
    std::array<xlnt::optional<xlnt::rich_text>, 3> result;

    if (hf_string.empty())
    {
        return result;
    }

    enum class hf_code
    {
        left_section, // &L
        center_section, // &C
        right_section, // &R
        current_page_number, // &P
        total_page_number, // &N
        font_size, // &#
        text_font_color, // &KRRGGBB or &KTTSNN
        text_strikethrough, // &S
        text_superscript, // &X
        text_subscript, // &Y
        date, // &D
        time, // &T
        picture_as_background, // &G
        text_single_underline, // &U
        text_double_underline, // &E
        workbook_file_path, // &Z
        workbook_file_name, // &F
        sheet_tab_name, // &A
        add_to_page_number, // &+
        subtract_from_page_number, // &-
        text_font_name, // &"font name,font type"
        bold_font_style, // &B
        italic_font_style, // &I
        outline_style, // &O
        shadow_style, // &H
        text // everything else
    };

    struct hf_token
    {
        hf_code code = hf_code::text;
        std::string value;
    };

    std::vector<hf_token> tokens;
    std::size_t position = 0;

    while (position < hf_string.size())
    {
        hf_token token;

        auto next_ampersand = hf_string.find('&', position + 1);
        token.value = hf_string.substr(position, next_ampersand - position);
        auto next_position = next_ampersand;

        if (hf_string[position] == '&')
        {
            token.value.clear();
            next_position = position + 2;
            auto first_code_char = hf_string[position + 1];

            if (first_code_char == '"')
            {
                auto end_quote_index = hf_string.find('"', position + 2);
                next_position = end_quote_index + 1;

                token.value = hf_string.substr(position + 2, end_quote_index - position - 2); // remove quotes
                token.code = hf_code::text_font_name;
            }
            else if (first_code_char == '&')
            {
                token.value = "&&"; // escaped ampersand
            }
            else if (first_code_char == 'L')
            {
                token.code = hf_code::left_section;
            }
            else if (first_code_char == 'C')
            {
                token.code = hf_code::center_section;
            }
            else if (first_code_char == 'R')
            {
                token.code = hf_code::right_section;
            }
            else if (first_code_char == 'P')
            {
                token.code = hf_code::current_page_number;
            }
            else if (first_code_char == 'N')
            {
                token.code = hf_code::total_page_number;
            }
            else if (std::string("0123456789").find(hf_string[position + 1]) != std::string::npos)
            {
                token.code = hf_code::font_size;
                next_position = hf_string.find_first_not_of("0123456789", position + 1);
                token.value = hf_string.substr(position + 1, next_position - position - 1);
            }
            else if (first_code_char == 'K')
            {
                if (hf_string[position + 4] == '+' || hf_string[position + 4] == '-')
                {
                    token.value = hf_string.substr(position + 2, 5);
                    next_position = position + 7;
                }
                else
                {
                    token.value = hf_string.substr(position + 2, 6);
                    next_position = position + 8;
                }

                token.code = hf_code::text_font_color;
            }
            else if (first_code_char == 'S')
            {
                token.code = hf_code::text_strikethrough;
            }
            else if (first_code_char == 'X')
            {
                token.code = hf_code::text_superscript;
            }
            else if (first_code_char == 'Y')
            {
                token.code = hf_code::text_subscript;
            }
            else if (first_code_char == 'D')
            {
                token.code = hf_code::date;
            }
            else if (first_code_char == 'T')
            {
                token.code = hf_code::time;
            }
            else if (first_code_char == 'G')
            {
                token.code = hf_code::picture_as_background;
            }
            else if (first_code_char == 'U')
            {
                token.code = hf_code::text_single_underline;
            }
            else if (first_code_char == 'E')
            {
                token.code = hf_code::text_double_underline;
            }
            else if (first_code_char == 'Z')
            {
                token.code = hf_code::workbook_file_path;
            }
            else if (first_code_char == 'F')
            {
                token.code = hf_code::workbook_file_name;
            }
            else if (first_code_char == 'A')
            {
                token.code = hf_code::sheet_tab_name;
            }
            else if (first_code_char == '+')
            {
                token.code = hf_code::add_to_page_number;
            }
            else if (first_code_char == '-')
            {
                token.code = hf_code::subtract_from_page_number;
            }
            else if (first_code_char == 'B')
            {
                token.code = hf_code::bold_font_style;
            }
            else if (first_code_char == 'I')
            {
                token.code = hf_code::italic_font_style;
            }
            else if (first_code_char == 'O')
            {
                token.code = hf_code::outline_style;
            }
            else if (first_code_char == 'H')
            {
                token.code = hf_code::shadow_style;
            }
        }

        position = next_position;
        tokens.push_back(token);
    }

    const auto parse_section = [&tokens, &result](hf_code code)
    {
        std::vector<hf_code> end_codes{hf_code::left_section, hf_code::center_section, hf_code::right_section};
        end_codes.erase(std::find(end_codes.begin(), end_codes.end(), code));

        std::size_t start_index = 0;

        while (start_index < tokens.size() && tokens[start_index].code != code)
        {
            ++start_index;
        }

        if (start_index == tokens.size())
        {
            return;
        }

        ++start_index; // skip the section code
        std::size_t end_index = start_index;

        while (end_index < tokens.size()
            && std::find(end_codes.begin(), end_codes.end(), tokens[end_index].code) == end_codes.end())
        {
            ++end_index;
        }

        xlnt::rich_text current_text;
        xlnt::rich_text_run current_run;

        // todo: all this nice parsing and the codes are just being turned back into text representations
        // It would be nice to create an interface for the library to read and write these codes

        for (auto i = start_index; i < end_index; ++i)
        {
            const auto &current_token = tokens[i];

            if (current_token.code == hf_code::text)
            {
                current_run.first = current_run.first + current_token.value;
                continue;
            }

            if (!current_run.first.empty())
            {
                current_text.add_run(current_run);
                current_run = xlnt::rich_text_run();
            }

            switch (current_token.code)
            {
            case hf_code::text:
                {
                    break; // already handled above
                }

            case hf_code::left_section:
                {
                    break; // used below
                }

            case hf_code::center_section:
                {
                    break; // used below
                }

            case hf_code::right_section:
                {
                    break; // used below
                }

            case hf_code::current_page_number:
                {
                    current_run.first = current_run.first + "&P";
                    break;
                }

            case hf_code::total_page_number:
                {
                    current_run.first = current_run.first + "&N";
                    break;
                }

            case hf_code::font_size:
                {
                    if (!current_run.second.is_set())
                    {
                        current_run.second = xlnt::font();
                    }

                    current_run.second.get().size(std::stod(current_token.value));

                    break;
                }

            case hf_code::text_font_color:
                {
                    if (current_token.value.size() == 6)
                    {
                        if (!current_run.second.is_set())
                        {
                            current_run.second = xlnt::font();
                        }

                        current_run.second.get().color(xlnt::rgb_color(current_token.value));
                    }

                    break;
                }

            case hf_code::text_strikethrough:
                {
                    break;
                }

            case hf_code::text_superscript:
                {
                    break;
                }

            case hf_code::text_subscript:
                {
                    break;
                }

            case hf_code::date:
                {
                    current_run.first = current_run.first + "&D";
                    break;
                }

            case hf_code::time:
                {
                    current_run.first = current_run.first + "&T";
                    break;
                }

            case hf_code::picture_as_background:
                {
                    current_run.first = current_run.first + "&G";
                    break;
                }

            case hf_code::text_single_underline:
                {
                    if (!current_run.second.is_set())
                    {
                        current_run.second = xlnt::font();
                    }
                    current_run.second.get().underline(font::underline_style::single);
                    break;
                }

            case hf_code::text_double_underline:
                {
                    break;
                }

            case hf_code::workbook_file_path:
                {
                    current_run.first = current_run.first + "&Z";
                    break;
                }

            case hf_code::workbook_file_name:
                {
                    current_run.first = current_run.first + "&F";
                    break;
                }

            case hf_code::sheet_tab_name:
                {
                    current_run.first = current_run.first + "&A";
                    break;
                }

            case hf_code::add_to_page_number:
                {
                    break;
                }

            case hf_code::subtract_from_page_number:
                {
                    break;
                }

            case hf_code::text_font_name:
                {
                    auto comma_index = current_token.value.find(',');
                    auto font_name = current_token.value.substr(0, comma_index);

                    if (!current_run.second.is_set())
                    {
                        current_run.second = xlnt::font();
                    }

                    if (font_name != "-")
                    {
                        current_run.second.get().name(font_name);
                    }

                    if (comma_index != std::string::npos)
                    {
                        auto font_type = current_token.value.substr(comma_index + 1);

                        if (font_type == "Bold")
                        {
                            current_run.second.get().bold(true);
                        }
                        else if (font_type == "Italic")
                        {
                            // TODO
                        }
                        else if (font_type == "BoldItalic")
                        {
                            current_run.second.get().bold(true);
                        }
                    }

                    break;
                }

            case hf_code::bold_font_style:
                {
                    if (!current_run.second.is_set())
                    {
                        current_run.second = xlnt::font();
                    }

                    current_run.second.get().bold(true);

                    break;
                }

            case hf_code::italic_font_style:
                {
                    break;
                }

            case hf_code::outline_style:
                {
                    break;
                }

            case hf_code::shadow_style:
                {
                    break;
                }
            }
        }

        if (!current_run.first.empty())
        {
            current_text.add_run(current_run);
        }

        auto location_index =
            static_cast<std::size_t>(code == hf_code::left_section ? 0 : code == hf_code::center_section ? 1 : 2);

        if (!current_text.plain_text().empty())
        {
            result[location_index] = current_text;
        }
    };

    parse_section(hf_code::left_section);
    parse_section(hf_code::center_section);
    parse_section(hf_code::right_section);

    return result;
}

std::string encode_header_footer(const rich_text &t, header_footer::location where)
{
    const auto location_code_map =
        std::unordered_map<header_footer::location, 
            std::string, scoped_enum_hash<header_footer::location>>
        {
                { header_footer::location::left, "&L" },
                { header_footer::location::center, "&C" },
                { header_footer::location::right, "&R" },
        };

    auto encoded = location_code_map.at(where);

    for (const auto &run : t.runs())
    {
        if (run.first.empty()) continue;

        if (run.second.is_set())
        {
            if (run.second.get().has_name())
            {
                encoded.push_back('&');
                encoded.push_back('"');
                encoded.append(run.second.get().name());
                encoded.push_back(',');

                if (run.second.get().bold())
                {
                    encoded.append("Bold");
                }
                else
                {
                    encoded.append("Regular");
                }
                // todo: BoldItalic?

                encoded.push_back('"');
            }
            else if (run.second.get().bold())
            {
                encoded.append("&B");
            }

            if (run.second.get().has_size())
            {
                encoded.push_back('&');
                encoded.append(serialize_number_to_string(run.second.get().size()));
            }
            if (run.second.get().underlined())
            {
                switch (run.second.get().underline()) 
                {
                case font::underline_style::single:
                case font::underline_style::single_accounting:
                    encoded.append("&U");
                    break;
                case font::underline_style::double_:
                case font::underline_style::double_accounting:
                    encoded.append("&E");
                    break;
                case font::underline_style::none:
                    break;
                }
            }
            if (run.second.get().has_color())
            {
                encoded.push_back('&');
                encoded.push_back('K');
                encoded.append(run.second.get().color().rgb().hex_string().substr(2));
            }
        }

        encoded.append(run.first);
    }

    return encoded;
};

} // namespace detail
} // namespace xlnt
