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

#include <detail/number_formatter.hpp>

namespace xlnt {
namespace detail {

const std::unordered_map<int, std::string> known_locales()
{
    const std::unordered_map<int, std::string> *all =
        new std::unordered_map<int, std::string>(
            {
                 { 0x401, "Arabic - Saudi Arabia" },
                 { 0x402, "Bulgarian" },
                 { 0x403, "Catalan" },
                 { 0x404, "Chinese - Taiwan" },
                 { 0x405, "Czech" },
                 { 0x406, "Danish" },
                 { 0x407, "German - Germany" },
                 { 0x408, "Greek" },
                 { 0x409, "English - United States" },
                 { 0x410, "Italian - Italy" },
                 { 0x411, "Japanese" },
                 { 0x412, "Korean" },
                 { 0x413, "Dutch - Netherlands" },
                 { 0x414, "Norwegian - Bokml" },
                 { 0x415, "Polish" },
                 { 0x416, "Portuguese - Brazil" },
                 { 0x417, "Raeto-Romance" },
                 { 0x418, "Romanian - Romania" },
                 { 0x419, "Russian" },
                 { 0x420, "Urdu" },
                 { 0x421, "Indonesian" },
                 { 0x422, "Ukrainian" },
                 { 0x423, "Belarusian" },
                 { 0x424, "Slovenian" },
                 { 0x425, "Estonian" },
                 { 0x426, "Latvian" },
                 { 0x427, "Lithuanian" },
                 { 0x428, "Tajik" },
                 { 0x429, "Farsi - Persian" },
                 { 0x430, "Sesotho (Sutu)" },
                 { 0x431, "Tsonga" },
                 { 0x432, "Setsuana" },
                 { 0x433, "Venda" },
                 { 0x434, "Xhosa" },
                 { 0x435, "Zulu" },
                 { 0x436, "Afrikaans" },
                 { 0x437, "Georgian" },
                 { 0x438, "Faroese" },
                 { 0x439, "Hindi" },
                 { 0x440, "Kyrgyz - Cyrillic" },
                 { 0x441, "Swahili" },
                 { 0x442, "Turkmen" },
                 { 0x443, "Uzbek - Latin" },
                 { 0x444, "Tatar" },
                 { 0x445, "Bengali - India" },
                 { 0x446, "Punjabi" },
                 { 0x447, "Gujarati" },
                 { 0x448, "Oriya" },
                 { 0x449, "Tamil" },
                 { 0x450, "Mongolian" },
                 { 0x451, "Tibetan" },
                 { 0x452, "Welsh" },
                 { 0x453, "Khmer" },
                 { 0x454, "Lao" },
                 { 0x455, "Burmese" },
                 { 0x456, "Galician" },
                 { 0x457, "Konkani" },
                 { 0x458, "Manipuri" },
                 { 0x459, "Sindhi" },
                 { 0x460, "Kashmiri" },
                 { 0x461, "Nepali" },
                 { 0x462, "Frisian - Netherlands" },
                 { 0x464, "Filipino" },
                 { 0x465, "Divehi; Dhivehi; Maldivian" },
                 { 0x466, "Edo" },
                 { 0x470, "Igbo - Nigeria" },
                 { 0x474, "Guarani - Paraguay" },
                 { 0x476, "Latin" },
                 { 0x477, "Somali" },
                 { 0x481, "Maori" },
                 { 0x801, "Arabic - Iraq" },
                 { 0x804, "Chinese - China" },
                 { 0x807, "German - Switzerland" },
                 { 0x809, "English - Great Britain" },
                 { 0x810, "Italian - Switzerland" },
                 { 0x813, "Dutch - Belgium" },
                 { 0x814, "Norwegian - Nynorsk" },
                 { 0x816, "Portuguese - Portugal" },
                 { 0x818, "Romanian - Moldova" },
                 { 0x819, "Russian - Moldova" },
                 { 0x843, "Uzbek - Cyrillic" },
                 { 0x845, "Bengali - Bangladesh" },
                 { 0x850, "Mongolian" },
                 { 0x1001, "Arabic - Libya" },
                 { 0x1004, "Chinese - Singapore" },
                 { 0x1007, "German - Luxembourg" },
                 { 0x1009, "English - Canada" },
                 { 0x1401, "Arabic - Algeria" },
                 { 0x1404, "Chinese - Macau SAR" },
                 { 0x1407, "German - Liechtenstein" },
                 { 0x1409, "English - New Zealand" },
                 { 0x1801, "Arabic - Morocco" },
                 { 0x1809, "English - Ireland" },
                 { 0x2001, "Arabic - Oman" },
                 { 0x2009, "English - Jamaica" },
                 { 0x2401, "Arabic - Yemen" },
                 { 0x2409, "English - Caribbean" },
                 { 0x2801, "Arabic - Syria" },
                 { 0x2809, "English - Belize" },
                 { 0x3001, "Arabic - Lebanon" },
                 { 0x3009, "English - Zimbabwe" },
                 { 0x3401, "Arabic - Kuwait" },
                 { 0x3409, "English - Phillippines" },
                 { 0x3801, "Arabic - United Arab Emirates" },
                 { 0x4001, "Arabic - Qatar" }
            });
    
    return *all;
}

bool condition::satisfied_by(long double number) const
{
    switch (type)
    {
    case condition_type::equal:
        return number == value;
    case condition_type::greater_or_equal:
        return number >= value;
    case condition_type::greater_than:
        return number > value;
    case condition_type::less_or_equal:
        return number <= value;
    case condition_type::less_than:
        return number < value;
    case condition_type::not_equal:
        return number != value;
    case condition_type::none:
        return false;
    }
}

number_format_parser::number_format_parser(const std::string &format_string)
{
    reset(format_string);
}

const std::vector<format_code> &number_format_parser::get_result() const
{
    return codes_;
}

void number_format_parser::reset(const std::string &format_string)
{
    format_string_ = format_string;
    position_ = 0;
    codes_.clear();
}

void number_format_parser::parse()
{
    auto token = parse_next_token();
    format_code section;
    template_part part;

    for (;;)
    {
        switch (token.type)
        {
        case number_format_token::token_type::end_section:
            codes_.push_back(section);
            section = format_code();

            break;

        case number_format_token::token_type::bad:
            throw std::runtime_error("bad format");

        case number_format_token::token_type::color:
            if (section.color != bracket_color::none
                || section.condition.type != condition::condition_type::none
                || section.locale != locale::none
                || !section.parts.empty())
            {
                throw std::runtime_error("color should be the first part of a format");
            }

            section.color = color_from_string(token.string);
            break;

        case number_format_token::token_type::locale:
        {
            if (section.locale != locale::none)
            {
                throw std::runtime_error("multiple locales");
            }

            auto parsed_locale = locale_from_string(token.string);
            section.locale = parsed_locale.first;

            if (!parsed_locale.second.empty())
            {
                part.type = template_part::template_type::text;
                part.string = parsed_locale.second;
                section.parts.push_back(part);
                part = template_part();
            }
            
            break;
        }

        case number_format_token::token_type::condition:
        {
            if (section.condition.type != condition::condition_type::none)
            {
                throw std::runtime_error("multiple conditions");
            }

            std::string value;

            if (token.string.front() == '<')
            {
                if (token.string[1] == '=')
                {
                    section.condition.type = condition::condition_type::less_or_equal;
                    value = token.string.substr(2);
                }
                else if (token.string[1] == '>')
                {
                    section.condition.type = condition::condition_type::not_equal;
                    value = token.string.substr(2);
                }
                else
                {
                    section.condition.type = condition::condition_type::less_than;
                    value = token.string.substr(1);
                }
            }
            else if (token.string.front() == '>')
            {
                if (token.string[1] == '=')
                {
                    section.condition.type = condition::condition_type::greater_or_equal;
                    value = token.string.substr(2);
                }
                else
                {
                    section.condition.type = condition::condition_type::greater_than;
                    value = token.string.substr(1);
                }
            }
            else if (token.string.front() == '=')
            {
                section.condition.type = condition::condition_type::equal;
                value = token.string.substr(1);
            }
            
            section.condition.value = std::stold(value);
            break;
        }

        case number_format_token::token_type::text:
            part.type = template_part::template_type::text;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;
            
        case number_format_token::token_type::fill:
            part.type = template_part::template_type::fill;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;

        case number_format_token::token_type::space:
            part.type = template_part::template_type::space;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;
        case number_format_token::token_type::number:
            part.type = template_part::template_type::general;
            part.placeholders = parse_placeholders(token.string);
            section.parts.push_back(part);
            part = template_part();

            break;
        case number_format_token::token_type::datetime:
            section.is_datetime = true;

            switch (token.string.front())
            {
            case '[':
                section.is_timedelta = true;

                if (token.string == "[h]" || token.string == "[hh]")
                {
                    part.type = template_part::template_type::elapsed_hours;
                    break;
                }
                else if (token.string == "[m]" || token.string == "[mm]")
                {
                    part.type = template_part::template_type::elapsed_minutes;
                    break;
                }
                else if (token.string == "[s]" || token.string == "[ss]")
                {
                    part.type = template_part::template_type::elapsed_seconds;
                    break;
                }
                else
                {
                    throw std::runtime_error("expected [h], [m], or [s]");
                }
            case 'm':
                if (token.string == "m")
                {
                    part.type = template_part::template_type::month_number;
                    break;
                }
                else if (token.string == "mm")
                {
                    part.type = template_part::template_type::month_number_leading_zero;
                    break;
                }
                else if (token.string == "mmm")
                {
                    part.type = template_part::template_type::month_abbreviation;
                    break;
                }
                else if (token.string == "mmmm")
                {
                    part.type = template_part::template_type::month_name;
                    break;
                }
                else if (token.string == "mmmmm")
                {
                    part.type = template_part::template_type::month_letter;
                    break;
                }
            case 'd':
                if (token.string == "d")
                {
                    part.type = template_part::template_type::day_number;
                    break;
                }
                else if (token.string == "dd")
                {
                    part.type = template_part::template_type::day_number_leading_zero;
                    break;
                }
                else if (token.string == "ddd")
                {
                    part.type = template_part::template_type::day_abbreviation;
                    break;
                }
                else if (token.string == "dddd")
                {
                    part.type = template_part::template_type::day_name;
                    break;
                }
            case 'y':
                if (token.string == "yy")
                {
                    part.type = template_part::template_type::year_short;
                    break;
                }
                else if (token.string == "yyyy")
                {
                    part.type = template_part::template_type::year_long;
                    break;
                }
            case 'h':
                if (token.string == "h")
                {
                    part.type = template_part::template_type::hour;
                    break;
                }
                else if (token.string == "hh")
                {
                    part.type = template_part::template_type::hour_leading_zero;
                    break;
                }
            case 's':
                if (token.string == "s")
                {
                    part.type = template_part::template_type::second;
                    break;
                }
                else if (token.string == "ss")
                {
                    part.type = template_part::template_type::second_leading_zero;
                    break;
                }
            case 'A':
                section.twelve_hour = true;

                if (token.string == "AM/PM")
                {
                    part.type = template_part::template_type::am_pm;
                    break;
                }
                else if (token.string == "A/P")
                {
                    part.type = template_part::template_type::a_p;
                    break;
                }
            }

            section.parts.push_back(part);
            part = template_part();
            
            break;
            
        case number_format_token::token_type::end:
            codes_.push_back(section);
            finalize();

            return;
            
        default:
            break;
        }

        token = parse_next_token();
    }

    throw std::runtime_error("bad format");
}

void number_format_parser::finalize()
{
    for (auto &code : codes_)
    {
        bool fix = false;
        bool leading_zero = false;
        std::size_t minutes_index = 0;

        for (std::size_t i = 0; i < code.parts.size(); ++i)
        {
            const auto &part = code.parts[i];

            if (part.type == template_part::template_type::month_number
                || part.type == template_part::template_type::month_number_leading_zero)
            {
                if (i < code.parts.size() - 2)
                {
                    const auto &next = code.parts[i + 1];
                    const auto &after_next = code.parts[i + 2];
                    
                    if (next.type == template_part::template_type::text
                        && next.string == ":"
                        && (after_next.type == template_part::template_type::second ||
                            after_next.type == template_part::template_type::second_leading_zero))
                    {
                        fix = true;
                        leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                        minutes_index = i;

                        break;
                    }
                }

                if (i > 1)
                {
                    const auto &previous = code.parts[i - 1];
                    const auto &before_previous = code.parts[i - 2];

                    if (previous.type == template_part::template_type::text
                        && previous.string == ":"
                        && (before_previous.type == template_part::template_type::hour ||
                            before_previous.type == template_part::template_type::hour_leading_zero))
                    {
                        fix = true;
                        leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                        minutes_index = i;
                    }
                }
            }
        }
        
        if (fix)
        {
            code.parts[minutes_index].type = leading_zero ?
                template_part::template_type::minute_leading_zero :
                template_part::template_type::minute;
        }
    }

    validate();
}

number_format_token number_format_parser::parse_next_token()
{
    number_format_token token;

    if (format_string_.size() <= position_)
    {
        token.type = number_format_token::token_type::end;
        return token;
    }

    auto current_char = format_string_[position_++];

    switch (current_char)
    {
    case '[':
        do
        {
            token.string.push_back(format_string_[position_++]);
        }
        while (format_string_[position_] != ']' && position_ < format_string_.size());

        if (position_ == format_string_.size())
        {
            throw std::runtime_error("missing ]");
        }

        if (token.string.empty())
        {
            throw std::runtime_error("empty []");
        }
        else if (token.string[0] == '<' || token.string[0] == '>' || token.string[0] == '=')
        {
            token.type = number_format_token::token_type::condition;
        }
        else if (token.string[0] == '$')
        {
            token.type = number_format_token::token_type::locale;
        }
        else if (token.string.size() <= 2 &&
            ((token.string == "h" || token.string == "hh") ||
            (token.string == "m" || token.string == "mm") ||
            (token.string == "s" || token.string == "ss")))
        {
            token.type = number_format_token::token_type::datetime;
            token.string = "[" + token.string + "]";
        }
        else
        {
            token.type = number_format_token::token_type::color;
            color_from_string(token.string);
        }

        ++position_;

        break;

    case '\\':
        token.type = number_format_token::token_type::text;
        token.string.push_back(format_string_[position_++]);

        break;

    case 'G':
        if (format_string_.substr(position_ - 1, 7) != "General")
        {
            throw std::runtime_error("expected General");
        }

        token.type = number_format_token::token_type::number;
        token.string = "General";
        position_ += 6;

        break;

    case '_':
        token.type = number_format_token::token_type::space;
        token.string.push_back(format_string_[position_++]);

        break;
        
    case '0':
    case '#':
    case '?':
    case '.':
        token.type = number_format_token::token_type::number;

        do
        {
            token.string.push_back(current_char);
            current_char = format_string_[position_];
            
            if (current_char != '.')
            {
                ++position_;
            }
        }
        while (current_char == '0' || current_char == '#' || current_char == '?' || current_char == ',');

        break;
        
    case 'y':
    case 'm':
    case 'd':
    case 'h':
    case 's':
        token.type = number_format_token::token_type::datetime;
        token.string.push_back(current_char);
        
        while (format_string_[position_] == current_char)
        {
            token.string.push_back(current_char);
            ++position_;
        }

        break;

    case 'A':
        token.type = number_format_token::token_type::datetime;

        if (format_string_.substr(position_ - 1, 5) == "AM/PM")
        {
            position_ += 4;
            token.string = "AM/PM";
        }
        else if (format_string_.substr(position_ - 1, 3) == "A/P")
        {
            position_ += 2;
            token.string = "A/P";
        }
        else
        {
            throw std::runtime_error("expected AM/PM or A/P");
        }

        break;

    case '"':
    {
        token.type = number_format_token::token_type::text;
        auto start = position_;
        auto end = format_string_.find('"', position_);

        while (end != std::string::npos && format_string_[end - 1] == '\\')
        {
            end = format_string_.find('"', position_);
        }

        token.string = format_string_.substr(start, end - start);
        position_ = end + 1;

        break;
    }

    case ';':
        token.type = number_format_token::token_type::end_section;
        break;
        
    case '(':
    case ')':
    case '-':
    case '+':
    case ':':
    case ' ':
    case '/':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;
        
    case '@':
        token.type = number_format_token::token_type::number;
        token.string.push_back(current_char);
        break;

    default:
        throw std::runtime_error("unexpected character");
    }

    return token;
}

void number_format_parser::validate()
{
    if (codes_.size() > 4)
    {
        throw std::runtime_error("too many format codes");
    }

    if (codes_.size() > 2)
    {
        if (codes_[0].condition.type != condition::condition_type::none &&
            codes_[1].condition.type != condition::condition_type::none &&
            codes_[2].condition.type != condition::condition_type::none)
        {
            throw std::runtime_error("format should have a maximum of two codes with conditions");
        }
    }
}

placeholders number_format_parser::parse_placeholders(const std::string &placeholders_string)
{
    placeholders p;

    if (placeholders_string == "General")
    {
        p.type = placeholders::placeholders_type::general;
        return p;
    }
    else if (placeholders_string == "@")
    {
        p.type = placeholders::placeholders_type::text;
        return p;
    }
    else if (placeholders_string.front() == '.')
    {
        p.type = placeholders::placeholders_type::fractional_part;
    }

    std::vector<std::size_t> comma_indices;

    for (std::size_t i = 0; i < placeholders_string.size(); ++i)
    {
        auto c = placeholders_string[i];

        if (c == '0')
        {
            ++p.num_zeros;
        }
        else if (c == '#')
        {
            ++p.num_optionals;
        }
        else if (c == '?')
        {
            ++p.num_spaces;
        }
        else if (c == ',')
        {
            comma_indices.push_back(i);
        }
    }
    
    if (!comma_indices.empty())
    {
        std::size_t i = placeholders_string.size();

        while (i == comma_indices.back())
        {
            ++p.thousands_scale;
            --i;
            comma_indices.pop_back();
        }
        
        p.use_comma_separator = !comma_indices.empty();
    }
    
    return p;
}

bracket_color number_format_parser::color_from_string(const std::string &color)
{
    switch (color[0])
    {
    case 'C':
        if (color == "Cyan")
        {
            return bracket_color::cyan;
        }
        else if (color.substr(0, 5) == "Color")
        {
            auto color_number = std::stoull(color.substr(5));

            if (color_number >= 1 && color_number <= 56)
            {
                return static_cast<bracket_color>(color_number);
            }
        }
    case 'B':
        if (color == "Black")
        {
            return bracket_color::black;
        }
        else if (color == "Blue")
        {
            return bracket_color::blue;
        }
    case 'G':
        if (color == "Green")
        {
            return bracket_color::black;
        }
    case 'W':
        if (color == "White")
        {
            return bracket_color::white;
        }
    case 'M':
        if (color == "Magenta")
        {
            return bracket_color::white;
        }
        return bracket_color::magenta;
    case 'Y':
        if (color == "Yellow")
        {
            return bracket_color::white;
        }
    case 'R':
        if (color == "Red")
        {
            return bracket_color::white;
        }
    default:
        throw std::runtime_error("bad color: " + color);
    }
}

std::pair<locale, std::string> number_format_parser::locale_from_string(const std::string &locale_string)
{
    auto hyphen_index = locale_string.find('-');

    if (locale_string.empty() || locale_string.front() != '$' || hyphen_index == std::string::npos)
    {
        throw std::runtime_error("bad locale: " + locale_string);
    }

    std::pair<locale, std::string> result;

    if (hyphen_index > 1)
    {
        result.second = locale_string.substr(1, hyphen_index - 1);
    }

    auto country_code_string = locale_string.substr(hyphen_index + 1);

    if (country_code_string.empty())
    {
        throw std::runtime_error("bad locale: " + locale_string);
    }
    
    for (auto c : country_code_string)
    {
        if (!((c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') || (c >= '0' && c <= '9')))
        {
            throw std::runtime_error("bad locale: " + locale_string);
        }
    }

    auto country_code = std::stoi(country_code_string, nullptr, 16);

    for (const auto &known_locale : known_locales())
    {
        if (known_locale.first == country_code)
        {
            result.first = static_cast<locale>(country_code);
            return result;
        }
    }

    throw std::runtime_error("unknown country code: " + country_code_string);
}

number_formatter::number_formatter(const std::string &format_string, xlnt::calendar calendar)
    : parser_(format_string),
      calendar_(calendar)
{
    parser_.parse();
    format_ = parser_.get_result();
}

std::string number_formatter::format_number(long double number)
{
    if (format_[0].condition.type != condition::condition_type::none)
    {
        if (format_[0].condition.satisfied_by(number))
        {
            return format_number(format_[0], number);
        }

        if (format_.size() == 1)
        {
            return std::string(11, '#');
        }
        
        if (format_[1].condition.type == condition::condition_type::none || format_[1].condition.satisfied_by(number))
        {
            return format_number(format_[1], number);
        }
        
        if (format_.size() == 2)
        {
            return std::string(11, '#');
        }
        
        return format_number(format_[2], number);
    }
    
    // no conditions, format based on sign:

    // 1 section, use for all
    if (format_.size() == 1)
    {
        return format_number(format_[0], number);
    }
    // 2 sections, first for positive and zero, second for negative
    else if (format_.size() == 2)
    {
        if (number >= 0)
        {
            return format_number(format_[0], number);
        }
        else
        {
            return format_number(format_[1], std::abs(number));
        }
    }
    // 3+ sections, first for positive, second for negative, third for zero
    else
    {
        if (number > 0)
        {
            return format_number(format_[0], number);
        }
        else if (number < 0)
        {
            return format_number(format_[1], std::abs(number));
        }
        else
        {
            return format_number(format_[2], number);
        }
    }
}

std::string number_formatter::format_text(const std::string &text)
{
    if (format_.size() < 4)
    {
        return format_text(format_[0], text);
    }

    return format_text(format_[3], text);
}

std::string number_formatter::fill_placeholders(const placeholders &p, long double number)
{
    if (p.type == placeholders::placeholders_type::general)
    {
        auto result = std::to_string(number);

        while (result.back() == '0')
        {
            result.pop_back();
        }

        if (result.back() == '.')
        {
            result.pop_back();
        }

        return result;
    }

    auto integer_part = static_cast<int>(number);

    if (p.type != placeholders::placeholders_type::fractional_part)
    {
        auto result = std::to_string(integer_part);
        
        while (result.size() < p.num_zeros)
        {
            result = "0" + result;
        }
        
        while (result.size() < p.num_zeros + p.num_spaces)
        {
            result = " " + result;
        }
        
        if (p.use_comma_separator)
        {
            std::vector<char> digits(result.rbegin(), result.rend());
            std::string temp;
            
            for (std::size_t i = 0; i < digits.size(); i++)
            {
                temp.push_back(digits[i]);

                if (i % 3 == 2)
                {
                    temp.push_back(',');
                }
            }
            
            result = std::string(temp.rbegin(), temp.rend());
        }

        return result;
    }
    else
    {
        auto fractional_part = number - integer_part;
        auto result = std::to_string(fractional_part).substr(1);

        while (result.back() == '0')
        {
            result.pop_back();
        }

        while (result.size() < p.num_zeros + 1)
        {
            result.push_back('0');
        }

        while (result.size() < p.num_zeros + p.num_spaces + 1)
        {
            result.push_back(' ');
        }

        return result;
    }
}

std::string number_formatter::format_number(const format_code &format, long double number)
{
    static const std::vector<std::string> *month_names =
    new std::vector<std::string>
    {
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December"
    };
    
    std::string result;

    if (number < 0)
    {
        result.push_back('-');
    }

    number = std::abs(number);

    xlnt::datetime dt(0, 0, 0);

    if (format.is_datetime)
    {
        dt = xlnt::datetime::from_number(number, calendar_);
    }

    for (const auto &part : format.parts)
    {
        switch (part.type)
        {
        case template_part::template_type::space:
            result.push_back(' ');
            break;
        case template_part::template_type::text:
            result.append(part.string);
            break;
        case template_part::template_type::general:
        {
            result.append(fill_placeholders(part.placeholders, number));
            break;
        }
        case template_part::template_type::day_number:
            result.append(std::to_string(dt.day));
            break;
        case template_part::template_type::day_number_leading_zero:
            if (dt.day < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.day));
            break;
        case template_part::template_type::month_abbreviation:
            result.append(month_names->at(dt.month - 1).substr(0, 3));
            break;
        case template_part::template_type::month_name:
            result.append(month_names->at(dt.month - 1));
            break;
        case template_part::template_type::month_number:
            result.append(std::to_string(dt.month));
            break;
        case template_part::template_type::month_number_leading_zero:
            if (dt.month < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.month));
            break;
        case template_part::template_type::year_short:
            if (dt.year % 1000 < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.year % 1000));
            break;
        case template_part::template_type::year_long:
            result.append(std::to_string(dt.year));
            break;
        case template_part::template_type::hour:
            result.append(std::to_string(format.twelve_hour ? dt.hour % 12 : dt.hour));
            break;
        case template_part::template_type::hour_leading_zero:
            if (format.twelve_hour ? dt.hour % 12 : dt.hour < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(format.twelve_hour ? dt.hour % 12 : dt.hour));
            break;
        case template_part::template_type::minute:
            result.append(std::to_string(dt.minute));
            break;
        case template_part::template_type::minute_leading_zero:
            if (dt.minute < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.minute));
            break;
        case template_part::template_type::second:
            result.append(std::to_string(dt.second));
            break;
        case template_part::template_type::second_leading_zero:
            if (dt.second < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.second));
            break;
        case template_part::template_type::am_pm:
            if (dt.hour < 12)
            {
                result.append("AM");
            }
            else
            {
                result.append("PM");
            }

            break;

        default:
            throw "unhandled";
        }
    }

    return result;
}

std::string number_formatter::format_text(const format_code &format, const std::string &text)
{
    std::string result;

    for (const auto &part : format.parts)
    {
        switch (part.type)
        {
        case template_part::template_type::text:
            result.append(part.string);
            break;

        case template_part::template_type::general:
            if (part.placeholders.type == placeholders::placeholders_type::general
                || part.placeholders.type == placeholders::placeholders_type::text)
            {
                result.append(text);
            }

            break;

        default:
            break;
        }
    }

    return result;
}

} // namespace detail
} // namespace xlnt
