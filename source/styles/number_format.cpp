// Copyright (c) 2014-2016 Thomas Fussell
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
#include <algorithm>
#include <cctype>
#include <unordered_map>
#include <vector>

#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/hash_combine.hpp>
#include <xlnt/styles/number_format.hpp>

namespace {

const std::unordered_map<std::size_t, std::string> &builtin_formats()
{
    static const std::unordered_map<std::size_t, std::string> *formats =
        new std::unordered_map<std::size_t, std::string>
        ({
            { 0, "General" },
            { 1, "0" },
            { 2, "0.00" },
            { 3, "#,##0" },
            { 4, "#,##0.00" },
            { 5, "\"$\"#,##0_);(\"$\"#,##0)" },
            { 6, "\"$\"#,##0_);[Red](\"$\"#,##0)" },
            { 7, "\"$\"#,##0.00_);(\"$\"#,##0.00)" },
            { 8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)" },
            { 9, "0%" },
            { 10, "0.00%" },
            { 11, "0.00E+00" },
            { 12, "# ?/?" },
            { 13, "# \?\?/??" }, // escape trigraph
            { 14, "mm-dd-yy" },
            { 15, "d-mmm-yy" },
            { 16, "d-mmm" },
            { 17, "mmm-yy" },
            { 18, "h:mm AM/PM" },
            { 19, "h:mm:ss AM/PM" },
            { 20, "h:mm" },
            { 21, "h:mm:ss" },
            { 22, "m/d/yy h:mm" },

            { 37, "#,##0_);(#,##0)" },
            { 38, "#,##0_);[Red](#,##0)" },
            { 39, "#,##0.00_);(#,##0.00)" },
            { 40, "#,##0.00_);[Red](#,##0.00)" },

            { 41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)" },
            { 42, "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)" },
            { 43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)" },

            { 44, "_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)" },
            { 45, "mm:ss" },
            { 46, "[h]:mm:ss" },
            { 47, "mmss.0" },
            { 48, "##0.0E+0" },
            { 49, "@" }

            // EXCEL differs from the standard in the following:
            //{14, "m/d/yyyy"},
            //{22, "m/d/yyyy h:mm"},
            //{37, "#,##0_);(#,##0)"},
            //{38, "#,##0_);[Red]"},
            //{39, "#,##0.00_);(#,##0.00)"},
            //{40, "#,##0.00_);[Red]"},
            //{47, "mm:ss.0"},
            //{55, "yyyy/mm/dd"}
        });

    return *formats;
}

enum class condition_type
{
    less_than,
    less_or_equal,
    equal,
    greater_than,
    greater_or_equal,
    invalid
};

struct section
{
    bool has_value = false;
    std::string value;
    bool has_color = false;
    std::string color;
    bool has_condition = false;
    condition_type condition = condition_type::invalid;
    std::string condition_value;
    bool has_locale = false;
    std::string locale;

    section &operator=(const section &other)
    {
        has_value = other.has_value;
        value = other.value;
        has_color = other.has_color;
        color = other.color;
        has_condition = other.has_condition;
        condition = other.condition;
        condition_value = other.condition_value;
        has_locale = other.has_locale;
        locale = other.locale;

        return *this;
    }
};

struct format_sections
{
    section first;
    section second;
    section third;
    section fourth;
};

// copied from named_range.cpp, keep in sync
/// <summary>
/// Return a vector containing string split at each delim.
/// </summary>
/// <remark>
/// This should maybe be in a utility header so it can be used elsewhere.
/// </remarks>
std::vector<std::string> split_string(const std::string &string, char delim)
{
    std::vector<std::string> split;
    std::string::size_type previous_index = 0;
    auto separator_index = string.find(delim);

    while (separator_index != std::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);
        split.push_back(part);

        previous_index = separator_index + 1;
        separator_index = string.find(delim, previous_index);
    }

    split.push_back(string.substr(previous_index));

    return split;
}

std::vector<std::string> split_string_any(const std::string &string, const std::string &delims)
{
    std::vector<std::string> split;
    std::string::size_type previous_index = 0;
    auto separator_index = string.find_first_of(delims);

    while (separator_index != std::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);

        if (!part.empty())
        {
            split.push_back(part);
        }

        previous_index = separator_index + 1;
        separator_index = string.find_first_of(delims, previous_index);
    }

    split.push_back(string.substr(previous_index));

    return split;
}

bool is_date_format(const std::string &format_string)
{
    auto not_in = format_string.find_first_not_of("/-:, mMyYdDhHsS");
    return not_in == std::string::npos;
}

bool is_valid_color(const std::string &color)
{
    static const std::vector<std::string> *colors =
        new std::vector<std::string>(
        {
            "Black",
            "Green"
            "White",
            "Blue",
            "Magenta",
            "Yellow",
            "Cyan",
            "Red"
        });

    auto compare_color = [&](const std::string &other) {
        if (color.size() != other.size()) return false;

        for (std::size_t i = 0; i < color.size(); i++)
        {
            if (std::toupper(color[i]) != std::toupper(other[i]))
            {
                return false;
            }
        }

        return true;
    };

    return std::find_if(colors->begin(), colors->end(), compare_color) != colors->end();
}

bool parse_condition(const std::string &string, section &s)
{
    s.has_condition = false;
    s.condition = condition_type::invalid;
    s.condition_value.clear();

    if (string[0] == '<')
    {
        s.has_condition = true;

        if (string[1] == '=')
        {
            s.condition = condition_type::less_or_equal;
            s.condition_value = string.substr(2);
        }
        else
        {
            s.condition = condition_type::less_than;
            s.condition_value = string.substr(1);
        }
    }
    if (string[0] == '>')
    {
        s.has_condition = true;

        if (string[1] == '=')
        {
            s.condition = condition_type::greater_or_equal;
            s.condition_value = string.substr(2);
        }
        else
        {
            s.condition = condition_type::greater_than;
            s.condition_value = string.substr(1);
        }
    }
    else if (string[0] == '=')
    {
        s.has_condition = true;
        s.condition = condition_type::equal;
        s.condition_value = string.substr(1);
    }

    return s.has_condition;
}

bool is_hex(char c)
{
    if (c >= 'A' || c <= 'F') return true;
    if (c >= '0' || c <= '9') return true;
    return false;
}

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

bool is_valid_locale(const std::string &locale_string)
{
    std::string country = locale_string.substr(locale_string.find('-') + 1);

    if (country.empty())
    {
        return false;
    }

    for (auto c : country)
    {
        if (!is_hex(static_cast<char>(std::toupper(c))))
        {
            return false;
        }
    }

    auto index = std::stoi(country, 0, 16);

    auto known_locales_ = known_locales();

    if (known_locales_.find(index) == known_locales_.end())
    {
        return false;
    }

    std::string beginning = locale_string.substr(0, locale_string.find('-'));

    if (beginning.empty() || beginning[0] != '$')
    {
        return false;
    }

    if (beginning.size() == 1)
    {
        return true;
    }

    beginning = beginning.substr(1);

    return true;
}

section parse_section(const std::string &section_string)
{
    section s;

    std::string format_part;
    std::string bracket_part;

    std::vector<std::string> bracket_parts;

    bool in_quotes = false;
    bool in_brackets = false;

    const std::vector<std::string> bracket_times = { "h", "hh", "m", "mm", "s", "ss" };

    for (std::size_t i = 0; i < section_string.size(); i++)
    {
        if (!in_quotes && section_string[i] == '"')
        {
            format_part.push_back(section_string[i]);
            in_quotes = true;
        }
        else if (in_quotes && section_string[i] == '"')
        {
            format_part.push_back(section_string[i]);

            if (i < section_string.size() - 1 && section_string[i + 1] != '"')
            {
                in_quotes = false;
            }
        }
        else if (!in_brackets && section_string[i] == '[')
        {
            in_brackets = true;

            for (auto bracket_time : bracket_times)
            {
                if (i < section_string.size() - bracket_time.size() &&
                    section_string.substr(i + 1, bracket_time.size()) == bracket_time)
                {
                    in_brackets = false;
                    break;
                }
            }
        }
        else if (in_brackets)
        {
            if (section_string[i] == ']')
            {
                in_brackets = false;

                if (is_valid_color(bracket_part))
                {
                    if (s.color.empty())
                    {
                        s.color = bracket_part;
                        s.has_color = true;
                    }
                    else
                    {
                        throw std::runtime_error("two colors");
                    }
                }
                else if (is_valid_locale(bracket_part))
                {
                    if (s.locale.empty())
                    {
                        s.locale = bracket_part;
                        s.has_locale = true;
                    }
                    else
                    {
                        throw std::runtime_error("two locales");
                    }
                }
                else if (s.has_condition || !parse_condition(bracket_part, s))
                {
                    throw std::runtime_error("invalid bracket format");
                }

                bracket_part.clear();
            }
            else
            {
                bracket_part.push_back(section_string[i]);
            }
        }
        else
        {
            format_part.push_back(section_string[i]);
        }
    }

    s.value = format_part;
    s.has_value = true;

    return s;
}

format_sections parse_format_sections(const std::string &combined)
{
    format_sections result = {};

    auto split = split_string(combined, ';');

    if (split.empty())
    {
        throw std::runtime_error("empty string");
    }

    result.first = parse_section(split[0]);

    if (!result.first.has_condition)
    {
        result.second = result.first;
        result.third = result.first;
    }

    if (split.size() > 1)
    {
        result.second = parse_section(split[1]);
    }

    if (split.size() > 2)
    {
        if (result.first.has_condition && !result.second.has_condition)
        {
            throw std::runtime_error("first two sections should have conditions");
        }

        result.third = parse_section(split[2]);

        if (result.third.has_condition)
        {
            throw std::runtime_error("third section shouldn't have a condition");
        }
    }

    if (split.size() > 3)
    {
        if (result.first.has_condition)
        {
            throw std::runtime_error("too many parts");
        }

        result.fourth = parse_section(split[3]);
    }

    if (split.size() > 4)
    {
        throw std::runtime_error("too many parts");
    }

    return result;
}

std::string format_section(long double number, const section &format, xlnt::calendar base_date)
{
    const std::string unquoted = "$+(:^'{<=-/)!&~}> ";
    std::string format_temp = format.value;
    std::string result;

    if (is_date_format(format.value))
    {
        const std::string date_unquoted = ",-/: ";

        const std::vector<std::string> dates = { "m", "mm", "mmm", "mmmmm", "mmmmmm", "d", "dd", "ddd", "dddd", "yy",
                                                 "yyyy", "h", "[h]", "hh", "m", "[m]", "mm", "s", "[s]", "ss", "AM/PM",
                                                 "am/pm", "A/P", "a/p" };
        
        const std::vector<std::string> MonthNames = { "January", "February", "March",
        "April", "May", "June", "July", "August", "September", "October", "November", "December" };


        auto split = split_string_any(format.value, date_unquoted);
        std::string::size_type index = 0, prev = 0;
        auto d = xlnt::datetime::from_number(number, base_date);
        bool processed_month = false;

        auto lower_string = [](const std::string &s) {
            std::string lower;
            lower.resize(s.size());
            for (std::size_t i = 0; i < s.size(); i++)
                lower[i] = static_cast<char>(std::tolower(s[i]));
            return lower;
        };

        for (auto part : split)
        {
            while (format.value.substr(index, part.size()) != part)
            {
                index++;
            }

            auto between = format.value.substr(prev, index - prev);
            result.append(between);

            part = lower_string(part);

            if (part == "m" && !processed_month)
            {
                result.append(std::to_string(d.month));
                processed_month = true;
            }
            else if (part == "mm" && !processed_month)
            {
                if (d.month < 10)
                {
                    result.append("0");
                }

                result.append(std::to_string(d.month));
                processed_month = true;
            }
            else if (part == "mmm" && !processed_month)
            {
                result.append(MonthNames.at(static_cast<std::size_t>(d.month - 1)).substr(0, 3));
                processed_month = true;
            }
            else if (part == "mmmm" && !processed_month)
            {
                result.append(MonthNames.at(static_cast<std::size_t>(d.month - 1)));
                processed_month = true;
            }
            else if (part == "d")
            {
                result.append(std::to_string(d.day));
            }
            else if (part == "dd")
            {
                if (d.day < 10)
                {
                    result.append("0");
                }

                result.append(std::to_string(d.day));
            }
            else if (part == "yyyy")
            {
                result.append(std::to_string(d.year));
            }

            else if (part == "h")
            {
                result.append(std::to_string(d.hour));
                processed_month = true;
            }
            else if (part == "hh")
            {
                if (d.hour < 10)
                {
                    result.append("0");
                }

                result.append(std::to_string(d.hour));
                processed_month = true;
            }
            else if (part == "m")
            {
                result.append(std::to_string(d.minute));
            }
            else if (part == "mm")
            {
                if (d.minute < 10)
                {
                    result.append("0");
                }

                result.append(std::to_string(d.minute));
            }
            else if (part == "s")
            {
                result.append(std::to_string(d.second));
            }
            else if (part == "ss")
            {
                if (d.second < 10)
                {
                    result.append("0");
                }

                result.append(std::to_string(d.second));
            }
            else if (part == "am/pm" || part == "a/p")
            {
                if (d.hour < 12)
                {
                    result.append("AM");
                }
                else
                {
                    result.append("PM");
                }
            }

            index += part.size();
            prev = index;
        }

        if (index < format.value.size())
        {
            result.append(format.value.substr(index));
        }
    }
    else if (format.value == "General" || format.value == "0")
    {
        if (number == static_cast<long long int>(number))
        {
            result = std::to_string(static_cast<long long int>(number));
        }
        else
        {
            result = std::to_string(number);
        }
    }
    else if (format.value.substr(0, 8) == "#,##0.00" || format.value.substr(0, 9) == "-#,##0.00")
    {
        if (format.value[0] == '-')
        {
            result = "-";
        }

        if (format.has_locale && format.locale == "$$-1009")
        {
            result += "CA$";
        }
        else if (format.has_locale && format.locale == "$€-407")
        {
            result += "€";
        }
        else
        {
            result += "$";
        }

        result += std::to_string(number < 0 ? -number : number);

        auto decimal_pos = result.find('.');

        if (decimal_pos != std::string::npos)
        {
            result[decimal_pos] = ',';
            decimal_pos += 3;

            while (decimal_pos < result.size())
            {
                result.pop_back();
            }
        }
    }

    return result;
}

std::string format_section(const std::string &text, const section &format)
{
    auto arobase_index = format.value.find('@');

    std::string first_part, middle_part, last_part;

    if (arobase_index != std::string::npos)
    {
        first_part = format.value.substr(0, arobase_index);
        middle_part = text;
        last_part = format.value.substr(arobase_index + 1);
    }
    else
    {
        first_part = format.value;
    }

    auto unquote = [](std::string &s) {
        if (!s.empty())
        {
            if (s.front() != '"' || s.back() != '"') return false;
            s = s.substr(0, s.size() - 2);
        }

        return true;
    };

    if (!unquote(first_part) || !unquote(last_part))
    {
        throw std::runtime_error(std::string("additional text must be enclosed in quotes: ") + format.value);
    }

    return first_part + middle_part + last_part;
}

std::string format_number(long double number, const std::string &format, xlnt::calendar base_date)
{
    auto sections = parse_format_sections(format);

    if (number > 0)
    {
        return format_section(number, sections.first, base_date);
    }
    else if (number < 0)
    {
        return format_section(number, sections.second, base_date);
    }

    // number == 0
    return format_section(number, sections.third, base_date);
}

std::string format_text(const std::string &text, const std::string &format)
{
    if (format == "General") return text;
    auto sections = parse_format_sections(format);
    return format_section(text, sections.fourth);
}

} // namespace

namespace xlnt {

const number_format number_format::general()
{
    static const number_format *format = new number_format(builtin_formats().at(0), 0);
    return *format;
}

const number_format number_format::text()
{
    static const number_format *format = new number_format(builtin_formats().at(49), 49);
    return *format;
}

const number_format number_format::number()
{
    static const number_format *format = new number_format(builtin_formats().at(1), 1);
    return *format;
}

const number_format number_format::number_00()
{
    static const number_format *format = new number_format(builtin_formats().at(2), 2);
    return *format;
}

const number_format number_format::number_comma_separated1()
{
    static const number_format *format = new number_format(builtin_formats().at(4), 4);
    return *format;
}

const number_format number_format::number_comma_separated2()
{
    static const number_format *format = new number_format("#,##0.00_-");
    return *format;
}

const number_format number_format::percentage()
{
    static const number_format *format = new number_format(builtin_formats().at(9), 9);
    return *format;
}

const number_format number_format::percentage_00()
{
    static const number_format *format = new number_format(builtin_formats().at(10), 10);
    return *format;
}

const number_format number_format::date_yyyymmdd2()
{
    static const number_format *format = new number_format("yyyy-mm-dd");
    return *format;
}

const number_format number_format::date_yyyymmdd()
{
    static const number_format *format = new number_format("yy-mm-dd");
    return *format;
}

const number_format number_format::date_ddmmyyyy()
{
    static const number_format *format = new number_format("dd/mm/yy");
    return *format;
}

const number_format number_format::date_dmyslash()
{
    static const number_format *format = new number_format("d/m/y");
    return *format;
}

const number_format number_format::date_dmyminus()
{
    static const number_format *format = new number_format("d-m-y");
    return *format;
}

const number_format number_format::date_dmminus()
{
    static const number_format *format = new number_format("d-m");
    return *format;
}

const number_format number_format::date_myminus()
{
    static const number_format *format = new number_format("m-y");
    return *format;
}

const number_format number_format::date_xlsx14()
{
    static const number_format *format = new number_format(builtin_formats().at(14), 14);
    return *format;
}

const number_format number_format::date_xlsx15()
{
    static const number_format *format = new number_format(builtin_formats().at(15), 15);
    return *format;
}

const number_format number_format::date_xlsx16()
{
    static const number_format *format = new number_format(builtin_formats().at(16), 16);
    return *format;
}

const number_format number_format::date_xlsx17()
{
    static const number_format *format = new number_format(builtin_formats().at(17), 17);
    return *format;
}

const number_format number_format::date_xlsx22()
{
    static const number_format *format = new number_format(builtin_formats().at(22), 22);
    return *format;
}

const number_format number_format::date_datetime()
{
    static const number_format *format = new number_format("yyyy-mm-dd h:mm:ss");
    return *format;
}

const number_format number_format::date_time1()
{
    static const number_format *format = new number_format(builtin_formats().at(18), 18);
    return *format;
}

const number_format number_format::date_time2()
{
    static const number_format *format = new number_format(builtin_formats().at(19), 19);
    return *format;
}

const number_format number_format::date_time3()
{
    static const number_format *format = new number_format(builtin_formats().at(20), 20);
    return *format;
}

const number_format number_format::date_time4()
{
    static const number_format *format = new number_format(builtin_formats().at(21), 21);
    return *format;
}

const number_format number_format::date_time5()
{
    static const number_format *format = new number_format(builtin_formats().at(45), 45);
    return *format;
}

const number_format number_format::date_time6()
{
    static const number_format *format = new number_format(builtin_formats().at(21), 21);
    return *format;
}

const number_format number_format::date_time7()
{
    static const number_format *format = new number_format("i:s.S");
    return *format;
}

const number_format number_format::date_time8()
{
    static const number_format *format = new number_format("h:mm:ss@");
    return *format;
}

const number_format number_format::date_timedelta()
{
    static const number_format *format = new number_format("[hh]:mm:ss");
    return *format;
}

const number_format number_format::date_yyyymmddslash()
{
    static const number_format *format = new number_format("yy/mm/dd@");
    return *format;
}

const number_format number_format::currency_usd_simple()
{
    static const number_format *format = new number_format("\"$\"#,##0.00_-");
    return *format;
}

const number_format number_format::currency_usd()
{
    static const number_format *format = new number_format("$#,##0_-");
    return *format;
}

const number_format number_format::currency_eur_simple()
{
    static const number_format *format = new number_format("[$EUR ]#,##0.00_-");
    return *format;
}

number_format::number_format() : number_format(general())
{
}

number_format::number_format(std::size_t id) : number_format(from_builtin_id(id))
{
}

number_format::number_format(const std::string &format_string) : id_set_(false), id_(0)
{
    set_format_string(format_string);
}

number_format::number_format(const std::string &format_string, std::size_t id) : id_set_(false), id_(0)
{
    set_format_string(format_string, id);
}

number_format number_format::from_builtin_id(std::size_t builtin_id)
{
    if (builtin_formats().find(builtin_id) == builtin_formats().end())
    {
        throw std::runtime_error("unknown id: " + std::to_string(builtin_id));
    }

    auto format_string = builtin_formats().at(builtin_id);
    return number_format(format_string, builtin_id);
}

std::string number_format::get_format_string() const
{
    return format_string_;
}

std::string number_format::to_hash_string() const
{
    std::string hash_string("number_format");
    
    hash_string.append(std::to_string(id_));
    hash_string.append(format_string_);

    return hash_string;
}


void number_format::set_format_string(const std::string &format_string)
{
    format_string_ = format_string;
    id_ = 0;
    id_set_ = false;

    for (const auto &pair : builtin_formats())
    {
        if (pair.second == format_string)
        {
            id_ = pair.first;
            id_set_ = true;
            break;
        }
    }
}

void number_format::set_format_string(const std::string &format_string, std::size_t id)
{
    format_string_ = format_string;
    id_ = id;
    id_set_ = true;
}

bool number_format::has_id() const
{
    return id_set_;
}

void number_format::set_id(std::size_t id)
{
    id_ = id;
    id_set_ = true;
}

std::size_t number_format::get_id() const
{
    if(!id_set_)
    {
        throw std::runtime_error("number format doesn't have an id");
    }
    
    return id_;
}

bool number_format::is_date_format() const
{
    if (format_string_ != "General")
    {
        try
        {
            auto sections = parse_format_sections(format_string_);
            return ::is_date_format(sections.first.value);
        }
        catch (std::exception)
        {
            return false;
        }
    }

    return false;
}

std::string number_format::format(const std::string &text) const
{
    return format_text(text, format_string_);
}

std::string number_format::format(long double number, calendar base_date) const
{
    return format_number(number, format_string_, base_date);
}

} // namespace xlnt
