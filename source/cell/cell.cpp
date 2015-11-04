#include <algorithm>
#include <cctype>
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/packaging/document_properties.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/string.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/worksheet/worksheet.hpp>

#include <detail/cell_impl.hpp>
#include <detail/comment_impl.hpp>

namespace {

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
    xlnt::string value;
    bool has_color = false;
	xlnt::string color;
    bool has_condition = false;
    condition_type condition = condition_type::invalid;
	xlnt::string condition_value;
    bool has_locale = false;
	xlnt::string locale;

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
std::vector<xlnt::string> split_string(const xlnt::string &string, char delim)
{
    std::vector<xlnt::string> split;
	xlnt::string::size_type previous_index = 0;
    auto separator_index = string.find(delim);

    while (separator_index != xlnt::string::npos)
    {
        auto part = string.substr(previous_index, separator_index - previous_index);
        split.push_back(part);

        previous_index = separator_index + 1;
        separator_index = string.find(delim, previous_index);
    }

    split.push_back(string.substr(previous_index));

    return split;
}

std::vector<xlnt::string> split_string_any(const xlnt::string &string, const xlnt::string &delims)
{
    std::vector<xlnt::string> split;
	xlnt::string::size_type previous_index = 0;
    auto separator_index = string.find_first_of(delims);

    while (separator_index != xlnt::string::npos)
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

bool is_date_format(const xlnt::string &format_string)
{
    auto not_in = format_string.find_first_not_of("/-:, mMyYdDhHsS");
    return not_in == xlnt::string::npos;
}

bool is_valid_color(const xlnt::string &color)
{
    static const std::vector<xlnt::string> *colors =
        new std::vector<xlnt::string>(
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

    auto compare_color = [&](const xlnt::string &other) {
        if (color.length() != other.length()) return false;

        for (std::size_t i = 0; i < color.length(); i++)
        {
            if (std::toupper(color[i].get()) != std::toupper(other[i].get()))
            {
                return false;
            }
        }

        return true;
    };

    return std::find_if(colors->begin(), colors->end(), compare_color) != colors->end();
}

bool parse_condition(const xlnt::string &string, section &s)
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

const std::unordered_map<int, xlnt::string> known_locales()
{
    const std::unordered_map<int, xlnt::string> *all =
        new std::unordered_map<int, xlnt::string>(
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

bool is_valid_locale(const xlnt::string &locale_string)
{
    auto country = locale_string.substr(locale_string.find('-') + 1);

    if (country.empty())
    {
        return false;
    }

    for (auto c : country)
    {
        if (!is_hex(static_cast<char>(std::toupper(c.get()))))
        {
            return false;
        }
    }

    auto index = country.to_hex();

    auto known_locales_ = known_locales();

    if (known_locales_.find(index) == known_locales_.end())
    {
        return false;
    }

    auto beginning = locale_string.substr(0, locale_string.find('-'));

    if (beginning.empty() || beginning[0] != '$')
    {
        return false;
    }

    if (beginning.length() == 1)
    {
        return true;
    }

    beginning = beginning.substr(1);

    return true;
}

section parse_section(const xlnt::string &section_string)
{
    section s;

    xlnt::string format_part;
    xlnt::string bracket_part;

    std::vector<xlnt::string> bracket_parts;

    bool in_quotes = false;
    bool in_brackets = false;

    const std::vector<xlnt::string> bracket_times = { "h", "hh", "m", "mm", "s", "ss" };

    for (std::size_t i = 0; i < section_string.length(); i++)
    {
        if (!in_quotes && section_string[i] == '"')
        {
            format_part.append(section_string[i]);
            in_quotes = true;
        }
        else if (in_quotes && section_string[i] == '"')
        {
            format_part.append(section_string[i]);

            if (i < section_string.length() - 1 && section_string[i + 1] != '"')
            {
                in_quotes = false;
            }
        }
        else if (!in_brackets && section_string[i] == '[')
        {
            in_brackets = true;

            for (auto bracket_time : bracket_times)
            {
                if (i < section_string.length() - bracket_time.length() &&
                    section_string.substr(i + 1, bracket_time.length()) == bracket_time)
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
                bracket_part.append(section_string[i]);
            }
        }
        else
        {
            format_part.append(section_string[i]);
        }
    }

    s.value = format_part;
    s.has_value = true;

    return s;
}

format_sections parse_format_sections(const xlnt::string &combined)
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

xlnt::string format_section(long double number, const section &format, xlnt::calendar base_date)
{
    const xlnt::string unquoted = "$+(:^'{<=-/)!&~}> ";
    xlnt::string format_temp = format.value;
    xlnt::string result;

    if (is_date_format(format.value))
    {
        const xlnt::string date_unquoted = ",-/: ";

        const std::vector<xlnt::string> dates = { "m", "mm", "mmm", "mmmmm", "mmmmmm", "d", "dd", "ddd", "dddd", "yy",
                                                 "yyyy", "h", "[h]", "hh", "m", "[m]", "mm", "s", "[s]", "ss", "AM/PM",
                                                 "am/pm", "A/P", "a/p" };
        
        const std::vector<xlnt::string> MonthNames = { "January", "February", "March",
        "April", "May", "June", "July", "August", "September", "October", "November", "December" };


        auto split = split_string_any(format.value, date_unquoted);
        xlnt::string::size_type index = 0, prev = 0;
        auto d = xlnt::datetime::from_number(number, base_date);
        bool processed_month = false;

        for (auto part : split)
        {
            while (format.value.substr(index, part.length()) != part)
            {
                index++;
            }

            auto between = format.value.substr(prev, index - prev);
            result.append(between);

			part = part.to_lower();

            if (part == "m" && !processed_month)
            {
                result.append(xlnt::string::from(d.month));
                processed_month = true;
            }
            else if (part == "mm" && !processed_month)
            {
                if (d.month < 10)
                {
                    result.append("0");
                }

                result.append(xlnt::string::from(d.month));
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
                result.append(xlnt::string::from(d.day));
            }
            else if (part == "dd")
            {
                if (d.day < 10)
                {
                    result.append("0");
                }

                result.append(xlnt::string::from(d.day));
            }
            else if (part == "yyyy")
            {
                result.append(xlnt::string::from(d.year));
            }

            else if (part == "h")
            {
                result.append(xlnt::string::from(d.hour));
                processed_month = true;
            }
            else if (part == "hh")
            {
                if (d.hour < 10)
                {
                    result.append("0");
                }

                result.append(xlnt::string::from(d.hour));
                processed_month = true;
            }
            else if (part == "m")
            {
                result.append(xlnt::string::from(d.minute));
            }
            else if (part == "mm")
            {
                if (d.minute < 10)
                {
                    result.append("0");
                }

                result.append(xlnt::string::from(d.minute));
            }
            else if (part == "s")
            {
                result.append(xlnt::string::from(d.second));
            }
            else if (part == "ss")
            {
                if (d.second < 10)
                {
                    result.append("0");
                }

                result.append(xlnt::string::from(d.second));
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

            index += part.length();
            prev = index;
        }

        if (index < format.value.length())
        {
            result.append(format.value.substr(index));
        }
    }
    else if (format.value == "General" || format.value == "0")
    {
        if (number == static_cast<long long int>(number))
        {
            result = xlnt::string::from(static_cast<long long int>(number));
        }
        else
        {
            result = xlnt::string::from(number);
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

        result += xlnt::string::from(number < 0 ? -number : number);

        auto decimal_pos = result.find('.');

        if (decimal_pos != xlnt::string::npos)
        {
            result[decimal_pos] = ',';
            decimal_pos += 3;

            while (decimal_pos < result.length())
            {
                result.remove(result.back());
            }
        }
    }

    return result;
}

xlnt::string format_section(const xlnt::string &text, const section &format)
{
    auto arobase_index = format.value.find('@');

    xlnt::string first_part, middle_part, last_part;

    if (arobase_index != xlnt::string::npos)
    {
        first_part = format.value.substr(0, arobase_index);
        middle_part = text;
        last_part = format.value.substr(arobase_index + 1);
    }
    else
    {
        first_part = format.value;
    }

    auto unquote = [](xlnt::string &s) {
        if (!s.empty())
        {
            if (s.front() != '"' || s.back() != '"') return false;
            s = s.substr(0, s.length() - 2);
        }

        return true;
    };

    if (!unquote(first_part) || !unquote(last_part))
    {
        throw std::runtime_error("additional text must be enclosed in quotes");
    }

    return first_part + middle_part + last_part;
}

xlnt::string format_number(long double number, const xlnt::string &format, xlnt::calendar base_date)
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

xlnt::string format_text(const xlnt::string &text, const xlnt::string &format)
{
    if (format == "General") return text;
    auto sections = parse_format_sections(format);
    return format_section(text, sections.fourth);
}
}

namespace xlnt {

const std::unordered_map<string, int> &cell::error_codes()
{
    static const std::unordered_map<string, int> *codes =
    new std::unordered_map<string, int>({ { "#NULL!", 0 }, { "#DIV/0!", 1 }, { "#VALUE!", 2 },
                                                                { "#REF!", 3 },  { "#NAME?", 4 },  { "#NUM!", 5 },
                                                                { "#N/A!", 6 } });

    return *codes;
};

cell::cell() : d_(nullptr)
{
}

cell::cell(detail::cell_impl *d) : d_(d)
{
}

cell::cell(worksheet worksheet, const cell_reference &reference) : d_(nullptr)
{
    cell self = worksheet.get_cell(reference);
    d_ = self.d_;
}

template <typename T>
cell::cell(worksheet worksheet, const cell_reference &reference, const T &initial_value) : cell(worksheet, reference)
{
    set_value(initial_value);
}

bool cell::garbage_collectible() const
{
    return !(get_data_type() != type::null || is_merged() || has_comment() || has_formula() || has_style());
}

template <>
void cell::set_value(bool b)
{
    d_->value_numeric_ = b ? 1 : 0;
    d_->type_ = type::boolean;
}

template <>
void cell::set_value(std::int8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::int16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::int32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::int64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::uint8_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::uint16_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::uint32_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(std::uint64_t i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

#ifdef _WIN32
template <>
void cell::set_value(unsigned long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

#ifdef __linux
template <>
void cell::set_value(long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(unsigned long long i)
{
    d_->value_numeric_ = static_cast<long double>(i);
    d_->type_ = type::numeric;
}
#endif

template <>
void cell::set_value(float f)
{
    d_->value_numeric_ = static_cast<long double>(f);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
void cell::set_value(long double d)
{
    d_->value_numeric_ = static_cast<long double>(d);
    d_->type_ = type::numeric;
}

template <>
void XLNT_FUNCTION cell::set_value(string s)
{
    d_->set_string(s, get_parent().get_parent().get_guess_types());

	if (get_data_type() == type::string && !s.empty())
	{
		get_parent().get_parent().add_shared_string(s);
	}
}

template <>
void XLNT_FUNCTION cell::set_value(char const *c)
{
    set_value(string(c));
}

template <>
void cell::set_value(cell c)
{
    d_->type_ = c.d_->type_;
    d_->value_numeric_ = c.d_->value_numeric_;
    d_->value_string_ = c.d_->value_string_;
    d_->hyperlink_ = c.d_->hyperlink_;
    d_->has_hyperlink_ = c.d_->has_hyperlink_;
    d_->formula_ = c.d_->formula_;
    d_->style_id_ = c.d_->style_id_;
    set_comment(c.get_comment());
}

template <>
void cell::set_value(date d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_yyyymmdd2());
}

template <>
void cell::set_value(datetime d)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = d.to_number(get_base_date());
    set_number_format(number_format::date_datetime());
}

template <>
void cell::set_value(time t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format::date_time6());
}

template <>
void cell::set_value(timedelta t)
{
    d_->type_ = type::numeric;
    d_->value_numeric_ = t.to_number();
    set_number_format(number_format::date_timedelta());
}

row_t cell::get_row() const
{
    return d_->row_;
}

column_t cell::get_column() const
{
    return d_->column_;
}

void cell::set_merged(bool merged)
{
    d_->is_merged_ = merged;
}

bool cell::is_merged() const
{
    return d_->is_merged_;
}

bool cell::is_date() const
{
    if (get_data_type() == type::numeric)
    {
        auto number_format = get_number_format().get_format_string();

        if (number_format != "General")
        {
            try
            {
                auto sections = parse_format_sections(number_format);
                return is_date_format(sections.first.value);
            }
            catch (std::exception)
            {
                return false;
            }
        }
    }

    return false;
}

cell_reference cell::get_reference() const
{
    return { d_->column_, d_->row_ };
}

bool cell::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool cell::operator==(const cell &comparand) const
{
    return d_ == comparand.d_;
}

cell &cell::operator=(const cell &rhs)
{
    *d_ = *rhs.d_;
    return *this;
}

bool operator<(cell left, cell right)
{
    return left.get_reference() < right.get_reference();
}

string cell::to_repr() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

relationship cell::get_hyperlink() const
{
    if (!d_->has_hyperlink_)
    {
        throw std::runtime_error("no hyperlink set");
    }

    return d_->hyperlink_;
}

bool cell::has_hyperlink() const
{
    return d_->has_hyperlink_;
}

void cell::set_hyperlink(const string &hyperlink)
{
    if (hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw data_type_exception();
    }

    d_->has_hyperlink_ = true;
    d_->hyperlink_ = worksheet(d_->parent_).create_relationship(relationship::type::hyperlink, hyperlink);

    if (get_data_type() == type::null)
    {
        set_value(hyperlink);
    }
}

void cell::set_formula(const string &formula)
{
    if (formula.length() == 0)
    {
        throw data_type_exception();
    }

    if (formula[0] == '=')
    {
        d_->formula_ = formula.substr(1);
    }
    else
    {
        d_->formula_ = formula;
    }
}

bool cell::has_formula() const
{
    return !d_->formula_.empty();
}

string cell::get_formula() const
{
    if (d_->formula_.empty())
    {
        throw data_type_exception();
    }

    return d_->formula_;
}

void cell::clear_formula()
{
    d_->formula_.clear();
}

void cell::set_comment(const xlnt::comment &c)
{
    if (c.d_ != d_->comment_.get())
    {
        throw xlnt::attribute_error();
    }

    if (!has_comment())
    {
        get_parent().increment_comments();
    }

    *get_comment().d_ = *c.d_;
}

void cell::clear_comment()
{
    if (has_comment())
    {
        get_parent().decrement_comments();
    }

    d_->comment_ = nullptr;
}

bool cell::has_comment() const
{
    return d_->comment_ != nullptr;
}

void cell::set_error(const string &error)
{
    if (error.length() == 0 || error[0] != '#')
    {
        throw data_type_exception();
    }

    d_->value_string_ = error;
    d_->type_ = type::error;
}

cell cell::offset(int column, int row)
{
    return get_parent().get_cell(cell_reference(d_->column_ + column, d_->row_ + row));
}

worksheet cell::get_parent()
{
    return worksheet(d_->parent_);
}

const worksheet cell::get_parent() const
{
    return worksheet(d_->parent_);
}

comment cell::get_comment()
{
    if (d_->comment_ == nullptr)
    {
        d_->comment_.reset(new detail::comment_impl());
        get_parent().increment_comments();
    }

    return comment(d_->comment_.get());
}

//TODO: this shares a lot of code with worksheet::get_point_pos, try to reduce repition
std::pair<int, int> cell::get_anchor() const
{
    static const double DefaultColumnWidth = 51.85;
    static const double DefaultRowHeight = 15.0;

    auto points_to_pixels = [](long double value, long double dpi)
    {
        return static_cast<int>(std::ceil(value * dpi / 72));
    };
    
    auto left_columns = d_->column_ - 1;
    int left_anchor = 0;
    auto default_width = points_to_pixels(DefaultColumnWidth, 96.0);

    for (column_t column_index = 1; column_index <= left_columns; column_index++)
    {
        if (get_parent().has_column_properties(column_index))
        {
            auto cdw = get_parent().get_column_properties(column_index).width;

            if (cdw > 0)
            {
                left_anchor += points_to_pixels(cdw, 96.0);
                continue;
            }
        }

        left_anchor += default_width;
    }

    auto top_rows = d_->row_ - 1;
    int top_anchor = 0;
    auto default_height = points_to_pixels(DefaultRowHeight, 96.0);

    for (row_t row_index = 1; row_index <= top_rows; row_index++)
    {
        if (get_parent().has_row_properties(row_index))
        {
            auto rdh = get_parent().get_row_properties(row_index).height;

            if (rdh > 0)
            {
                top_anchor += points_to_pixels(rdh, 96.0);
                continue;
            }
        }

        top_anchor += default_height;
    }

    return { left_anchor, top_anchor };
}

cell::type cell::get_data_type() const
{
    return d_->type_;
}

void cell::set_data_type(type t)
{
    d_->type_ = t;
}

const number_format &cell::get_number_format() const
{
    if (d_->has_style_)
    {
        return get_parent().get_parent().get_number_format(d_->style_id_);
    }
    else
    {
        return get_parent().get_parent().get_number_formats().front();
    }
}

const font &cell::get_font() const
{
    return get_parent().get_parent().get_font(d_->style_id_);
}

const fill &cell::get_fill() const
{
    return get_parent().get_parent().get_fill(d_->style_id_);
}

const border &cell::get_border() const
{
    return get_parent().get_parent().get_border(d_->style_id_);
}

const alignment &cell::get_alignment() const
{
    return get_parent().get_parent().get_alignment(d_->style_id_);
}

const protection &cell::get_protection() const
{
    return get_parent().get_parent().get_protection(d_->style_id_);
}

bool cell::pivot_button() const
{
    return get_parent().get_parent().get_pivot_button(d_->style_id_);
}

bool cell::quote_prefix() const
{
    return get_parent().get_parent().get_quote_prefix(d_->style_id_);
}

void cell::clear_value()
{
    d_->value_numeric_ = 0;
    d_->value_string_.clear();
    d_->formula_.clear();
    d_->type_ = cell::type::null;
}

template <>
bool cell::get_value() const
{
    return d_->value_numeric_ != 0;
}

template <>
std::int8_t cell::get_value() const
{
    return static_cast<std::int8_t>(d_->value_numeric_);
}

template <>
std::int16_t cell::get_value() const
{
    return static_cast<std::int16_t>(d_->value_numeric_);
}

template <>
std::int32_t cell::get_value() const
{
    return static_cast<std::int32_t>(d_->value_numeric_);
}

template <>
std::int64_t cell::get_value() const
{
    return static_cast<std::int64_t>(d_->value_numeric_);
}

template <>
std::uint8_t cell::get_value() const
{
    return static_cast<std::uint8_t>(d_->value_numeric_);
}

template <>
std::uint16_t cell::get_value() const
{
    return static_cast<std::uint16_t>(d_->value_numeric_);
}

template <>
std::uint32_t cell::get_value() const
{
    return static_cast<std::uint32_t>(d_->value_numeric_);
}

template <>
std::uint64_t cell::get_value() const
{
    return static_cast<std::uint64_t>(d_->value_numeric_);
}

#ifdef __linux
template <>
long long int cell::get_value() const
{
    return static_cast<long long int>(d_->value_numeric_);
}
#endif

template <>
float cell::get_value() const
{
    return static_cast<float>(d_->value_numeric_);
}

template <>
double cell::get_value() const
{
    return static_cast<double>(d_->value_numeric_);
}

template <>
long double cell::get_value() const
{
    return d_->value_numeric_;
}

template <>
time cell::get_value() const
{
    return time::from_number(d_->value_numeric_);
}

template <>
datetime cell::get_value() const
{
    return datetime::from_number(d_->value_numeric_, get_base_date());
}

template <>
date cell::get_value() const
{
    return date::from_number(static_cast<int>(d_->value_numeric_), get_base_date());
}

template <>
timedelta cell::get_value() const
{
    return timedelta::from_number(d_->value_numeric_);
}

void cell::set_number_format(const number_format &number_format_)
{
    d_->has_style_ = true;
    d_->style_id_ = get_parent().get_parent().set_number_format(number_format_, d_->style_id_);
}

template <>
string cell::get_value() const
{
    return d_->value_string_;
}

bool cell::has_value() const
{
    return d_->type_ != cell::type::null;
}

string cell::to_string() const
{
    auto nf = get_number_format();

    switch (get_data_type())
    {
    case cell::type::null:
        return "";
    case cell::type::numeric:
        return format_number(get_value<long double>(), nf.get_format_string(), get_base_date());
    case cell::type::string:
    case cell::type::formula:
    case cell::type::error:
        return format_text(get_value<string>(), nf.get_format_string());
    case cell::type::boolean:
        return get_value<long double>() == 0 ? "FALSE" : "TRUE";
    default:
        return "";
    }
}

std::size_t cell::get_style_id() const
{
    return d_->style_id_;
}

bool cell::has_style() const
{
    return d_->has_style_;
}

void cell::set_style_id(std::size_t style_id)
{
    d_->style_id_ = style_id;
    d_->has_style_ = true;
}

calendar cell::get_base_date() const
{
    return get_parent().get_parent().get_properties().excel_base_date;
}

} // namespace xlnt
