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
#include <unordered_map>
#include <vector>

#include <xlnt/utils/datetime.hpp>

namespace xlnt {
namespace detail {

enum class format_color
{
    black,
    blue,
    cyan,
    green,
    magenta,
    red,
    white,
    yellow,
    color1,
    color2,
    color3,
    color4,
    color5,
    color6,
    color7,
    color8,
    color9,
    color10,
    color11,
    color12,
    color13,
    color14,
    color15,
    color16,
    color17,
    color18,
    color19,
    color20,
    color21,
    color22,
    color23,
    color24,
    color25,
    color26,
    color27,
    color28,
    color29,
    color30,
    color31,
    color32,
    color33,
    color34,
    color35,
    color36,
    color37,
    color38,
    color39,
    color40,
    color41,
    color42,
    color43,
    color44,
    color45,
    color46,
    color47,
    color48,
    color49,
    color50,
    color51,
    color52,
    color53,
    color54,
    color55,
    color56
};

enum class format_locale
{
    arabic_saudi_arabia = 0x401,
    bulgarian = 0x402,
    catalan = 0x403,
    chinese_taiwan = 0x404,
    czech = 0x405,
    danish = 0x406,
    german_germany = 0x407,
    greek = 0x408,
    english_united_states = 0x409,
    italian_italy = 0x410,
    japanese = 0x411,
    korean = 0x412,
    dutch_netherlands = 0x413,
    norwegian_bokml = 0x414,
    polish = 0x415,
    portuguese_brazil = 0x416,
    raeto_romance = 0x417,
    romanian_romania = 0x418,
    russian = 0x419,
    urdu = 0x420,
    indonesian = 0x421,
    ukrainian = 0x422,
    belarusian = 0x423,
    slovenian = 0x424,
    estonian = 0x425,
    latvian = 0x426,
    lithuanian = 0x427,
    tajik = 0x428,
    farsi_persian = 0x429,
    sesotho_sutu = 0x430,
    tsonga = 0x431,
    setsuana = 0x432,
    venda = 0x433,
    xhosa = 0x434,
    zulu = 0x435,
    afrikaans = 0x436,
    georgian = 0x437,
    faroese = 0x438,
    hindi = 0x439,
    kyrgyz_cyrillic = 0x440,
    swahili = 0x441,
    turkmen = 0x442,
    uzbek_latin = 0x443,
    tatar = 0x444,
    bengali_india = 0x445,
    punjabi = 0x446,
    gujarati = 0x447,
    oriya = 0x448,
    tamil = 0x449,
    mongolian = 0x450,
    tibetan = 0x451,
    welsh = 0x452,
    khmer = 0x453,
    lao = 0x454,
    burmese = 0x455,
    galician = 0x456,
    konkani = 0x457,
    manipuri = 0x458,
    sindhi = 0x459,
    kashmiri = 0x460,
    nepali = 0x461,
    frisian_netherlands = 0x462,
    filipino = 0x464,
    divehi_dhivehi_maldivian = 0x465,
    edo = 0x466,
    igbo_nigeria = 0x470,
    guarani_paraguay = 0x474,
    latin = 0x476,
    somali = 0x477,
    maori = 0x481,
    arabic_iraq = 0x801,
    chinese_china = 0x804,
    german_switzerland = 0x807,
    english_great_britain = 0x809,
    italian_switzerland = 0x810,
    dutch_belgium = 0x813,
    norwegian_nynorsk = 0x814,
    portuguese_portugal = 0x816,
    romanian_moldova = 0x818,
    russian_moldova = 0x819,
    uzbek_cyrillic = 0x843,
    bengali_bangladesh = 0x845,
    mongolian2 = 0x850,
    arabic_libya = 0x1001,
    chinese_singapore = 0x1004,
    german_luxembourg = 0x1007,
    english_canada = 0x1009,
    arabic_algeria = 0x1401,
    chinese_macau_sar = 0x1404,
    german_liechtenstein = 0x1407,
    english_new_zealand = 0x1409,
    arabic_morocco = 0x1801,
    english_ireland = 0x1809,
    arabic_oman = 0x2001,
    english_jamaica = 0x2009,
    arabic_yemen = 0x2401,
    english_caribbean = 0x2409,
    arabic_syria = 0x2801,
    english_belize = 0x2809,
    arabic_lebanon = 0x3001,
    english_zimbabwe = 0x3009,
    arabic_kuwait = 0x3401,
    english_phillippines = 0x3409,
    arabic_united_arab_emirates = 0x3801,
    arabic_qatar = 0x4001
};

// TODO this really shouldn't be exported...
struct XLNT_API format_condition
{
    enum class condition_type
    {
        less_than,
        less_or_equal,
        equal,
        not_equal,
        greater_than,
        greater_or_equal
    } type = condition_type::not_equal;

    double value = 0.0;

    bool satisfied_by(double number) const;
};

struct format_placeholders
{
    enum class placeholders_type
    {
        general,
        text,
        integer_only,
        integer_part,
        fractional_part,
        fraction_integer,
        fraction_numerator,
        fraction_denominator,
        scientific_significand,
        scientific_exponent_plus,
        scientific_exponent_minus
    } type = placeholders_type::general;

    bool use_comma_separator = false;
    bool percentage = false;
    bool scientific = false;

    std::size_t num_zeros = 0; // 0
    std::size_t num_optionals = 0; // #
    std::size_t num_spaces = 0; // ?
    std::size_t thousands_scale = 0;
};

struct number_format_token
{
    enum class token_type
    {
        color,
        locale,
        condition,
        text,
        fill,
        space,
        number,
        datetime,
        end_section,
        end
    } type = token_type::end;

    std::string string;
};

struct template_part
{
    enum class template_type
    {
        text,
        fill,
        space,
        general,
        month_number,
        month_number_leading_zero,
        month_abbreviation,
        month_name,
        month_letter,
        day_number,
        day_number_leading_zero,
        day_abbreviation,
        day_name,
        year_short,
        year_long,
        hour,
        hour_leading_zero,
        minute,
        minute_leading_zero,
        second,
        second_fractional,
        second_leading_zero,
        second_leading_zero_fractional,
        am_pm,
        a_p,
        elapsed_hours,
        elapsed_minutes,
        elapsed_seconds
    } type = template_type::general;

    std::string string;
    format_placeholders placeholders;
};

struct format_code
{
    bool has_color = false;
    format_color color = format_color::black;
    bool has_locale = false;
    format_locale locale = format_locale::english_united_states;
    bool has_condition = false;
    format_condition condition;
    bool is_datetime = false;
    bool is_timedelta = false;
    bool twelve_hour = false;
    std::vector<template_part> parts;
};

class number_format_parser
{
public:
    number_format_parser(const std::string &format_string);
    const std::vector<format_code> &result() const;
    void reset(const std::string &format_string);
    void parse();

private:
    void finalize();
    void validate();

    number_format_token parse_next_token();

    format_placeholders parse_placeholders(const std::string &placeholders_string);
    format_color color_from_string(const std::string &color);
    std::pair<format_locale, std::string> locale_from_string(const std::string &locale_string);

    std::size_t position_ = 0;
    std::string format_string_;
    std::vector<format_code> codes_;
};

class XLNT_API number_formatter
{
public:
    number_formatter(const std::string &format_string, xlnt::calendar calendar);
    std::string format_number(double number);
    std::string format_text(const std::string &text);

private:
    std::string fill_placeholders(const format_placeholders &p, double number);
    std::string fill_fraction_placeholders(const format_placeholders &numerator,
        const format_placeholders &denominator, double number, bool improper);
    std::string fill_scientific_placeholders(const format_placeholders &integer_part,
        const format_placeholders &fractional_part, const format_placeholders &exponent_part,
        double number);
    std::string format_number(const format_code &format, double number);
    std::string format_text(const format_code &format, const std::string &text);

    number_format_parser parser_;
    std::vector<format_code> format_;
    xlnt::calendar calendar_;
};

} // namespace detail
} // namespace xlnt
