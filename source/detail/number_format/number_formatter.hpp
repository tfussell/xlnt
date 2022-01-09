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

#include <string>
#include <unordered_map>
#include <vector>

#include <xlnt/utils/datetime.hpp>
#include <xlnt/utils/numeric.hpp>

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
    arabic = 0x1,
    bulgarian = 0x2,
    catalan = 0x3,
    chinese_simplified = 0x4,
    chinese_simplified_legacy = 0x4,
    czech = 0x5,
    danish = 0x6,
    german = 0x7,
    greek = 0x8,
    english = 0x9,
    spanish = 0xA,
    finnish = 0xB,
    french = 0xC,
    hebrew = 0xD,
    hungarian = 0xE,
    icelandic = 0xF,
    italian = 0x10,
    japanese = 0x11,
    korean = 0x12,
    dutch = 0x13,
    norwegian = 0x14,
    polish = 0x15,
    portuguese = 0x16,
    romansh = 0x17,
    romanian = 0x18,
    russian = 0x19,
    croatian = 0x1A,
    slovak = 0x1B,
    albanian = 0x1C,
    swedish = 0x1D,
    thai = 0x1E,
    turkish = 0x1F,
    urdu = 0x20,
    indonesian = 0x21,
    ukrainian = 0x22,
    belarusian = 0x23,
    slovenian = 0x24,
    estonian = 0x25,
    latvian = 0x26,
    lithuanian = 0x27,
    tajik = 0x28,
    persian = 0x29,
    vietnamese = 0x2A,
    armenian = 0x2B,
    azerbaijani = 0x2C,
    basque = 0x2D,
    upper_sorbian = 0x2E,
    macedonian_fyrom = 0x2F,
    southern_sotho = 0x30,
    tsonga = 0x31,
    setswana = 0x32,
    venda = 0x33,
    isixhosa = 0x34,
    isizulu = 0x35,
    afrikaans = 0x36,
    georgian = 0x37,
    faroese = 0x38,
    hindi = 0x39,
    maltese = 0x3A,
    sami_northern = 0x3B,
    irish = 0x3C,
    yiddish = 0x3D,
    malay = 0x3E,
    kazakh = 0x3F,
    kyrgyz = 0x40,
    kiswahili = 0x41,
    turkmen = 0x42,
    uzbek = 0x43,
    tatar = 0x44,
    bangla = 0x45,
    punjabi = 0x46,
    gujarati = 0x47,
    odia = 0x48,
    tamil = 0x49,
    telugu = 0x4A,
    kannada = 0x4B,
    malayalam = 0x4C,
    assamese = 0x4D,
    marathi = 0x4E,
    sanskrit = 0x4F,
    mongolian = 0x50,
    tibetan = 0x51,
    welsh = 0x52,
    khmer = 0x53,
    lao = 0x54,
    burmese = 0x55,
    galician = 0x56,
    konkani = 0x57,
    manipuri = 0x58,
    sindhi = 0x59,
    syriac = 0x5A,
    sinhala = 0x5B,
    cherokee = 0x5C,
    inuktitut = 0x5D,
    amharic = 0x5E,
    tamazight = 0x5F,
    kashmiri = 0x60,
    nepali = 0x61,
    frisian = 0x62,
    pashto = 0x63,
    filipino = 0x64,
    divehi = 0x65,
    edo = 0x66,
    fulah = 0x67,
    hausa = 0x68,
    ibibio = 0x69,
    yoruba = 0x6A,
    quechua = 0x6B,
    sesotho_sa_leboa = 0x6C,
    bashkir = 0x6D,
    luxembourgish = 0x6E,
    greenlandic = 0x6F,
    igbo = 0x70,
    kanuri = 0x71,
    oromo = 0x72,
    tigrinya = 0x73,
    guarani = 0x74,
    hawaiian = 0x75,
    latin = 0x76,
    somali = 0x77,
    yi = 0x78,
    papiamento = 0x79,
    mapudungun = 0x7A,
    mohawk = 0x7C,
    breton = 0x7E,
    invariant_language_invariant_country = 0x7F,
    uyghur = 0x80,
    maori = 0x81,
    occitan = 0x82,
    corsican = 0x83,
    alsatian = 0x84,
    sakha = 0x85,
    kiche = 0x86,
    kinyarwanda = 0x87,
    wolof = 0x88,
    dari = 0x8C,
    scottish_gaelic = 0x91,
    central_kurdish = 0x92,
    arabic_saudi_arabia = 0x401,
    bulgarian_bulgaria = 0x402,
    catalan_catalan = 0x403,
    chinese_traditional_taiwan = 0x404,
    czech_czech_republic = 0x405,
    danish_denmark = 0x406,
    german_germany = 0x407,
    greek_greece = 0x408,
    english_united_states = 0x409,
    finnish_finland = 0x40B,
    french_france = 0x40C,
    hebrew_israel = 0x40D,
    hungarian_hungary = 0x40E,
    icelandic_iceland = 0x40F,
    italian_italy = 0x410,
    japanese_japan = 0x411,
    korean_korea = 0x412,
    dutch_netherlands = 0x413,
    norwegian_bokml_norway = 0x414,
    polish_poland = 0x415,
    portuguese_brazil = 0x416,
    romansh_switzerland = 0x417,
    romanian_romania = 0x418,
    russian_russia = 0x419,
    croatian_croatia = 0x41A,
    slovak_slovakia = 0x41B,
    albanian_albania = 0x41C,
    swedish_sweden = 0x41D,
    thai_thailand = 0x41E,
    turkish_turkey = 0x41F,
    urdu_islamic_republic_of_pakistan = 0x420,
    indonesian_indonesia = 0x421,
    ukrainian_ukraine = 0x422,
    belarusian_belarus = 0x423,
    slovenian_slovenia = 0x424,
    estonian_estonia = 0x425,
    latvian_latvia = 0x426,
    lithuanian_lithuania = 0x427,
    tajik_cyrillic_tajikistan = 0x428,
    persian_iran = 0x429,
    vietnamese_vietnam = 0x42A,
    armenian_armenia = 0x42B,
    azerbaijani_latin_azerbaijan = 0x42C,
    basque_basque = 0x42D,
    upper_sorbian_germany = 0x42E,
    macedonian_former_yugoslav_republic_of_macedonia = 0x42F,
    southern_sotho_south_africa = 0x430,
    tsonga_south_africa = 0x431,
    setswana_south_africa = 0x432,
    venda_south_africa = 0x433,
    isixhosa_south_africa = 0x434,
    isizulu_south_africa = 0x435,
    afrikaans_south_africa = 0x436,
    georgian_georgia = 0x437,
    faroese_faroe_islands = 0x438,
    hindi_india = 0x439,
    maltese_malta = 0x43A,
    sami_northern_norway = 0x43B,
    yiddish_world = 0x43D,
    malay_malaysia = 0x43E,
    kazakh_kazakhstan = 0x43F,
    kyrgyz_kyrgyzstan = 0x440,
    kiswahili_kenya = 0x441,
    turkmen_turkmenistan = 0x442,
    uzbek_latin_uzbekistan = 0x443,
    tatar_russia = 0x444,
    bangla_india = 0x445,
    punjabi_india = 0x446,
    gujarati_india = 0x447,
    odia_india = 0x448,
    tamil_india = 0x449,
    telugu_india = 0x44A,
    kannada_india = 0x44B,
    malayalam_india = 0x44C,
    assamese_india = 0x44D,
    marathi_india = 0x44E,
    sanskrit_india = 0x44F,
    mongolian_cyrillic_mongolia = 0x450,
    tibetan_prc = 0x451,
    welsh_united_kingdom = 0x452,
    khmer_cambodia = 0x453,
    lao_lao_p_d_r = 0x454,
    burmese_myanmar = 0x455,
    galician_galician = 0x456,
    konkani_india = 0x457,
    manipuri_india = 0x458,
    sindhi_devanagari_india = 0x459,
    syriac_syria = 0x45A,
    sinhala_sri_lanka = 0x45B,
    cherokee_cherokee = 0x45C,
    inuktitut_syllabics_canada = 0x45D,
    amharic_ethiopia = 0x45E,
    central_atlas_tamazight_arabic_morocco = 0x45F,
    kashmiri_perso_arabic = 0x460,
    nepali_nepal = 0x461,
    frisian_netherlands = 0x462,
    pashto_afghanistan = 0x463,
    filipino_philippines = 0x464,
    divehi_maldives = 0x465,
    edo_nigeria = 0x466,
    fulah_nigeria = 0x467,
    hausa_latin_nigeria = 0x468,
    ibibio_nigeria = 0x469,
    yoruba_nigeria = 0x46A,
    quechua_bolivia = 0x46B,
    sesotho_sa_leboa_south_africa = 0x46C,
    bashkir_russia = 0x46D,
    luxembourgish_luxembourg = 0x46E,
    greenlandic_greenland = 0x46F,
    igbo_nigeria = 0x470,
    kanuri_nigeria = 0x471,
    oromo_ethiopia = 0x472,
    tigrinya_ethiopia = 0x473,
    guarani_paraguay = 0x474,
    hawaiian_united_states = 0x475,
    latin_world = 0x476,
    somali_somalia = 0x477,
    yi_prc = 0x478,
    papiamento_caribbean = 0x479,
    mapudungun_chile = 0x47A,
    mohawk_mohawk = 0x47C,
    breton_france = 0x47E,
    uyghur_prc = 0x480,
    maori_new_zealand = 0x481,
    occitan_france = 0x482,
    corsican_france = 0x483,
    alsatian_france = 0x484,
    sakha_russia = 0x485,
    kiche_guatemala = 0x486,
    kinyarwanda_rwanda = 0x487,
    wolof_senegal = 0x488,
    dari_afghanistan = 0x48C,
    scottish_gaelic_united_kingdom = 0x491,
    central_kurdish_iraq = 0x492,
    arabic_iraq = 0x801,
    valencian_spain = 0x803,
    chinese_simplified_prc = 0x804,
    german_switzerland = 0x807,
    english_united_kingdom = 0x809,
    spanish_mexico = 0x80A,
    french_belgium = 0x80C,
    italian_switzerland = 0x810,
    dutch_belgium = 0x813,
    norwegian_nynorsk_norway = 0x814,
    portuguese_portugal = 0x816,
    romanian_moldova = 0x818,
    russian_moldova = 0x819,
    swedish_finland = 0x81D,
    urdu_india = 0x820,
    azerbaijani_cyrillic_azerbaijan = 0x82C,
    lower_sorbian_germany = 0x82E,
    setswana_botswana = 0x832,
    sami_northern_sweden = 0x83B,
    irish_ireland = 0x83C,
    malay_brunei_darussalam = 0x83E,
    uzbek_cyrillic_uzbekistan = 0x843,
    bangla_bangladesh = 0x845,
    punjabi_islamic_republic_of_pakistan = 0x846,
    tamil_sri_lanka = 0x849,
    mongolian_traditional_mongolian_prc = 0x850,
    sindhi_islamic_republic_of_pakistan = 0x859,
    inuktitut_latin_canada = 0x85D,
    tamazight_latin_algeria = 0x85F,
    kashmiri_devanagari_india = 0x860,
    nepali_india = 0x861,
    fulah_latin_senegal = 0x867,
    quechua_ecuador = 0x86B,
    tigrinya_eritrea = 0x873,
    arabic_egypt = 0xC01,
    chinese_traditional_hong_kong_s_a_r = 0xC04,
    german_austria = 0xC07,
    english_australia = 0xC09,
    spanish_spain = 0xC0A,
    french_canada = 0xC0C,
    sami_northern_finland = 0xC3B,
    mongolian_traditional_mongolian_mongolia = 0xC50,
    dzongkha_bhutan = 0xC51,
    quechua_peru = 0xC6B,
    arabic_libya = 0x1001,
    chinese_simplified_singapore = 0x1004,
    german_luxembourg = 0x1007,
    english_canada = 0x1009,
    spanish_guatemala = 0x100A,
    french_switzerland = 0x100C,
    croatian_latin_bosnia_and_herzegovina = 0x101A,
    sami_lule_norway = 0x103B,
    central_atlas_tamazight_tifinagh_morocco = 0x105F,
    arabic_algeria = 0x1401,
    chinese_traditional_macao_s_a_r = 0x1404,
    german_liechtenstein = 0x1407,
    english_new_zealand = 0x1409,
    spanish_costa_rica = 0x140A,
    french_luxembourg = 0x140C,
    bosnian_latin_bosnia_and_herzegovina = 0x141A,
    sami_lule_sweden = 0x143B,
    arabic_morocco = 0x1801,
    english_ireland = 0x1809,
    spanish_panama = 0x180A,
    french_monaco = 0x180C,
    serbian_latin_bosnia_and_herzegovina = 0x181A,
    sami_southern_norway = 0x183B,
    arabic_tunisia = 0x1C01,
    english_south_africa = 0x1C09,
    spanish_dominican_republic = 0x1C0A,
    french_caribbean = 0x1C0C,
    serbian_cyrillic_bosnia_and_herzegovina = 0x1C1A,
    sami_southern_sweden = 0x1C3B,
    arabic_oman = 0x2001,
    english_jamaica = 0x2009,
    spanish_venezuela = 0x200A,
    french_reunion = 0x200C,
    bosnian_cyrillic_bosnia_and_herzegovina = 0x201A,
    sami_skolt_finland = 0x203B,
    arabic_yemen = 0x2401,
    english_caribbean = 0x2409,
    spanish_colombia = 0x240A,
    french_congo_drc = 0x240C,
    serbian_latin_serbia = 0x241A,
    sami_inari_finland = 0x243B,
    arabic_syria = 0x2801,
    english_belize = 0x2809,
    spanish_peru = 0x280A,
    french_senegal = 0x280C,
    serbian_cyrillic_serbia = 0x281A,
    arabic_jordan = 0x2C01,
    english_trinidad_and_tobago = 0x2C09,
    spanish_argentina = 0x2C0A,
    french_cameroon = 0x2C0C,
    serbian_latin_montenegro = 0x2C1A,
    arabic_lebanon = 0x3001,
    english_zimbabwe = 0x3009,
    spanish_ecuador = 0x300A,
    french_cote_divoire = 0x300C,
    serbian_cyrillic_montenegro = 0x301A,
    arabic_kuwait = 0x3401,
    english_philippines = 0x3409,
    spanish_chile = 0x340A,
    french_mali = 0x340C,
    arabic_u_a_e = 0x3801,
    english_indonesia = 0x3809,
    spanish_uruguay = 0x380A,
    french_morocco = 0x380C,
    arabic_bahrain = 0x3C01,
    english_hong_kong_sar = 0x3C09,
    spanish_paraguay = 0x3C0A,
    french_haiti = 0x3C0C,
    arabic_qatar = 0x4001,
    english_india = 0x4009,
    spanish_bolivia = 0x400A,
    english_malaysia = 0x4409,
    spanish_el_salvador = 0x440A,
    english_singapore = 0x4809,
    spanish_honduras = 0x480A,
    spanish_nicaragua = 0x4C0A,
    spanish_puerto_rico = 0x500A,
    spanish_united_states = 0x540A,
    spanish_latin_america = 0x580A,
    spanish_cuba = 0x5C0A,
    bosnian_cyrillic = 0x641A,
    bosnian_latin = 0x681A,
    serbian_cyrillic = 0x6C1A,
    serbian_latin = 0x701A,
    sami_inari = 0x703B,
    azerbaijani_cyrillic = 0x742C,
    sami_skolt = 0x743B,
    chinese = 0x7804,
    norwegian_nynorsk = 0x7814,
    bosnian = 0x781A,
    azerbaijani_latin = 0x782C,
    sami_southern = 0x783B,
    uzbek_cyrillic = 0x7843,
    mongolian_cyrillic = 0x7850,
    inuktitut_syllabics = 0x785D,
    tamazight_tifinagh = 0x785F,
    chinese_traditional = 0x7C04,
    chinese_traditional_legacy = 0x7C04,
    norwegian_bokml = 0x7C14,
    serbian = 0x7C1A,
    tajik_cyrillic = 0x7C28,
    lower_sorbian = 0x7C2E,
    sami_lule = 0x7C3B,
    uzbek_latin = 0x7C43,
    punjabi_arabic = 0x7C46,
    mongolian_traditional_mongolian = 0x7C50,
    sindhi_arabic = 0x7C59,
    inuktitut_latin = 0x7C5D,
    tamazight_latin = 0x7C5F,
    fulah_latin = 0x7C67,
    hausa_latin = 0x7C68,
    central_kurdish_arabic = 0x7C92,
    system_default_time = 0xF400,
    system_default_long_date = 0xF800
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
    xlnt::detail::number_serialiser serialiser_;
};

} // namespace detail
} // namespace xlnt
