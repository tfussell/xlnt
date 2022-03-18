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

#include <algorithm>
#include <cctype>
#include <cmath>
#include <limits>

#include <xlnt/utils/exceptions.hpp>
#include <xlnt/utils/numeric.hpp>
#include <detail/default_case.hpp>
#include <detail/number_format/number_formatter.hpp>

namespace {

const std::unordered_map<int, std::string> known_locales()
{
    static const std::unordered_map<int, std::string> all = std::unordered_map<int, std::string>(
        {
            {0x1, "Arabic"},
            {0x2, "Bulgarian"},
            {0x3, "Catalan"},
            {0x4, "Chinese (Simplified)"},
            {0x4, "Chinese (Simplified) Legacy"},
            {0x5, "Czech"},
            {0x6, "Danish"},
            {0x7, "German"},
            {0x8, "Greek"},
            {0x9, "English"},
            {0xA, "Spanish"},
            {0xB, "Finnish"},
            {0xC, "French"},
            {0xD, "Hebrew"},
            {0xE, "Hungarian"},
            {0xF, "Icelandic"},
            {0x10, "Italian"},
            {0x11, "Japanese"},
            {0x12, "Korean"},
            {0x13, "Dutch"},
            {0x14, "Norwegian"},
            {0x15, "Polish"},
            {0x16, "Portuguese"},
            {0x17, "Romansh"},
            {0x18, "Romanian"},
            {0x19, "Russian"},
            {0x1A, "Croatian"},
            {0x1B, "Slovak"},
            {0x1C, "Albanian"},
            {0x1D, "Swedish"},
            {0x1E, "Thai"},
            {0x1F, "Turkish"},
            {0x20, "Urdu"},
            {0x21, "Indonesian"},
            {0x22, "Ukrainian"},
            {0x23, "Belarusian"},
            {0x24, "Slovenian"},
            {0x25, "Estonian"},
            {0x26, "Latvian"},
            {0x27, "Lithuanian"},
            {0x28, "Tajik"},
            {0x29, "Persian"},
            {0x2A, "Vietnamese"},
            {0x2B, "Armenian"},
            {0x2C, "Azerbaijani"},
            {0x2D, "Basque"},
            {0x2E, "Upper Sorbian"},
            {0x2F, "Macedonian (FYROM)"},
            {0x30, "Southern Sotho"},
            {0x31, "Tsonga"},
            {0x32, "Setswana"},
            {0x33, "Venda"},
            {0x34, "isiXhosa"},
            {0x35, "isiZulu"},
            {0x36, "Afrikaans"},
            {0x37, "Georgian"},
            {0x38, "Faroese"},
            {0x39, "Hindi"},
            {0x3A, "Maltese"},
            {0x3B, "Sami (Northern)"},
            {0x3C, "Irish"},
            {0x3D, "Yiddish"},
            {0x3E, "Malay"},
            {0x3F, "Kazakh"},
            {0x40, "Kyrgyz"},
            {0x41, "Kiswahili"},
            {0x42, "Turkmen"},
            {0x43, "Uzbek"},
            {0x44, "Tatar"},
            {0x45, "Bangla"},
            {0x46, "Punjabi"},
            {0x47, "Gujarati"},
            {0x48, "Odia"},
            {0x49, "Tamil"},
            {0x4A, "Telugu"},
            {0x4B, "Kannada"},
            {0x4C, "Malayalam"},
            {0x4D, "Assamese"},
            {0x4E, "Marathi"},
            {0x4F, "Sanskrit"},
            {0x50, "Mongolian"},
            {0x51, "Tibetan"},
            {0x52, "Welsh"},
            {0x53, "Khmer"},
            {0x54, "Lao"},
            {0x55, "Burmese"},
            {0x56, "Galician"},
            {0x57, "Konkani"},
            {0x58, "Manipuri"},
            {0x59, "Sindhi"},
            {0x5A, "Syriac"},
            {0x5B, "Sinhala"},
            {0x5C, "Cherokee"},
            {0x5D, "Inuktitut"},
            {0x5E, "Amharic"},
            {0x5F, "Tamazight"},
            {0x60, "Kashmiri"},
            {0x61, "Nepali"},
            {0x62, "Frisian"},
            {0x63, "Pashto"},
            {0x64, "Filipino"},
            {0x65, "Divehi"},
            {0x66, "Edo"},
            {0x67, "Fulah"},
            {0x68, "Hausa"},
            {0x69, "Ibibio"},
            {0x6A, "Yoruba"},
            {0x6B, "Quechua"},
            {0x6C, "Sesotho sa Leboa"},
            {0x6D, "Bashkir"},
            {0x6E, "Luxembourgish"},
            {0x6F, "Greenlandic"},
            {0x70, "Igbo"},
            {0x71, "Kanuri"},
            {0x72, "Oromo"},
            {0x73, "Tigrinya"},
            {0x74, "Guarani"},
            {0x75, "Hawaiian"},
            {0x76, "Latin"},
            {0x77, "Somali"},
            {0x78, "Yi"},
            {0x79, "Papiamento"},
            {0x7A, "Mapudungun"},
            {0x7C, "Mohawk"},
            {0x7E, "Breton"},
            {0x7F, "Invariant Language (Invariant Country)"},
            {0x80, "Uyghur"},
            {0x81, "Maori"},
            {0x82, "Occitan"},
            {0x83, "Corsican"},
            {0x84, "Alsatian"},
            {0x85, "Sakha"},
            {0x86, "K’iche’"},
            {0x87, "Kinyarwanda"},
            {0x88, "Wolof"},
            {0x8C, "Dari"},
            {0x91, "Scottish Gaelic"},
            {0x92, "Central Kurdish"},
            {0x401, "Arabic (Saudi Arabia)"},
            {0x402, "Bulgarian (Bulgaria)"},
            {0x403, "Catalan (Catalan)"},
            {0x404, "Chinese (Traditional, Taiwan)"},
            {0x405, "Czech (Czech Republic)"},
            {0x406, "Danish (Denmark)"},
            {0x407, "German (Germany)"},
            {0x408, "Greek (Greece)"},
            {0x409, "English (United States)"},
            {0x40B, "Finnish (Finland)"},
            {0x40C, "French (France)"},
            {0x40D, "Hebrew (Israel)"},
            {0x40E, "Hungarian (Hungary)"},
            {0x40F, "Icelandic (Iceland)"},
            {0x410, "Italian (Italy)"},
            {0x411, "Japanese (Japan)"},
            {0x412, "Korean (Korea)"},
            {0x413, "Dutch (Netherlands)"},
            {0x414, "Norwegian, Bokmål (Norway)"},
            {0x415, "Polish (Poland)"},
            {0x416, "Portuguese (Brazil)"},
            {0x417, "Romansh (Switzerland)"},
            {0x418, "Romanian (Romania)"},
            {0x419, "Russian (Russia)"},
            {0x41A, "Croatian (Croatia)"},
            {0x41B, "Slovak (Slovakia)"},
            {0x41C, "Albanian (Albania)"},
            {0x41D, "Swedish (Sweden)"},
            {0x41E, "Thai (Thailand)"},
            {0x41F, "Turkish (Turkey)"},
            {0x420, "Urdu (Islamic Republic of Pakistan)"},
            {0x421, "Indonesian (Indonesia)"},
            {0x422, "Ukrainian (Ukraine)"},
            {0x423, "Belarusian (Belarus)"},
            {0x424, "Slovenian (Slovenia)"},
            {0x425, "Estonian (Estonia)"},
            {0x426, "Latvian (Latvia)"},
            {0x427, "Lithuanian (Lithuania)"},
            {0x428, "Tajik (Cyrillic, Tajikistan)"},
            {0x429, "Persian (Iran)"},
            {0x42A, "Vietnamese (Vietnam)"},
            {0x42B, "Armenian (Armenia)"},
            {0x42C, "Azerbaijani (Latin, Azerbaijan)"},
            {0x42D, "Basque (Basque)"},
            {0x42E, "Upper Sorbian (Germany)"},
            {0x42F, "Macedonian (Former Yugoslav Republic of Macedonia)"},
            {0x430, "Southern Sotho (South Africa)"},
            {0x431, "Tsonga (South Africa)"},
            {0x432, "Setswana (South Africa)"},
            {0x433, "Venda (South Africa)"},
            {0x434, "isiXhosa (South Africa)"},
            {0x435, "isiZulu (South Africa)"},
            {0x436, "Afrikaans (South Africa)"},
            {0x437, "Georgian (Georgia)"},
            {0x438, "Faroese (Faroe Islands)"},
            {0x439, "Hindi (India)"},
            {0x43A, "Maltese (Malta)"},
            {0x43B, "Sami, Northern (Norway)"},
            {0x43D, "Yiddish (World)"},
            {0x43E, "Malay (Malaysia)"},
            {0x43F, "Kazakh (Kazakhstan)"},
            {0x440, "Kyrgyz (Kyrgyzstan)"},
            {0x441, "Kiswahili (Kenya)"},
            {0x442, "Turkmen (Turkmenistan)"},
            {0x443, "Uzbek (Latin, Uzbekistan)"},
            {0x444, "Tatar (Russia)"},
            {0x445, "Bangla (India)"},
            {0x446, "Punjabi (India)"},
            {0x447, "Gujarati (India)"},
            {0x448, "Odia (India)"},
            {0x449, "Tamil (India)"},
            {0x44A, "Telugu (India)"},
            {0x44B, "Kannada (India)"},
            {0x44C, "Malayalam (India)"},
            {0x44D, "Assamese (India)"},
            {0x44E, "Marathi (India)"},
            {0x44F, "Sanskrit (India)"},
            {0x450, "Mongolian (Cyrillic, Mongolia)"},
            {0x451, "Tibetan (PRC)"},
            {0x452, "Welsh (United Kingdom)"},
            {0x453, "Khmer (Cambodia)"},
            {0x454, "Lao (Lao P.D.R.)"},
            {0x455, "Burmese (Myanmar)"},
            {0x456, "Galician (Galician)"},
            {0x457, "Konkani (India)"},
            {0x458, "Manipuri (India)"},
            {0x459, "Sindhi (Devanagari, India)"},
            {0x45A, "Syriac (Syria)"},
            {0x45B, "Sinhala (Sri Lanka)"},
            {0x45C, "Cherokee (Cherokee)"},
            {0x45D, "Inuktitut (Syllabics, Canada)"},
            {0x45E, "Amharic (Ethiopia)"},
            {0x45F, "Central Atlas Tamazight (Arabic, Morocco)"},
            {0x460, "Kashmiri (Perso-Arabic)"},
            {0x461, "Nepali (Nepal)"},
            {0x462, "Frisian (Netherlands)"},
            {0x463, "Pashto (Afghanistan)"},
            {0x464, "Filipino (Philippines)"},
            {0x465, "Divehi (Maldives)"},
            {0x466, "Edo (Nigeria)"},
            {0x467, "Fulah (Nigeria)"},
            {0x468, "Hausa (Latin, Nigeria)"},
            {0x469, "Ibibio (Nigeria)"},
            {0x46A, "Yoruba (Nigeria)"},
            {0x46B, "Quechua (Bolivia)"},
            {0x46C, "Sesotho sa Leboa (South Africa)"},
            {0x46D, "Bashkir (Russia)"},
            {0x46E, "Luxembourgish (Luxembourg)"},
            {0x46F, "Greenlandic (Greenland)"},
            {0x470, "Igbo (Nigeria)"},
            {0x471, "Kanuri (Nigeria)"},
            {0x472, "Oromo (Ethiopia)"},
            {0x473, "Tigrinya (Ethiopia)"},
            {0x474, "Guarani (Paraguay)"},
            {0x475, "Hawaiian (United States)"},
            {0x476, "Latin (World)"},
            {0x477, "Somali (Somalia)"},
            {0x478, "Yi (PRC)"},
            {0x479, "Papiamento (Caribbean)"},
            {0x47A, "Mapudungun (Chile)"},
            {0x47C, "Mohawk (Mohawk)"},
            {0x47E, "Breton (France)"},
            {0x480, "Uyghur (PRC)"},
            {0x481, "Maori (New Zealand)"},
            {0x482, "Occitan (France)"},
            {0x483, "Corsican (France)"},
            {0x484, "Alsatian (France)"},
            {0x485, "Sakha (Russia)"},
            {0x486, "K’iche’ (Guatemala)"},
            {0x487, "Kinyarwanda (Rwanda)"},
            {0x488, "Wolof (Senegal)"},
            {0x48C, "Dari (Afghanistan)"},
            {0x491, "Scottish Gaelic (United Kingdom)"},
            {0x492, "Central Kurdish (Iraq)"},
            {0x801, "Arabic (Iraq)"},
            {0x803, "Valencian (Spain)"},
            {0x804, "Chinese (Simplified, PRC)"},
            {0x807, "German (Switzerland)"},
            {0x809, "English (United Kingdom)"},
            {0x80A, "Spanish (Mexico)"},
            {0x80C, "French (Belgium)"},
            {0x810, "Italian (Switzerland)"},
            {0x813, "Dutch (Belgium)"},
            {0x814, "Norwegian, Nynorsk (Norway)"},
            {0x816, "Portuguese (Portugal)"},
            {0x818, "Romanian (Moldova)"},
            {0x819, "Russian (Moldova)"},
            {0x81D, "Swedish (Finland)"},
            {0x820, "Urdu (India)"},
            {0x82C, "Azerbaijani (Cyrillic, Azerbaijan)"},
            {0x82E, "Lower Sorbian (Germany)"},
            {0x832, "Setswana (Botswana)"},
            {0x83B, "Sami, Northern (Sweden)"},
            {0x83C, "Irish (Ireland)"},
            {0x83E, "Malay (Brunei Darussalam)"},
            {0x843, "Uzbek (Cyrillic, Uzbekistan)"},
            {0x845, "Bangla (Bangladesh)"},
            {0x846, "Punjabi (Islamic Republic of Pakistan)"},
            {0x849, "Tamil (Sri Lanka)"},
            {0x850, "Mongolian (Traditional Mongolian, PRC)"},
            {0x859, "Sindhi (Islamic Republic of Pakistan)"},
            {0x85D, "Inuktitut (Latin, Canada)"},
            {0x85F, "Tamazight (Latin, Algeria)"},
            {0x860, "Kashmiri (Devanagari, India)"},
            {0x861, "Nepali (India)"},
            {0x867, "Fulah (Latin, Senegal)"},
            {0x86B, "Quechua (Ecuador)"},
            {0x873, "Tigrinya (Eritrea)"},
            {0xC01, "Arabic (Egypt)"},
            {0xC04, "Chinese (Traditional, Hong Kong S.A.R.)"},
            {0xC07, "German (Austria)"},
            {0xC09, "English (Australia)"},
            {0xC0A, "Spanish (Spain)"},
            {0xC0C, "French (Canada)"},
            {0xC3B, "Sami, Northern (Finland)"},
            {0xC50, "Mongolian (Traditional Mongolian, Mongolia)"},
            {0xC51, "Dzongkha (Bhutan)"},
            {0xC6B, "Quechua (Peru)"},
            {0x1001, "Arabic (Libya)"},
            {0x1004, "Chinese (Simplified, Singapore)"},
            {0x1007, "German (Luxembourg)"},
            {0x1009, "English (Canada)"},
            {0x100A, "Spanish (Guatemala)"},
            {0x100C, "French (Switzerland)"},
            {0x101A, "Croatian (Latin, Bosnia and Herzegovina)"},
            {0x103B, "Sami, Lule (Norway)"},
            {0x105F, "Central Atlas Tamazight (Tifinagh, Morocco)"},
            {0x1401, "Arabic (Algeria)"},
            {0x1404, "Chinese (Traditional, Macao S.A.R.)"},
            {0x1407, "German (Liechtenstein)"},
            {0x1409, "English (New Zealand)"},
            {0x140A, "Spanish (Costa Rica)"},
            {0x140C, "French (Luxembourg)"},
            {0x141A, "Bosnian (Latin, Bosnia and Herzegovina)"},
            {0x143B, "Sami, Lule (Sweden)"},
            {0x1801, "Arabic (Morocco)"},
            {0x1809, "English (Ireland)"},
            {0x180A, "Spanish (Panama)"},
            {0x180C, "French (Monaco)"},
            {0x181A, "Serbian (Latin, Bosnia and Herzegovina)"},
            {0x183B, "Sami, Southern (Norway)"},
            {0x1C01, "Arabic (Tunisia)"},
            {0x1C09, "English (South Africa)"},
            {0x1C0A, "Spanish (Dominican Republic)"},
            {0x1C0C, "French (Caribbean)"},
            {0x1C1A, "Serbian (Cyrillic, Bosnia and Herzegovina)"},
            {0x1C3B, "Sami, Southern (Sweden)"},
            {0x2001, "Arabic (Oman)"},
            {0x2009, "English (Jamaica)"},
            {0x200A, "Spanish (Venezuela)"},
            {0x200C, "French (Reunion)"},
            {0x201A, "Bosnian (Cyrillic, Bosnia and Herzegovina)"},
            {0x203B, "Sami, Skolt (Finland)"},
            {0x2401, "Arabic (Yemen)"},
            {0x2409, "English (Caribbean)"},
            {0x240A, "Spanish (Colombia)"},
            {0x240C, "French (Congo DRC)"},
            {0x241A, "Serbian (Latin, Serbia)"},
            {0x243B, "Sami, Inari (Finland)"},
            {0x2801, "Arabic (Syria)"},
            {0x2809, "English (Belize)"},
            {0x280A, "Spanish (Peru)"},
            {0x280C, "French (Senegal)"},
            {0x281A, "Serbian (Cyrillic, Serbia)"},
            {0x2C01, "Arabic (Jordan)"},
            {0x2C09, "English (Trinidad and Tobago)"},
            {0x2C0A, "Spanish (Argentina)"},
            {0x2C0C, "French (Cameroon)"},
            {0x2C1A, "Serbian (Latin, Montenegro)"},
            {0x3001, "Arabic (Lebanon)"},
            {0x3009, "English (Zimbabwe)"},
            {0x300A, "Spanish (Ecuador)"},
            {0x300C, "French (Côte d’Ivoire)"},
            {0x301A, "Serbian (Cyrillic, Montenegro)"},
            {0x3401, "Arabic (Kuwait)"},
            {0x3409, "English (Philippines)"},
            {0x340A, "Spanish (Chile)"},
            {0x340C, "French (Mali)"},
            {0x3801, "Arabic (U.A.E.)"},
            {0x3809, "English (Indonesia)"},
            {0x380A, "Spanish (Uruguay)"},
            {0x380C, "French (Morocco)"},
            {0x3C01, "Arabic (Bahrain)"},
            {0x3C09, "English (Hong Kong SAR)"},
            {0x3C0A, "Spanish (Paraguay)"},
            {0x3C0C, "French (Haiti)"},
            {0x4001, "Arabic (Qatar)"},
            {0x4009, "English (India)"},
            {0x400A, "Spanish (Bolivia)"},
            {0x4409, "English (Malaysia)"},
            {0x440A, "Spanish (El Salvador)"},
            {0x4809, "English (Singapore)"},
            {0x480A, "Spanish (Honduras)"},
            {0x4C0A, "Spanish (Nicaragua)"},
            {0x500A, "Spanish (Puerto Rico)"},
            {0x540A, "Spanish (United States)"},
            {0x580A, "Spanish (Latin America)"},
            {0x5C0A, "Spanish (Cuba)"},
            {0x641A, "Bosnian (Cyrillic)"},
            {0x681A, "Bosnian (Latin)"},
            {0x6C1A, "Serbian (Cyrillic)"},
            {0x701A, "Serbian (Latin)"},
            {0x703B, "Sami (Inari)"},
            {0x742C, "Azerbaijani (Cyrillic)"},
            {0x743B, "Sami (Skolt)"},
            {0x7804, "Chinese"},
            {0x7814, "Norwegian (Nynorsk)"},
            {0x781A, "Bosnian"},
            {0x782C, "Azerbaijani (Latin)"},
            {0x783B, "Sami (Southern)"},
            {0x7843, "Uzbek (Cyrillic)"},
            {0x7850, "Mongolian (Cyrillic)"},
            {0x785D, "Inuktitut (Syllabics)"},
            {0x785F, "Tamazight (Tifinagh)"},
            {0x7C04, "Chinese (Traditional)"},
            {0x7C04, "Chinese (Traditional) Legacy"},
            {0x7C14, "Norwegian (Bokmål)"},
            {0x7C1A, "Serbian"},
            {0x7C28, "Tajik (Cyrillic)"},
            {0x7C2E, "Lower Sorbian"},
            {0x7C3B, "Sami (Lule)"},
            {0x7C43, "Uzbek (Latin)"},
            {0x7C46, "Punjabi (Arabic)"},
            {0x7C50, "Mongolian (Traditional Mongolian)"},
            {0x7C59, "Sindhi (Arabic)"},
            {0x7C5C, "Cherokee (Cherokee)"},
            {0x7C5D, "Inuktitut (Latin)"},
            {0x7C5F, "Tamazight (Latin)"},
            {0x7C67, "Fulah (Latin)"},
            {0x7C68, "Hausa (Latin)"},
            {0x7C86, "K’iche’"},
            {0x7C92, "Central Kurdish (Arabic)"},
            {0xF400, "System Default for Time"},
            {0xF800, "System Default for Long Date"},
        });

    return all;
}

[[noreturn]] void unhandled_case_error()
{
    throw xlnt::exception("unhandled");
}

void unhandled_case(bool error)
{
    if (error)
    {
        unhandled_case_error();
    }
}

} // namespace

namespace xlnt {
namespace detail {

bool format_condition::satisfied_by(double number) const
{
    switch (type)
    {
    case condition_type::greater_or_equal:
        return number >= value;
    case condition_type::greater_than:
        return number > value;
    case condition_type::less_or_equal:
        return number <= value;
    case condition_type::less_than:
        return number < value;
    case condition_type::not_equal:
        return std::fabs(number - value) != 0.0;
    case condition_type::equal:
        return std::fabs(number - value) == 0.0;
    }

    default_case(false);
}

number_format_parser::number_format_parser(const std::string &format_string)
{
    reset(format_string);
}

const std::vector<format_code> &number_format_parser::result() const
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
        case number_format_token::token_type::end_section: {
            codes_.push_back(section);
            section = format_code();

            break;
        }

        case number_format_token::token_type::color: {
            if (section.has_color || section.has_condition || section.has_locale || !section.parts.empty())
            {
                throw xlnt::exception("color should be the first part of a format");
            }

            section.has_color = true;
            section.color = color_from_string(token.string);

            break;
        }

        case number_format_token::token_type::locale: {
            if (section.has_locale)
            {
                throw xlnt::exception("multiple locales");
            }

            section.has_locale = true;
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

        case number_format_token::token_type::condition: {
            if (section.has_condition)
            {
                throw xlnt::exception("multiple conditions");
            }

            section.has_condition = true;
            std::string value;

            if (token.string.front() == '<')
            {
                if (token.string[1] == '=')
                {
                    section.condition.type = format_condition::condition_type::less_or_equal;
                    value = token.string.substr(2);
                }
                else if (token.string[1] == '>')
                {
                    section.condition.type = format_condition::condition_type::not_equal;
                    value = token.string.substr(2);
                }
                else
                {
                    section.condition.type = format_condition::condition_type::less_than;
                    value = token.string.substr(1);
                }
            }
            else if (token.string.front() == '>')
            {
                if (token.string[1] == '=')
                {
                    section.condition.type = format_condition::condition_type::greater_or_equal;
                    value = token.string.substr(2);
                }
                else
                {
                    section.condition.type = format_condition::condition_type::greater_than;
                    value = token.string.substr(1);
                }
            }
            else if (token.string.front() == '=')
            {
                section.condition.type = format_condition::condition_type::equal;
                value = token.string.substr(1);
            }

            detail::number_serialiser ser;
            section.condition.value = ser.deserialise(value);
            break;
        }

        case number_format_token::token_type::text: {
            part.type = template_part::template_type::text;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;
        }

        case number_format_token::token_type::fill: {
            part.type = template_part::template_type::fill;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;
        }

        case number_format_token::token_type::space: {
            part.type = template_part::template_type::space;
            part.string = token.string;
            section.parts.push_back(part);
            part = template_part();

            break;
        }

        case number_format_token::token_type::number: {
            part.type = template_part::template_type::general;
            part.placeholders = parse_placeholders(token.string);
            section.parts.push_back(part);
            part = template_part();

            break;
        }

        case number_format_token::token_type::datetime: {
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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

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

                unhandled_case(true);
                break;

            default:
                unhandled_case(true);
                break;
            }

            section.parts.push_back(part);
            part = template_part();

            break;
        }

        case number_format_token::token_type::end: {
            codes_.push_back(section);
            finalize();

            return;
        }
        }

        token = parse_next_token();
    }
}

void number_format_parser::finalize()
{
    for (auto &code : codes_)
    {
        bool fix = false;
        bool leading_zero = false;
        std::size_t minutes_index = 0;

        bool integer_part = false;
        bool fractional_part = false;
        std::size_t integer_part_index = 0;

        bool percentage = false;

        bool exponent = false;
        std::size_t exponent_index = 0;

        bool fraction = false;
        std::size_t fraction_denominator_index = 0;
        std::size_t fraction_numerator_index = 0;

        bool seconds = false;
        bool fractional_seconds = false;
        std::size_t seconds_index = 0;

        for (std::size_t i = 0; i < code.parts.size(); ++i)
        {
            const auto &part = code.parts[i];

            if (i > 0 && i + 1 < code.parts.size() && part.type == template_part::template_type::text
                && part.string == "/"
                && code.parts[i - 1].placeholders.type == format_placeholders::placeholders_type::integer_part
                && code.parts[i + 1].placeholders.type == format_placeholders::placeholders_type::integer_part)
            {
                fraction = true;
                fraction_numerator_index = i - 1;
                fraction_denominator_index = i + 1;
            }

            if (part.placeholders.type == format_placeholders::placeholders_type::integer_part)
            {
                integer_part = true;
                integer_part_index = i;
            }
            else if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
            {
                fractional_part = true;
            }
            else if (part.placeholders.type == format_placeholders::placeholders_type::scientific_exponent_plus
                || part.placeholders.type == format_placeholders::placeholders_type::scientific_exponent_minus)
            {
                exponent = true;
                exponent_index = i;
            }

            if (part.placeholders.percentage)
            {
                percentage = true;
            }

            if (part.type == template_part::template_type::second
                || part.type == template_part::template_type::second_leading_zero)
            {
                seconds = true;
                seconds_index = i;
            }

            if (seconds && part.placeholders.type == format_placeholders::placeholders_type::fractional_part)
            {
                fractional_seconds = true;
            }

            // TODO this block needs improvement
            if (part.type == template_part::template_type::month_number
                || part.type == template_part::template_type::month_number_leading_zero)
            {
                if (code.parts.size() > 1 && i < code.parts.size() - 2)
                {
                    const auto &next = code.parts[i + 1];
                    const auto &after_next = code.parts[i + 2];

                    if ((next.type == template_part::template_type::second
                            || next.type == template_part::template_type::second_leading_zero)
                        || (next.type == template_part::template_type::text && next.string == ":"
                            && (after_next.type == template_part::template_type::second
                                || after_next.type == template_part::template_type::second_leading_zero)))
                    {
                        fix = true;
                        leading_zero = part.type == template_part::template_type::month_number_leading_zero;
                        minutes_index = i;
                    }
                }

                if (!fix && i > 1)
                {
                    const auto &previous = code.parts[i - 1];
                    const auto &before_previous = code.parts[i - 2];

                    if (previous.type == template_part::template_type::text && previous.string == ":"
                        && (before_previous.type == template_part::template_type::hour_leading_zero
                            || before_previous.type == template_part::template_type::hour))
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
            code.parts[minutes_index].type =
                leading_zero ? template_part::template_type::minute_leading_zero : template_part::template_type::minute;
        }

        if (integer_part && !fractional_part)
        {
            code.parts[integer_part_index].placeholders.type = format_placeholders::placeholders_type::integer_only;
        }

        if (integer_part && fractional_part && percentage)
        {
            code.parts[integer_part_index].placeholders.percentage = true;
        }

        if (exponent)
        {
            const auto &next = code.parts[exponent_index + 1];
            auto temp = code.parts[exponent_index].placeholders.type;
            code.parts[exponent_index].placeholders = next.placeholders;
            code.parts[exponent_index].placeholders.type = temp;
            code.parts.erase(code.parts.begin() + static_cast<std::ptrdiff_t>(exponent_index + 1));

            for (std::size_t i = 0; i < code.parts.size(); ++i)
            {
                code.parts[i].placeholders.scientific = true;
            }
        }

        if (fraction)
        {
            code.parts[fraction_numerator_index].placeholders.type =
                format_placeholders::placeholders_type::fraction_numerator;
            code.parts[fraction_denominator_index].placeholders.type =
                format_placeholders::placeholders_type::fraction_denominator;

            for (std::size_t i = 0; i < code.parts.size(); ++i)
            {
                if (code.parts[i].placeholders.type == format_placeholders::placeholders_type::integer_part)
                {
                    code.parts[i].placeholders.type = format_placeholders::placeholders_type::fraction_integer;
                }
            }
        }

        if (fractional_seconds)
        {
            if (code.parts[seconds_index].type == template_part::template_type::second)
            {
                code.parts[seconds_index].type = template_part::template_type::second_fractional;
            }
            else
            {
                code.parts[seconds_index].type = template_part::template_type::second_leading_zero_fractional;
            }
        }
    }

    validate();
}

number_format_token number_format_parser::parse_next_token()
{
    number_format_token token;

    auto to_lower = [](char c) {
        return static_cast<char>(std::tolower(static_cast<std::uint8_t>(c)));
    };

    if (format_string_.size() <= position_)
    {
        token.type = number_format_token::token_type::end;
        return token;
    }

    auto current_char = format_string_[position_++];

    switch (current_char)
    {
    case '[':
        if (position_ == format_string_.size())
        {
            throw xlnt::exception("missing ]");
        }

        if (format_string_[position_] == ']')
        {
            throw xlnt::exception("empty []");
        }

        do
        {
            token.string.push_back(format_string_[position_++]);
        } while (position_ < format_string_.size() && format_string_[position_] != ']');

        if (token.string[0] == '<' || token.string[0] == '>' || token.string[0] == '=')
        {
            token.type = number_format_token::token_type::condition;
        }
        else if (token.string[0] == '$')
        {
            token.type = number_format_token::token_type::locale;
        }
        else if (token.string.size() <= 2
            && ((token.string == "h" || token.string == "hh") || (token.string == "m" || token.string == "mm")
                || (token.string == "s" || token.string == "ss")))
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
            throw xlnt::exception("expected General");
        }

        token.type = number_format_token::token_type::number;
        token.string = "General";
        position_ += 6;

        break;

    case '_':
        token.type = number_format_token::token_type::space;
        token.string.push_back(format_string_[position_++]);

        break;

    case '*':
        token.type = number_format_token::token_type::fill;
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
            current_char = format_string_[position_++];
        } while (current_char == '0' || current_char == '#' || current_char == '?' || current_char == ',');

        --position_;

        if (current_char == '%')
        {
            token.string.push_back('%');
            ++position_;
        }

        break;

    case 'y':
    case 'Y':
    case 'm':
    case 'M':
    case 'd':
    case 'D':
    case 'h':
    case 'H':
    case 's':
    case 'S':
        token.type = number_format_token::token_type::datetime;
        token.string.push_back(to_lower(current_char));

        while (format_string_[position_] == current_char)
        {
            token.string.push_back(to_lower(current_char));
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
            throw xlnt::exception("expected AM/PM or A/P");
        }

        break;

    case '"': {
        token.type = number_format_token::token_type::text;
        auto start = position_;
        auto end = format_string_.find('"', position_);

        while (end != std::string::npos && format_string_[end - 1] == '\\')
        {
            token.string.append(format_string_.substr(start, end - start - 1));
            token.string.push_back('"');
            position_ = end + 1;
            start = position_;
            end = format_string_.find('"', position_);
        }

        if (end != start)
        {
            token.string.append(format_string_.substr(start, end - start));
        }

        position_ = end + 1;

        break;
    }

    case ';':
        token.type = number_format_token::token_type::end_section;
        break;

    case '(':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case ')':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case '-':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case '+':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case ':':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case ' ':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case '/':
        token.type = number_format_token::token_type::text;
        token.string.push_back(current_char);

        break;

    case '@':
        token.type = number_format_token::token_type::number;
        token.string.push_back(current_char);
        break;

    case 'E':
        token.type = number_format_token::token_type::number;
        token.string.push_back(current_char);
        current_char = format_string_[position_++];

        if (current_char == '+' || current_char == '-')
        {
            token.string.push_back(current_char);
            break;
        }

        break;

    default:
        throw xlnt::exception("unexpected character");
    }

    return token;
}

void number_format_parser::validate()
{
    if (codes_.size() > 4)
    {
        throw xlnt::exception("too many format codes");
    }

    if (codes_.size() > 2)
    {
        if (codes_[0].has_condition && codes_[1].has_condition && codes_[2].has_condition)
        {
            throw xlnt::exception("format should have a maximum of two codes with conditions");
        }
    }
}

format_placeholders number_format_parser::parse_placeholders(const std::string &placeholders_string)
{
    format_placeholders p;

    if (placeholders_string == "General")
    {
        p.type = format_placeholders::placeholders_type::general;
        return p;
    }
    else if (placeholders_string == "@")
    {
        p.type = format_placeholders::placeholders_type::text;
        return p;
    }
    else if (placeholders_string.front() == '.')
    {
        p.type = format_placeholders::placeholders_type::fractional_part;
    }
    else if (placeholders_string.front() == 'E')
    {
        p.type = placeholders_string[1] == '+' ? format_placeholders::placeholders_type::scientific_exponent_plus
                                               : format_placeholders::placeholders_type::scientific_exponent_minus;
        return p;
    }
    else
    {
        p.type = format_placeholders::placeholders_type::integer_part;
    }

    if (placeholders_string.back() == '%')
    {
        p.percentage = true;
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
        std::size_t i = placeholders_string.size() - 1;

        while (!comma_indices.empty() && i == comma_indices.back())
        {
            ++p.thousands_scale;
            --i;
            comma_indices.pop_back();
        }

        p.use_comma_separator = !comma_indices.empty();
    }

    return p;
}

format_color number_format_parser::color_from_string(const std::string &color)
{
    switch (color[0])
    {
    case 'C':
        if ((color == "Cyan") || (color == "CYAN"))
        {
            return format_color::cyan;
        }
        else if ((color.substr(0, 5) == "Color") || (color.substr(0, 5) == "COLOR"))
        {
            auto color_number = std::stoull(color.substr(5));

            if (color_number >= 1 && color_number <= 56)
            {
                return static_cast<format_color>(color_number);
            }
        }

        unhandled_case(true);
        break;

    case 'B':
        if ((color == "Black") || (color == "BLACK"))
        {
            return format_color::black;
        }
        else if ((color == "Blue") || (color == "BLUE"))
        {
            return format_color::blue;
        }

        unhandled_case(true);
        break;

    case 'G':
        if ((color == "Green") || (color == "GREEN"))
        {
            return format_color::green;
        }

        unhandled_case(true);
        break;

    case 'W':
        if ((color == "White") || (color == "WHITE"))
        {
            return format_color::white;
        }

        unhandled_case(true);
        break;

    case 'M':
        if ((color == "Magenta") || (color == "MAGENTA"))
        {
            return format_color::magenta;
        }

        unhandled_case(true);
        break;

    case 'Y':
        if ((color == "Yellow") || (color == "YELLOW"))
        {
            return format_color::yellow;
        }

        unhandled_case(true);
        break;

    case 'R':
        if ((color == "Red") || (color == "RED"))
        {
            return format_color::red;
        }

        unhandled_case(true);
        break;

    default:
        unhandled_case(true);
    }

    unhandled_case_error();
}

std::pair<format_locale, std::string> number_format_parser::locale_from_string(const std::string &locale_string)
{
    auto hyphen_index = locale_string.find('-');

    if (locale_string.empty() || locale_string.front() != '$' || hyphen_index == std::string::npos)
    {
        throw xlnt::exception("bad locale: " + locale_string);
    }

    std::pair<format_locale, std::string> result;

    if (hyphen_index > 1)
    {
        result.second = locale_string.substr(1, hyphen_index - 1);
    }

    auto country_code_string = locale_string.substr(hyphen_index + 1);

    if (country_code_string.empty())
    {
        throw xlnt::exception("bad locale: " + locale_string);
    }

    for (auto c : country_code_string)
    {
        if (!((c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') || (c >= '0' && c <= '9')))
        {
            throw xlnt::exception("bad locale: " + locale_string);
        }
    }

    auto country_code = std::stoi(country_code_string, nullptr, 16);
    country_code &= 0xFFFF;

    for (const auto &known_locale : known_locales())
    {
        if (known_locale.first == country_code)
        {
            result.first = static_cast<format_locale>(country_code);
            return result;
        }
    }

    throw xlnt::exception("unknown country code: " + country_code_string);
}

number_formatter::number_formatter(const std::string &format_string, xlnt::calendar calendar)
    : parser_(format_string), calendar_(calendar)
{
    parser_.parse();
    format_ = parser_.result();
}

std::string number_formatter::format_number(double number)
{
    if (format_[0].has_condition)
    {
        if (format_[0].condition.satisfied_by(number))
        {
            return format_number(format_[0], number);
        }

        if (format_.size() == 1)
        {
            return std::string(11, '#');
        }

        if (!format_[1].has_condition || format_[1].condition.satisfied_by(number))
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
            return format_number(format_[1], std::fabs(number));
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
            return format_number(format_[1], std::fabs(number));
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
        format_code temp;
        template_part temp_part;
        temp_part.type = template_part::template_type::general;
        temp_part.placeholders.type = format_placeholders::placeholders_type::general;
        temp.parts.push_back(temp_part);
        return format_text(temp, text);
    }

    return format_text(format_[3], text);
}

std::string number_formatter::fill_placeholders(const format_placeholders &p, double number)
{
    std::string result;

    if (p.type == format_placeholders::placeholders_type::general
        || p.type == format_placeholders::placeholders_type::text)
    {
        auto s = serialiser_.serialise_short(number);
        while (s.size() > 1 && s.back() == '0')
        {
            s.pop_back();
        }
        if (s.back() == '.')
        {
            s.pop_back();
        }
        return s;
    }

    if (p.percentage)
    {
        number *= 100;
    }

    if (p.thousands_scale > 0)
    {
        number /= std::pow(1000.0, p.thousands_scale);
    }

    auto integer_part = static_cast<int>(number);

    if (p.type == format_placeholders::placeholders_type::integer_only
        || p.type == format_placeholders::placeholders_type::integer_part
        || p.type == format_placeholders::placeholders_type::fraction_integer)
    {
        result = std::to_string(integer_part);

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

        if (p.percentage && p.type == format_placeholders::placeholders_type::integer_only)
        {
            result.push_back('%');
        }
    }
    else if (p.type == format_placeholders::placeholders_type::fractional_part)
    {
        auto fractional_part = number - integer_part;
        result = std::fabs(fractional_part) < std::numeric_limits<double>::min()
            ? std::string(".")
            : serialiser_.serialise_short(fractional_part).substr(1);

        while (result.back() == '0' || result.size() > (p.num_zeros + p.num_optionals + p.num_spaces + 1))
        {
            result.pop_back();
        }

        
        if (result.size() < p.num_zeros + 1)
        {
            result.resize(p.num_zeros + 1, '0');
        }

        if (result.size() < p.num_zeros + p.num_optionals + p.num_spaces + 1)
        {
            result.resize(p.num_zeros + p.num_optionals + p.num_spaces + 1, ' ');
        }

        if (p.percentage)
        {
            result.push_back('%');
        }
    }

    return result;
}

std::string number_formatter::fill_scientific_placeholders(const format_placeholders &integer_part,
    const format_placeholders &fractional_part, const format_placeholders &exponent_part, double number)
{
    std::size_t logarithm = 0;

    if (number != 0.0)
    {
        logarithm = static_cast<std::size_t>(std::log10(number));

        if (integer_part.num_zeros + integer_part.num_optionals > 1)
        {
            logarithm = integer_part.num_zeros + integer_part.num_optionals;
        }
    }

    number /= std::pow(10.0, logarithm);

    auto integer = static_cast<int>(number);
    auto fraction = number - integer;

    std::string integer_string = std::to_string(integer);

    if (number == 0.0)
    {
        integer_string = std::string(integer_part.num_zeros + integer_part.num_optionals, '0');
    }

    std::string fractional_string = serialiser_.serialise_short(fraction).substr(1, fractional_part.num_zeros + fractional_part.num_optionals + 1);
    std::string exponent_string = std::to_string(logarithm);

    while (exponent_string.size() < fractional_part.num_zeros)
    {
        exponent_string.insert(0, "0");
    }

    if (exponent_part.type == format_placeholders::placeholders_type::scientific_exponent_plus)
    {
        exponent_string.insert(0, "E+");
    }
    else
    {
        exponent_string.insert(0, "E");
    }

    return integer_string + fractional_string + exponent_string;
}

std::string number_formatter::fill_fraction_placeholders(const format_placeholders & /*numerator*/,
    const format_placeholders &denominator, double number, bool /*improper*/)
{
    auto fractional_part = number - static_cast<int>(number);
    auto original_fractional_part = fractional_part;
    fractional_part *= 10;

    while (std::abs(fractional_part - static_cast<int>(fractional_part)) > 0.000001
        && std::abs(fractional_part - static_cast<int>(fractional_part)) < 0.999999)
    {
        fractional_part *= 10;
    }

    fractional_part = static_cast<int>(fractional_part);
    auto denominator_digits = denominator.num_zeros + denominator.num_optionals + denominator.num_spaces;
    //    auto denominator_digits = static_cast<std::size_t>(std::ceil(std::log10(fractional_part)));

    auto lower = static_cast<int>(std::pow(10, denominator_digits - 1));
    auto upper = static_cast<int>(std::pow(10, denominator_digits));
    auto best_denominator = lower;
    auto best_difference = 1000.0;

    for (int i = lower; i < upper; ++i)
    {
        auto numerator_full = original_fractional_part * i;
        auto numerator_rounded = static_cast<int>(std::round(numerator_full));
        auto difference = std::fabs(original_fractional_part - (numerator_rounded / static_cast<double>(i)));

        if (difference < best_difference)
        {
            best_difference = difference;
            best_denominator = i;
        }
    }

    auto numerator_rounded = static_cast<int>(std::round(original_fractional_part * best_denominator));
    return std::to_string(numerator_rounded) + "/" + std::to_string(best_denominator);
}

std::string number_formatter::format_number(const format_code &format, double number)
{
    static const std::vector<std::string> month_names = std::vector<std::string>{"January", "February", "March",
        "April", "May", "June", "July", "August", "September", "October", "November", "December"};

    static const std::vector<std::string> day_names =
        std::vector<std::string>{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};

    std::string result;

    if (number < 0)
    {
        result.push_back('-');

        if (format.is_datetime)
        {
            return std::string(11, '#');
        }
    }

    number = std::fabs(number);

    xlnt::datetime dt(0, 1, 0);
    std::size_t hour = 0;

    if (format.is_datetime)
    {
        if (number != 0.0)
        {
            dt = xlnt::datetime::from_number(number, calendar_);
        }

        hour = static_cast<std::size_t>(dt.hour);

        if (format.twelve_hour)
        {
            hour %= 12;

            if (hour == 0)
            {
                hour = 12;
            }
        }
    }

    bool improper_fraction = true;
    std::size_t fill_index = 0;
    bool fill = false;
    std::string fill_character;

    for (std::size_t i = 0; i < format.parts.size(); ++i)
    {
        const auto &part = format.parts[i];

        switch (part.type)
        {
        case template_part::template_type::space: {
            result.push_back(' ');
            break;
        }

        case template_part::template_type::text: {
            result.append(part.string);
            break;
        }

        case template_part::template_type::fill: {
            fill = true;
            fill_index = result.size();
            fill_character = part.string;
            break;
        }

        case template_part::template_type::general: {
            if (part.placeholders.type == format_placeholders::placeholders_type::fractional_part
                && (format.is_datetime || format.is_timedelta))
            {
                auto digits = std::min(
                    static_cast<std::size_t>(6), part.placeholders.num_zeros + part.placeholders.num_optionals);
                auto denominator = static_cast<int>(std::pow(10.0, digits));
                auto fractional_seconds = dt.microsecond / 1.0E6 * denominator;
                fractional_seconds = std::round(fractional_seconds) / denominator;
                result.append(fill_placeholders(part.placeholders, fractional_seconds));
                break;
            }

            if (part.placeholders.type == format_placeholders::placeholders_type::fraction_integer)
            {
                improper_fraction = false;
            }

            if (part.placeholders.type == format_placeholders::placeholders_type::fraction_numerator)
            {
                i += 2;

                if (number == 0.0)
                {
                    result.pop_back();
                    break;
                }

                result.append(fill_fraction_placeholders(
                    part.placeholders, format.parts[i].placeholders, number, improper_fraction));
            }
            else if (part.placeholders.scientific
                && part.placeholders.type == format_placeholders::placeholders_type::integer_part)
            {
                auto integer_part = part.placeholders;
                ++i;
                auto fractional_part = format.parts[i++].placeholders;
                auto exponent_part = format.parts[i++].placeholders;
                result.append(fill_scientific_placeholders(integer_part, fractional_part, exponent_part, number));
            }
            else
            {
                result.append(fill_placeholders(part.placeholders, number));
            }

            break;
        }

        case template_part::template_type::day_number: {
            result.append(std::to_string(dt.day));
            break;
        }

        case template_part::template_type::day_number_leading_zero: {
            if (dt.day < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.day));
            break;
        }

        case template_part::template_type::month_abbreviation: {
            result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1).substr(0, 3));
            break;
        }

        case template_part::template_type::month_name: {
            result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1));
            break;
        }

        case template_part::template_type::month_number: {
            result.append(std::to_string(dt.month));
            break;
        }

        case template_part::template_type::month_number_leading_zero: {
            if (dt.month < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.month));
            break;
        }

        case template_part::template_type::year_short: {
            if (dt.year % 1000 < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.year % 1000));
            break;
        }

        case template_part::template_type::year_long: {
            result.append(std::to_string(dt.year));
            break;
        }

        case template_part::template_type::hour: {
            result.append(std::to_string(hour));
            break;
        }

        case template_part::template_type::hour_leading_zero: {
            if (hour < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(hour));
            break;
        }

        case template_part::template_type::minute: {
            result.append(std::to_string(dt.minute));
            break;
        }

        case template_part::template_type::minute_leading_zero: {
            if (dt.minute < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.minute));
            break;
        }

        case template_part::template_type::second: {
            result.append(std::to_string(dt.second + (dt.microsecond > 500000 ? 1 : 0)));
            break;
        }

        case template_part::template_type::second_fractional: {
            result.append(std::to_string(dt.second));
            break;
        }

        case template_part::template_type::second_leading_zero: {
            if ((dt.second + (dt.microsecond > 500000 ? 1 : 0)) < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.second + (dt.microsecond > 500000 ? 1 : 0)));
            break;
        }

        case template_part::template_type::second_leading_zero_fractional: {
            if (dt.second < 10)
            {
                result.push_back('0');
            }

            result.append(std::to_string(dt.second));
            break;
        }

        case template_part::template_type::am_pm: {
            if (dt.hour < 12)
            {
                result.append("AM");
            }
            else
            {
                result.append("PM");
            }

            break;
        }

        case template_part::template_type::a_p: {
            if (dt.hour < 12)
            {
                result.append("A");
            }
            else
            {
                result.append("P");
            }

            break;
        }

        case template_part::template_type::elapsed_hours: {
            result.append(std::to_string(24 * static_cast<int>(number) + dt.hour));
            break;
        }

        case template_part::template_type::elapsed_minutes: {
            result.append(std::to_string(24 * 60 * static_cast<int>(number)
                + (60 * dt.hour) + dt.minute));
            break;
        }

        case template_part::template_type::elapsed_seconds: {
            result.append(std::to_string(24 * 60 * 60 * static_cast<int>(number)
                + (60 * 60 * dt.hour) + (60 * dt.minute) + dt.second));
            break;
        }

        case template_part::template_type::month_letter: {
            result.append(month_names.at(static_cast<std::size_t>(dt.month) - 1).substr(0, 1));
            break;
        }

        case template_part::template_type::day_abbreviation: {
            result.append(day_names.at(static_cast<std::size_t>(dt.weekday())).substr(0, 3));
            break;
        }

        case template_part::template_type::day_name: {
            result.append(day_names.at(static_cast<std::size_t>(dt.weekday())));
            break;
        }
        }
    }

    const std::size_t width = 11;

    if (fill && result.size() < width)
    {
        auto remaining = width - result.size();

        std::string fill_string(remaining, fill_character.front());
        // TODO: A UTF-8 character could be multiple bytes

        result = result.substr(0, fill_index) + fill_string + result.substr(fill_index);
    }

    return result;
}

std::string number_formatter::format_text(const format_code &format, const std::string &text)
{
    std::string result;
    bool any_text_part = false;

    for (const auto &part : format.parts)
    {
        if (part.type == template_part::template_type::text)
        {
            result.append(part.string);
            any_text_part = true;
        }
        else if (part.type == template_part::template_type::general)
        {
            if (part.placeholders.type == format_placeholders::placeholders_type::general
                || part.placeholders.type == format_placeholders::placeholders_type::text)
            {
                result.append(text);
                any_text_part = true;
            }
        }
    }

    if (!format.parts.empty() && !any_text_part)
    {
        return text;
    }

    return result;
}

} // namespace detail
} // namespace xlnt
