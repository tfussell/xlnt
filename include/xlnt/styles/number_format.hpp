// Copyright (c) 2014 Thomas Fussell
// Copyright (c) 2010-2014 openpyxl
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

namespace xlnt {

class number_format
{
public:
    enum class format
    {
        general,
        text,
        number,
        number00,
        number_comma_separated1,
        number_comma_separated2,
        percentage,
        percentage00,
        date_yyyymmdd2,
        date_yyyymmdd,
        date_ddmmyyyy,
        date_dmyslash,
        date_dmyminus,
        date_dmminus,
        date_myminus,
        date_xlsx14,
        date_xlsx15,
        date_xlsx16,
        date_xlsx17,
        date_xlsx22,
        date_datetime,
        date_time1,
        date_time2,
        date_time3,
        date_time4,
        date_time5,
        date_time6,
        date_time7,
        date_time8,
        date_timedelta,
        date_yyyymmddslash,
        currency_usd_simple,
        currency_usd,
        currency_eur_simple
    };
    
    std::unordered_map<format, std::string> format_strings =
    {
        {format::general, "General"},
        {format::text, "@"},
    {format::number, "0"},
    {format::number_00, "0.00"},
    {format::NUMBER_COMMA_SEPARATED1, "#,##0.00"},
    {format::NUMBER_COMMA_SEPARATED2, "#,##0.00_-"},
    {format::PERCENTAGE, "0%"},
    {format::PERCENTAGE_00, "0.00%"},
    {format::DATE_YYYYMMDD2, "yyyy-mm-dd"},
    {format::DATE_YYYYMMDD, "yy-mm-dd"},
    {format::DATE_DDMMYYYY, "dd/mm/yy"},
    {format::DATE_DMYSLASH, "d/m/y"},
    {format::DATE_DMYMINUS, "d-m-y"},
    {format::DATE_DMMINUS, "d-m"},
    {format::DATE_MYMINUS, "m-y"},
    {format::DATE_XLSX14, "mm-dd-yy"},
    {format::DATE_XLSX15, "d-mmm-yy"},
    {format::DATE_XLSX16, "d-mmm"},
    {format::DATE_XLSX17, "mmm-yy"},
    {format::DATE_XLSX22, "m/d/yy h:mm"},
    {format::DATE_DATETIME, "d/m/y h:mm"},
    {format::DATE_TIME1, "h:mm AM/PM"},
    {format::DATE_TIME2, "h:mm:ss AM/PM"},
    {format::DATE_TIME3, "h:mm"},
    {format::DATE_TIME4, "h:mm:ss"},
    {format::DATE_TIME5, "mm:ss"},
    {format::DATE_TIME6, "h:mm:ss"},
    {format::DATE_TIME7, "i:s.S"},
    {format::DATE_TIME8, "h:mm:ss@"},
    {format::DATE_TIMEDELTA, "[hh]:mm:ss"},
    {format::DATE_YYYYMMDDSLASH, "yy/mm/dd@"},
    {format::CURRENCY_USD_SIMPLE, "\"$\"#,##0.00_-"},
    {format::CURRENCY_USD, "$#,##0_-"},
        {format::CURRENCY_EUR_SIMPLE, "[$EUR ]#,##0.00_-"}
    };
    
    static const std::unordered_map<int, std::string> builtin_formats =
    {
        {0, "General"},
        {1, "0"},
        {2, "0.00"},
        {3, "#,##0"},
        {4, "#,##0.00"},
        {5, "\"$\"#,##0_);(\"$\"#,##0)"},
        {6, "\"$\"#,##0_);[Red](\"$\"#,##0)"},
        {7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"},
        {8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"},
        {9, "0%"},
        {10, "0.00%"},
        {11, "0.00E+00"},
        {12, "# ?/?"},
        {13, "# ??/??"},
        {14, "mm-dd-yy"},
        {15, "d-mmm-yy"},
        {16, "d-mmm"},
        {17, "mmm-yy"},
        {18, "h:mm AM/PM"},
        {19, "h:mm:ss AM/PM"},
        {20, "h:mm"},
        {21, "h:mm:ss"},
        {22, "m/d/yy h:mm"},

        {37, "#,##0_);(#,##0)"},
        {38, "#,##0_);[Red](#,##0)"},
        {39, "#,##0.00_);(#,##0.00)"},
        {40, "#,##0.00_);[Red](#,##0.00)"},
        
        {41, "_(* #,##0_);_(* \(#,##0\\);_(* \"-\"_);_(@_)"},
        {42, "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)"},
        {43, "_(* #,##0.00_);_(* \(#,##0.00\\);_(* \"-\"??_);_(@_)"},
        
        {44, "_(\"$\"* #,##0.00_)_(\"$\"* \(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)"},
        {45, "mm:ss"},
        {46, "[h]:mm:ss"},
        {47, "mmss.0"},
        {48, "##0.0E+0"},
        {49, "@"}
        };

    static const std::unordered_map<std::string, int> reversed_builtin_formats =
    {
        {"General", 0},
        {"0", 1},
        {"0.00", 2},
        {"#,##0", 3},
        {"#,##0.00", 4},
        {"\"$\"#,##0_);(\"$\"#,##0)", 5},
        {"\"$\"#,##0_);[Red](\"$\"#,##0)", 6},
        {"\"$\"#,##0.00_);(\"$\"#,##0.00)", 7},
        {"\"$\"#,##0.00_);[Red](\"$\"#,##0.00)", 8},
        {"0%", 9},
        {"0.00%", 10},
        {"0.00E+00", 11},
        {"# ?/?", 12},
        {"# ??/??", 13},
        {"mm-dd-yy", 14},
        {"d-mmm-yy", 15},
        {"d-mmm", 16},
        {"mmm-yy", 17},
        {"h:mm AM/PM", 18},
        {"h:mm:ss AM/PM", 19},
        {"h:mm", 20},
        {"h:mm:ss", 21},
        {"m/d/yy h:mm", 22},
        
        {"#,##0_);(#,##0)", 37},
        {"#,##0_);[Red](#,##0)", 38},
        {"#,##0.00_);(#,##0.00)", 39},
        {"#,##0.00_);[Red](#,##0.00)", 40},
        
        {"_(* #,##0_);_(* \(#,##0\\);_(* \"-\"_);_(@_)", 41},
        {"_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)", 42},
        {"_(* #,##0.00_);_(* \(#,##0.00\\);_(* \"-\"??_);_(@_)", 43},
        
        {"_(\"$\"* #,##0.00_)_(\"$\"* \(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)", 44},
        {"mm:ss", 45},
        {"[h]:mm:ss", 46},
        {"mmss.0", 47},
        {"##0.0E+0", 48},
        {"@", 49}
    };
    
    static std::string builtin_format_code(int index);
    
    static bool is_date_format(const std::string &format);
    static bool is_builtin(const std::string &format);
    
    format get_format_code() const { return format_code_; }
    void set_format_code(format format_code) { format_code_ = format_code; }
    void set_format_code(const std::string &format_code) { custom_format_code_ = format_code; }
    
private:
    std::string custom_format_code_ = "";
    format format_code_ = format::general;
    int format_index_ = 0;
};

} // namespace xlnt
