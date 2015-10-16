// Copyright (c) 2015 Thomas Fussell
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
#pragma once

#include <string>
#include <unordered_map>
#include <utility>

namespace xlnt {

class number_format
{
public:
    enum class format
    {
        general,
        text,
        number,
        number_00,
        number_comma_separated1,
        number_comma_separated2,
        percentage,
        percentage_00,
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
        currency_eur_simple,
        unknown
    };

    struct format_hash
    {
        std::size_t operator()(format f) const
        {
	    return std::hash<int>()((int)f);
	}
    };
    
    static const std::unordered_map<format, std::string, format_hash> &format_strings();    
    static const std::unordered_map<int, std::string> &builtin_formats();
    static const std::unordered_map<std::string, int> &reversed_builtin_formats();    

    static std::string builtin_format_code(int index);
    static format lookup_format(int code);
    
    static bool is_date_format(const std::string &format);
    static bool is_builtin(const std::string &format);
    
    number_format() : format_code_(format::general), format_index_(0) {}
    number_format(format code) : format_code_(code), format_index_(reversed_builtin_formats().at(format_strings().at(code))) {}
    
    format get_format_code() const { return format_code_; }
    
    void set_format_code(format format_code, int index = -1)
    {
        format_code_ = format_code;
        
        if(format_code_ != format::unknown)
        {
            set_format_code_string(format_strings().at(format_code), index);
        }
    }
    
    void set_format_code_string(const std::string &format_code, int index)
    {
        custom_format_code_ = format_code;
        format_index_ = index;
        
        const auto &reversed = reversed_builtin_formats();
        auto match = reversed.find(format_code);
        
        format_code_ = format::unknown;
        
        if(match != reversed.end())
        {
            format_index_ = match->second;
            
            for(const auto &p : format_strings())
            {
                if(p.second == format_code)
                {
                    format_code_ = p.first;
                    break;
                }
            }
        }
        
    }
    std::string get_format_code_string() const;

private:
    std::string custom_format_code_ = "";
    format format_code_ = format::general;
    int format_index_ = 0;
};

} // namespace xlnt
