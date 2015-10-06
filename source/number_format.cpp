#include <algorithm>

#include <xlnt/styles/number_format.hpp>

namespace xlnt {
    
const std::unordered_map<int, std::string> &number_format::builtin_formats()
{
    static const std::unordered_map<int, std::string> formats =
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
        {13, "# \?\?/??"}, //escape trigraph
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

        {41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)"},
        {42, "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)"},
        {43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"},

        {44, "_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)"},
        {45, "mm:ss"},
        {46, "[h]:mm:ss"},
        {47, "mmss.0"},
        {48, "##0.0E+0"},
        {49, "@"}

        //EXCEL differs from the standard in the following:
        //{14, "m/d/yyyy"},
        //{22, "m/d/yyyy h:mm"},
        //{37, "#,##0_);(#,##0)"},
        //{38, "#,##0_);[Red]"},
        //{39, "#,##0.00_);(#,##0.00)"},
        //{40, "#,##0.00_);[Red]"},
        //{47, "mm:ss.0"},
        //{55, "yyyy/mm/dd"}
    };

    return formats;
}

const std::unordered_map<number_format::format, std::string, number_format::format_hash> &number_format::format_strings()
{
    static const std::unordered_map<number_format::format, std::string, number_format::format_hash> strings =
    {
        {format::general, builtin_formats().at(0)},
        {format::text, builtin_formats().at(49)},
        {format::number, builtin_formats().at(1)},
        {format::number_00, builtin_formats().at(2)},
        {format::number_comma_separated1, builtin_formats().at(4)},
        {format::number_comma_separated2, "#,##0.00_-"},
        {format::percentage, builtin_formats().at(9)},
        {format::percentage_00, builtin_formats().at(10)},
        {format::date_yyyymmdd2, "yyyy-mm-dd"},
        {format::date_yyyymmdd, "yy-mm-dd"},
        {format::date_ddmmyyyy, "dd/mm/yy"},
        {format::date_dmyslash, "d/m/y"},
        {format::date_dmyminus, "d-m-y"},
        {format::date_dmminus, "d-m"},
        {format::date_myminus, "m-y"},
        {format::date_xlsx14, builtin_formats().at(14)},
        {format::date_xlsx15, builtin_formats().at(15)},
        {format::date_xlsx16, builtin_formats().at(16)},
        {format::date_xlsx17, builtin_formats().at(17)},
        {format::date_xlsx22, builtin_formats().at(22)},
        {format::date_datetime, "yyyy-mm-dd h:mm:ss"},
        {format::date_time1, builtin_formats().at(18)},
        {format::date_time2, builtin_formats().at(19)},
        {format::date_time3, builtin_formats().at(20)},
        {format::date_time4, builtin_formats().at(21)},
        {format::date_time5, builtin_formats().at(45)},
        {format::date_time6, builtin_formats().at(21)},
        {format::date_time7, "i:s.S"},
        {format::date_time8, "h:mm:ss@"},
        {format::date_timedelta, "[hh]:mm:ss"},
        {format::date_yyyymmddslash, "yy/mm/dd@"},
        {format::currency_usd_simple, "\"$\"#,##0.00_-"},
        {format::currency_usd, "$#,##0_-"},
        {format::currency_eur_simple, "[$EUR ]#,##0.00_-"}
    };

    return strings;
}

const std::unordered_map<std::string, int> &number_format::reversed_builtin_formats()
{
    static std::unordered_map<std::string, int> formats;

    for(auto format_pair : builtin_formats())
    {
        formats[format_pair.second] = format_pair.first;
    }
    
    return formats;
}
    
number_format::format number_format::lookup_format(int code)
{
    if(builtin_formats().find(code) == builtin_formats().end())
    {
        return format::unknown;
    }
    
    auto format_string = builtin_formats().at(code);
    auto match = std::find_if(format_strings().begin(), format_strings().end(), [&](const std::pair<format, std::string> &p) { return p.second == format_string; });
    
    if(match == format_strings().end())
    {
        return format::unknown;
    }
    
    return match->first;
}

std::string number_format::get_format_code_string()
{
    if(format_code_ == format::unknown)
    {
        return custom_format_code_;
    }

    return format_strings().at(format_code_);
}

} // namespace xlnt
