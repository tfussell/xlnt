#include <algorithm>

#include <xlnt/styles/number_format.hpp>

namespace xlnt {

const std::unordered_map<number_format::format, std::string, number_format::format_hash> &number_format::format_strings()
{
    static const std::unordered_map<number_format::format, std::string, number_format::format_hash> strings =
    {
        {format::general, "General"},
        {format::text, "@"},
        {format::number, "0"},
        {format::number_00, "0.00"},
        {format::number_comma_separated1, "#,##0.00"},
        {format::number_comma_separated2, "#,##0.00_-"},
        {format::percentage, "0%"},
        {format::percentage_00, "0.00%"},
        {format::date_yyyymmdd2, "yyyy-mm-dd"},
        {format::date_yyyymmdd, "yy-mm-dd"},
        {format::date_ddmmyyyy, "dd/mm/yy"},
        {format::date_dmyslash, "d/m/y"},
        {format::date_dmyminus, "d-m-y"},
        {format::date_dmminus, "d-m"},
        {format::date_myminus, "m-y"},
        {format::date_xlsx14, "mm-dd-yy"},
        {format::date_xlsx15, "d-mmm-yy"},
        {format::date_xlsx16, "d-mmm"},
        {format::date_xlsx17, "mmm-yy"},
        {format::date_xlsx22, "m/d/yy h:mm"},
        {format::date_datetime, "d/m/y h:mm"},
        {format::date_time1, "h:mm AM/PM"},
        {format::date_time2, "h:mm:ss AM/PM"},
        {format::date_time3, "h:mm"},
        {format::date_time4, "h:mm:ss"},
        {format::date_time5, "mm:ss"},
        {format::date_time6, "h:mm:ss"},
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

const std::unordered_map<std::string, int> &number_format::reversed_builtin_formats()
{
    static const std::unordered_map<std::string, int> formats =
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
        {"# \?\?/??", 13}, //escape trigraph
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

        {"_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)", 41},
        {"_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)", 42},
        {"_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)", 43},

        {"_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)", 44},
        {"mm:ss", 45},
        {"[h]:mm:ss", 46},
        {"mmss.0", 47},
        {"##0.0E+0", 48},
        {"@", 49}
    };

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

} // namespace xlnt
