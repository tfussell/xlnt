#include <algorithm>
#include <regex>

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

std::size_t number_format::hash() const
{
    std::size_t seed = id_;
    hash_combine(seed, format_string_);

    return seed;
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

} // namespace xlnt
