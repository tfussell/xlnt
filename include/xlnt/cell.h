#pragma once

#include <memory>
#include <string>
#include <unordered_map>

#include "types.h"

namespace xlnt {
    
class cell_reference;
class date;
class datetime;
class time;
class worksheet;

namespace detail {    
struct cell_impl;
} // namespace detail
    
typedef std::string comment;
    
/// <summary>
/// Alignment options for use in styles.
/// </summary>
struct alignment
{
    enum class horizontal_alignment
    {
        general,
        left,
        right,
        center,
        center_continuous,
        justify
    };
    
    enum class vertical_alignment
    {
        bottom,
        top,
        center,
        justify
    };
    
    horizontal_alignment horizontal = horizontal_alignment::general;
    vertical_alignment vertical = vertical_alignment::bottom;
    int text_rotation = 0;
    bool wrap_text = false;
    bool shrink_to_fit = false;
    int indent = 0;
};

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
    
    static const std::unordered_map<int, std::string> builtin_formats;
    
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

struct color
{
    static const color black;
    static const color white;
    static const color red;
    static const color darkred;
    static const color blue;
    static const color darkblue;
    static const color green;
    static const color darkgreen;
    static const color yellow;
    static const color darkyellow;
    
    color(int index)
    {
        this->index = index;
    }
    
    int index;
};

class font
{
    enum class underline
    {
        none,
        double_,
        double_accounting,
        single,
        single_accounting
    };
    
    /*    std::string name = "Calibri";
     int size = 11;
     bool bold = false;
     bool italic = false;
     bool superscript = false;
     bool subscript = false;
     underline underline = underline::none;
     bool strikethrough = false;
     color color = color::black;*/
};

class fill
{
public:
    enum class type
    {
        none,
        solid,
        gradient_linear,
        gradient_path,
        pattern_darkdown,
        pattern_darkgray,
        pattern_darkgrid,
        pattern_darkhorizontal,
        pattern_darktrellis,
        pattern_darkup,
        pattern_darkvertical,
        pattern_gray0625,
        pattern_gray125,
        pattern_lightdown,
        pattern_lightgray,
        pattern_lightgrid,
        pattern_lighthorizontal,
        pattern_lighttrellis,
        pattern_lightup,
        pattern_lightvertical,
        pattern_mediumgray,
    };
    
    type type_ = type::none;
    int rotation = 0;
    color start_color = color::white;
    color end_color = color::black;
};

class borders
{
    struct border
    {
        enum class style
        {
            none,
            dashdot,
            dashdotdot,
            dashed,
            dotted,
            double_,
            hair,
            medium,
            mediumdashdot,
            mediumdashdotdot,
            mediumdashed,
            slantdashdot,
            thick,
            thin
        };
        
        style style_ = style::none;
        color color_ = color::black;
    };
    
    enum class diagonal_direction
    {
        none,
        up,
        down,
        both
    };
    
    border left;
    border right;
    border top;
    border bottom;
    border diagonal;
    //    diagonal_direction diagonal_direction = diagonal_direction::none;
    border all_borders;
    border outline;
    border inside;
    border vertical;
    border horizontal;
};

class protection
{
public:
    enum class type
    {
        inherit,
        protected_,
        unprotected
    };
    
    type locked;
    type hidden;
};

class style
{
public:
    style(bool static_ = false) : static_(static_) {}
    
    style copy() const;
    
    font get_font() const;
    void set_font(font font);
    
    fill get_fill() const;
    void set_fill(fill fill);
    
    borders get_borders() const;
    void set_borders(borders borders);
    
    alignment get_alignment() const;
    void set_alignment(alignment alignment);
    
    number_format &get_number_format() { return number_format_; }
    const number_format &get_number_format() const { return number_format_; }
    void set_number_format(number_format number_format);
    
    protection get_protection() const;
    void set_protection(protection protection);
    
private:
    style(const style &rhs);
    
    bool static_ = false;
    font font_;
    fill fill_;
    borders borders_;
    alignment alignment_;
    number_format number_format_;
    protection protection_;
};

/// <summary>
/// Describes cell associated properties.
/// </summary>
/// <remarks>
/// Properties of interest include style, type, value, and address.
/// The Cell class is required to know its value and type, display options,
/// and any other features of an Excel cell.Utilities for referencing
/// cells using Excel's 'A1' column/row nomenclature are also provided.
/// </remarks>
class cell
{
public:    
    enum class type
    {
        null,
	numeric,
	string,
	formula,
	boolean,
	error
    };

    static const std::unordered_map<std::string, int> ErrorCodes;
    
    static std::string check_string(const std::string &value);
    static std::string check_numeric(const std::string &value);
    static std::string check_error(const std::string &value);
    
    cell();
    cell(worksheet ws, const cell_reference &reference, const std::string &initial_value = std::string());
    
    std::string get_value() const;
    
    std::string get_column() const;
    row_t get_row() const;
    
    std::string to_string() const;
    
    void set_explicit_value(const std::string &value, type data_type);
    void set_explicit_value(int value, type data_type);
    void set_explicit_value(double value, type data_type);
    void set_explicit_value(const date &value, type data_type);
    void set_explicit_value(const time &value, type data_type);
    void set_explicit_value(const datetime &value, type data_type);
    type data_type_for_value(const std::string &value);
    
    bool is_merged() const;
    void set_merged(bool merged);
    
    std::string get_hyperlink() const;
    void set_hyperlink(const std::string &value);
    
    void set_number_format(const std::string &format_code);
    
    bool has_style() const;
    
    style &get_style();
    const style &get_style() const;
    
    type get_data_type() const;
    
    cell_reference get_reference() const;
    
    bool is_date() const;
    
    //std::pair<int, int> get_anchor() const;
    
    comment get_comment() const;
    void set_comment(comment comment);

    std::string get_formula() const;
    void set_formula(const std::string &formula);

    std::string get_error() const;
    void set_error(const std::string &error);

    void set_null();
    
    cell &operator=(const cell &rhs);
    cell &operator=(bool value);
    cell &operator=(int value);
    cell &operator=(double value);
    cell &operator=(const std::string &value);
    cell &operator=(const char *value);
    cell &operator=(const date &value);
    cell &operator=(const time &value);
    cell &operator=(const datetime &value);
    
    bool operator==(const cell &comparand) const { return d_ == comparand.d_; }
    bool operator==(std::nullptr_t) const;
    bool operator==(bool comparand) const;
    bool operator==(int comparand) const;
    bool operator==(double comparand) const;
    bool operator==(const std::string &comparand) const;
    bool operator==(const char *comparand) const;
    bool operator==(const date &comparand) const;
    bool operator==(const time &comparand) const;
    bool operator==(const datetime &comparand) const;

    friend bool operator==(std::nullptr_t, const cell &cell);
    friend bool operator==(bool comparand, const cell &cell);
    friend bool operator==(int comparand, const cell &cell);
    friend bool operator==(double comparand, const cell &cell);
    friend bool operator==(const std::string &comparand, const cell &cell);
    friend bool operator==(const char *comparand, const cell &cell);
    friend bool operator==(const date &comparand, const cell &cell);
    friend bool operator==(const time &comparand, const cell &cell);
    friend bool operator==(const datetime &comparand, const cell &cell);
    
private:
    friend class worksheet;
    cell(detail::cell_impl *d);
    detail::cell_impl *d_;
};

inline std::ostream &operator<<(std::ostream &stream, const xlnt::cell &cell)
{
    return stream << cell.to_string();
}
    
} // namespace xlnt
