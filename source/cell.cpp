#include <algorithm>
#include <locale>
#include <sstream>

#include <xlnt/cell/cell.hpp>
#include <xlnt/cell/cell_reference.hpp>
#include <xlnt/cell/comment.hpp>
#include <xlnt/cell/value.hpp>
#include <xlnt/common/datetime.hpp>
#include <xlnt/common/relationship.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/common/exceptions.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/document_properties.hpp>

#include "detail/cell_impl.hpp"

namespace {

std::vector<std::string> split_string(const std::string &string, char delim = ' ')
{
    std::stringstream ss(string);
    std::string part;
    std::vector<std::string> parts;
    while(std::getline(ss, part, delim))
    {
        parts.push_back(part);
    }
    return parts;
}

xlnt::value::type data_type_for_value(const std::string &value)
{
    if(value.empty())
    {
        return xlnt::value::type::null;
    }

    if(value[0] == '0')
    {
        if(value.length() > 1)
        {
            if(value[1] == '.' || (value.length() > 2 && (value[1] == 'e' || value[1] == 'E')))
            {
                auto first_non_number = std::find_if(value.begin() + 2, value.end(),
                    [](char c) { return !std::isdigit(c, std::locale::classic()); });
                if(first_non_number == value.end())
                {
                    return xlnt::value::type::numeric;
                }
            }
            auto split = split_string(value, ':');
            if(split.size() == 2 || split.size() == 3)
            {
                for(auto part : split)
                {
                    try
                    {
                        std::stoi(part);
                    }
                    catch(std::invalid_argument)
                    {
                        return xlnt::value::type::string;
                    }
                }
                return xlnt::value::type::numeric;
            }
            else
            {
                return xlnt::value::type::string;
            }
        }
        return xlnt::value::type::numeric;
    }
    else if(value[0] == '#')
    {
        return xlnt::value::type::error;
    }
    else
    {
        char *p;
        strtod(value.c_str(), &p);
        if(*p != 0)
        {
            static const std::vector<std::string> possible_booleans = {"TRUE", "true", "FALSE", "false"};
            if(std::find(possible_booleans.begin(), possible_booleans.end(), value) != possible_booleans.end())
            {
                return xlnt::value::type::boolean;
            }
            if(value.back() == '%')
            {
                strtod(value.substr(0, value.length() - 1).c_str(), &p);
                if(*p == 0)
                {
                    return xlnt::value::type::numeric;
                }
            }
            return xlnt::value::type::string;
        }
        else
        {
            return xlnt::value::type::numeric;
        }
    }
}

} // namespace

namespace xlnt {
    
const xlnt::color xlnt::color::black(0);
const xlnt::color xlnt::color::white(1);
    
const std::unordered_map<std::string, int> cell::ErrorCodes =
{
    {"#NULL!", 0},
    {"#DIV/0!", 1},
    {"#VALUE!", 2},
    {"#REF!", 3},
    {"#NAME?", 4},
    {"#NUM!", 5},
    {"#N/A!", 6}
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
    
cell::cell(worksheet worksheet, const cell_reference &reference, const value &initial_value) : cell(worksheet, reference)
{
    set_value(initial_value);
}

bool cell::garbage_collectible() const
{
    return !(get_value().get_type() != value::type::null || is_merged() || has_comment() || has_formula());
}

value &cell::get_value()
{
    return d_->value_;
}

const value &cell::get_value() const
{
    return d_->value_;
}

void cell::set_value(const value &v)
{
    d_->value_ = v;
}

// return s after checking encoding, size, and illegal characters
std::string check_string(std::string s)
{
	if (s.size() == 0)
	{
		return s;
	}

	// check encoding?

	if (s.size() > 32767)
	{
		s = s.substr(0, 32767); // max string length in Excel
	}

	for (unsigned char c : s)
	{
		if (c <= 8 || c == 11 || c == 12 || (c >= 14 && c <= 31))
		{
			throw "illegal characters";
		}
	}

	return s;
}

std::pair<bool, long double> cast_numeric(const std::string &s)
{
	const char *str = s.c_str();
	char *str_end = nullptr;
	auto result = std::strtold(str, &str_end);
	if (str_end != str + s.size()) return{ false, 0 };
	return{ true, result };
}

std::pair<bool, long double> cast_percentage(const std::string &s)
{
	if (s.back() == '%')
	{
		auto number = cast_numeric(s.substr(0, s.size() - 1));

		if (number.first)
		{
			return{ true, number.second / 100 };
		}
	}

	return { false, 0 };
}

std::pair<bool, time> cast_time(const std::string &s)
{
	time result;

	try
	{
		auto last_colon = s.find_last_of(':');
		if (last_colon == std::string::npos) return { false, result };
		double seconds = std::stod(s.substr(last_colon + 1));
		result.second = static_cast<int>(seconds);
		result.microsecond = static_cast<int>((seconds - static_cast<double>(result.second)) * 1e6);

		auto first_colon = s.find_first_of(':');

		if (first_colon == last_colon)
		{
			auto decimal_pos = s.find('.');
			if (decimal_pos != std::string::npos)
			{
				result.minute = std::stoi(s.substr(0, first_colon));
			}
			else
			{
				result.hour = std::stoi(s.substr(0, first_colon));
				result.minute = result.second;
				result.second = 0;
			}
		}
		else
		{
			result.hour = std::stoi(s.substr(0, first_colon));
			result.minute = std::stoi(s.substr(first_colon + 1, last_colon - first_colon - 1));
		}
	}
	catch (std::invalid_argument)
	{
		return{ false, result };
	}

	return { true, result };
}

void cell::set_value(const std::string &s)
{
	d_->is_date_ = false;
	auto temp = check_string(s);
	d_->value_ = value(temp);

	if (temp.size() > 1 && temp.front() == '=')
	{
		d_->formula_ = temp;
	}

    if(get_parent().get_parent().get_guess_types())
    {
		auto percentage = cast_percentage(s);
		if (percentage.first)
		{
			d_->value_ = value(percentage.second);
			get_style().get_number_format().set_format_code(xlnt::number_format::format::percentage);
		}
		else
		{
			auto time = cast_time(s);
			if (time.first)
			{
				set_value(time.second);
			}

			else
			{
				auto numeric = cast_numeric(s);
				if (numeric.first)
				{
					d_->value_ = value(numeric.second);
				}
			}
		}
    }
}

void cell::set_value(const char *s)
{
    set_value(std::string(s));
}

void cell::set_value(bool b)
{
    d_->value_ = value(b);
}

void cell::set_value(std::int8_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::int16_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::int32_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::int64_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::uint8_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::uint16_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::uint32_t i)
{
	d_->value_ = value(i);
}

void cell::set_value(std::uint64_t i)
{
	d_->value_ = value(i);
}

#ifdef _WIN32
void cell::set_value(unsigned long i)
{
	d_->value_ = value(i);
}
#endif

#ifdef __linux__
void cell::set_value(long long i)
{
  d_->value_ = value(i);
}

void cell::set_value(unsigned long long i)
{
  d_->value_ = value(i);
}
#endif

void cell::set_value(double d)
{
    d_->value_ = value(d);
}

void cell::set_value(const date &d)
{
    d_->is_date_ = true;
    auto date_format_code = xlnt::number_format::lookup_format(14);
    auto number_format = xlnt::number_format(date_format_code);
    get_style().set_number_format(number_format);
    auto base_date = get_parent().get_parent().get_properties().excel_base_date;
    set_value(d.to_number(base_date));
}

void cell::set_value(const datetime &d)
{
    d_->is_date_ = true;
    auto date_format_code = xlnt::number_format::lookup_format(22);
    auto number_format = xlnt::number_format(date_format_code);
    get_style().set_number_format(number_format);
    auto base_date = get_parent().get_parent().get_properties().excel_base_date;
    set_value(d.to_number(base_date));
}

void cell::set_value(const time &t)
{
    d_->is_date_ = true;
    set_value(t.to_number());
}

void cell::set_value(const timedelta &t)
{
    d_->is_date_ = true;
    set_value(t.to_number());
}

bool cell::has_style() const
{
    return d_->style_ != nullptr;
}

row_t cell::get_row() const
{
    return d_->row_ + 1;
}

std::string cell::get_column() const
{
    return cell_reference::column_string_from_index(d_->column_ + 1);
}

void cell::set_merged(bool merged)
{
    d_->merged = merged;
}

bool cell::is_merged() const
{
    return d_->merged;
}

bool cell::is_date() const
{
    return d_->is_date_ || (d_->style_ != nullptr && get_style().get_number_format().get_format_code() == number_format::format::date_xlsx14);
}

cell_reference cell::get_reference() const
{
    return {d_->column_, d_->row_};
}

bool cell::operator==(std::nullptr_t) const
{
    return d_ == nullptr;
}

bool cell::operator==(const cell &comparand) const
{
    if(comparand == nullptr)
    {
        return d_ == nullptr;
    }

    return d_->value_ == comparand.d_->value_;
}

style &cell::get_style()
{
    if(d_->style_ == nullptr)
    {
        d_->style_ = new style();
    }
    return *d_->style_;
}

const style &cell::get_style() const
{
    if(d_->style_ == nullptr)
    {
        d_->style_ = new style();
    }
    return *d_->style_;
}
    
void cell::set_style(const xlnt::style &s)
{
    get_style() = s;
}

cell &cell::operator=(const cell &rhs)
{
    d_ = rhs.d_;
    return *this;
}

bool operator<(cell left, cell right)
{
    return left.get_reference() < right.get_reference();
}

std::string cell::to_string() const
{
    return "<Cell " + worksheet(d_->parent_).get_title() + "." + get_reference().to_string() + ">";
}

relationship cell::get_hyperlink() const
{
    if(!d_->has_hyperlink_)
    {
        throw std::runtime_error("no hyperlink set");
    }

    return d_->hyperlink_;
}

bool cell::has_hyperlink() const
{
    return d_->has_hyperlink_;
}

void cell::set_hyperlink(const std::string &hyperlink)
{
    if(hyperlink.length() == 0 || std::find(hyperlink.begin(), hyperlink.end(), ':') == hyperlink.end())
    {
        throw data_type_exception();
    }

    d_->has_hyperlink_ = true;
    d_->hyperlink_ = worksheet(d_->parent_).create_relationship(relationship::type::hyperlink, hyperlink);

    if(get_value().is(value::type::null))
    {
        set_value(hyperlink);
    }
}

void cell::set_formula(const std::string &formula)
{
    if(formula.length() == 0)
    {
        throw data_type_exception();
    }

    d_->formula_ = formula;
}

bool cell::has_formula() const
{
    return !d_->formula_.empty();
}

std::string cell::get_formula() const
{
    if(d_->formula_.empty())
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
    if(!has_comment())
    {
        get_parent().increment_comments();
    }
    
    d_->comment_ = c;
}

void cell::clear_comment()
{
    if(has_comment())
    {
        get_parent().decrement_comments();
    }
    
    d_->comment_ = comment();
}

bool cell::has_comment() const
{
    return d_->comment_.get_text() != "";
}

void cell::set_error(const std::string &error)
{
    if(error.length() == 0 || error[0] != '#')
    {
        throw data_type_exception();
    }

    set_value(value::error(error));
}

cell cell::offset(column_t column, row_t row)
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

comment cell::get_comment() const
{
    return d_->comment_;
}

std::pair<int, int> cell::get_anchor() const
{
    static const double DefaultColumnWidth = 51.85;
    static const double DefaultRowHeight = 15.0;

    auto points_to_pixels = [](double value, double dpi) { return (int)std::ceil(value * dpi / 72); };

    auto left_columns = d_->column_;
    auto &column_dimensions = get_parent().get_column_dimensions();
    int left_anchor = 0;
    auto default_width = points_to_pixels(DefaultColumnWidth, 96.0);

    for(int column_index = 1; column_index <= (int)left_columns; column_index++)
    {
        if(column_dimensions.find(column_index) != column_dimensions.end())
        {
            auto cdw = column_dimensions.at(column_index);

            if(cdw > 0)
            {
                left_anchor += points_to_pixels(cdw, 96.0);
                continue;
            }
        }

        left_anchor += default_width;
    }

    auto top_rows = d_->row_;
    auto &row_dimensions = get_parent().get_row_dimensions();
    int top_anchor = 0;
    auto default_height = points_to_pixels(DefaultRowHeight, 96.0);

    for(int row_index = 1; row_index <= (int)top_rows; row_index++)
    {
        if(row_dimensions.find(row_index) != row_dimensions.end())
        {
            auto rdh = row_dimensions.at(row_index);

            if(rdh > 0)
            {
                top_anchor += points_to_pixels(rdh, 96.0);
                continue;
            }
        }

        top_anchor += default_height;
    }

    return {left_anchor, top_anchor};
}

} // namespace xlnt
