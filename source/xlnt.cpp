#include <algorithm>
#include <array>
#include <cassert>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <locale>
#include <sstream>

#include "xlnt.h"
#include "../third-party/pugixml/src/pugixml.hpp"

namespace xlnt {

#ifdef _WIN32
#include <Windows.h>
#include <Shlwapi.h>
void file::copy(const std::string &source, const std::string &destination, bool overwrite)
{
    assert(source.size() + 1 < MAX_PATH);
    assert(destination.size() + 1 < MAX_PATH);

    std::wstring source_wide(source.begin(), source.end());
    std::wstring destination_wide(destination.begin(), destination.end());

    BOOL result = CopyFile(source_wide.c_str(), destination_wide.c_str(), !overwrite);

    if(result == 0)
    {
        switch(GetLastError())
        {
        case ERROR_ACCESS_DENIED: throw std::runtime_error("Access is denied");
        case ERROR_ENCRYPTION_FAILED: throw std::runtime_error("The specified file could not be encrypted");
        case ERROR_FILE_NOT_FOUND: throw std::runtime_error("The source file wasn't found");
        default:
            if(!overwrite)
            {
                throw std::runtime_error("The destination file already exists");
            }
            throw std::runtime_error("Unknown error");
        }
    }
}

bool file::exists(const std::string &path)
{
    std::wstring path_wide(path.begin(), path.end());
    return PathFileExists(path_wide.c_str()) && !PathIsDirectory(path_wide.c_str());
}

#else

#include <sys/stat.h>
void file::copy(const std::string &source, const std::string &destination, bool overwrite)
{
    if(!overwrite && exists(destination))
    {
	throw std::runtime_error("destination file already exists and overwrite==false");
    }

    std::ifstream src(source, std::ios::binary);
    std::ofstream dst(destination, std::ios::binary);
    
    dst << src.rdbuf();
}

bool file::exists(const std::string &path)
{
    struct stat fileAtt;
    
    if (stat(path.c_str(), &fileAtt) != 0)
    {
	throw std::runtime_error("stat failed");
    }
    
    return S_ISREG(fileAtt.st_mode);
}

#endif //_WIN32


zip_file::zip_file(const std::string &filename, file_mode mode, file_access access)
  : zip_file_(nullptr),
    unzip_file_(nullptr),
    current_state_(state::closed),
    filename_(filename),
    modified_(false),
//    mode_(mode),
    access_(access)
{
    switch(mode)
    {
    case file_mode::open:
        read_all();
        break;
    case file_mode::open_or_create:
        if(file_exists(filename))
        {
            read_all();
        }
        else
        {
            flush(true);
        }
        break;
    case file_mode::create:
        flush(true);
        break;
    case file_mode::create_new:
        if(file_exists(filename))
        {
            throw std::runtime_error("file exists");
        }
        flush(true);
        break;
    case file_mode::truncate:
        if((int)access & (int)file_access::read)
        {
            throw std::runtime_error("cannot read from file opened with file_mode truncate");
        }
        flush(true);
        break;
    case file_mode::append:
        read_all();
        break;
    }
}

zip_file::~zip_file()
{
    change_state(state::closed);
}

std::string zip_file::get_file_contents(const std::string &filename)
{
    return files_[filename];
}

void zip_file::set_file_contents(const std::string &filename, const std::string &contents)
{
    if(!has_file(filename) || files_[filename] != contents)
    {
        modified_ = true;
    }

    files_[filename] = contents;
}

void zip_file::delete_file(const std::string &filename)
{
    files_.erase(filename);
}

bool zip_file::has_file(const std::string &filename)
{
    return files_.find(filename) != files_.end();
}

void zip_file::flush(bool force_write)
{
    if(modified_ || force_write)
    {
        write_all();
    }
}

void zip_file::read_all()
{
    if(!((int)access_ & (int)file_access::read))
    {
        throw std::runtime_error("don't have read access");
    }

    change_state(state::read);

    int result = unzGoToFirstFile(unzip_file_);

    std::array<char, 1000> file_name_buffer = {{'\0'}};
    std::vector<char> file_buffer;

    while(result == UNZ_OK)
    {
        unz_file_info file_info;
        file_name_buffer.fill('\0');

        result = unzGetCurrentFileInfo(unzip_file_, &file_info, file_name_buffer.data(), 
            static_cast<uLong>(file_name_buffer.size()), nullptr, 0, nullptr, 0);

        if(result != UNZ_OK)
        {
            throw result;
        }

        result = unzOpenCurrentFile(unzip_file_);

        if(result != UNZ_OK)
        {
            throw result;
        }

        if(file_buffer.size() < file_info.uncompressed_size + 1)
        {
            file_buffer.resize(file_info.uncompressed_size + 1);
        }
        file_buffer[file_info.uncompressed_size] = '\0';

        result = unzReadCurrentFile(unzip_file_, file_buffer.data(), file_info.uncompressed_size);

        if(result != static_cast<int>(file_info.uncompressed_size))
        {
            throw result;
        }

        std::string current_filename(file_name_buffer.begin(), file_name_buffer.begin() + file_info.size_filename);
        std::string contents(file_buffer.begin(), file_buffer.begin() + file_info.uncompressed_size);

        if(current_filename.back() != '/')
        {
            files_[current_filename] = contents;
        }
        else
        {
            directories_.push_back(current_filename);
        }

        result = unzCloseCurrentFile(unzip_file_);

        if(result != UNZ_OK)
        {
            throw result;
        }

        result = unzGoToNextFile(unzip_file_);
    }
}

void zip_file::write_all()
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }

    change_state(state::write, false);

    for(auto file : files_)
    {
        write_to_zip(file.first, file.second, true);
    }

    modified_ = false;
}

std::string zip_file::read_from_zip(const std::string &filename)
{
    if(!((int)access_ & (int)file_access::read))
    {
        throw std::runtime_error("don't have read access");
    }

    change_state(state::read);

    auto result = unzLocateFile(unzip_file_, filename.c_str(), 1);

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = unzOpenCurrentFile(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    unz_file_info file_info;
    std::array<char, 1000> file_name_buffer;
    std::array<char, 1000> extra_field_buffer;
    std::array<char, 1000> comment_buffer;

    unzGetCurrentFileInfo(unzip_file_, &file_info,
        file_name_buffer.data(), static_cast<uLong>(file_name_buffer.size()),
        extra_field_buffer.data(), static_cast<uLong>(extra_field_buffer.size()),
        comment_buffer.data(), static_cast<uLong>(comment_buffer.size()));

    std::vector<char> file_buffer(file_info.uncompressed_size + 1, '\0');
    result = unzReadCurrentFile(unzip_file_, file_buffer.data(), file_info.uncompressed_size);

    if(result != static_cast<int>(file_info.uncompressed_size))
    {
        throw result;
    }

    result = unzCloseCurrentFile(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    return std::string(file_buffer.begin(), file_buffer.end());
}

void zip_file::write_to_zip(const std::string &filename, const std::string &content, bool append)
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }

    change_state(state::write, append);

    zip_fileinfo file_info;

    int result = zipOpenNewFileInZip(zip_file_, filename.c_str(), &file_info, nullptr, 0, nullptr, 0, nullptr, Z_DEFLATED, Z_DEFAULT_COMPRESSION);

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = zipWriteInFileInZip(zip_file_, content.data(), static_cast<int>(content.size()));

    if(result != UNZ_OK)
    {
        throw result;
    }

    result = zipCloseFileInZip(zip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }
}

void zip_file::change_state(state new_state, bool append)
{
    if(new_state == current_state_ && append)
    {
        return;
    }

    switch(new_state)
    {
    case state::closed:
        if(current_state_ == state::write)
        {
            stop_write();
        }
        else if(current_state_ == state::read)
        {
            stop_read();
        }
        break;
    case state::read:
        if(current_state_ == state::write)
        {
            stop_write();
        }
        start_read();
        break;
    case state::write:
        if(current_state_ == state::read)
        {
            stop_read();
        }
        if(current_state_ != state::write)
        {
            start_write(append);
        }
        break;
    default:
        throw std::runtime_error("bad enum");
    }

    current_state_ = new_state;
}

bool zip_file::file_exists(const std::string& name)
{
    std::ifstream f(name.c_str());
    return f.good();
}

void zip_file::start_read()
{
    if(unzip_file_ != nullptr || zip_file_ != nullptr)
    {
        throw std::runtime_error("bad state");
    }

    unzip_file_ = unzOpen(filename_.c_str());

    if(unzip_file_ == nullptr)
    {
        throw std::runtime_error("bad or non-existant file");
    }
}

void zip_file::stop_read()
{
    if(unzip_file_ == nullptr)
    {
        throw std::runtime_error("bad state");
    }

    int result = unzClose(unzip_file_);

    if(result != UNZ_OK)
    {
        throw result;
    }

    unzip_file_ = nullptr;
}

void zip_file::start_write(bool append)
{
    if(unzip_file_ != nullptr || zip_file_ != nullptr)
    {
        throw std::runtime_error("bad state");
    }

    int append_status;

    if(append)
    {
        if(!file_exists(filename_))
        {
            throw std::runtime_error("can't append to non-existent file");
        }
        append_status = APPEND_STATUS_ADDINZIP;
    }
    else
    {
        append_status = APPEND_STATUS_CREATE;
    }

    zip_file_ = zipOpen(filename_.c_str(), append_status);

    if(zip_file_ == nullptr)
    {
        if(append)
        {
            throw std::runtime_error("couldn't append to zip file");
        }
        else
        {
            throw std::runtime_error("couldn't create zip file");
        }
    }
}

void zip_file::stop_write()
{
    if(zip_file_ == nullptr)
    {
        throw std::runtime_error("bad state");
    }

    flush();

    int result = zipClose(zip_file_, nullptr);

    if(result != UNZ_OK)
    {
        throw result;
    }

    zip_file_ = nullptr;
}

const xlnt::color xlnt::color::black(0);
const xlnt::color xlnt::color::white(1);

struct cell_struct
{
    cell_struct(worksheet_struct *ws, int column, int row)
        : type(cell::type::null), parent_worksheet(ws),
        column(column), row(row),
        hyperlink_rel("invalid", "")
    {

    }

    std::string to_string() const;

    cell::type type;

    union
    {
        long double numeric_value;
        bool bool_value;
    };

    
    std::string error_value; 
    tm date_value;
    std::string string_value;
    std::string formula_value;
    worksheet_struct *parent_worksheet;
    int column;
    int row;
    style style;
    relationship hyperlink_rel;
    bool merged;
};

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

cell::cell() : root_(nullptr)
{
}

cell::cell(worksheet &worksheet, const std::string &column, int row) : root_(nullptr)
{
    cell self = worksheet.cell(column + std::to_string(row));
    root_ = self.root_;
}


cell::cell(worksheet &worksheet, const std::string &column, int row, const std::string &initial_value) : root_(nullptr)
{
    cell self = worksheet.cell(column + std::to_string(row));
    root_ = self.root_;
    *this = initial_value;
}

cell::cell(cell_struct *root) : root_(root)
{
}

std::string cell::get_value() const
{
    switch(root_->type)
    {
    case type::string:
        return root_->string_value;
    case type::numeric:
        return std::to_string(root_->numeric_value);
    case type::formula:
        return root_->formula_value;
    case type::error:
        return root_->error_value;
    case type::null:
        return "";
    case type::date:
        return "00:00:00";
    case type::boolean:
        return root_->bool_value ? "TRUE" : "FALSE";
    default:
        throw std::runtime_error("bad enum");
    }
}

int cell::get_row() const
{
    return root_->row + 1;
}

std::string cell::get_column() const
{
    return get_column_letter(root_->column + 1);
}

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

cell::type cell::data_type_for_value(const std::string &value)
{
    if(value.empty())
    {
        return type::null;
    }

    if(value[0] == '=')
    {
        return type::formula;
    }
    else if(value[0] == '0')
    {
        if(value.length() > 1)
        {
            if(value[1] == '.' || (value.length() > 2 && (value[1] == 'e' || value[1] == 'E')))
            {
                auto first_non_number = std::find_if(value.begin() + 2, value.end(),
                    [](char c) { return !std::isdigit(c, std::locale::classic()); });
                if(first_non_number == value.end())
                {
                    return type::numeric;
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
                        return type::string;
                    }
                }
                return type::date;
            }
            else
            {
                return type::string;
            }
        }
        return type::numeric;
    }
    else if(value[0] == '#')
    {
        return type::error;
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
                return type::boolean;
            }
            return type::string;
        }
        else
        {
            return type::numeric;
        }
    }
}

void cell::set_explicit_value(const std::string &value, type data_type)
{
    root_->type = data_type;
    switch(data_type)
    {
    case type::formula: root_->formula_value = value; return;
    case type::date: root_->date_value.tm_hour = std::stoi(value); return;
    case type::error: root_->error_value = value; return;
    case type::boolean: root_->bool_value = value == "true"; return;
    case type::null: return;
    case type::numeric: root_->numeric_value = std::stod(value); return;
    case type::string: root_->string_value = value; return;
    default: throw std::runtime_error("bad enum");
    }
}

void cell::set_hyperlink(const std::string &url)
{
    root_->type = type::hyperlink;
    root_->hyperlink_rel = worksheet(root_->parent_worksheet).create_relationship("hyperlink", url);
}

void cell::set_merged(bool merged)
{
    root_->merged = merged;
}

bool cell::get_merged() const
{
    return root_->merged;
}

bool cell::bind_value()
{
    root_->type = type::null;
    return true;
}

bool cell::bind_value(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return true;
}

bool cell::bind_value(const std::string &value)
{
    //Given a value, infer type and display options.
    root_->type = data_type_for_value(value);
    return true;
}

bool cell::bind_value(const char *value)
{
    return bind_value(std::string(value));
}

bool cell::bind_value(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return true;
}

bool cell::bind_value(const tm &value)
{
    root_->type = type::date;
    root_->date_value = value;
    return true;
}

coordinate cell::coordinate_from_string(const std::string &coord_string)
{
    // Convert a coordinate string like 'B12' to a tuple ('B', 12)
    bool column_part = true;
    coordinate result;

    for(auto character : coord_string)
    {
        char upper = std::toupper(character, std::locale::classic());

        if(std::isalpha(character, std::locale::classic()))
        {
            if(column_part)
            {
                result.column.append(1, upper);
            }
            else
            {
                throw bad_cell_coordinates(coord_string);
            }
        }
        else
        {
            if(column_part)
            {
                column_part = false;
            }
            else if(!(std::isdigit(character, std::locale::classic()) || character == '$'))
            {
                throw bad_cell_coordinates(coord_string);
            }
        }
    }

    std::string row_string = coord_string.substr(result.column.length());
    if(row_string[0] == '$')
    {
        row_string = row_string.substr(1);
    }
    result.row = std::stoi(row_string);

    if(result.row < 1)
    {
        throw bad_cell_coordinates(result.row, cell::column_index_from_string(result.column));
    }

    return result;
}

int cell::column_index_from_string(const std::string &column_string)
{
    if(column_string.length() > 3 || column_string.empty())
    {
        throw std::runtime_error("column must be one to three characters");
    }

    int column_index = 0;
    int place = 1;

    for(int i = static_cast<int>(column_string.length()) - 1; i >= 0; i--)
    {
        if(!std::isalpha(column_string[i], std::locale::classic()))
        {
            throw std::runtime_error("column must contain only letters in the range A-Z");
        }

        column_index += (std::toupper(column_string[i], std::locale::classic()) - 'A' + 1) * place;
        place *= 26;
    }

    return column_index;
}

// Convert a column number into a column letter (3 -> 'C')
// Right shift the column col_idx by 26 to find column letters in reverse
// order.These numbers are 1 - based, and can be converted to ASCII
// ordinals by adding 64.
std::string cell::get_column_letter(int column_index)
{
    // these indicies corrospond to A->ZZZ and include all allowed
    // columns
    if(column_index < 1 || column_index > 18278)
    {
	auto msg = "Column index out of bounds: " + std::to_string(column_index);
	throw std::runtime_error(msg);
    }
    
    auto temp = column_index;
    std::string column_letter = "";
    
    while(temp > 0)
    {
	int quotient = temp / 26, remainder = temp % 26;
	
	// check for exact division and borrow if needed
	if(remainder == 0)
	{
	    quotient -= 1;
	    remainder = 26;
	}
	
	column_letter = std::string(1, char(remainder + 64)) + column_letter;
	temp = quotient;
    }
    
    return column_letter;
}

bool cell::is_date() const
{
    return root_->type == type::date;
}

std::string cell::get_address() const
{
    return get_column_letter(root_->column + 1) + std::to_string(root_->row + 1);
}

std::string cell::get_coordinate() const
{
    return get_address();
}

std::string cell::get_hyperlink_rel_id() const
{
    return root_->hyperlink_rel.get_id();
}

bool cell::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}
    
bool cell::operator==(int comparand) const
{
    return root_->type == type::numeric && root_->numeric_value == comparand;
}

bool cell::operator==(double comparand) const
{
    return root_->type == type::numeric && root_->numeric_value == comparand;
}
    
bool cell::operator==(const std::string &comparand) const
{
    if(root_->type == type::hyperlink)
    {
        return root_->hyperlink_rel.get_target_uri() == comparand;
    }
    if(root_->type == type::string)
    {
        return root_->string_value == comparand;
    }
    return false;
}

bool cell::operator==(const char *comparand) const
{
    return *this == std::string(comparand);
}

bool cell::operator==(const tm &comparand) const
{
    return root_->type == cell::type::date && root_->date_value.tm_hour == comparand.tm_hour;
}

bool operator==(int comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const char *comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const std::string &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

bool operator==(const tm &comparand, const xlnt::cell &cell)
{
    return cell == comparand;
}

style &cell::get_style()
{
    return root_->style;
}

const style &cell::get_style() const
{
    return root_->style;
}

std::string cell::absolute_coordinate(const std::string &absolute_address)
{
    // Convert a coordinate to an absolute coordinate string (B12 -> $B$12)
    auto colon_index = absolute_address.find(':');
    if(colon_index != std::string::npos)
    {
        return absolute_coordinate(absolute_address.substr(0, colon_index)) + ":"
            + absolute_coordinate(absolute_address.substr(colon_index + 1));
    }
    else
    {
        auto coord = coordinate_from_string(absolute_address);
        return std::string("$") + coord.column + "$" + std::to_string(coord.row);
    }
}

xlnt::cell::type cell::get_data_type() const
{
    return root_->type;
}

xlnt::cell cell::get_offset(int row_offset, int column_offset)
{
    return worksheet(root_->parent_worksheet).cell(root_->column + column_offset, root_->row + row_offset);
}

cell &cell::operator=(const cell &rhs)
{
    root_ = rhs.root_;
    return *this;
}

cell &cell::operator=(int value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(double value)
{
    root_->type = type::numeric;
    root_->numeric_value = value;
    return *this;
}

cell &cell::operator=(bool value)
{
    root_->type = type::boolean;
    root_->bool_value = value;
    return *this;
}


cell &cell::operator=(const std::string &value)
{
    root_->type = data_type_for_value(value);

    switch(root_->type)
    {
    case type::date:
        {
        root_->date_value = std::tm();
        auto split = split_string(value, ':');
        root_->date_value.tm_hour = std::stoi(split[0]);
        root_->date_value.tm_min = std::stoi(split[1]);
        if(split.size() > 2)
        {
            root_->date_value.tm_sec = std::stoi(split[2]);
        }
        break;
        }
    case type::formula:
        root_->formula_value = value;
        break;
    case type::numeric:
        root_->numeric_value = std::stod(value);
        break;
    case type::boolean:
        root_->bool_value = value == "TRUE" || value == "true";
        break;
    case type::error:
        root_->error_value = value;
        break;
    case type::string:
        root_->string_value = value;
        break;
    case type::null:
        break;
    default:
        throw std::runtime_error("bad enum");
    }

    return *this;
}

cell &cell::operator=(const char *value)
{
    return *this = std::string(value);
}

cell &cell::operator=(const tm &value)
{
    root_->type = type::date;
    root_->date_value = value;
    return *this;
}

std::string cell::to_string() const
{
    return root_->to_string();
}

struct worksheet_struct
{
    worksheet_struct(workbook &parent_workbook, const std::string &title)
        : parent_(parent_workbook), title_(title), freeze_panes_(nullptr)
    {

    }

    ~worksheet_struct()
    {
        clear();
    }

    void clear()
    {
        while(!cell_map_.empty())
        {
            delete cell_map_.begin()->second.root_;
            cell_map_.erase(cell_map_.begin()->first);
        }
    }

    void garbage_collect()
    {
        std::vector<std::string> null_cell_keys;

        for(auto key_cell_pair : cell_map_)
        {
            if(key_cell_pair.second.get_data_type() == cell::type::null)
            {
                null_cell_keys.push_back(key_cell_pair.first);
            }
        }

        while(!null_cell_keys.empty())
        {
            cell_map_.erase(null_cell_keys.back());
            null_cell_keys.pop_back();
        }
    }

    std::list<cell> get_cell_collection()
    {
        std::list<xlnt::cell> cells;
        for(auto cell : cell_map_)
        {
            cells.push_front(xlnt::cell(cell.second));
        }
        return cells;
    }

    std::string get_title() const
    {
        return title_;
    }

    void set_title(const std::string &title)
    {
        title_ = title;
    }

    cell get_freeze_panes() const
    {
        return freeze_panes_;
    }

    void set_freeze_panes(cell top_left_cell)
    {
        if(top_left_cell.get_address() == "A1")
        {
            unfreeze_panes();
        }
        else
        {
            freeze_panes_ = top_left_cell;
        }
    }

    void set_freeze_panes(const std::string &top_left_coordinate)
    {
        freeze_panes_ = cell(top_left_coordinate);
    }

    void unfreeze_panes()
    {
        freeze_panes_ = xlnt::cell(nullptr);
    }

    xlnt::cell cell(const std::string &coordinate)
    {
        if(cell_map_.find(coordinate) == cell_map_.end())
        {
            auto coord = xlnt::cell::coordinate_from_string(coordinate);
            cell_struct *cell = new xlnt::cell_struct(this, xlnt::cell::column_index_from_string(coord.column) - 1, coord.row - 1);
            cell_map_[coordinate] = xlnt::cell(cell);
        }

        return cell_map_[coordinate];
    }

    xlnt::cell cell(int column, int row)
    {
        return cell(xlnt::cell::get_column_letter(column + 1) + std::to_string(row + 1));
    }

    int get_highest_row() const
    {
        int highest = 1;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, cell.second.get_row());
        }
        return highest;
    }

    int get_highest_column() const
    {
        int highest = 1;
        for(auto cell : cell_map_)
        {
            highest = (std::max)(highest, xlnt::cell::column_index_from_string(cell.second.get_column()));
        }
        return highest;
    }

    int get_lowest_row() const
    {
        if(cell_map_.empty())
        {
            return 1;
        }

        int lowest = cell_map_.begin()->second.get_row();

        for(auto cell : cell_map_)
        {
            lowest = (std::min)(lowest, cell.second.get_row());
        }

        return lowest;
    }

    int get_lowest_column() const
    {
        if(cell_map_.empty())
        {
            return 1;
        }

        int lowest = xlnt::cell::column_index_from_string(cell_map_.begin()->second.get_column());

        for(auto cell : cell_map_)
        {
            lowest = (std::min)(lowest, xlnt::cell::column_index_from_string(cell.second.get_column()));
        }

        return lowest;
    }

    std::string calculate_dimension() const
    {
        int lowest_column = get_lowest_column();
        std::string lowest_column_letter = xlnt::cell::get_column_letter(lowest_column);
        int lowest_row = get_lowest_row();

        int highest_column = get_highest_column();
        std::string highest_column_letter = xlnt::cell::get_column_letter(highest_column);
        int highest_row = get_highest_row();

        return lowest_column_letter + std::to_string(lowest_row) + ":" + highest_column_letter + std::to_string(highest_row);
    }

    xlnt::range range(const std::string &range_string, int row_offset, int column_offset)
    {
        xlnt::range r;
        
        if(parent_.has_named_range(range_string, worksheet(this)))
        {
            return range(parent_.get_named_range(range_string, worksheet(this)).get_range_string(), row_offset, column_offset);
        }

        auto colon_index = range_string.find(':');

        if(colon_index != std::string::npos)
        {
            auto min_range = range_string.substr(0, colon_index);
            auto max_range = range_string.substr(colon_index + 1);
            auto min_coord = xlnt::cell::coordinate_from_string(min_range);
            auto max_coord = xlnt::cell::coordinate_from_string(max_range);

            if(column_offset != 0)
            {
                min_coord.column = xlnt::cell::get_column_letter(xlnt::cell::column_index_from_string(min_coord.column) + column_offset);
                max_coord.column = xlnt::cell::get_column_letter(xlnt::cell::column_index_from_string(max_coord.column) + column_offset);
            }
            
            if(row_offset != 0)
            {
                min_coord.row = min_coord.row + row_offset;
                max_coord.row = max_coord.row + row_offset;
            }

            std::unordered_map<int, std::string> column_cache;

            for(int i = xlnt::cell::column_index_from_string(min_coord.column);
                i <= xlnt::cell::column_index_from_string(max_coord.column); i++)
            {
                column_cache[i] = xlnt::cell::get_column_letter(i);
            }
            for(int row = min_coord.row; row <= max_coord.row; row++)
            {
                r.push_back(std::vector<xlnt::cell>());
                for(int column = xlnt::cell::column_index_from_string(min_coord.column);
                    column <= xlnt::cell::column_index_from_string(max_coord.column); column++)
                {
                    std::string coordinate = column_cache[column] + std::to_string(row);
                    r.back().push_back(cell(coordinate));
                }
            }
        }
        else
        {
            r.push_back(std::vector<xlnt::cell>());
            r.back().push_back(cell(range_string));
        }

        return r;
    }

    relationship create_relationship(const std::string &relationship_type, const std::string &target_uri)
    {
        std::string r_id = "rId" + std::to_string(relationships_.size() + 1);
        relationships_.push_back(relationship(relationship_type, r_id, target_uri));
        return relationships_.back();
    }

    //void add_chart(chart chart);

    void merge_cells(const std::string &range_string)
    {
        merged_cells_.push_back(range_string);
        bool first = true;
        for(auto row : range(range_string, 0, 0))
        {
            for(auto cell : row)
            {
                cell.set_merged(true);
                if(!first)
                {
                    cell.bind_value();
                }
                first = false;
            }
        }
    }

    void merge_cells(int start_row, int start_column, int end_row, int end_column)
    {
        auto range_string = xlnt::cell::get_column_letter(start_column + 1) + std::to_string(start_row + 1) + ":"
            + xlnt::cell::get_column_letter(end_column + 1) + std::to_string(end_row + 1);
        merge_cells(range_string);
    }

    void unmerge_cells(const std::string &range_string)
    {
        auto match = std::find(merged_cells_.begin(), merged_cells_.end(), range_string);
        if(match == merged_cells_.end())
        {
            throw std::runtime_error("cells not merged");
        }
        merged_cells_.erase(match);
        for(auto row : range(range_string, 0, 0))
        {
            for(auto cell : row)
            {
                cell.set_merged(false);
            }
        }
    }

    void unmerge_cells(int start_row, int start_column, int end_row, int end_column)
    {
        auto range_string = xlnt::cell::get_column_letter(start_column + 1) + std::to_string(start_row + 1) + ":"
            + xlnt::cell::get_column_letter(end_column + 1) + std::to_string(end_row + 1);
        merge_cells(range_string);
    }

    void append(const std::vector<std::string> &cells)
    {
        int row = get_highest_row() - 1;
        if(cell_map_.size() != 0)
        {
            row++;
        }
        int column = 0;
        for(auto cell : cells)
        {
            this->cell(column++, row) = cell;
        }
    }

    void append(const std::unordered_map<std::string, std::string> &cells)
    {
        int row = get_highest_row() - 1;
        if(cell_map_.size() != 0)
        {
            row++;
        }
        for(auto cell : cells)
        {
            int column = xlnt::cell::column_index_from_string(cell.first) - 1;
            this->cell(column, row) = cell.second;
        }
    }

    void append(const std::unordered_map<int, std::string> &cells)
    {
        int row = get_highest_row() - 1;
        if(cell_map_.size() != 0)
        {
            row++;
        }
        for(auto cell : cells)
        {
            this->cell(cell.first, row) = cell.second;
        }
    }

    xlnt::range rows()
    {
        return range(calculate_dimension(), 0, 0);
    }

    xlnt::range columns()
    {
        int max_row = get_highest_row();
        xlnt::range cols;
        for(int col_idx = 0; col_idx < get_highest_column(); col_idx++)
        {
            cols.push_back(std::vector<xlnt::cell>());
            std::string col = xlnt::cell::get_column_letter(col_idx + 1);
            std::string range_string = col + "1:" + col + std::to_string(max_row + 1);
            for(auto row : range(range_string, 0, 0))
            {
                cols.back().push_back(row[0]);
            }
        }
        return cols;
    }

    void operator=(const worksheet_struct &other) = delete;

    workbook &parent_;
    std::string title_;
    xlnt::cell freeze_panes_;
    std::unordered_map<std::string, xlnt::cell> cell_map_;
    std::vector<relationship> relationships_;
    page_setup page_setup_;
    std::string auto_filter_;
    margins page_margins_;
    std::vector<std::string> merged_cells_;
};
    
worksheet::worksheet() : root_(nullptr)
{
    
}

worksheet::worksheet(worksheet_struct *root) : root_(root)
{
}

worksheet::worksheet(workbook &parent)
{
    *this = parent.create_sheet();
}

std::vector<std::string> worksheet::get_merged_ranges() const
{
    return root_->merged_cells_;
}
    
margins &worksheet::get_page_margins()
{
    return root_->page_margins_;
}

void worksheet::set_auto_filter(const std::string &range_string)
{
    std::string upper;
    std::transform(range_string.begin(), range_string.end(), std::back_inserter(upper), 
        [](char c) { return std::toupper(c, std::locale::classic()); });
    root_->auto_filter_ = upper;
}

void worksheet::set_auto_filter(const xlnt::range &range)
{
    auto first = range[0][0].get_address();
    auto last = range.back().back().get_address();
    if(first == last)
    {
        set_auto_filter(first);
    }
    else
    {
        set_auto_filter(first + ":" + last);
    }
}

std::string worksheet::get_auto_filter() const
{
    return root_->auto_filter_;
}

bool worksheet::has_auto_filter() const
{
    return root_->auto_filter_ != "";
}

void worksheet::unset_auto_filter()
{
    root_->auto_filter_ = "";
}
    
page_setup &worksheet::get_page_setup()
{
    return root_->page_setup_;
}

std::string worksheet::to_string() const
{
    return "<Worksheet \"" + root_->title_ + "\">";
}

workbook &worksheet::get_parent() const
{
    return root_->parent_;
}

void worksheet::garbage_collect()
{
    root_->garbage_collect();
}

std::list<cell> worksheet::get_cell_collection()
{
    return root_->get_cell_collection();
}

std::string worksheet::get_title() const
{
    if(root_ == nullptr)
    {
        throw std::runtime_error("null worksheet");
    }
    return root_->title_;
}

void worksheet::set_title(const std::string &title)
{
    root_->title_ = title;
}

cell worksheet::get_freeze_panes() const
{
    return root_->freeze_panes_;
}

void worksheet::set_freeze_panes(xlnt::cell top_left_cell)
{
    root_->set_freeze_panes(top_left_cell);
}

void worksheet::set_freeze_panes(const std::string &top_left_coordinate)
{
    root_->set_freeze_panes(top_left_coordinate);
}

void worksheet::unfreeze_panes()
{
    root_->unfreeze_panes();
}

xlnt::cell worksheet::cell(const std::string &coordinate)
{
    return root_->cell(coordinate);
}

xlnt::cell worksheet::cell(int row, int column)
{
    return root_->cell(row, column);
}

int worksheet::get_highest_row() const
{
    return root_->get_highest_row();
}

int worksheet::get_highest_column() const
{
    return root_->get_highest_column();
}

std::string worksheet::calculate_dimension() const
{
    return root_->calculate_dimension();
}

range worksheet::range(const std::string &range_string, int row_offset, int column_offset)
{
    return root_->range(range_string, row_offset, column_offset);
}

range worksheet::range(const std::string &range_string)
{
    return root_->range(range_string, 0, 0);
}

std::vector<relationship> worksheet::get_relationships()
{
    return root_->relationships_;
}

relationship worksheet::create_relationship(const std::string &relationship_type, const std::string &target_uri)
{
    return root_->create_relationship(relationship_type, target_uri);
}

//void worksheet::add_chart(chart chart);

void worksheet::merge_cells(const std::string &range_string)
{
    root_->merge_cells(range_string);
}

void worksheet::merge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->merge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::unmerge_cells(const std::string &range_string)
{
    root_->unmerge_cells(range_string);
}

void worksheet::unmerge_cells(int start_row, int start_column, int end_row, int end_column)
{
    root_->unmerge_cells(start_row, start_column, end_row, end_column);
}

void worksheet::append(const std::vector<std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, std::string> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<int, std::string> &cells)
{
    root_->append(cells);
}

xlnt::range worksheet::rows() const
{
    return root_->rows();
}

xlnt::range worksheet::columns() const
{
    return root_->columns();
}

bool worksheet::operator==(const worksheet &other) const
{
    return root_ == other.root_;
}

bool worksheet::operator!=(const worksheet &other) const
{
    return root_ != other.root_;
}

bool worksheet::operator==(std::nullptr_t) const
{
    return root_ == nullptr;
}

bool worksheet::operator!=(std::nullptr_t) const
{
    return root_ != nullptr;
}


void worksheet::operator=(const worksheet &other)
{
    root_ = other.root_;
}

cell worksheet::operator[](const std::string &address)
{
    return cell(address);
}

std::string write_content_types(const std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> &content_types)
{
    /*std::set<std::string> seen;

    pugi::xml_node root;

    if(wb.has_vba_archive())
    {
        root = fromstring(wb.get_vba_archive().read(ARC_CONTENT_TYPES));

        for(auto elem : root.findall("{" + CONTYPES_NS + "}Override"))
        {
            seen.insert(elem.attrib["PartName"]);
        }
    }
    else
    {
        root = Element("{" + CONTYPES_NS + "}Types");

        for(auto content_type : static_content_types_config)
        {
            if(setting_type == "Override")
            {
                tag = "{" + CONTYPES_NS + "}Override";
                attrib = {"PartName": "/" + name};
            }
            else
            {
                tag = "{" + CONTYPES_NS + "}Default";
                attrib = {"Extension": name};
            }

            attrib["ContentType"] = content_type;
            SubElement(root, tag, attrib);
        }
    }

    int drawing_id = 1;
    int chart_id = 1;
    int comments_id = 1;
    int sheet_id = 0;

    for(auto sheet : wb)
    {
        std::string name = "/xl/worksheets/sheet" + std::to_string(sheet_id) + ".xml";

        if(seen.find(name) == seen.end())
        {
            SubElement(root, "{" + CONTYPES_NS + "}Override", {{"PartName", name},
            {"ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}});
        }

        if(sheet._charts || sheet._images)
        {
            name = "/xl/drawings/drawing" + drawing_id + ".xml";

            if(seen.find(name) == seen.end())
            {
                SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                    "ContentType" : "application/vnd.openxmlformats-officedocument.drawing+xml"});
            }

            drawing_id += 1;

            for(auto chart : sheet._charts)
            {
                name = "/xl/charts/chart%d.xml" % chart_id;

                if(seen.find(name) == seen.end())
                {
                    SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                        "ContentType" : "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"});
                }

                chart_id += 1;

                if(chart._shapes)
                {
                    name = "/xl/drawings/drawing%d.xml" % drawing_id;

                    if(seen.find(name) == seen.end())
                    {
                        SubElement(root, "{%s}Override" % CONTYPES_NS, {"PartName" : name,
                            "ContentType" : "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"});
                    }

                    drawing_id += 1;
                }
            }
        }
        if(sheet.get_comment_count() > 0)
        {
            SubElement(root, "{%s}Override" % CONTYPES_NS,
            {"PartName": "/xl/comments%d.xml" % comments_id,
            "ContentType" : "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"});
            comments_id += 1;
        }
    }*/

    pugi::xml_document doc;
    auto root_node = doc.append_child("Types");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/content-types");

    for(auto default_type : content_types.first)
    {
        auto xml_node = root_node.append_child("Default");
        xml_node.append_attribute("Extension").set_value(default_type.first.c_str());
        xml_node.append_attribute("ContentType").set_value(default_type.second.c_str());
    }
    
    for(auto override_type : content_types.second)
    {
        auto theme_node = root_node.append_child("Override");
        theme_node.append_attribute("PartName").set_value(override_type.first.c_str());
        theme_node.append_attribute("ContentType").set_value(override_type.second.c_str());
    }

    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::string write_relationships(const std::unordered_map<std::string, std::pair<std::string, std::string>> &relationships)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("Relationships");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/relationships");

    for(auto relationship : relationships)
    {
        auto app_props_node = root_node.append_child("Relationship");
        app_props_node.append_attribute("Id").set_value(relationship.first.c_str());
        app_props_node.append_attribute("Type").set_value(relationship.second.first.c_str());
        app_props_node.append_attribute("Target").set_value(relationship.second.second.c_str());
    }

    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

std::unordered_map<std::string, std::pair<std::string, std::string>> read_relationships(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());

    auto root_node = doc.child("Relationships");

    std::unordered_map<std::string, std::pair<std::string, std::string>> relationships;

    for(auto relationship : root_node.children("Relationship"))
    {
        std::string id = relationship.attribute("Id").as_string();
        std::string type = relationship.attribute("Type").as_string();
        std::string target = relationship.attribute("Target").as_string();
        relationships[id] = std::make_pair(type, target);
    }

    return relationships;
}

std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> read_content_types(const std::string &content)
{
    pugi::xml_document doc;
    doc.load(content.c_str());

    auto root_node = doc.child("Types");

    std::unordered_map<std::string, std::string> default_types;

    for(auto child : root_node.children("Default"))
    {
        default_types[child.attribute("Extension").as_string()] = child.attribute("ContentType").as_string();
    }

    std::unordered_map<std::string, std::string> override_types;

    for(auto child : root_node.children("Override"))
    {
        override_types[child.attribute("PartName").as_string()] = child.attribute("ContentType").as_string();
    }

    return std::make_pair(default_types, override_types);
}

workbook::workbook(optimized optimized) 
  : optimized_write_(optimized==optimized::write),
    optimized_read_(optimized==optimized::read), 
    active_sheet_index_(0)
{
    if(!optimized_write_)
    {
        auto *worksheet = new worksheet_struct(*this, "Sheet1");
        worksheets_.push_back(worksheet);
    }
}

workbook::~workbook()
{
    clear();
}

worksheet workbook::get_sheet_by_name(const std::string &name)
{
    auto match = std::find_if(worksheets_.begin(), worksheets_.end(), [&](const worksheet &w) { return w.get_title() == name; });
    if(match != worksheets_.end())
    {
        return worksheet(*match);
    }
    return worksheet(nullptr);
}

worksheet workbook::get_active_sheet()
{
    return worksheets_[active_sheet_index_];
}
    
bool workbook::has_named_range(const std::string &name, xlnt::worksheet ws) const
{
    for(auto named_range : named_ranges_)
    {
        if(named_range.first == name && named_range.second.get_scope() == ws)
        {
            return true;
        }
    }
    return false;
}

worksheet workbook::create_sheet()
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    std::string title = "Sheet1";
    int index = 1;
    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }
    auto *worksheet = new worksheet_struct(*this, title);
    worksheets_.push_back(worksheet);
    return get_sheet_by_name(title);
}
    
void workbook::add_sheet(xlnt::worksheet worksheet)
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            throw std::runtime_error("worksheet already in workbook");
        }
    }
    worksheets_.push_back(worksheet);
}
    
void workbook::add_sheet(xlnt::worksheet worksheet, std::size_t index)
{
    if(optimized_read_)
    {
        throw std::runtime_error("this is a read-only workbook");
    }
    
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            throw std::runtime_error("worksheet already in workbook");
        }
    }
    worksheets_.push_back(worksheet);
    std::swap(worksheets_[index], worksheets_.back());
}
    
int workbook::get_index(xlnt::worksheet worksheet)
{
    int i = 0;
    for(auto ws : *this)
    {
        if(worksheet == ws)
        {
            return i;
        }
        i++;
    }
    throw std::runtime_error("worksheet isn't owned by this workbook");
}

void workbook::create_named_range(const std::string &name, worksheet range_owner, const std::string &range_string)
{
    for(auto ws : worksheets_)
    {
        if(ws == range_owner)
        {
            named_ranges_[name] = named_range(range_owner, range_string);
            return;
        }
    }
    
    throw std::runtime_error("worksheet isn't owned by this workbook");
}
    
void workbook::add_named_range(const std::string &name, named_range named_range)
{
    for(auto ws : worksheets_)
    {
        if(ws == named_range.get_scope())
        {
            named_ranges_[name] = named_range;
            return;
        }
    }
    
    throw std::runtime_error("worksheet isn't owned by this workbook");
}
    
std::vector<named_range> workbook::get_named_ranges()
{
    std::vector<named_range> named_ranges;
    for(auto named_range : named_ranges_)
    {
        named_ranges.push_back(named_range.second);
    }
    return named_ranges;
}
    
void workbook::remove_named_range(named_range named_range)
{
    std::string key_match = "";
    for(auto r : named_ranges_)
    {
        if(r.second == named_range)
        {
            key_match = r.first;
        }
    }
    if(key_match == "")
    {
        throw std::runtime_error("range not found in worksheet");
    }
    named_ranges_.erase(key_match);
}
    
named_range workbook::get_named_range(const std::string &name, worksheet ws)
{
    for(auto named_range : named_ranges_)
    {
        if(named_range.first == name && named_range.second.get_scope() == ws)
        {
            return named_range.second;
        }
    }
    throw std::runtime_error("doesn't exist");
}
    
std::string determine_document_type(const std::unordered_map<std::string, std::pair<std::string, std::string>> &root_relationships,
                                    const std::unordered_map<std::string, std::string> &override_types)
{
    auto relationship_match = std::find_if(root_relationships.begin(), root_relationships.end(),
                                           [](const std::pair<std::string, std::pair<std::string, std::string>> &v)
                                           { return v.second.first == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"; });
    std::string type;
    
    if(relationship_match != root_relationships.end())
    {
        std::string office_document_relationship = relationship_match->second.second;
        
        if(office_document_relationship[0] != '/')
        {
            office_document_relationship = std::string("/") + office_document_relationship;
        }
        
        type = override_types.at(office_document_relationship);
    }
    
    if(type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    {
        return "excel";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    {
        return "powerpoint";
    }
    else if(type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    {
        return "word";
    }
    
    return "unsupported";
}

void workbook::load(const std::string &filename)
{
    zip_file f(filename, file_mode::open);
    //auto core_properties = read_core_properties();
    //auto app_properties = read_app_properties();
    auto root_relationships = read_relationships(f.get_file_contents("_rels/.rels"));
    auto content_types = read_content_types(f.get_file_contents("[Content_Types].xml"));
    
    auto type = determine_document_type(root_relationships, content_types.second);
    
    if(type != "excel")
    {
        throw std::runtime_error("unsupported document type: " + filename);
    }
    
    auto workbook_relationships = read_relationships(f.get_file_contents("xl/_rels/workbook.xml.rels"));

    pugi::xml_document doc;
    doc.load(f.get_file_contents("xl/workbook.xml").c_str());

    auto root_node = doc.child("workbook");
    auto sheets_node = root_node.child("sheets");

    while(!worksheets_.empty())
    {
        remove_sheet(worksheets_.front());
    }

    std::vector<std::string> shared_strings;
    if(f.has_file("xl/sharedStrings.xml"))
    {
        shared_strings = xlnt::reader::read_shared_string(f.get_file_contents("xl/sharedStrings.xml"));
    }

    for(auto sheet_node : sheets_node.children("sheet"))
    {
        auto relation_id = sheet_node.attribute("r:id").as_string();
        auto ws = create_sheet(sheet_node.attribute("name").as_string());
        std::string sheet_filename("xl/");
        sheet_filename += workbook_relationships[relation_id].second;
        xlnt::reader::read_worksheet(ws, f.get_file_contents(sheet_filename).c_str(), shared_strings);
    }
}

void workbook::remove_sheet(worksheet ws)
{
    auto match_iter = std::find(worksheets_.begin(), worksheets_.end(), ws);
    if(match_iter == worksheets_.end())
    {
        throw std::runtime_error("worksheet not owned by this workbook");
    }
    delete match_iter->root_;
    worksheets_.erase(match_iter);
}

worksheet workbook::create_sheet(std::size_t index)
{
    auto ws = create_sheet();
    if(index != worksheets_.size())
    {
        std::swap(worksheets_[index], worksheets_.back());
    }
    return ws;
}
    
worksheet workbook::create_sheet(std::size_t index, const std::string &title)
{
    auto ws = create_sheet(index);
    ws.set_title(title);
    return ws;
}

worksheet workbook::create_sheet(const std::string &title)
{
    if(title.length() > 31)
    {
        throw bad_sheet_title(title);
    }

    if(std::find_if(title.begin(), title.end(), 
        [](char c) { return c == '*' || c == ':' || c == '/' || c == '\\' || c == '?' || c == '[' || c == ']'; }) != title.end())
    {
        throw bad_sheet_title(title);
    }
    
    if(get_sheet_by_name(title) != nullptr)
    {
        throw std::runtime_error("sheet exists");
    }

    auto *worksheet = new worksheet_struct(*this, title);
    worksheets_.push_back(worksheet);
    return xlnt::worksheet(worksheet);
}

std::vector<worksheet>::iterator workbook::begin()
{
    return worksheets_.begin();
}

std::vector<worksheet>::iterator workbook::end()
{
    return worksheets_.end();
}

std::vector<std::string> workbook::get_sheet_names() const
{
    std::vector<std::string> names;
    for(auto &ws : worksheets_)
    {
        names.push_back(ws.get_title());
    }
    return names;
}

worksheet workbook::operator[](const std::string &name)
{
    for(auto sheet : worksheets_)
    {
        if(sheet.get_title() == name)
        {
            return sheet;
        }
    }
    throw std::runtime_error("sheet not found");
}

worksheet workbook::operator[](int index)
{
    return worksheets_[index];
}

void workbook::clear()
{
    while(!worksheets_.empty())
    {
        delete worksheets_.back().root_;
        worksheets_.pop_back();
    }
}

void workbook::save(const std::string &filename)
{
    zip_file f(filename, file_mode::create, file_access::write);

    std::pair<std::unordered_map<std::string, std::string>, std::unordered_map<std::string, std::string>> content_types = 
    {
        {
            {"rels", "application/vnd.openxmlformats-package.relationships+xml"},
            {"xml", "application/xml"}
        },
        {
            {"/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"},
            {"/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"},
            {"/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"},
            {"/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"},
            {"/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"},
            {"/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"}
        }
    };

    f.set_file_contents("[Content_Types].xml", write_content_types(content_types));

    std::unordered_map<std::string, std::pair<std::string, std::string>> root_rels = 
    {
        {"rId3", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"}},
        {"rId2", {"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"}},
        {"rId1", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml"}}
    };
    f.set_file_contents("_rels/.rels", write_relationships(root_rels));

    std::unordered_map<std::string, std::pair<std::string, std::string>> workbook_rels =
    {
        {"rId1", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet1.xml"}},
        {"rId2", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"}},
        {"rId3", {"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"}}
    };
    f.set_file_contents("xl/_rels/workbook.xml.rels", write_relationships(workbook_rels));

    pugi::xml_document doc;
    auto root_node = doc.append_child("workbook");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    auto sheets_node = root_node.append_child("sheets");

    int i = 0;
    for(auto ws : *this)
    {
        auto sheet_node = sheets_node.append_child("sheet");
        sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
        sheet_node.append_attribute("sheetId").set_value(std::to_string(i + 1).c_str());
        sheet_node.append_attribute("r:id").set_value((std::string("rId") + std::to_string(i + 1)).c_str());
        std::string filename = "xl/worksheets/sheet";
        f.set_file_contents(filename + std::to_string(i + 1) + ".xml", xlnt::writer::write_worksheet(ws));
        i++;
    }

    std::stringstream ss;
    doc.save(ss);

    f.set_file_contents("xl/workbook.xml", ss.str());
}

std::string cell_struct::to_string() const
{
    return "<Cell " + parent_worksheet->title_ + "." + xlnt::cell::get_column_letter(column + 1) + std::to_string(row + 1) + ">";
}

bool workbook::operator==(const workbook &rhs) const
{
    if(optimized_write_ == rhs.optimized_write_
        && optimized_read_ == rhs.optimized_read_
        && guess_types_ == rhs.guess_types_
        && data_only_ == rhs.data_only_
        && active_sheet_index_ == rhs.active_sheet_index_
        && encoding_ == rhs.encoding_)
    {
        if(worksheets_.size() != rhs.worksheets_.size())
        {
            return false;
        }

        for(std::size_t i = 0; i < worksheets_.size(); i++)
        {
            if(worksheets_[i] != rhs.worksheets_[i])
            {
                return false;
            }
        }

        /*
        if(named_ranges_.size() != rhs.named_ranges_.size())
        {
            return false;
        }

        for(int i = 0; i < named_ranges_.size(); i++)
        {
            if(named_ranges_[i] != rhs.named_ranges_[i])
            {
                return false;
            }
        }

        if(relationships_.size() != rhs.relationships_.size())
        {
            return false;
        }

        for(int i = 0; i < relationships_.size(); i++)
        {
            if(relationships_[i] != rhs.relationships_[i])
            {
                return false;
            }
        }

        if(drawings_.size() != rhs.drawings_.size())
        {
            return false;
        }

        for(int i = 0; i < drawings_.size(); i++)
        {
            if(drawings_[i] != rhs.drawings_[i])
            {
                return false;
            }
        }
        */

        return true;
    }
    return false;
}

void sheet_protection::set_password(const std::string &password)
{
    hashed_password_ = hash_password(password);
}

template< typename T >
std::string int_to_hex(T i)
{
    std::stringstream stream;
    stream << std::hex << i;
    return stream.str();
}

std::string sheet_protection::hash_password(const std::string &plaintext_password)
{
    int password = 0x0000;
    int i = 1;

    for(auto character : plaintext_password)
    {
        int value = character << i;
        int rotated_bits = value >> 15;
        value &= 0x7fff;
        password ^= (value | rotated_bits);
        i++;
    }

    password ^= plaintext_password.size();
    password ^= 0xCE4B;

    std::string hashed = int_to_hex(password);
    std::transform(hashed.begin(), hashed.end(), hashed.begin(), [](char c){ return std::toupper(c, std::locale::classic()); });
    return hashed;
}

int string_table::operator[](const std::string &key) const
{
    for(std::size_t i = 0; i < strings_.size(); i++)
    {
        if(key == strings_[i])
        {
            return (int)i;
        }
    }
    throw std::runtime_error("bad string");
}

void string_table_builder::add(const std::string &string)
{
    for(std::size_t i = 0; i < table_.strings_.size(); i++)
    {
        if(string == table_.strings_[i])
        {
            return;
        }
    }
    table_.strings_.push_back(string);
}
    
void read_worksheet(worksheet ws, const pugi::xml_node &root_node, const std::vector<std::string> &string_table)
{
    auto dimension_node = root_node.child("dimension");
    std::string dimension = dimension_node.attribute("ref").as_string();
    auto sheet_data_node = root_node.child("sheetData");
    auto merge_cells_node = root_node.child("mergeCells");

    if(merge_cells_node != nullptr)
    {
        int count = merge_cells_node.attribute("count").as_int();

        for(auto merge_cell_node : merge_cells_node.children("mergeCell"))
        {
            ws.merge_cells(merge_cell_node.attribute("ref").as_string());
            count--;
        }

        if(count != 0)
        {
            throw std::runtime_error("mismatch between count and actual number of merged cells");
        }
    }

    for(auto row_node : sheet_data_node.children("row"))
    {
        int row_index = row_node.attribute("r").as_int();
        std::string span_string = row_node.attribute("spans").as_string();
        auto colon_index = span_string.find(':');
        int min_column = std::stoi(span_string.substr(0, colon_index));
        int max_column = std::stoi(span_string.substr(colon_index + 1));

        for(int i = min_column; i < max_column + 1; i++)
        {
            std::string address = xlnt::cell::get_column_letter(i) + std::to_string(row_index);
            auto cell_node = row_node.find_child_by_attribute("c", "r", address.c_str());

            if(cell_node != nullptr)
            {
                if(cell_node.attribute("t") != nullptr && std::string(cell_node.attribute("t").as_string()) == "inlineStr") // inline string
                {
                    ws.cell(address) = cell_node.child("is").child("t").text().as_string();
                }
                else if(cell_node.attribute("t") != nullptr && std::string(cell_node.attribute("t").as_string()) == "s") // shared string
                {
                    ws.cell(address) = string_table.at(cell_node.child("v").text().as_int());
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "2") // date
                {
                    tm date = {0};
                    date.tm_year = 1900;
                    int days = cell_node.child("v").text().as_int();
                    while(days > 365)
                    {
                        date.tm_year += 1;
                        days -= 365;
                    }
                    date.tm_yday = days;
                    while(days > 30)
                    {
                        date.tm_mon += 1;
                        days -= 30;
                    }
                    date.tm_mday = days;
                    ws.cell(address) = date;
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "3") // time
                {
                    tm date = {0};
                    double remaining = cell_node.child("v").text().as_double() * 24;
                    date.tm_hour = (int)(remaining);
                    remaining -= date.tm_hour;
                    remaining *= 60;
                    date.tm_min = (int)(remaining);
                    remaining -= date.tm_min;
                    remaining *= 60;
                    date.tm_sec = (int)(remaining);
                    remaining -= date.tm_sec;
                    if(remaining > 0.5)
                    {
                        date.tm_sec++;
                        if(date.tm_sec > 59)
                        {
                            date.tm_sec = 0;
                            date.tm_min++;
                            if(date.tm_min > 59)
                            {
                                date.tm_min = 0;
                                date.tm_hour++;
                            }
                        }
                    }
                    ws.cell(address) = date;
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "4") // decimal
                {
                    ws.cell(address) = cell_node.child("v").text().as_double();
                }
                else if(cell_node.attribute("s") != nullptr && std::string(cell_node.attribute("s").as_string()) == "1") // percent
                {
                    ws.cell(address) = cell_node.child("v").text().as_double();
                }
                else if(cell_node.child("v") != nullptr)
                {
                    ws.cell(address) = cell_node.child("v").text().as_string();
                }
            }
        }
    }
}

void xlnt::reader::read_worksheet(worksheet ws, const std::string &xml_string, const std::vector<std::string> &string_table)
{
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    xlnt::read_worksheet(ws, doc.child("worksheet"), string_table);
}

worksheet xlnt::reader::read_worksheet(std::istream &handle, xlnt::workbook &wb, const std::string &title, const std::vector<std::string> &string_table)
{
    auto ws = wb.create_sheet();
    ws.set_title(title);
    pugi::xml_document doc;
    doc.load(handle);
    xlnt::read_worksheet(ws, doc.child("worksheet"), string_table);
    return ws;
}

std::vector<std::string> reader::read_shared_string(const std::string &xml_string)
{
    std::vector<std::string> shared_strings;
    pugi::xml_document doc;
    doc.load(xml_string.c_str());
    auto root_node = doc.child("sst");
    //int count = root_node.attribute("count").as_int();
    int unique_count = root_node.attribute("uniqueCount").as_int();

    for(auto si_node : root_node)
    {
        shared_strings.push_back(si_node.child("t").text().as_string());
    }

    if(unique_count != shared_strings.size())
    {
        throw std::runtime_error("counts don't match");
    }

    return shared_strings;
}
    
std::string xlnt::writer::write_worksheet(xlnt::worksheet ws, const std::vector<std::string> &string_table)
{
    pugi::xml_document doc;
    auto root_node = doc.append_child("worksheet");
    root_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    root_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    auto dimension_node = root_node.append_child("dimension");
    int width = (std::max)(1, ws.get_highest_column());
    std::string width_letter = xlnt::cell::get_column_letter(width);
    int height = (std::max)(1, ws.get_highest_row());
    auto dimension = width_letter + std::to_string(height);
    dimension_node.append_attribute("ref").set_value(dimension.c_str());
    auto sheet_views_node = root_node.append_child("sheetViews");
    auto sheet_view_node = sheet_views_node.append_child("sheetView");
    sheet_view_node.append_attribute("tabSelected").set_value(1);
    sheet_view_node.append_attribute("workbookViewId").set_value(0);

    auto selection_node = sheet_view_node.append_child("selection");

    std::string active_cell = "B2";
    selection_node.append_attribute("activeCell").set_value(active_cell.c_str());
    selection_node.append_attribute("sqref").set_value(active_cell.c_str());

    auto sheet_format_pr_node = root_node.append_child("sheetFormatPr");
    sheet_format_pr_node.append_attribute("defaultRowHeight").set_value(15);

    auto sheet_data_node = root_node.append_child("sheetData");
    for(auto row : ws.rows())
    {
        if(row.empty())
        {
            continue;
        }

        int min = (int)row.size();
        int max = 0;
        bool any_non_null = false;

        for(auto cell : row)
        {
            min = (std::min)(min, cell::column_index_from_string(cell.get_column()));
            max = (std::max)(max, cell::column_index_from_string(cell.get_column()));

            if(cell.get_data_type() != cell::type::null)
            {
                any_non_null = true;
            }
        }

        if(!any_non_null)
        {
            continue;
        }

        auto row_node = sheet_data_node.append_child("row");
        row_node.append_attribute("r").set_value(row.front().get_row());

        row_node.append_attribute("spans").set_value((std::to_string(min) + ":" + std::to_string(max)).c_str());
        row_node.append_attribute("x14ac:dyDescent").set_value(0.25);

        for(auto cell : row)
        {
            if(cell.get_data_type() != cell::type::null || cell.get_merged())
            {
                auto cell_node = row_node.append_child("c");
                cell_node.append_attribute("r").set_value(cell.get_address().c_str());

                if(cell.get_data_type() == cell::type::string)
                {
                    int match_index = -1;
                    for(int i = 0; i < string_table.size(); i++)
                    {
                        if(string_table[i] == cell.get_value())
                        {
                            match_index = i;
                            break;
                        }
                    }

                    if(match_index == -1)
                    {
                        cell_node.append_attribute("t").set_value("inlineStr");
                        auto inline_string_node = cell_node.append_child("is");
                        inline_string_node.append_child("t").set_value(cell.get_value().c_str());
                    }
                    else
                    {
                        cell_node.append_attribute("t").set_value("s");
                        auto value_node = cell_node.append_child("v");
                        value_node.text().set(match_index);
                    }
                }
                else
                {
                    if(cell.get_data_type() != cell::type::null)
                    {
                        auto value_node = cell_node.append_child("v");
                        value_node.text().set(cell.get_value().c_str());
                    }
                }
            }
        }
    }

    if(!ws.get_merged_ranges().empty())
    {
        auto merge_cells_node = root_node.append_child("mergeCells");
        merge_cells_node.append_attribute("count").set_value(ws.get_merged_ranges().size());

        for(auto merged_range : ws.get_merged_ranges())
        {
            auto merge_cell_node = merge_cells_node.append_child("mergeCell");
            merge_cell_node.append_attribute("ref").set_value(merged_range.c_str());
        }
    }

    if(!ws.get_page_margins().is_default())
    {
        auto page_margins_node = root_node.append_child("pageMargins");

        page_margins_node.append_attribute("left").set_value(ws.get_page_margins().get_left());
        page_margins_node.append_attribute("right").set_value(ws.get_page_margins().get_right());
        page_margins_node.append_attribute("top").set_value(ws.get_page_margins().get_top());
        page_margins_node.append_attribute("bottom").set_value(ws.get_page_margins().get_bottom());
        page_margins_node.append_attribute("header").set_value(ws.get_page_margins().get_header());
        page_margins_node.append_attribute("footer").set_value(ws.get_page_margins().get_footer());
    }

    if(!ws.get_page_setup().is_default())
    {
        auto page_setup_node = root_node.append_child("pageSetup");

        std::string orientation_string = ws.get_page_setup().get_orientation() == xlnt::page_setup::orientation::landscape ? "landscape" : "portrait";
        page_setup_node.append_attribute("orientation").set_value(orientation_string.c_str());
        page_setup_node.append_attribute("paperSize").set_value((int)ws.get_page_setup().get_paper_size());
        page_setup_node.append_attribute("fitToHeight").set_value(ws.get_page_setup().fit_to_height() ? 1 : 0);
        page_setup_node.append_attribute("fitToWidth").set_value(ws.get_page_setup().fit_to_width() ? 1 : 0);

        auto page_set_up_pr_node = root_node.append_child("pageSetUpPr");
        page_set_up_pr_node.append_attribute("fitToPage").set_value(ws.get_page_setup().fit_to_page() ? 1 : 0);
    }

    std::stringstream ss;
    doc.save(ss);

    return ss.str();
}
    
std::string xlnt::writer::write_worksheet(xlnt::worksheet ws)
{
    return write_worksheet(ws, std::vector<std::string>());
}
    
bool named_range::operator==(const xlnt::named_range &comparand) const
{
    return comparand.parent_worksheet_ == parent_worksheet_ && comparand.range_string_ == range_string_;
}
    
    std::string xlnt::writer::write_theme()
    {
        /*
        pugi::xml_document doc;
        auto theme_node = doc.append_child("a:theme");
        theme_node.append_attribute("xmlns:a").set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
        theme_node.append_attribute("name").set_value("Office Theme");
        auto theme_elements_node = theme_node.append_child("a:themeElements");
        auto clr_scheme_node = theme_elements_node.append_child("a:clrScheme");
        clr_scheme_node.append_attribute("name").set_value("Office");
        
        struct scheme_element
        {
            std::string name;
            std::string sub_element_name;
            std::string val;
        };
        
        std::vector<scheme_element> scheme_elements =
        {
            {"a:dk1", "a:sysClr", "windowText"}
        };
        
        for(auto element : scheme_elements)
        {
            auto element_node = clr_scheme_node.append_child("a:dk1");
            element_node.append_child(element.sub_element_name.c_str()).append_attribute("val").set_value(element.val.c_str());
        }
        
        std::stringstream ss;
        doc.print(ss);
        return ss.str();
        */
        std::array<unsigned char, 6994> data =
        {
            60, 63, 120, 109, 108, 32, 118, 101, 114, 115,
            105, 111, 110, 61, 34, 49, 46, 48, 34, 32,
            101, 110, 99, 111, 100, 105, 110, 103, 61, 34,
            85, 84, 70, 45, 56, 34, 32, 115, 116, 97,
            110, 100, 97, 108, 111, 110, 101, 61, 34, 121,
            101, 115, 34, 63, 62, 10, 60, 97, 58, 116,
            104, 101, 109, 101, 32, 120, 109, 108, 110, 115,
            58, 97, 61, 34, 104, 116, 116, 112, 58, 47,
            47, 115, 99, 104, 101, 109, 97, 115, 46, 111,
            112, 101, 110, 120, 109, 108, 102, 111, 114, 109,
            97, 116, 115, 46, 111, 114, 103, 47, 100, 114,
            97, 119, 105, 110, 103, 109, 108, 47, 50, 48,
            48, 54, 47, 109, 97, 105, 110, 34, 32, 110,
            97, 109, 101, 61, 34, 79, 102, 102, 105, 99,
            101, 32, 84, 104, 101, 109, 101, 34, 62, 60,
            97, 58, 116, 104, 101, 109, 101, 69, 108, 101,
            109, 101, 110, 116, 115, 62, 60, 97, 58, 99,
            108, 114, 83, 99, 104, 101, 109, 101, 32, 110,
            97, 109, 101, 61, 34, 79, 102, 102, 105, 99,
            101, 34, 62, 60, 97, 58, 100, 107, 49, 62,
            60, 97, 58, 115, 121, 115, 67, 108, 114, 32,
            118, 97, 108, 61, 34, 119, 105, 110, 100, 111,
            119, 84, 101, 120, 116, 34, 32, 108, 97, 115,
            116, 67, 108, 114, 61, 34, 48, 48, 48, 48,
            48, 48, 34, 47, 62, 60, 47, 97, 58, 100,
            107, 49, 62, 60, 97, 58, 108, 116, 49, 62,
            60, 97, 58, 115, 121, 115, 67, 108, 114, 32,
            118, 97, 108, 61, 34, 119, 105, 110, 100, 111,
            119, 34, 32, 108, 97, 115, 116, 67, 108, 114,
            61, 34, 70, 70, 70, 70, 70, 70, 34, 47,
            62, 60, 47, 97, 58, 108, 116, 49, 62, 60,
            97, 58, 100, 107, 50, 62, 60, 97, 58, 115,
            114, 103, 98, 67, 108, 114, 32, 118, 97, 108,
            61, 34, 49, 70, 52, 57, 55, 68, 34, 47,
            62, 60, 47, 97, 58, 100, 107, 50, 62, 60,
            97, 58, 108, 116, 50, 62, 60, 97, 58, 115,
            114, 103, 98, 67, 108, 114, 32, 118, 97, 108,
            61, 34, 69, 69, 69, 67, 69, 49, 34, 47,
            62, 60, 47, 97, 58, 108, 116, 50, 62, 60,
            97, 58, 97, 99, 99, 101, 110, 116, 49, 62,
            60, 97, 58, 115, 114, 103, 98, 67, 108, 114,
            32, 118, 97, 108, 61, 34, 52, 70, 56, 49,
            66, 68, 34, 47, 62, 60, 47, 97, 58, 97,
            99, 99, 101, 110, 116, 49, 62, 60, 97, 58,
            97, 99, 99, 101, 110, 116, 50, 62, 60, 97,
            58, 115, 114, 103, 98, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 67, 48, 53, 48, 52, 68,
            34, 47, 62, 60, 47, 97, 58, 97, 99, 99,
            101, 110, 116, 50, 62, 60, 97, 58, 97, 99,
            99, 101, 110, 116, 51, 62, 60, 97, 58, 115,
            114, 103, 98, 67, 108, 114, 32, 118, 97, 108,
            61, 34, 57, 66, 66, 66, 53, 57, 34, 47,
            62, 60, 47, 97, 58, 97, 99, 99, 101, 110,
            116, 51, 62, 60, 97, 58, 97, 99, 99, 101,
            110, 116, 52, 62, 60, 97, 58, 115, 114, 103,
            98, 67, 108, 114, 32, 118, 97, 108, 61, 34,
            56, 48, 54, 52, 65, 50, 34, 47, 62, 60,
            47, 97, 58, 97, 99, 99, 101, 110, 116, 52,
            62, 60, 97, 58, 97, 99, 99, 101, 110, 116,
            53, 62, 60, 97, 58, 115, 114, 103, 98, 67,
            108, 114, 32, 118, 97, 108, 61, 34, 52, 66,
            65, 67, 67, 54, 34, 47, 62, 60, 47, 97,
            58, 97, 99, 99, 101, 110, 116, 53, 62, 60,
            97, 58, 97, 99, 99, 101, 110, 116, 54, 62,
            60, 97, 58, 115, 114, 103, 98, 67, 108, 114,
            32, 118, 97, 108, 61, 34, 70, 55, 57, 54,
            52, 54, 34, 47, 62, 60, 47, 97, 58, 97,
            99, 99, 101, 110, 116, 54, 62, 60, 97, 58,
            104, 108, 105, 110, 107, 62, 60, 97, 58, 115,
            114, 103, 98, 67, 108, 114, 32, 118, 97, 108,
            61, 34, 48, 48, 48, 48, 70, 70, 34, 47,
            62, 60, 47, 97, 58, 104, 108, 105, 110, 107,
            62, 60, 97, 58, 102, 111, 108, 72, 108, 105,
            110, 107, 62, 60, 97, 58, 115, 114, 103, 98,
            67, 108, 114, 32, 118, 97, 108, 61, 34, 56,
            48, 48, 48, 56, 48, 34, 47, 62, 60, 47,
            97, 58, 102, 111, 108, 72, 108, 105, 110, 107,
            62, 60, 47, 97, 58, 99, 108, 114, 83, 99,
            104, 101, 109, 101, 62, 60, 97, 58, 102, 111,
            110, 116, 83, 99, 104, 101, 109, 101, 32, 110,
            97, 109, 101, 61, 34, 79, 102, 102, 105, 99,
            101, 34, 62, 60, 97, 58, 109, 97, 106, 111,
            114, 70, 111, 110, 116, 62, 60, 97, 58, 108,
            97, 116, 105, 110, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 67, 97, 109, 98, 114,
            105, 97, 34, 47, 62, 60, 97, 58, 101, 97,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 34, 47, 62, 60, 97, 58, 99, 115, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 74,
            112, 97, 110, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 239, 188, 173, 239, 188,
            179, 32, 239, 188, 176, 227, 130, 180, 227, 130,
            183, 227, 131, 131, 227, 130, 175, 34, 47, 62,
            60, 97, 58, 102, 111, 110, 116, 32, 115, 99,
            114, 105, 112, 116, 61, 34, 72, 97, 110, 103,
            34, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 235, 167, 145, 236, 157, 128, 32, 234,
            179, 160, 235, 148, 149, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 72, 97, 110, 115, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            229, 174, 139, 228, 189, 147, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 72, 97, 110, 116, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 230, 150, 176, 231, 180, 176, 230, 152, 142,
            233, 171, 148, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 65, 114, 97, 98, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 84, 105,
            109, 101, 115, 32, 78, 101, 119, 32, 82, 111,
            109, 97, 110, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 72, 101, 98, 114, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 84, 105,
            109, 101, 115, 32, 78, 101, 119, 32, 82, 111,
            109, 97, 110, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 84, 104, 97, 105, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 84, 97,
            104, 111, 109, 97, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 69, 116, 104, 105, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 78,
            121, 97, 108, 97, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 66, 101, 110, 103, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 86,
            114, 105, 110, 100, 97, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 71, 117, 106, 114, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            83, 104, 114, 117, 116, 105, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 75, 104, 109, 114, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 111, 111, 108, 66, 111, 114, 97, 110,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 75,
            110, 100, 97, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 84, 117, 110, 103, 97,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 71,
            117, 114, 117, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 82, 97, 97, 118, 105,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 67,
            97, 110, 115, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 69, 117, 112, 104, 101,
            109, 105, 97, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 67, 104, 101, 114, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 80, 108,
            97, 110, 116, 97, 103, 101, 110, 101, 116, 32,
            67, 104, 101, 114, 111, 107, 101, 101, 34, 47,
            62, 60, 97, 58, 102, 111, 110, 116, 32, 115,
            99, 114, 105, 112, 116, 61, 34, 89, 105, 105,
            105, 34, 32, 116, 121, 112, 101, 102, 97, 99,
            101, 61, 34, 77, 105, 99, 114, 111, 115, 111,
            102, 116, 32, 89, 105, 32, 66, 97, 105, 116,
            105, 34, 47, 62, 60, 97, 58, 102, 111, 110,
            116, 32, 115, 99, 114, 105, 112, 116, 61, 34,
            84, 105, 98, 116, 34, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 77, 105, 99, 114,
            111, 115, 111, 102, 116, 32, 72, 105, 109, 97,
            108, 97, 121, 97, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 84, 104, 97, 97, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 77,
            86, 32, 66, 111, 108, 105, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 68, 101, 118, 97, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 97, 110, 103, 97, 108, 34, 47, 62,
            60, 97, 58, 102, 111, 110, 116, 32, 115, 99,
            114, 105, 112, 116, 61, 34, 84, 101, 108, 117,
            34, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 71, 97, 117, 116, 97, 109, 105, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 84, 97,
            109, 108, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 76, 97, 116, 104, 97, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 83, 121,
            114, 99, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 69, 115, 116, 114, 97, 110,
            103, 101, 108, 111, 32, 69, 100, 101, 115, 115,
            97, 34, 47, 62, 60, 97, 58, 102, 111, 110,
            116, 32, 115, 99, 114, 105, 112, 116, 61, 34,
            79, 114, 121, 97, 34, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 75, 97, 108, 105,
            110, 103, 97, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 77, 108, 121, 109, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 75, 97,
            114, 116, 105, 107, 97, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 76, 97, 111, 111, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            68, 111, 107, 67, 104, 97, 109, 112, 97, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 83, 105,
            110, 104, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 73, 115, 107, 111, 111, 108,
            97, 32, 80, 111, 116, 97, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 77, 111, 110, 103, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 111, 110, 103, 111, 108, 105, 97, 110,
            32, 66, 97, 105, 116, 105, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 86, 105, 101, 116, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 84, 105, 109, 101, 115, 32, 78, 101, 119,
            32, 82, 111, 109, 97, 110, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 85, 105, 103, 104, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 105, 99, 114, 111, 115, 111, 102, 116,
            32, 85, 105, 103, 104, 117, 114, 34, 47, 62,
            60, 47, 97, 58, 109, 97, 106, 111, 114, 70,
            111, 110, 116, 62, 60, 97, 58, 109, 105, 110,
            111, 114, 70, 111, 110, 116, 62, 60, 97, 58,
            108, 97, 116, 105, 110, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 67, 97, 108, 105,
            98, 114, 105, 34, 47, 62, 60, 97, 58, 101,
            97, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 34, 47, 62, 60, 97, 58, 99, 115,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 34, 47, 62, 60, 97, 58, 102, 111, 110,
            116, 32, 115, 99, 114, 105, 112, 116, 61, 34,
            74, 112, 97, 110, 34, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 239, 188, 173, 239,
            188, 179, 32, 239, 188, 176, 227, 130, 180, 227,
            130, 183, 227, 131, 131, 227, 130, 175, 34, 47,
            62, 60, 97, 58, 102, 111, 110, 116, 32, 115,
            99, 114, 105, 112, 116, 61, 34, 72, 97, 110,
            103, 34, 32, 116, 121, 112, 101, 102, 97, 99,
            101, 61, 34, 235, 167, 145, 236, 157, 128, 32,
            234, 179, 160, 235, 148, 149, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 72, 97, 110, 115, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 229, 174, 139, 228, 189, 147, 34, 47, 62,
            60, 97, 58, 102, 111, 110, 116, 32, 115, 99,
            114, 105, 112, 116, 61, 34, 72, 97, 110, 116,
            34, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 230, 150, 176, 231, 180, 176, 230, 152,
            142, 233, 171, 148, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 65, 114, 97, 98, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 65,
            114, 105, 97, 108, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 72, 101, 98, 114, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 65,
            114, 105, 97, 108, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 84, 104, 97, 105, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 84,
            97, 104, 111, 109, 97, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 69, 116, 104, 105, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            78, 121, 97, 108, 97, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 66, 101, 110, 103, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            86, 114, 105, 110, 100, 97, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 71, 117, 106, 114, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 83, 104, 114, 117, 116, 105, 34, 47, 62,
            60, 97, 58, 102, 111, 110, 116, 32, 115, 99,
            114, 105, 112, 116, 61, 34, 75, 104, 109, 114,
            34, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 68, 97, 117, 110, 80, 101, 110, 104,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 75,
            110, 100, 97, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 84, 117, 110, 103, 97,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 71,
            117, 114, 117, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 82, 97, 97, 118, 105,
            34, 47, 62, 60, 97, 58, 102, 111, 110, 116,
            32, 115, 99, 114, 105, 112, 116, 61, 34, 67,
            97, 110, 115, 34, 32, 116, 121, 112, 101, 102,
            97, 99, 101, 61, 34, 69, 117, 112, 104, 101,
            109, 105, 97, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 67, 104, 101, 114, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 80, 108,
            97, 110, 116, 97, 103, 101, 110, 101, 116, 32,
            67, 104, 101, 114, 111, 107, 101, 101, 34, 47,
            62, 60, 97, 58, 102, 111, 110, 116, 32, 115,
            99, 114, 105, 112, 116, 61, 34, 89, 105, 105,
            105, 34, 32, 116, 121, 112, 101, 102, 97, 99,
            101, 61, 34, 77, 105, 99, 114, 111, 115, 111,
            102, 116, 32, 89, 105, 32, 66, 97, 105, 116,
            105, 34, 47, 62, 60, 97, 58, 102, 111, 110,
            116, 32, 115, 99, 114, 105, 112, 116, 61, 34,
            84, 105, 98, 116, 34, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 77, 105, 99, 114,
            111, 115, 111, 102, 116, 32, 72, 105, 109, 97,
            108, 97, 121, 97, 34, 47, 62, 60, 97, 58,
            102, 111, 110, 116, 32, 115, 99, 114, 105, 112,
            116, 61, 34, 84, 104, 97, 97, 34, 32, 116,
            121, 112, 101, 102, 97, 99, 101, 61, 34, 77,
            86, 32, 66, 111, 108, 105, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 68, 101, 118, 97, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 97, 110, 103, 97, 108, 34, 47, 62,
            60, 97, 58, 102, 111, 110, 116, 32, 115, 99,
            114, 105, 112, 116, 61, 34, 84, 101, 108, 117,
            34, 32, 116, 121, 112, 101, 102, 97, 99, 101,
            61, 34, 71, 97, 117, 116, 97, 109, 105, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 84, 97,
            109, 108, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 76, 97, 116, 104, 97, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 83, 121,
            114, 99, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 69, 115, 116, 114, 97, 110,
            103, 101, 108, 111, 32, 69, 100, 101, 115, 115,
            97, 34, 47, 62, 60, 97, 58, 102, 111, 110,
            116, 32, 115, 99, 114, 105, 112, 116, 61, 34,
            79, 114, 121, 97, 34, 32, 116, 121, 112, 101,
            102, 97, 99, 101, 61, 34, 75, 97, 108, 105,
            110, 103, 97, 34, 47, 62, 60, 97, 58, 102,
            111, 110, 116, 32, 115, 99, 114, 105, 112, 116,
            61, 34, 77, 108, 121, 109, 34, 32, 116, 121,
            112, 101, 102, 97, 99, 101, 61, 34, 75, 97,
            114, 116, 105, 107, 97, 34, 47, 62, 60, 97,
            58, 102, 111, 110, 116, 32, 115, 99, 114, 105,
            112, 116, 61, 34, 76, 97, 111, 111, 34, 32,
            116, 121, 112, 101, 102, 97, 99, 101, 61, 34,
            68, 111, 107, 67, 104, 97, 109, 112, 97, 34,
            47, 62, 60, 97, 58, 102, 111, 110, 116, 32,
            115, 99, 114, 105, 112, 116, 61, 34, 83, 105,
            110, 104, 34, 32, 116, 121, 112, 101, 102, 97,
            99, 101, 61, 34, 73, 115, 107, 111, 111, 108,
            97, 32, 80, 111, 116, 97, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 77, 111, 110, 103, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 111, 110, 103, 111, 108, 105, 97, 110,
            32, 66, 97, 105, 116, 105, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 86, 105, 101, 116, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 65, 114, 105, 97, 108, 34, 47, 62, 60,
            97, 58, 102, 111, 110, 116, 32, 115, 99, 114,
            105, 112, 116, 61, 34, 85, 105, 103, 104, 34,
            32, 116, 121, 112, 101, 102, 97, 99, 101, 61,
            34, 77, 105, 99, 114, 111, 115, 111, 102, 116,
            32, 85, 105, 103, 104, 117, 114, 34, 47, 62,
            60, 47, 97, 58, 109, 105, 110, 111, 114, 70,
            111, 110, 116, 62, 60, 47, 97, 58, 102, 111,
            110, 116, 83, 99, 104, 101, 109, 101, 62, 60,
            97, 58, 102, 109, 116, 83, 99, 104, 101, 109,
            101, 32, 110, 97, 109, 101, 61, 34, 79, 102,
            102, 105, 99, 101, 34, 62, 60, 97, 58, 102,
            105, 108, 108, 83, 116, 121, 108, 101, 76, 115,
            116, 62, 60, 97, 58, 115, 111, 108, 105, 100,
            70, 105, 108, 108, 62, 60, 97, 58, 115, 99,
            104, 101, 109, 101, 67, 108, 114, 32, 118, 97,
            108, 61, 34, 112, 104, 67, 108, 114, 34, 47,
            62, 60, 47, 97, 58, 115, 111, 108, 105, 100,
            70, 105, 108, 108, 62, 60, 97, 58, 103, 114,
            97, 100, 70, 105, 108, 108, 32, 114, 111, 116,
            87, 105, 116, 104, 83, 104, 97, 112, 101, 61,
            34, 49, 34, 62, 60, 97, 58, 103, 115, 76,
            115, 116, 62, 60, 97, 58, 103, 115, 32, 112,
            111, 115, 61, 34, 48, 34, 62, 60, 97, 58,
            115, 99, 104, 101, 109, 101, 67, 108, 114, 32,
            118, 97, 108, 61, 34, 112, 104, 67, 108, 114,
            34, 62, 60, 97, 58, 116, 105, 110, 116, 32,
            118, 97, 108, 61, 34, 53, 48, 48, 48, 48,
            34, 47, 62, 60, 97, 58, 115, 97, 116, 77,
            111, 100, 32, 118, 97, 108, 61, 34, 51, 48,
            48, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 115, 99, 104, 101, 109, 101, 67, 108, 114,
            62, 60, 47, 97, 58, 103, 115, 62, 60, 97,
            58, 103, 115, 32, 112, 111, 115, 61, 34, 51,
            53, 48, 48, 48, 34, 62, 60, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 112, 104, 67, 108, 114, 34,
            62, 60, 97, 58, 116, 105, 110, 116, 32, 118,
            97, 108, 61, 34, 51, 55, 48, 48, 48, 34,
            47, 62, 60, 97, 58, 115, 97, 116, 77, 111,
            100, 32, 118, 97, 108, 61, 34, 51, 48, 48,
            48, 48, 48, 34, 47, 62, 60, 47, 97, 58,
            115, 99, 104, 101, 109, 101, 67, 108, 114, 62,
            60, 47, 97, 58, 103, 115, 62, 60, 97, 58,
            103, 115, 32, 112, 111, 115, 61, 34, 49, 48,
            48, 48, 48, 48, 34, 62, 60, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 112, 104, 67, 108, 114, 34,
            62, 60, 97, 58, 116, 105, 110, 116, 32, 118,
            97, 108, 61, 34, 49, 53, 48, 48, 48, 34,
            47, 62, 60, 97, 58, 115, 97, 116, 77, 111,
            100, 32, 118, 97, 108, 61, 34, 51, 53, 48,
            48, 48, 48, 34, 47, 62, 60, 47, 97, 58,
            115, 99, 104, 101, 109, 101, 67, 108, 114, 62,
            60, 47, 97, 58, 103, 115, 62, 60, 47, 97,
            58, 103, 115, 76, 115, 116, 62, 60, 97, 58,
            108, 105, 110, 32, 97, 110, 103, 61, 34, 49,
            54, 50, 48, 48, 48, 48, 48, 34, 32, 115,
            99, 97, 108, 101, 100, 61, 34, 49, 34, 47,
            62, 60, 47, 97, 58, 103, 114, 97, 100, 70,
            105, 108, 108, 62, 60, 97, 58, 103, 114, 97,
            100, 70, 105, 108, 108, 32, 114, 111, 116, 87,
            105, 116, 104, 83, 104, 97, 112, 101, 61, 34,
            49, 34, 62, 60, 97, 58, 103, 115, 76, 115,
            116, 62, 60, 97, 58, 103, 115, 32, 112, 111,
            115, 61, 34, 48, 34, 62, 60, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 112, 104, 67, 108, 114, 34,
            62, 60, 97, 58, 115, 104, 97, 100, 101, 32,
            118, 97, 108, 61, 34, 53, 49, 48, 48, 48,
            34, 47, 62, 60, 97, 58, 115, 97, 116, 77,
            111, 100, 32, 118, 97, 108, 61, 34, 49, 51,
            48, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 115, 99, 104, 101, 109, 101, 67, 108, 114,
            62, 60, 47, 97, 58, 103, 115, 62, 60, 97,
            58, 103, 115, 32, 112, 111, 115, 61, 34, 56,
            48, 48, 48, 48, 34, 62, 60, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 112, 104, 67, 108, 114, 34,
            62, 60, 97, 58, 115, 104, 97, 100, 101, 32,
            118, 97, 108, 61, 34, 57, 51, 48, 48, 48,
            34, 47, 62, 60, 97, 58, 115, 97, 116, 77,
            111, 100, 32, 118, 97, 108, 61, 34, 49, 51,
            48, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 115, 99, 104, 101, 109, 101, 67, 108, 114,
            62, 60, 47, 97, 58, 103, 115, 62, 60, 97,
            58, 103, 115, 32, 112, 111, 115, 61, 34, 49,
            48, 48, 48, 48, 48, 34, 62, 60, 97, 58,
            115, 99, 104, 101, 109, 101, 67, 108, 114, 32,
            118, 97, 108, 61, 34, 112, 104, 67, 108, 114,
            34, 62, 60, 97, 58, 115, 104, 97, 100, 101,
            32, 118, 97, 108, 61, 34, 57, 52, 48, 48,
            48, 34, 47, 62, 60, 97, 58, 115, 97, 116,
            77, 111, 100, 32, 118, 97, 108, 61, 34, 49,
            51, 53, 48, 48, 48, 34, 47, 62, 60, 47,
            97, 58, 115, 99, 104, 101, 109, 101, 67, 108,
            114, 62, 60, 47, 97, 58, 103, 115, 62, 60,
            47, 97, 58, 103, 115, 76, 115, 116, 62, 60,
            97, 58, 108, 105, 110, 32, 97, 110, 103, 61,
            34, 49, 54, 50, 48, 48, 48, 48, 48, 34,
            32, 115, 99, 97, 108, 101, 100, 61, 34, 48,
            34, 47, 62, 60, 47, 97, 58, 103, 114, 97,
            100, 70, 105, 108, 108, 62, 60, 47, 97, 58,
            102, 105, 108, 108, 83, 116, 121, 108, 101, 76,
            115, 116, 62, 60, 97, 58, 108, 110, 83, 116,
            121, 108, 101, 76, 115, 116, 62, 60, 97, 58,
            108, 110, 32, 119, 61, 34, 57, 53, 50, 53,
            34, 32, 99, 97, 112, 61, 34, 102, 108, 97,
            116, 34, 32, 99, 109, 112, 100, 61, 34, 115,
            110, 103, 34, 32, 97, 108, 103, 110, 61, 34,
            99, 116, 114, 34, 62, 60, 97, 58, 115, 111,
            108, 105, 100, 70, 105, 108, 108, 62, 60, 97,
            58, 115, 99, 104, 101, 109, 101, 67, 108, 114,
            32, 118, 97, 108, 61, 34, 112, 104, 67, 108,
            114, 34, 62, 60, 97, 58, 115, 104, 97, 100,
            101, 32, 118, 97, 108, 61, 34, 57, 53, 48,
            48, 48, 34, 47, 62, 60, 97, 58, 115, 97,
            116, 77, 111, 100, 32, 118, 97, 108, 61, 34,
            49, 48, 53, 48, 48, 48, 34, 47, 62, 60,
            47, 97, 58, 115, 99, 104, 101, 109, 101, 67,
            108, 114, 62, 60, 47, 97, 58, 115, 111, 108,
            105, 100, 70, 105, 108, 108, 62, 60, 97, 58,
            112, 114, 115, 116, 68, 97, 115, 104, 32, 118,
            97, 108, 61, 34, 115, 111, 108, 105, 100, 34,
            47, 62, 60, 47, 97, 58, 108, 110, 62, 60,
            97, 58, 108, 110, 32, 119, 61, 34, 50, 53,
            52, 48, 48, 34, 32, 99, 97, 112, 61, 34,
            102, 108, 97, 116, 34, 32, 99, 109, 112, 100,
            61, 34, 115, 110, 103, 34, 32, 97, 108, 103,
            110, 61, 34, 99, 116, 114, 34, 62, 60, 97,
            58, 115, 111, 108, 105, 100, 70, 105, 108, 108,
            62, 60, 97, 58, 115, 99, 104, 101, 109, 101,
            67, 108, 114, 32, 118, 97, 108, 61, 34, 112,
            104, 67, 108, 114, 34, 47, 62, 60, 47, 97,
            58, 115, 111, 108, 105, 100, 70, 105, 108, 108,
            62, 60, 97, 58, 112, 114, 115, 116, 68, 97,
            115, 104, 32, 118, 97, 108, 61, 34, 115, 111,
            108, 105, 100, 34, 47, 62, 60, 47, 97, 58,
            108, 110, 62, 60, 97, 58, 108, 110, 32, 119,
            61, 34, 51, 56, 49, 48, 48, 34, 32, 99,
            97, 112, 61, 34, 102, 108, 97, 116, 34, 32,
            99, 109, 112, 100, 61, 34, 115, 110, 103, 34,
            32, 97, 108, 103, 110, 61, 34, 99, 116, 114,
            34, 62, 60, 97, 58, 115, 111, 108, 105, 100,
            70, 105, 108, 108, 62, 60, 97, 58, 115, 99,
            104, 101, 109, 101, 67, 108, 114, 32, 118, 97,
            108, 61, 34, 112, 104, 67, 108, 114, 34, 47,
            62, 60, 47, 97, 58, 115, 111, 108, 105, 100,
            70, 105, 108, 108, 62, 60, 97, 58, 112, 114,
            115, 116, 68, 97, 115, 104, 32, 118, 97, 108,
            61, 34, 115, 111, 108, 105, 100, 34, 47, 62,
            60, 47, 97, 58, 108, 110, 62, 60, 47, 97,
            58, 108, 110, 83, 116, 121, 108, 101, 76, 115,
            116, 62, 60, 97, 58, 101, 102, 102, 101, 99,
            116, 83, 116, 121, 108, 101, 76, 115, 116, 62,
            60, 97, 58, 101, 102, 102, 101, 99, 116, 83,
            116, 121, 108, 101, 62, 60, 97, 58, 101, 102,
            102, 101, 99, 116, 76, 115, 116, 62, 60, 97,
            58, 111, 117, 116, 101, 114, 83, 104, 100, 119,
            32, 98, 108, 117, 114, 82, 97, 100, 61, 34,
            52, 48, 48, 48, 48, 34, 32, 100, 105, 115,
            116, 61, 34, 50, 48, 48, 48, 48, 34, 32,
            100, 105, 114, 61, 34, 53, 52, 48, 48, 48,
            48, 48, 34, 32, 114, 111, 116, 87, 105, 116,
            104, 83, 104, 97, 112, 101, 61, 34, 48, 34,
            62, 60, 97, 58, 115, 114, 103, 98, 67, 108,
            114, 32, 118, 97, 108, 61, 34, 48, 48, 48,
            48, 48, 48, 34, 62, 60, 97, 58, 97, 108,
            112, 104, 97, 32, 118, 97, 108, 61, 34, 51,
            56, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 115, 114, 103, 98, 67, 108, 114, 62, 60,
            47, 97, 58, 111, 117, 116, 101, 114, 83, 104,
            100, 119, 62, 60, 47, 97, 58, 101, 102, 102,
            101, 99, 116, 76, 115, 116, 62, 60, 47, 97,
            58, 101, 102, 102, 101, 99, 116, 83, 116, 121,
            108, 101, 62, 60, 97, 58, 101, 102, 102, 101,
            99, 116, 83, 116, 121, 108, 101, 62, 60, 97,
            58, 101, 102, 102, 101, 99, 116, 76, 115, 116,
            62, 60, 97, 58, 111, 117, 116, 101, 114, 83,
            104, 100, 119, 32, 98, 108, 117, 114, 82, 97,
            100, 61, 34, 52, 48, 48, 48, 48, 34, 32,
            100, 105, 115, 116, 61, 34, 50, 51, 48, 48,
            48, 34, 32, 100, 105, 114, 61, 34, 53, 52,
            48, 48, 48, 48, 48, 34, 32, 114, 111, 116,
            87, 105, 116, 104, 83, 104, 97, 112, 101, 61,
            34, 48, 34, 62, 60, 97, 58, 115, 114, 103,
            98, 67, 108, 114, 32, 118, 97, 108, 61, 34,
            48, 48, 48, 48, 48, 48, 34, 62, 60, 97,
            58, 97, 108, 112, 104, 97, 32, 118, 97, 108,
            61, 34, 51, 53, 48, 48, 48, 34, 47, 62,
            60, 47, 97, 58, 115, 114, 103, 98, 67, 108,
            114, 62, 60, 47, 97, 58, 111, 117, 116, 101,
            114, 83, 104, 100, 119, 62, 60, 47, 97, 58,
            101, 102, 102, 101, 99, 116, 76, 115, 116, 62,
            60, 47, 97, 58, 101, 102, 102, 101, 99, 116,
            83, 116, 121, 108, 101, 62, 60, 97, 58, 101,
            102, 102, 101, 99, 116, 83, 116, 121, 108, 101,
            62, 60, 97, 58, 101, 102, 102, 101, 99, 116,
            76, 115, 116, 62, 60, 97, 58, 111, 117, 116,
            101, 114, 83, 104, 100, 119, 32, 98, 108, 117,
            114, 82, 97, 100, 61, 34, 52, 48, 48, 48,
            48, 34, 32, 100, 105, 115, 116, 61, 34, 50,
            51, 48, 48, 48, 34, 32, 100, 105, 114, 61,
            34, 53, 52, 48, 48, 48, 48, 48, 34, 32,
            114, 111, 116, 87, 105, 116, 104, 83, 104, 97,
            112, 101, 61, 34, 48, 34, 62, 60, 97, 58,
            115, 114, 103, 98, 67, 108, 114, 32, 118, 97,
            108, 61, 34, 48, 48, 48, 48, 48, 48, 34,
            62, 60, 97, 58, 97, 108, 112, 104, 97, 32,
            118, 97, 108, 61, 34, 51, 53, 48, 48, 48,
            34, 47, 62, 60, 47, 97, 58, 115, 114, 103,
            98, 67, 108, 114, 62, 60, 47, 97, 58, 111,
            117, 116, 101, 114, 83, 104, 100, 119, 62, 60,
            47, 97, 58, 101, 102, 102, 101, 99, 116, 76,
            115, 116, 62, 60, 97, 58, 115, 99, 101, 110,
            101, 51, 100, 62, 60, 97, 58, 99, 97, 109,
            101, 114, 97, 32, 112, 114, 115, 116, 61, 34,
            111, 114, 116, 104, 111, 103, 114, 97, 112, 104,
            105, 99, 70, 114, 111, 110, 116, 34, 62, 60,
            97, 58, 114, 111, 116, 32, 108, 97, 116, 61,
            34, 48, 34, 32, 108, 111, 110, 61, 34, 48,
            34, 32, 114, 101, 118, 61, 34, 48, 34, 47,
            62, 60, 47, 97, 58, 99, 97, 109, 101, 114,
            97, 62, 60, 97, 58, 108, 105, 103, 104, 116,
            82, 105, 103, 32, 114, 105, 103, 61, 34, 116,
            104, 114, 101, 101, 80, 116, 34, 32, 100, 105,
            114, 61, 34, 116, 34, 62, 60, 97, 58, 114,
            111, 116, 32, 108, 97, 116, 61, 34, 48, 34,
            32, 108, 111, 110, 61, 34, 48, 34, 32, 114,
            101, 118, 61, 34, 49, 50, 48, 48, 48, 48,
            48, 34, 47, 62, 60, 47, 97, 58, 108, 105,
            103, 104, 116, 82, 105, 103, 62, 60, 47, 97,
            58, 115, 99, 101, 110, 101, 51, 100, 62, 60,
            97, 58, 115, 112, 51, 100, 62, 60, 97, 58,
            98, 101, 118, 101, 108, 84, 32, 119, 61, 34,
            54, 51, 53, 48, 48, 34, 32, 104, 61, 34,
            50, 53, 52, 48, 48, 34, 47, 62, 60, 47,
            97, 58, 115, 112, 51, 100, 62, 60, 47, 97,
            58, 101, 102, 102, 101, 99, 116, 83, 116, 121,
            108, 101, 62, 60, 47, 97, 58, 101, 102, 102,
            101, 99, 116, 83, 116, 121, 108, 101, 76, 115,
            116, 62, 60, 97, 58, 98, 103, 70, 105, 108,
            108, 83, 116, 121, 108, 101, 76, 115, 116, 62,
            60, 97, 58, 115, 111, 108, 105, 100, 70, 105,
            108, 108, 62, 60, 97, 58, 115, 99, 104, 101,
            109, 101, 67, 108, 114, 32, 118, 97, 108, 61,
            34, 112, 104, 67, 108, 114, 34, 47, 62, 60,
            47, 97, 58, 115, 111, 108, 105, 100, 70, 105,
            108, 108, 62, 60, 97, 58, 103, 114, 97, 100,
            70, 105, 108, 108, 32, 114, 111, 116, 87, 105,
            116, 104, 83, 104, 97, 112, 101, 61, 34, 49,
            34, 62, 60, 97, 58, 103, 115, 76, 115, 116,
            62, 60, 97, 58, 103, 115, 32, 112, 111, 115,
            61, 34, 48, 34, 62, 60, 97, 58, 115, 99,
            104, 101, 109, 101, 67, 108, 114, 32, 118, 97,
            108, 61, 34, 112, 104, 67, 108, 114, 34, 62,
            60, 97, 58, 116, 105, 110, 116, 32, 118, 97,
            108, 61, 34, 52, 48, 48, 48, 48, 34, 47,
            62, 60, 97, 58, 115, 97, 116, 77, 111, 100,
            32, 118, 97, 108, 61, 34, 51, 53, 48, 48,
            48, 48, 34, 47, 62, 60, 47, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 62, 60,
            47, 97, 58, 103, 115, 62, 60, 97, 58, 103,
            115, 32, 112, 111, 115, 61, 34, 52, 48, 48,
            48, 48, 34, 62, 60, 97, 58, 115, 99, 104,
            101, 109, 101, 67, 108, 114, 32, 118, 97, 108,
            61, 34, 112, 104, 67, 108, 114, 34, 62, 60,
            97, 58, 116, 105, 110, 116, 32, 118, 97, 108,
            61, 34, 52, 53, 48, 48, 48, 34, 47, 62,
            60, 97, 58, 115, 104, 97, 100, 101, 32, 118,
            97, 108, 61, 34, 57, 57, 48, 48, 48, 34,
            47, 62, 60, 97, 58, 115, 97, 116, 77, 111,
            100, 32, 118, 97, 108, 61, 34, 51, 53, 48,
            48, 48, 48, 34, 47, 62, 60, 47, 97, 58,
            115, 99, 104, 101, 109, 101, 67, 108, 114, 62,
            60, 47, 97, 58, 103, 115, 62, 60, 97, 58,
            103, 115, 32, 112, 111, 115, 61, 34, 49, 48,
            48, 48, 48, 48, 34, 62, 60, 97, 58, 115,
            99, 104, 101, 109, 101, 67, 108, 114, 32, 118,
            97, 108, 61, 34, 112, 104, 67, 108, 114, 34,
            62, 60, 97, 58, 115, 104, 97, 100, 101, 32,
            118, 97, 108, 61, 34, 50, 48, 48, 48, 48,
            34, 47, 62, 60, 97, 58, 115, 97, 116, 77,
            111, 100, 32, 118, 97, 108, 61, 34, 50, 53,
            53, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 115, 99, 104, 101, 109, 101, 67, 108, 114,
            62, 60, 47, 97, 58, 103, 115, 62, 60, 47,
            97, 58, 103, 115, 76, 115, 116, 62, 60, 97,
            58, 112, 97, 116, 104, 32, 112, 97, 116, 104,
            61, 34, 99, 105, 114, 99, 108, 101, 34, 62,
            60, 97, 58, 102, 105, 108, 108, 84, 111, 82,
            101, 99, 116, 32, 108, 61, 34, 53, 48, 48,
            48, 48, 34, 32, 116, 61, 34, 45, 56, 48,
            48, 48, 48, 34, 32, 114, 61, 34, 53, 48,
            48, 48, 48, 34, 32, 98, 61, 34, 49, 56,
            48, 48, 48, 48, 34, 47, 62, 60, 47, 97,
            58, 112, 97, 116, 104, 62, 60, 47, 97, 58,
            103, 114, 97, 100, 70, 105, 108, 108, 62, 60,
            97, 58, 103, 114, 97, 100, 70, 105, 108, 108,
            32, 114, 111, 116, 87, 105, 116, 104, 83, 104,
            97, 112, 101, 61, 34, 49, 34, 62, 60, 97,
            58, 103, 115, 76, 115, 116, 62, 60, 97, 58,
            103, 115, 32, 112, 111, 115, 61, 34, 48, 34,
            62, 60, 97, 58, 115, 99, 104, 101, 109, 101,
            67, 108, 114, 32, 118, 97, 108, 61, 34, 112,
            104, 67, 108, 114, 34, 62, 60, 97, 58, 116,
            105, 110, 116, 32, 118, 97, 108, 61, 34, 56,
            48, 48, 48, 48, 34, 47, 62, 60, 97, 58,
            115, 97, 116, 77, 111, 100, 32, 118, 97, 108,
            61, 34, 51, 48, 48, 48, 48, 48, 34, 47,
            62, 60, 47, 97, 58, 115, 99, 104, 101, 109,
            101, 67, 108, 114, 62, 60, 47, 97, 58, 103,
            115, 62, 60, 97, 58, 103, 115, 32, 112, 111,
            115, 61, 34, 49, 48, 48, 48, 48, 48, 34,
            62, 60, 97, 58, 115, 99, 104, 101, 109, 101,
            67, 108, 114, 32, 118, 97, 108, 61, 34, 112,
            104, 67, 108, 114, 34, 62, 60, 97, 58, 115,
            104, 97, 100, 101, 32, 118, 97, 108, 61, 34,
            51, 48, 48, 48, 48, 34, 47, 62, 60, 97,
            58, 115, 97, 116, 77, 111, 100, 32, 118, 97,
            108, 61, 34, 50, 48, 48, 48, 48, 48, 34,
            47, 62, 60, 47, 97, 58, 115, 99, 104, 101,
            109, 101, 67, 108, 114, 62, 60, 47, 97, 58,
            103, 115, 62, 60, 47, 97, 58, 103, 115, 76,
            115, 116, 62, 60, 97, 58, 112, 97, 116, 104,
            32, 112, 97, 116, 104, 61, 34, 99, 105, 114,
            99, 108, 101, 34, 62, 60, 97, 58, 102, 105,
            108, 108, 84, 111, 82, 101, 99, 116, 32, 108,
            61, 34, 53, 48, 48, 48, 48, 34, 32, 116,
            61, 34, 53, 48, 48, 48, 48, 34, 32, 114,
            61, 34, 53, 48, 48, 48, 48, 34, 32, 98,
            61, 34, 53, 48, 48, 48, 48, 34, 47, 62,
            60, 47, 97, 58, 112, 97, 116, 104, 62, 60,
            47, 97, 58, 103, 114, 97, 100, 70, 105, 108,
            108, 62, 60, 47, 97, 58, 98, 103, 70, 105,
            108, 108, 83, 116, 121, 108, 101, 76, 115, 116,
            62, 60, 47, 97, 58, 102, 109, 116, 83, 99,
            104, 101, 109, 101, 62, 60, 47, 97, 58, 116,
            104, 101, 109, 101, 69, 108, 101, 109, 101, 110,
            116, 115, 62, 60, 97, 58, 111, 98, 106, 101,
            99, 116, 68, 101, 102, 97, 117, 108, 116, 115,
            47, 62, 60, 97, 58, 101, 120, 116, 114, 97,
            67, 108, 114, 83, 99, 104, 101, 109, 101, 76,
            115, 116, 47, 62, 60, 47, 97, 58, 116, 104,
            101, 109, 101, 62
        };

        return std::string(data.begin(), data.end());;
    }

}
