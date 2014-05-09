#include <algorithm>
#include <array>
#include <fstream>
#include <iostream>
#include <sstream>

#include "xlnt.h"

namespace xlnt {

struct part_struct
{
    part_struct(package_impl &package, const std::string &uri_part, const std::string &mime_type = "", compression_option compression = compression_option::NotCompressed)
        : package_(package),
        uri_(uri_part),
        content_type_(mime_type),
        compression_option_(compression)
    {}

    part_struct(package_impl &package, const std::string &uri, opcContainer *container)
        : package_(package),
        uri_(uri),
        container_(container)
    {}

    relationship create_relationship(const std::string &target_uri, target_mode target_mode, const std::string &relationship_type);

    void delete_relationship(const std::string &id);

    relationship get_relationship(const std::string &id);

    relationship_collection get_relationships();

    relationship_collection get_relationship_by_type(const std::string &relationship_type);

    std::string read()
    {
        std::string ss;
        auto part_stream = opcContainerOpenInputStream(container_, (xmlChar*)uri_.c_str());
        std::array<xmlChar, 1024> buffer;
        auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
        if(bytes_read > 0)
        {
            ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            while(bytes_read == buffer.size())
            {
                auto bytes_read = opcContainerReadInputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(buffer.size()));
                ss.append(std::string(buffer.begin(), buffer.begin() + bytes_read));
            }
        }
        opcContainerCloseInputStream(part_stream);
        return ss;
    }

    void write(const std::string &data)
    {
        auto name = uri_;
        auto name_pointer = name.c_str();

        auto part_stream = opcContainerCreateOutputStream(container_, (xmlChar*)name_pointer, opcCompressionOption_t::OPC_COMPRESSIONOPTION_NORMAL);

        std::stringstream ss(data);
        std::array<xmlChar, 1024> buffer;

        while(ss)
        {
            ss.get((char*)buffer.data(), 1024);
            auto count = ss.gcount();
            if(count > 0)
            {
                opcContainerWriteOutputStream(part_stream, buffer.data(), static_cast<opc_uint32_t>(count));
            }
        }
        opcContainerCloseOutputStream(part_stream);
    }

    /// <summary>
    /// Returns a value that indicates whether this part owns a relationship with a specified Id.
    /// </summary>
    bool relationship_exists(const std::string &id) const;

    /// <summary>
    /// Returns true if the given Id string is a valid relationship identifier.
    /// </summary>
    bool is_valid_xml_id(const std::string &id);

    void operator=(const part_struct &other);

    compression_option compression_option_;
    std::string content_type_;
    package_impl &package_;
    std::string uri_;
    opcContainer *container_;
};

part::part(package_impl &package, const std::string &uri, opcContainer *container) 
    : root_(new part_struct(package, uri, container))
{

}

std::string part::get_content_type() const
{
    return "";
}

std::string part::read()
{
    if(root_ == nullptr)
    {
        return "";
    }

    return root_->read();
}

void part::write(const std::string &data)
{
    if(root_ == nullptr)
    {
        return;
    }

    return root_->write(data);
}

bool part::operator==(const part &comparand) const
{
    return root_ == comparand.root_;
}

bool part::operator==(const nullptr_t &) const
{
    return root_ == nullptr;
}

struct package_impl
{
    package *parent_;
    opcContainer *opc_container_;
    std::iostream &stream_;
    std::fstream file_stream_;
    file_mode package_mode_;
    file_access package_access_;
    std::vector<xmlChar> container_buffer_;
    bool open_;
    bool streaming_;

    file_access get_file_open_access() const
    {
        if(!open_)
        {
            throw std::runtime_error("The package is not open");
        }

        return package_access_;
    }

    package_impl(std::iostream &stream, file_mode package_mode, file_access package_access)
        : stream_(stream), container_buffer_(4096), package_mode_(package_mode), package_access_(package_access)
    {
        open_container();
    }

    package_impl(const std::string &path, file_mode package_mode, file_access package_access)
        : stream_(file_stream_), container_buffer_(4096), package_mode_(package_mode), package_access_(package_access)
    {
        switch(package_mode)
        {
        case file_mode::Append:
            switch(package_access)
            {
            case file_access::Read: throw std::runtime_error("Append can only be used with file_access.Write");
            case file_access::ReadWrite: throw std::runtime_error("Append can only be used with file_access.Write");
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::app | std::ios::out);
                break;
            }
            break;
        case file_mode::Create:
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            }
            break;
        case file_mode::CreateNew:
            if(!file::exists(path))
            {
                throw std::runtime_error("File already exists");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            }
            break;
        case file_mode::Open:
            if(!file::exists(path))
            {
                throw std::runtime_error("Can't open non-existent file");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            }
            break;
        case file_mode::OpenOrCreate:
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::out);
                break;
            }
            break;
        case file_mode::Truncate:
            if(!file::exists(path))
            {
                throw std::runtime_error("Can't truncate non-existent file");
            }
            switch(package_access)
            {
            case file_access::Read:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::in);
                break;
            case file_access::ReadWrite:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::in | std::ios::out);
                break;
            case file_access::Write:
                file_stream_.open(path, std::ios::binary | std::ios::trunc | std::ios::out);
                break;
            }
            break;
        }

        open_container();
    }

    void open_container()
    {
        opcContainerOpenMode m;

        switch(package_access_)
        {
        case file_access::Read:
            m = opcContainerOpenMode::OPC_OPEN_READ_ONLY;
            break;
        case file_access::ReadWrite:
            m = opcContainerOpenMode::OPC_OPEN_READ_WRITE;
            break;
        case file_access::Write:
            m = opcContainerOpenMode::OPC_OPEN_WRITE_ONLY;
            break;
        default:
            throw std::runtime_error("unknown file access");
        }

        opc_container_ = opcContainerOpenIO(&read_callback, &write_callback,
            &close_callback, &seek_callback,
            &trim_callback, &flush_callback, this, 4096, m, this);
        open_ = true;
    }

    ~package_impl()
    {
        close();
    }

    void close()
    {
        if(open_)
        {
            open_ = false;
            opcContainerClose(opc_container_, opcContainerCloseMode::OPC_CLOSE_NOW);
            opc_container_ = nullptr;
        }
    }

    part create_part(const std::string &part_uri, const std::string &/*content_type*/, compression_option /*compression*/)
    {
        return part(new part_struct(*this, part_uri, opc_container_));
    }

    void delete_part(const std::string &/*part_uri*/)
    {

    }

    void flush()
    {
        stream_.flush();
    }

    part get_part(const std::string &part_uri)
    {
        return part(*this, part_uri, opc_container_);
    }

    part_collection get_parts()
    {
        return part_collection();
    }

    int write(char *buffer, int length)
    {
        stream_.read(buffer, length);
        auto bytes_read = stream_.gcount();
        return static_cast<int>(bytes_read);
    }

    int read(const char *buffer, int length)
    {
        auto before = stream_.tellp();
        stream_.write(buffer, length);
        auto bytes_written = stream_.tellp() - before;
        return static_cast<int>(bytes_written);
    }

    std::ios::pos_type seek(std::ios::pos_type offset)
    {
        stream_.clear();
        stream_.seekg(offset);
        auto current_position = stream_.tellg();
        if(stream_.eof())
        {
            std::cout << "eof" << std::endl;
        }
        if(stream_.fail())
        {
            std::cout << "fail" << std::endl;
        }
        if(stream_.bad())
        {
            std::cout << "bad" << std::endl;
        }
        return current_position;
    }

    int trim(std::ios::pos_type /*new_size*/)
    {
        return 0;
    }

    static int read_callback(void *context, char *buffer, int length)
    {
        auto object = static_cast<package_impl *>(context);
        return object->write(buffer, length);
    }

    static int write_callback(void *context, const char *buffer, int length)
    {
        auto object = static_cast<package_impl *>(context);
        return object->read(buffer, length);
    }

    static int close_callback(void *context)
    {
        auto object = static_cast<package_impl *>(context);
        object->close();
        return 0;
    }

    static opc_ofs_t seek_callback(void *context, opc_ofs_t ofs)
    {
        auto object = static_cast<package_impl *>(context);
        return static_cast<opc_ofs_t>(object->seek(ofs));
    }

    static int trim_callback(void *context, opc_ofs_t new_size)
    {
        auto object = static_cast<package_impl *>(context);
        return object->trim(new_size);
    }

    static int flush_callback(void *context)
    {
        auto object = static_cast<package_impl *>(context);
        object->flush();
        return 0;
    }
};

file_access package::get_file_open_access() const
{
    return impl_->get_file_open_access();
}

package package::open(std::iostream &stream, file_mode package_mode, file_access package_access)
{
    return package(stream, package_mode, package_access);
}

package package::open(const std::string &path, file_mode package_mode, file_access package_access, file_share /*package_share*/)
{
    return package(path, package_mode, package_access);
}

package::package(std::iostream &stream, file_mode package_mode, file_access package_access)
    : impl_(new package_impl(stream, package_mode, package_access))
{
    open_container();
}

package::package(const std::string &path, file_mode package_mode, file_access package_access)
    : impl_(new package_impl(path, package_mode, package_access))
{

}

void package::open_container()
{
    impl_->open_container();
}

void package::close()
{
    impl_->close();
}

part package::create_part(const std::string &part_uri, const std::string &content_type, compression_option compression)
{
    return impl_->create_part(part_uri, content_type, compression);
}

void package::delete_part(const std::string &part_uri)
{
    impl_->delete_part(part_uri);
}

void package::flush()
{
    impl_->flush();
}

part package::get_part(const std::string &part_uri)
{
    return impl_->get_part(part_uri);
}

part_collection package::get_parts()
{
    return impl_->get_parts();
}

int package::write(char *buffer, int length)
{
    return impl_->write(buffer, length);
}

int package::read(const char *buffer, int length)
{
    return impl_->read(buffer, length);
}

std::ios::pos_type package::seek(std::ios::pos_type offset)
{
    return impl_->seek(offset);
}

int package::trim(std::ios::pos_type new_size)
{
    return impl_->trim(new_size);
}

bool package::operator==(const package &comparand) const
{
    return impl_ == comparand.impl_;
}

bool package::operator==(const nullptr_t &) const
{
    return impl_ == nullptr;
}


struct cell_struct
{
    cell_struct(worksheet_struct *ws, const std::string &column, int row)
        : type(cell::type::null), parent_worksheet(ws),
        column(xlnt::cell::column_index_from_string(column) - 1), row(row)
    {

    }

    std::string to_string() const
    {
        switch(type)
        {
        case cell::type::numeric: return std::to_string(numeric_value);
        case cell::type::boolean: return bool_value ? "true" : "false";
        case cell::type::string: return string_value;
        default: throw std::runtime_error("bad enum");
        }
    }

    cell::type type;

    union
    {
        long double numeric_value;
        bool bool_value;
    };

    tm date_value;
    std::string string_value;
    worksheet_struct *parent_worksheet;
    int column;
    int row;
};

cell::cell() : root_(nullptr)
{
}

cell::cell(cell_struct *root) : root_(root)
{
}

coordinate cell::coordinate_from_string(const std::string &address)
{
    if(address == "A1")
    {
        return{"A", 1};
    }

    return{"A", 1};
}

int cell::column_index_from_string(const std::string &column_string)
{
    return column_string[0] - 'A';
}

std::string cell::get_column_letter(int column_index)
{
    return std::string(1, (char)('A' + column_index - 1));
}

bool cell::is_date() const
{
    return root_->type == type::date;
}

bool operator==(const char *comparand, const cell &cell)
{
    return std::string(comparand) == cell;
}

bool operator==(const std::string &comparand, const cell &cell)
{
    return cell.root_->type == cell::type::string && cell.root_->string_value == comparand;
}

bool operator==(const tm &comparand, const cell &cell)
{
    return cell.root_->type == cell::type::date && cell.root_->date_value.tm_hour == comparand.tm_hour;
}

std::string cell::absolute_coordinate(const std::string &absolute_address)
{
    return absolute_address;
}

cell::type cell::get_data_type()
{
    return root_->type;
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

cell &cell::operator=(const std::string &value)
{
    root_->type = type::string;
    root_->string_value = value;
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

cell::~cell()
{
    delete root_;
}

std::string cell::to_string() const
{
    return root_->to_string();
}


namespace {

    class not_implemented : public std::runtime_error
    {
    public:
        not_implemented() : std::runtime_error("error: not implemented")
        {

        }
    };

    // Convert a column number into a column letter (3 -> 'C')
    // Right shift the column col_idx by 26 to find column letters in reverse
    // order.These numbers are 1 - based, and can be converted to ASCII
    // ordinals by adding 64.
    std::string get_column_letter(int column_index)
    {
        // these indicies corrospond to A->ZZZ and include all allowed
        // columns
        if(column_index < 1 || column_index > 18278)
        {
            auto msg = "Column index out of bounds: " + column_index;
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

}

struct worksheet_struct
{
    worksheet_struct(workbook &parent_workbook, const std::string &title)
        : parent_(parent_workbook), title_(title), freeze_panes_(nullptr)
    {

    }

    void garbage_collect()
    {
        throw not_implemented();
    }

    std::set<cell> get_cell_collection()
    {
        throw not_implemented();
    }

    std::string get_title() const
    {
        throw not_implemented();
    }

    void set_title(const std::string &title)
    {
        title_ = title;
    }

    cell get_freeze_panes() const
    {
        throw not_implemented();
    }

    void set_freeze_panes(cell top_left_cell)
    {
        throw not_implemented();
    }

    void set_freeze_panes(const std::string &top_left_coordinate)
    {
        freeze_panes_ = cell(top_left_coordinate);
    }

    void unfreeze_panes()
    {
        throw not_implemented();
    }

    xlnt::cell cell(const std::string &coordinate)
    {
        if(cell_map_.find(coordinate) == cell_map_.end())
        {
            auto coord = xlnt::cell::coordinate_from_string(coordinate);
            cell_struct *cell = new xlnt::cell_struct(this, coord.column, coord.row);
            cell_map_[coordinate] = xlnt::cell(cell);
        }

        return cell_map_[coordinate];
    }

    xlnt::cell cell(int row, int column)
    {
        return cell(get_column_letter(column + 1) + std::to_string(row + 1));
    }

    int get_highest_row() const
    {
        throw not_implemented();
    }

    int get_highest_column() const
    {
        throw not_implemented();
    }

    std::string calculate_dimension() const
    {
        throw not_implemented();
    }

    range range(const std::string &range_string, int /*row_offset*/, int /*column_offset*/)
    {
        auto colon_index = range_string.find(':');
        if(colon_index != std::string::npos)
        {
            auto min_range = range_string.substr(0, colon_index);
            auto max_range = range_string.substr(colon_index + 1);
            xlnt::range r;
            r.push_back(std::vector<xlnt::cell>());
            r[0].push_back(cell(min_range));
            r[0].push_back(cell(max_range));
            return r;
        }
        throw not_implemented();
    }

    relationship create_relationship(const std::string &relationship_type)
    {
        relationships_.push_back(relationship(relationship_type));
        return relationships_.back();
    }

    //void add_chart(chart chart);

    void merge_cells(const std::string &/*range_string*/)
    {
        throw not_implemented();
    }

    void merge_cells(int /*start_row*/, int /*start_column*/, int /*end_row*/, int /*end_column*/)
    {
        throw not_implemented();
    }

    void unmerge_cells(const std::string &/*range_string*/)
    {
        throw not_implemented();
    }

    void unmerge_cells(int /*start_row*/, int /*start_column*/, int /*end_row*/, int /*end_column*/)
    {
        throw not_implemented();
    }

    void append(const std::vector<xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cell_map_["a"] = cell;
        }
    }

    void append(const std::unordered_map<std::string, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cell_map_[cell.first] = cell.second;
        }
    }

    void append(const std::unordered_map<int, xlnt::cell> &cells)
    {
        for(auto cell : cells)
        {
            cell_map_[get_column_letter(cell.first + 1)] = cell.second;
        }
    }

    xlnt::range rows() const
    {
        throw not_implemented();
    }

    xlnt::range columns() const
    {
        throw not_implemented();
    }

    void operator=(const worksheet_struct &other) = delete;

    workbook &parent_;
    std::string title_;
    xlnt::cell freeze_panes_;
    std::unordered_map<std::string, xlnt::cell> cell_map_;
    std::vector<relationship> relationships_;
};

worksheet::worksheet(worksheet_struct *root) : root_(root)
{
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

std::set<cell> worksheet::get_cell_collection()
{
    return root_->get_cell_collection();
}

std::string worksheet::get_title() const
{
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

relationship worksheet::create_relationship(const std::string &relationship_type)
{
    return root_->create_relationship(relationship_type);
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

void worksheet::append(const std::vector<xlnt::cell> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<std::string, xlnt::cell> &cells)
{
    root_->append(cells);
}

void worksheet::append(const std::unordered_map<int, xlnt::cell> &cells)
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

bool worksheet::operator==(nullptr_t) const
{
    return root_ == nullptr;
}

bool worksheet::operator!=(nullptr_t) const
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

workbook::workbook() : active_worksheet_(nullptr)
{
    auto ws = create_sheet();
    ws.set_title("Sheet1");
    active_worksheet_ = ws;
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

worksheet workbook::get_active()
{
    return active_worksheet_;
}

worksheet workbook::create_sheet()
{
    std::string title = "Sheet1";
    int index = 1;
    worksheet current = get_sheet_by_name(title);
    while(get_sheet_by_name(title) != nullptr)
    {
        title = "Sheet" + std::to_string(++index);
    }
    auto *worksheet = new worksheet_struct(*this, title);
    worksheets_.push_back(worksheet);
    return get_sheet_by_name(title);
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
    return get_sheet_by_name(name);
}

void workbook::save(const std::string &filename)
{
    auto package = package::open(filename);
    package.close();
}

}