#include <array>
#include <fstream>

#include "common/zip_file.hpp"

namespace xlnt {

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
    
}
