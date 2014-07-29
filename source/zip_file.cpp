#include <array>
#include <fstream>

#include "common/zip_file.hpp"
#include "common/exceptions.hpp"

namespace xlnt {

zip_file::zip_file(const std::string &filename, file_mode mode, file_access access)
  : current_state_(state::closed),
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

std::string zip_file::get_file_contents(const std::string &filename) const
{
    return files_.at(filename);
}

std::string dirname(const std::string &filename)
{
    auto last_sep_index = filename.find_last_of('/');

    if(last_sep_index != std::string::npos)
    {
        return filename.substr(0, last_sep_index + 1);
    }

    return "";
}

void zip_file::set_file_contents(const std::string &filename, const std::string &contents)
{
    if(!has_file(filename) || files_[filename] != contents)
    {
        modified_ = true;
    }

    auto dir = dirname(filename);

    while(dir != "")
    {
        auto matching_directory = std::find(directories_.begin(), directories_.end(), dir);

        if(matching_directory == directories_.end())
        {
            directories_.push_back(dir);
        }

        dir = dirname(dir.substr(0, dir.length() - 1));
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
    
    auto num_files = zip_file_.m_total_files;
    mz_uint i = 0;
    
    for(;i < num_files; i++)
    {
        mz_zip_archive_file_stat file_info;
        
        if(!mz_zip_reader_file_stat(&zip_file_, i, &file_info))
        {
            throw std::runtime_error("stat failed");
        }
        
        std::string current_filename(file_info.m_filename, file_info.m_filename + strlen(file_info.m_filename));
        
        if(mz_zip_reader_is_file_a_directory(&zip_file_, i))
        {
            directories_.push_back(current_filename);
            continue;
        }
        
        files_[current_filename] = read_from_zip(current_filename);
    }
}

void zip_file::write_all()
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }
    
    change_state(state::write, false);
    
    for(auto directory : directories_)
    {
        write_directory_to_zip(directory, true);
    }

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
    
    auto num_files = (std::size_t)mz_zip_reader_get_num_files(&zip_file_);
    mz_uint i = 0;
    
    for(;i < num_files; i++)
    {
        mz_zip_archive_file_stat file_info;
    
        if(!mz_zip_reader_file_stat(&zip_file_, i, &file_info))
        {
            throw std::runtime_error("stat failed");
        }
    
        std::string current_filename(file_info.m_filename, file_info.m_filename + strlen(file_info.m_filename));
        
        if(filename == current_filename)
        {
            break;
        }
    }
    
    if(i == num_files)
    {
        throw std::runtime_error("not found");
    }
    
    if(mz_zip_reader_is_file_a_directory(&zip_file_, i))
    {
        directories_.push_back(filename);
        return "";
    }
    
    std::size_t uncomp_size = 0;
    char archive_filename[260];
    std::fill(archive_filename, archive_filename + 260, '\0');
    std::copy(filename.begin(), filename.begin() + filename.length(), archive_filename);
    char *data = (char *)mz_zip_reader_extract_file_to_heap(&zip_file_, archive_filename, &uncomp_size, 0);
    
    if(data == nullptr)
    {
        throw std::runtime_error("extract failed");
    }
    
    std::string content(data, data + uncomp_size);
    mz_free(data);
    
    return content;
}

void zip_file::write_directory_to_zip(const std::string &name, bool append)
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }

    change_state(state::write, append);
    
    if(!mz_zip_writer_add_mem(&zip_file_, name.c_str(), nullptr, 0, MZ_BEST_COMPRESSION))
    {
        throw std::runtime_error("write directory failed");
    }
}

void zip_file::write_to_zip(const std::string &filename, const std::string &content, bool append)
{
    if(!((int)access_ & (int)file_access::write))
    {
        throw std::runtime_error("don't have write access");
    }
    
    change_state(state::write, append);
    
    auto status = mz_zip_writer_add_mem(&zip_file_, filename.c_str(), content.c_str(), content.length(), MZ_BEST_COMPRESSION);
    
    if(!status)
    {
        throw std::runtime_error("write failed");
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
    if(!mz_zip_reader_init_file(&zip_file_, filename_.c_str(), 0))
    {
        throw invalid_file_exception(filename_);
    }
}

void zip_file::stop_read()
{
    if(!mz_zip_reader_end(&zip_file_))
    {
        throw std::runtime_error("");
    }
}

void zip_file::start_write(bool append)
{
    if(append && !file_exists(filename_))
    {
        throw std::runtime_error("can't append to non-existent file");
    }
    
    if(!mz_zip_writer_init_file(&zip_file_, filename_.c_str(), 0))
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
    flush();
    
    if(!mz_zip_writer_finalize_archive(&zip_file_))
    {
        throw std::runtime_error("");
    }
    
    if(!mz_zip_writer_end(&zip_file_))
    {
        throw std::runtime_error("");
    }
}
    
}
