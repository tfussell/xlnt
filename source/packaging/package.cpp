#include <array>
#include <functional>
#include <iostream>

#include "package.h"
#include "file.h"
#include "part.h"

namespace xlnt {

class package_impl
{
public:
    package *parent_;
    opcContainer *opc_container_;
    std::iostream &stream_;
    std::fstream file_stream_;
    file_mode package_mode_;
    file_access package_access_;
    std::vector<xmlChar> container_buffer_;
    bool open_;
    package_properties package_properties_;
    bool streaming_;

    file_access get_file_open_access() const
    {
        if(!open_)
        {
            throw std::runtime_error("The package is not open");
        }

        return package_access_;
    }

    package_properties get_package_properties() const
    {
        if(!open_)
        {
            throw std::runtime_error("The package is not open");
        }

        return package_properties_;
    }

    package_impl(package *parent, std::iostream &stream, file_mode package_mode, file_access package_access) 
        : parent_(parent), stream_(stream), container_buffer_(4096), package_mode_(package_mode), package_access_(package_access)
    {
        open_container();
    }

    package_impl(package *parent, const std::string &path, file_mode package_mode, file_access package_access) 
        : parent_(parent), stream_(file_stream_), container_buffer_(4096), package_mode_(package_mode), package_access_(package_access)
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

    part create_part(const uri &/*part_uri*/, const std::string &/*content_type*/, compression_option /*compression*/)
    {
        return part();
    }

    void delete_part(const uri &/*part_uri*/)
    {

    }

    void flush()
    {
        stream_.flush();
    }

    part get_part(const uri &part_uri)
    {
        return part(*parent_, part_uri, opc_container_);
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

package_properties package::get_package_properties() const
{
    return impl_->get_package_properties();
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
    : impl_(new package_impl(this, stream, package_mode, package_access))
{
    open_container();
}

package::package(const std::string &path, file_mode package_mode, file_access package_access) 
    : impl_(new package_impl(this, path, package_mode, package_access))
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

part package::create_part(const uri &part_uri, const std::string &content_type, compression_option compression)
{
    return impl_->create_part(part_uri, content_type, compression);
}

void package::delete_part(const uri &part_uri)
{
    impl_->delete_part(part_uri);
}

void package::flush()
{
    impl_->flush();
}

part package::get_part(const uri &part_uri)
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
    return impl_.get() == comparand.impl_.get();
}

bool package::operator==(const nullptr_t &) const
{
    return impl_.get() == nullptr;
}

} // namespace xlnt
