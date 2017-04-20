#pragma once

#include <array>
#include <cstdio>
#include <string>

#include <detail/external/include_windows.hpp>
#include <helpers/path_helper.hpp>

namespace {

static std::string create_temporary_filename()
{
    return "temp.xlsx";
}

} // namespace 

class temporary_file
{
public:
    temporary_file() : path_(create_temporary_filename())
    {
        if(path_.exists())
        {
            std::remove(path_.string().c_str());
        }
    }

    ~temporary_file()
    {
        std::remove(path_.string().c_str());
    }

    xlnt::path get_path() const { return path_; }

private:
    const xlnt::path path_;
};
