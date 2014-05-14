#pragma once

#include <array>
#include <cstdio>
#include <string>

class TemporaryFile
{
public:
    static std::string CreateTemporaryFilename()
    {
        return "C:/Users/taf656/Desktop/xlnt.xlsx";
    }

    TemporaryFile() : filename_(CreateTemporaryFilename())
    {

    }

    ~TemporaryFile()
    {
        remove(filename_.c_str());
    }

    std::string GetFilename() const { return filename_; }

private:
    const std::string filename_;
};