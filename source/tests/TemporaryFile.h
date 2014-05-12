#pragma once

#include <array>
#include <cstdio>
#include <string>

class TemporaryFile
{
public:
    static std::string CreateTemporaryFilename()
    {
        std::array<char, L_tmpnam> buffer;
        tmpnam(buffer.data());
        return std::string(buffer.begin(), buffer.end());
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