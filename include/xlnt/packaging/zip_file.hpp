#pragma once

#include <cstdint>
#include <iostream>
#include <memory>
#include <sstream>
#include <vector>

#include <xlnt/utils/string.hpp>

#include <xlnt/xlnt_config.hpp>

// Note: this comes from https://github.com/tfussell/miniz-cpp

struct mz_zip_archive_tag;

namespace xlnt {

/// <summary>
/// Information about a specific file in zip_file.
/// </summary>
struct XLNT_CLASS zip_info
{
    /// <summary>
    /// A struct representing a particular date and time.
    /// </summary>
    struct date_time_t
    {
        int year;
        int month;
        int day;
        int hours;
        int minutes;
        int seconds;
    };

    /// <summary>
    /// Default constructor for zip_info.
    /// </summary>
    zip_info();

    date_time_t date_time;
    string filename;
    string comment;
    string extra;
    uint16_t create_system;
    uint16_t create_version;
    uint16_t extract_version;
    uint16_t flag_bits;
    std::size_t volume;
    uint32_t internal_attr;
    uint32_t external_attr;
    std::size_t header_offset;
    uint32_t crc;
    std::size_t compress_size;
    std::size_t file_size;
};

/// <summary>
/// A compressed archive file that exists in memory which can read
/// or write to and from the filesystem, std::iostreams, and byte vectors.
/// </summary>
class XLNT_CLASS zip_file
{
  public:
    zip_file();
    zip_file(const string &filename);
    zip_file(const std::vector<unsigned char> &bytes);
    zip_file(std::istream &stream);
    ~zip_file();

    // to/from file
    void load(const string &filename);
    void save(const string &filename);

    // to/from byte vector
    void load(const std::vector<unsigned char> &bytes);
    void save(std::vector<unsigned char> &bytes);

    // to/from iostream
    void load(std::istream &stream);
    void save(std::ostream &stream);

    void reset();

    bool has_file(const string &name);
    bool has_file(const zip_info &name);

    zip_info getinfo(const string &name);

    std::vector<zip_info> infolist();
    std::vector<string> namelist();

    std::ostream &open(const string &name);
    std::ostream &open(const zip_info &name);

    void extract(const string &name);
    void extract(const string &name, const string &path);
    void extract(const zip_info &name);
    void extract(const zip_info &name, const string &path);

    void extractall();
    void extractall(const string &path);
    void extractall(const string &path, const std::vector<string> &members);
    void extractall(const string &path, const std::vector<zip_info> &members);

    void printdir();
    void printdir(std::ostream &stream);

    string read(const string &name);
    string read(const zip_info &name);

    std::pair<bool, string> testzip();

    void write(const string &filename);
    void write(const string &filename, const string &arcname);

    void writestr(const string &arcname, const string &bytes);
    void writestr(const zip_info &arcname, const string &bytes);

    string get_filename() const
    {
        return filename_;
    }

    string comment;

  private:
    void start_read();
    void start_write();

    void append_comment();
    void remove_comment();

    zip_info getinfo(int index);

    std::unique_ptr<mz_zip_archive_tag> archive_;
    std::vector<char> buffer_;
    std::stringstream open_stream_;
    string filename_;
};

} // namespace xlnt
