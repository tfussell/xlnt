#pragma once

#include <vector>

#include <xlnt/utils/string.hpp>

#include "xlnt_config.hpp"

namespace xlnt {

class XLNT_CLASS default_type
{
  public:
    default_type();
    default_type(const string &extension, const string &content_type);
    default_type(const default_type &other);
    default_type &operator=(const default_type &other);

    string get_extension() const
    {
        return extension_;
    }
    string get_content_type() const
    {
        return content_type_;
    }

  private:
    string extension_;
    string content_type_;
};

class XLNT_CLASS override_type
{
  public:
    override_type();
    override_type(const string &extension, const string &content_type);
    override_type(const override_type &other);
    override_type &operator=(const override_type &other);

    string get_part_name() const
    {
        return part_name_;
    }
    string get_content_type() const
    {
        return content_type_;
    }

  private:
    string part_name_;
    string content_type_;
};

class XLNT_CLASS manifest
{
  public:
    void add_default_type(const string &extension, const string &content_type);
    void add_override_type(const string &part_name, const string &content_type);

    const std::vector<default_type> &get_default_types() const
    {
        return default_types_;
    }
    const std::vector<override_type> &get_override_types() const
    {
        return override_types_;
    }

    bool has_default_type(const string &extension) const;
    bool has_override_type(const string &part_name) const;

    string get_default_type(const string &extension) const;
    string get_override_type(const string &part_name) const;

  private:
    std::vector<default_type> default_types_;
    std::vector<override_type> override_types_;
};

} // namespace xlnt
