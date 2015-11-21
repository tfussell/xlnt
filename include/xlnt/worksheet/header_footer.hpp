#pragma once

#include <cstdint>
#include <string>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/worksheet/footer.hpp>
#include <xlnt/worksheet/header.hpp>

namespace xlnt {

/// <summary>
/// Worksheet header and footer for all six possible positions.
/// </summary>
class XLNT_CLASS header_footer
{
  public:
    header_footer();

    header &get_left_header()
    {
        return left_header_;
    }
    header &get_center_header()
    {
        return center_header_;
    }
    header &get_right_header()
    {
        return right_header_;
    }
    footer &get_left_footer()
    {
        return left_footer_;
    }
    footer &get_center_footer()
    {
        return center_footer_;
    }
    footer &get_right_footer()
    {
        return right_footer_;
    }

    bool is_default_header() const
    {
        return left_header_.is_default() && center_header_.is_default() && right_header_.is_default();
    }
    bool is_default_footer() const
    {
        return left_footer_.is_default() && center_footer_.is_default() && right_footer_.is_default();
    }
    bool is_default() const
    {
        return is_default_header() && is_default_footer();
    }

  private:
    header left_header_, right_header_, center_header_;
    footer left_footer_, right_footer_, center_footer_;
};

} // namespace xlnt
