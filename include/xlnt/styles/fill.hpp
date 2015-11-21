// Copyright (c) 2014-2015 Thomas Fussell
// Copyright (c) 2010-2015 openpyxl
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file
#pragma once

#include <xlnt/xlnt_config.hpp>
#include <xlnt/styles/color.hpp>
#include <xlnt/utils/hash_combine.hpp>

namespace xlnt {

/// <summary>
/// Describes the fill style of a particular cell.
/// </summary>
class XLNT_CLASS fill
{
  public:
    enum class type
    {
        none,
        solid,
        gradient,
        pattern
    };

    enum class gradient_type
    {
        linear,
        path
    };

    enum class pattern_type
    {
        none,
        solid,
        darkdown,
        darkgray,
        darkgrid,
        darkhorizontal,
        darktrellis,
        darkup,
        darkvertical,
        gray0625,
        gray125,
        lightdown,
        lightgray,
        lightgrid,
        lighthorizontal,
        lighttrellis,
        lightup,
        lightvertical,
        mediumgray,
    };

    type get_type() const
    {
        return type_;
    }
    void set_type(type t)
    {
        type_ = t;
    }

    std::string get_pattern_type_string() const
    {
        if (type_ != type::pattern)
        {
            throw std::runtime_error("not pattern fill");
        }

        switch (pattern_type_)
        {
        case pattern_type::none:
            return "none";
        case pattern_type::solid:
            return "solid";
        case pattern_type::darkdown:
            return "darkdown";
        case pattern_type::darkgray:
            return "darkgray";
        case pattern_type::darkgrid:
            return "darkgrid";
        case pattern_type::darkhorizontal:
            return "darkhorizontal";
        case pattern_type::darktrellis:
            return "darktrellis";
        case pattern_type::darkup:
            return "darkup";
        case pattern_type::darkvertical:
            return "darkvertical";
        case pattern_type::gray0625:
            return "gray0625";
        case pattern_type::gray125:
            return "gray125";
        case pattern_type::lightdown:
            return "lightdown";
        case pattern_type::lightgray:
            return "lightgray";
        case pattern_type::lightgrid:
            return "lightgrid";
        case pattern_type::lighthorizontal:
            return "lighthorizontal";
        case pattern_type::lighttrellis:
            return "lighttrellis";
        case pattern_type::lightup:
            return "lightup";
        case pattern_type::lightvertical:
            return "lightvertical";
        case pattern_type::mediumgray:
            return "mediumgray";
        default:
            throw std::runtime_error("invalid type");
        }
    }

    std::string get_gradient_type_string() const
    {
        if (type_ != type::gradient)
        {
            throw std::runtime_error("not gradient fill");
        }

        switch (gradient_type_)
        {
        case gradient_type::linear:
            return "linear";
        case gradient_type::path:
            return "path";
        default:
            throw std::runtime_error("invalid type");
        }
    }

    pattern_type get_pattern_type() const
    {
        return pattern_type_;
    }

    void set_pattern_type(pattern_type t)
    {
        type_ = type::pattern;
        pattern_type_ = t;
    }

    void set_gradient_type(gradient_type t)
    {
        type_ = type::gradient;
        gradient_type_ = t;
    }

    void set_start_color(const color &c)
    {
        start_color_ = c;
        start_color_assigned_ = true;
    }

    void set_end_color(const color &c)
    {
        end_color_ = c;
        end_color_assigned_ = true;
    }

    void set_foreground_color(const color &c)
    {
        foreground_color_ = c;
        foreground_color_assigned_ = true;
    }

    void set_background_color(const color &c)
    {
        background_color_ = c;
        background_color_assigned_ = true;
    }

    color get_foreground_color() const
    {
        return foreground_color_;
    }

    color get_background_color() const
    {
        return background_color_;
    }

    bool has_foreground_color() const
    {
        return foreground_color_assigned_;
    }

    bool has_background_color() const
    {
        return background_color_assigned_;
    }

    bool operator==(const fill &other) const
    {
        return hash() == other.hash();
    }

    void set_rotation(double rotation)
    {
        rotation_ = rotation;
    }

    double get_rotation() const
    {
        return rotation_;
    }

    std::size_t hash() const
    {
        auto seed = static_cast<std::size_t>(type_);

        if (type_ == type::pattern)
        {
            hash_combine(seed, static_cast<std::size_t>(pattern_type_));
            hash_combine(seed, foreground_color_assigned_);

            if (foreground_color_assigned_)
            {
                hash_combine(seed, static_cast<std::size_t>(foreground_color_.get_type()));

                switch (foreground_color_.get_type())
                {
                case color::type::auto_:
                    hash_combine(seed, static_cast<std::size_t>(foreground_color_.get_auto()));
                    break;
                case color::type::indexed:
                    hash_combine(seed, static_cast<std::size_t>(foreground_color_.get_index()));
                    break;
                case color::type::theme:
                    hash_combine(seed, static_cast<std::size_t>(foreground_color_.get_theme()));
                    break;
                case color::type::rgb:
                    break;
                }
            }

            hash_combine(seed, background_color_assigned_);

            if (background_color_assigned_)
            {
                hash_combine(seed, static_cast<std::size_t>(background_color_.get_type()));

                switch (foreground_color_.get_type())
                {
                case color::type::auto_:
                    hash_combine(seed, static_cast<std::size_t>(background_color_.get_auto()));
                    break;
                case color::type::indexed:
                    hash_combine(seed, static_cast<std::size_t>(background_color_.get_index()));
                    break;
                case color::type::theme:
                    hash_combine(seed, static_cast<std::size_t>(background_color_.get_theme()));
                    break;
                case color::type::rgb:
                    break;
                }
            }
        }
        else if (type_ == type::gradient)
        {
            hash_combine(seed, static_cast<std::size_t>(gradient_type_));

            if (gradient_type_ == gradient_type::path)
            {
                hash_combine(seed, gradient_path_left_);
                hash_combine(seed, gradient_path_right_);
                hash_combine(seed, gradient_path_top_);
                hash_combine(seed, gradient_path_bottom_);
            }
            else if (gradient_type_ == gradient_type::linear)
            {
                hash_combine(seed, rotation_);
            }
        }
        else if (type_ == type::solid)
        {
            //            hash_combine(seed, static_cast<std::size_t>());
        }

        return seed;
    }

    double get_gradient_left() const
    {
        return gradient_path_left_;
    }
    double get_gradient_right() const
    {
        return gradient_path_right_;
    }
    double get_gradient_top() const
    {
        return gradient_path_top_;
    }
    double get_gradient_bottom() const
    {
        return gradient_path_bottom_;
    }

  private:
    type type_ = type::none;
    pattern_type pattern_type_;
    gradient_type gradient_type_;
    double rotation_ = 0;
    bool foreground_color_assigned_ = false;
    color foreground_color_ = color::black();
    bool background_color_assigned_ = false;
    color background_color_ = color::white();
    bool start_color_assigned_ = false;
    color start_color_ = color::white();
    bool end_color_assigned_ = false;
    color end_color_ = color::black();
    double gradient_path_left_ = 0;
    double gradient_path_right_ = 0;
    double gradient_path_top_ = 0;
    double gradient_path_bottom_ = 0;
};

} // namespace xlnt
