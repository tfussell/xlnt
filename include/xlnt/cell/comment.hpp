#pragma once

#include <string>

#include "xlnt_config.hpp"

namespace xlnt {

class cell;
namespace detail { struct comment_impl; }

/// <summary>
/// A comment can be applied to a cell to provide extra information.
/// </summary>
class XLNT_CLASS comment
{
  public:
    /// <summary>
    /// The default constructor makes an invalid comment without a parent cell.
    /// </summary>
    comment();
    
    /// <summary>
    /// Constructs a comment applied to the given cell, parent, and with the comment
    /// text and author set to the provided respective values.
    comment(cell parent, const std::string &text, const std::string &auth);
    
    ~comment();
    
    /// <summary>
    /// Return the text that will be displayed for this comment.
    /// </summary>
    std::string get_text() const;
    
    /// <summary>
    /// Return the author of this comment.
    /// </summary>
    std::string get_author() const;

    /// <summary>
    /// True if the comments point to the same sell (false if
    /// they are different cells but identical comments). Note
    /// that a cell can only have one comment and a comment
    /// can only be applied to one cell.
    /// </summary>
    bool operator==(const comment &other) const;

  private:
    friend class cell; // cell needs access to private constructor
    
    /// <summary>
    /// Construct a comment from an implementation of a comment.
    /// </summary>
    comment(detail::comment_impl *d);
    
    /// <summary>
    /// Pointer to the implementation of this comment.
    /// This allows comments to be passed by value while
    /// retaining the ability to modify the parent cell.
    /// </summary>
    detail::comment_impl *d_;
};

} // namespace xlnt
