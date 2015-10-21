#pragma once

#include <string>

namespace xlnt {
    
class cell;
    
namespace detail {
    
struct comment_impl;

} // namespace detail

/// <summary>
/// A comment can be applied to a cell to provide extra information.
/// </summary>
class comment
{
public:
    /// <summary>
    /// The default constructor makes an invalid comment without a parent cell.
    /// </summary>
    comment();
    comment(cell parent, const std::string &text, const std::string &auth);
    ~comment();
    
    std::string get_text() const;
    std::string get_author() const;
    
    /// <summary>
    /// True if the comments point to the same sell.
    /// Note that a cell can only have one comment and a comment
    /// can only be applied to one cell.
    /// </summary>
    bool operator==(const comment &other) const;

private:
    friend class cell;
    comment(detail::comment_impl *d);
    detail::comment_impl *d_;
};

} // namespace xlnt
