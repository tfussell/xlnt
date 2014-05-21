#pragma once

#include <string>

namespace xlnt {
    
/// <summary>
/// Specifies whether the target of a relationship is inside or outside the Package.
/// </summary>
enum class target_mode
{
    /// <summary>
    /// The relationship references a part that is inside the package.
    /// </summary>
    external,
    /// <summary>
    /// The relationship references a resource that is external to the package.
    /// </summary>
    internal
};

/// <summary>
/// Represents an association between a source Package or part, and a target object which can be a part or external resource.
/// </summary>
class relationship
{
public:
    enum class type
    {
        hyperlink,
        drawing,
        worksheet,
        sharedStrings,
        styles,
        theme
    };
    
    relationship(const std::string &type, const std::string &r_id = "", const std::string &target_uri = "");
    
    /// <summary>
    /// gets a string that identifies the relationship.
    /// </summary>
    std::string get_id() const { return id_; }
    
    /// <summary>
    /// gets the URI of the package or part that owns the relationship.
    /// </summary>
    std::string get_source_uri() const { return source_uri_; }
    
    /// <summary>
    /// gets a value that indicates whether the target of the relationship is or External to the Package.
    /// </summary>
    target_mode get_target_mode() const { return target_mode_; }
    
    /// <summary>
    /// gets the URI of the target resource of the relationship.
    /// </summary>
    std::string get_target_uri() const { return target_uri_; }
    
private:
    relationship(const std::string &id, const std::string &relationship_type_, const std::string &source_uri, target_mode target_mode, const std::string &target_uri);
    //relationship &operator=(const relationship &rhs) = delete;
    
    type type_;
    std::string id_;
    std::string source_uri_;
    std::string target_uri_;
    target_mode target_mode_;
};
    
} // namespace xlnt
