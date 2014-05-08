#pragma once

#include <memory>
#include <string>
#include <vector>

#include "target_mode.h"
#include "uri.h"

namespace xlnt {

class package;

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

    relationship(const std::string &type) : source_uri_(""), target_uri_("")
    {
        if(type == "hyperlink")
        {
            type_ = type::hyperlink;
        }
    }

	/// <summary>
	/// gets a string that identifies the relationship.
	/// </summary>
	std::string get_id() const { return id_; }

	/// <summary>
	/// gets the URI of the package or part that owns the relationship.
	/// </summary>
	uri getsource_uri() const { return source_uri_; }

	/// <summary>
	/// gets a value that indicates whether the target of the relationship is or External to the Package.
	/// </summary>
	target_mode get_target_mode() const { return target_mode_; }

	/// <summary>
	/// gets the URI of the target resource of the relationship.
	/// </summary>
	uri get_target_uri() const { return target_uri_; }

private:
	friend class package;

	relationship(const std::string &id, package &package, const std::string &relationship_type_, uri source_uri, target_mode target_mode, uri target_uri);
    relationship &operator=(const relationship &rhs) = delete;

	std::string id_;
	type type_;
	uri source_uri_;
	target_mode target_mode_;
	uri target_uri_;
};

typedef std::vector<std::shared_ptr<relationship>> relationship_collection;

} // namespace xlnt
