#pragma once

#include <array>
#include <sstream>
#include <opc/opc.h>

#include "uri.h"
#include "compression_option.h"
#include "relationship.h"
#include "target_mode.h"
#include "file_mode.h"
#include "file_access.h"

namespace xlnt {

class part_impl;

/// <summary>
/// Represents a part that is stored in a ZipPackage.
/// </summary>
class part
{
public:
    /// <summary>
    /// Initializes a new instance of the part class with a specified parent Package, part URI, MIME content type, and compression_option.
    /// </summary>
    part(package &package, const uri &uri_part, const std::string &mime_type = "", compression_option compression = compression_option::NotCompressed);

    part &operator=(const part &) = delete;

    /// <summary>
    /// gets the compression option of the part content stream.
    /// </summary>
    compression_option get_compression_option() const;

    /// <summary>
    /// gets the MIME type of the content stream.
    /// </summary>
    std::string get_content_type() const;

    /// <summary>
    /// gets the parent Package of the part.
    /// </summary>
    package &get_package() const;

    /// <summary>
    /// gets the URI of the part.
    /// </summary>
    uri get_uri() const;

    /// <summary>
    /// Creates a part-level relationship between this part to a specified target part or external resource.
    /// </summary>
    std::shared_ptr<relationship> create_relationship(const uri &target_uri, target_mode target_mode, const std::string &relationship_type);

    /// <summary>
    /// Deletes a specified part-level relationship.
    /// </summary>
    void delete_relationship(const std::string &id);

    /// <summary>
    /// Returns the relationship that has a specified Id.
    /// </summary>
    relationship get_relationship(const std::string &id);

    /// <summary>
    /// Returns a collection of all the relationships that are owned by this part.
    /// </summary>
    relationship_collection get_relationships();

    /// <summary>
    /// Returns a collection of the relationships that match a specified RelationshipType.
    /// </summary>
    relationship_collection get_relationship_by_type(const std::string &relationship_type);

    /// <summary>
    /// Returns all the content of this part.
    /// </summary>
    std::string read();

    /// <summary>
    /// Writes the given data to the part stream.
    /// </summary>
    void write(const std::string &data);

    /// <summary>
    /// Returns a value that indicates whether this part owns a relationship with a specified Id.
    /// </summary>
    bool relationship_exists(const std::string &id) const;

    bool operator==(const part &comparand) const;
    bool operator==(const nullptr_t &) const;

private:
    friend class package_impl;

    part();
    part(package &package, const uri &uri, opcContainer *container);

    std::shared_ptr<part_impl> impl_;
};

typedef std::vector<part> part_collection;

} // namespace xlnt
