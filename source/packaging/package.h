#pragma once

#include <fstream>
#include <opc/opc.h>
#include <opc/container.h>

#include "compression_option.h"
#include "file_access.h"
#include "file_mode.h"
#include "file_share.h"
#include "opc_callback_handler.h"
#include "package_properties.h"
#include "part.h"
#include "relationship.h"
#include "target_mode.h"
#include "uri.h"

namespace xlnt {

class package_impl;

/// <summary>
/// Implements a derived subclass of the abstract Package base class—the ZipPackage class uses a 
/// ZIP archive as the container store. This class should not be inherited.
/// </summary>
class package
{
public:
    /// <summary>
    /// Opens a package with a given IO stream, file mode, and file access setting.
    /// </summary>
    static package open(std::iostream &stream, file_mode package_mode, file_access package_access);

    /// <summary>
    /// Opens a package at a given path using a given file mode, file access, and file share setting.
    /// </summary>
    static package open(const std::string &path, file_mode package_mode = file_mode::OpenOrCreate, 
        file_access package_access = file_access::ReadWrite, file_share package_share = file_share::None);

    /// <summary>
    /// Saves and closes the package plus all underlying part streams.
    /// </summary>
    void close();

    /// <summary>
    /// Creates a new part with a given URI, content type, and compression option.
    /// </summary>
    part create_part(const uri &part_uri, const std::string &content_type, compression_option compression = compression_option::Normal);

    /// <summary>
    /// Creates a package-level relationship to a part with a given URI, target mode, relationship type, and identifier (ID).
    /// </summary>
    relationship create_relationship(const uri &part_uri, target_mode target_mode, const std::string &relationship_type, const std::string &id);

    /// <summary>
    /// Deletes a part with a given URI from the package.
    /// </summary>
    void delete_part(const uri &part_uri);

    /// <summary>
    /// Deletes a package-level relationship.
    /// </summary>
    void delete_relationship(const std::string &id);

    /// <summary>
    /// Saves the contents of all parts and relationships that are contained in the package.
    /// </summary>
    void flush();

    /// <summary>
    /// Returns the part with a given URI.
    /// </summary>
    part get_part(const uri &part_uri);

    /// <summary>
    /// Returns a collection of all the parts in the package.
    /// </summary>
    part_collection get_parts();

    /// <summary>
    /// 
    /// </summary>
    relationship get_relationship(const std::string &id);

    /// <summary>
    /// Returns the package-level relationship with a given identifier.
    /// </summary>
    relationship_collection get_relationships();

    /// <summary>
    /// Returns a collection of all the package-level relationships that match a given RelationshipType.
    /// </summary>
    relationship_collection get_relationships(const std::string &relationship_type);

    /// <summary>
    /// Indicates whether a part with a given URI is in the package.
    /// </summary>
    bool part_exists(const uri &part_uri);

    /// <summary>
    /// Indicates whether a package-level relationship with a given ID is contained in the package.
    /// </summary>
    bool relationship_exists(const std::string &id);

    /// <summary>
    /// gets the file access setting for the package.
    /// </summary>
    file_access get_file_open_access() const;

    /// <summary>
    /// gets the core properties of the package.
    /// </summary>
    package_properties get_package_properties() const;

    bool operator==(const package &comparand) const;
    bool operator==(const nullptr_t &) const;

private:
	friend opc_callback_handler;

	package(std::iostream &stream, file_mode package_mode, file_access package_access);
	package(const std::string &path, file_mode package_mode, file_access package_access);

	void open_container();

	int write(char *buffer, int length);

	int read(const char *buffer, int length);

	std::ios::pos_type seek(std::ios::pos_type ofs);

	int trim(std::ios::pos_type new_size);

    std::shared_ptr<package_impl> impl_;
};

} // namespace xlnt
