// Copyright (c) 2014-2017 Thomas Fussell
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

#include <string>
#include <unordered_map>

#include <xlnt/xlnt_config.hpp>
#include <xlnt/packaging/relationship.hpp>
#include <xlnt/utils/path.hpp>

namespace xlnt {

/// <summary>
/// The manifest keeps track of all files in the OOXML package and
/// their type and relationships.
/// </summary>
class XLNT_API manifest
{
public:
    /// <summary>
    /// Unregisters all default and override type and all relationships and known parts.
    /// </summary>
    void clear();

    /// <summary>
    /// Returns the path to all internal package parts registered as a source
    /// or target of a relationship.
    /// </summary>
    std::vector<path> parts() const;

    // Relationships

    /// <summary>
    /// Returns true if the manifest contains a relationship with the given type with part as the source.
    /// </summary>
    bool has_relationship(const path &source, relationship_type type) const;

    /// <summary>
    /// Returns true if the manifest contains a relationship with the given type with part as the source.
    /// </summary>
    bool has_relationship(const path &source, const std::string &rel_id) const;

    /// <summary>
    /// Returns the relationship with "source" as the source and with a type of "type".
    /// Throws a key_not_found exception if no such relationship is found.
    /// </summary>
    class relationship relationship(const path &source, relationship_type type) const;

    /// <summary>
    /// Returns the relationship with "source" as the source and with an ID of "rel_id".
    /// Throws a key_not_found exception if no such relationship is found.
    /// </summary>
    class relationship relationship(const path &source, const std::string &rel_id) const;

    /// <summary>
    /// Returns all relationship with "source" as the source.
    /// </summary>
    std::vector<xlnt::relationship> relationships(const path &source) const;

    /// <summary>
    /// Returns all relationships with "source" as the source and with a type of "type".
    /// </summary>
    std::vector<xlnt::relationship> relationships(const path &source, relationship_type type) const;

    /// <summary>
    /// Returns the canonical path of the chain of relationships by traversing through rels
    /// and forming the absolute combined path.
    /// </summary>
    path canonicalize(const std::vector<xlnt::relationship> &rels) const;

    /// <summary>
    /// Registers a new relationship by specifying all of the relationship properties explicitly.
    /// </summary>
    std::string register_relationship(const uri &source, relationship_type type, const uri &target, target_mode mode);

    /// <summary>
    /// Registers a new relationship already constructed elsewhere.
    /// </summary>
    std::string register_relationship(const class relationship &rel);

    /// <summary>
    /// Delete the relationship with the given id from source part. Returns a mapping
    /// of relationship IDs since IDs are shifted down. For example, if there are three
    /// relationships for part a.xml like [rId1, rId2, rId3] and rId2 is deleted, the
    /// resulting map would look like [rId3->rId2].
    /// </summary>
    std::unordered_map<std::string, std::string> unregister_relationship(const uri &source, const std::string &rel_id);

    // Content Types

    /// <summary>
    /// Given the path to a part, returns the content type of the part as a string.
    /// </summary>
    std::string content_type(const path &part) const;

    // Default Content Types

    /// <summary>
    /// Returns true if a default content type for the extension has been registered.
    /// </summary>
    bool has_default_type(const std::string &extension) const;

    /// <summary>
    /// Returns a vector of all extensions with registered default content types.
    /// </summary>
    std::vector<std::string> extensions_with_default_types() const;

    /// <summary>
    /// Returns the registered default content type for parts of the given extension.
    /// </summary>
    std::string default_type(const std::string &extension) const;

    /// <summary>
    /// Associates the given extension with the given content type.
    /// </summary>
    void register_default_type(const std::string &extension, const std::string &type);

    /// <summary>
    /// Unregisters the default content type for the given extension.
    /// </summary>
    void unregister_default_type(const std::string &extension);

    // Override Content Types

    /// <summary>
    /// Returns true if a content type overriding the default content type has been registered
    /// for the given part.
    /// </summary>
    bool has_override_type(const path &part) const;

    /// <summary>
    /// Returns the override content type registered for the given part.
    /// Throws key_not_found exception if no override type was found.
    /// </summary>
    std::string override_type(const path &part) const;

    /// <summary>
    /// Returns the path of every part in this manifest with an overriden content type.
    /// </summary>
    std::vector<path> parts_with_overriden_types() const;

    /// <summary>
    /// Overrides any default type registered for the part's extension with the given content type.
    /// </summary>
    void register_override_type(const path &part, const std::string &type);

    /// <summary>
    /// Unregisters the overriding content type of the given part.
    /// </summary>
    void unregister_override_type(const path &part);

private:
    /// <summary>
    /// Returns the lowest rId for the given part that hasn't already been registered.
    /// </summary>
    std::string next_relationship_id(const path &part) const;

    /// <summary>
    /// The map of extensions to default content types.
    /// </summary>
    std::unordered_map<std::string, std::string> default_content_types_;

    /// <summary>
    /// The map of package parts to overriding content types.
    /// </summary>
    std::unordered_map<path, std::string> override_content_types_;

    /// <summary>
    /// The map of package parts to their registered relationships.
    /// </summary>
    std::unordered_map<path, std::unordered_map<std::string, xlnt::relationship>> relationships_;
};

} // namespace xlnt
