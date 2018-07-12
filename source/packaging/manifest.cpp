// Copyright (c) 2014-2018 Thomas Fussell
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

#include <algorithm>
#include <unordered_set>

#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace xlnt {

void manifest::clear()
{
    default_content_types_.clear();
    override_content_types_.clear();
    relationships_.clear();
}

path manifest::canonicalize(const std::vector<xlnt::relationship> &rels) const
{
    xlnt::path relative;

    for (auto component : rels)
    {
        if (component == rels.back())
        {
            relative = relative.append(component.target().path());
        }
        else
        {
            relative = relative.append(component.target().path().parent());
        }
    }

    std::vector<std::string> absolute_parts;

    for (const auto &component : relative.split())
    {
        if (component == ".") continue;

        if (component == "..")
        {
            absolute_parts.pop_back();
            continue;
        }

        absolute_parts.push_back(component);
    }

    xlnt::path result;

    for (const auto &component : absolute_parts)
    {
        result = result.append(component);
    }

    return result;
}

bool manifest::has_relationship(const path &path, relationship_type type) const
{
    auto rels = relationships_.find(path);
    if (rels == relationships_.end())
    {
        return false;
    }
    return rels->second.end() != std::find_if(rels->second.begin(), rels->second.end(), 
                                              [type](const std::pair<std::string, xlnt::relationship> &rel) { return rel.second.type() == type; });
}

bool manifest::has_relationship(const path &path, const std::string &rel_id) const
{
    auto rels = relationships_.find(path);
    if (rels == relationships_.end()) 
    {
        return false;
    }
    return rels->second.find(rel_id) != rels->second.end();
}

relationship manifest::relationship(const path &part, relationship_type type) const
{
    if (relationships_.find(part) == relationships_.end()) throw key_not_found();

    for (const auto &rel : relationships_.at(part))
    {
        if (rel.second.type() == type) return rel.second;
    }

    throw key_not_found();
}

std::vector<xlnt::relationship> manifest::relationships(const path &part, relationship_type type) const
{
    std::vector<xlnt::relationship> matches;

    if (has_relationship(part, type))
    {
        for (const auto &rel : relationships_.at(part))
        {
            if (rel.second.type() == type)
            {
                matches.push_back(rel.second);
            }
        }
    }

    return matches;
}

std::string manifest::content_type(const path &part) const
{
    auto absolute = part.resolve(path("/"));

    if (has_override_type(absolute))
    {
        return override_type(absolute);
    }

    if (has_default_type(part.extension()))
    {
        return default_type(part.extension());
    }

    throw key_not_found();
}

void manifest::register_override_type(const path &part, const std::string &content_type)
{
    override_content_types_[part] = content_type;
}

void manifest::unregister_override_type(const path &part)
{
    override_content_types_.erase(part);
}

std::vector<path> manifest::parts_with_overriden_types() const
{
    std::vector<path> overriden;

    for (const auto &part : override_content_types_)
    {
        overriden.push_back(part.first);
    }

    return overriden;
}

std::vector<relationship> manifest::relationships(const path &part) const
{
    if (relationships_.find(part) == relationships_.end())
    {
        return {};
    }

    std::vector<xlnt::relationship> relationships;

    for (const auto &rel : relationships_.at(part))
    {
        relationships.push_back(rel.second);
    }

    return relationships;
}

relationship manifest::relationship(const path &part, const std::string &rel_id) const
{
    if (relationships_.find(part) == relationships_.end())
    {
        throw key_not_found();
    }

    for (const auto &rel : relationships_.at(part))
    {
        if (rel.second.id() == rel_id)
        {
            return rel.second;
        }
    }

    throw key_not_found();
}

std::vector<path> manifest::parts() const
{
    std::unordered_set<path> parts;

    for (const auto &part_rels : relationships_)
    {
        parts.insert(part_rels.first);

        for (const auto &rel : part_rels.second)
        {
            if (rel.second.target_mode() == target_mode::internal)
            {
                parts.insert(rel.second.target().path());
            }
        }
    }

    return std::vector<path>(parts.begin(), parts.end());
}

std::string manifest::register_relationship(const uri &source,
    relationship_type type, const uri &target, target_mode mode)
{
    xlnt::relationship rel(next_relationship_id(source.path()), type, source, target, mode);
    return register_relationship(rel);
}

std::string manifest::register_relationship(const class relationship &rel)
{
    relationships_[rel.source().path()][rel.id()] = rel;
    return rel.id();
}

std::unordered_map<std::string, std::string> manifest::unregister_relationship(const uri &source, const std::string &rel_id)
{
    // This shouldn't happen, but just in case...
    if (rel_id.substr(0, 3) != "rId" || rel_id.size() < 4)
    {
        throw xlnt::invalid_parameter();
    }

    std::unordered_map<std::string, std::string> id_map;
    auto rel_index = static_cast<std::size_t>(std::stoull(rel_id.substr(3)));
    auto &part_rels = relationships_.at(source.path());

    for (auto i = rel_index; i <= part_rels.size() + 1; ++i)
    {
        auto old_id = "rId" + std::to_string(i);

        // Don't re-add the relationship to be deleted
        if (i > rel_index)
        {
            // Shift all relationships with IDs greater than the deleted one
            // down by one (e.g. rId7->rId6).
            auto new_id = "rId" + std::to_string(i - 1);
            const auto &old_rel = part_rels.at(old_id);
            register_relationship(xlnt::relationship(new_id, old_rel.type(),
                old_rel.source(), old_rel.target(), old_rel.target_mode()));
            id_map[old_id] = new_id;
        }

        part_rels.erase(old_id);
    }

    return id_map;
}

bool manifest::has_default_type(const std::string &extension) const
{
    return default_content_types_.find(extension) != default_content_types_.end();
}

std::vector<std::string> manifest::extensions_with_default_types() const
{
    std::vector<std::string> extensions;

    for (const auto &extension_type_pair : default_content_types_)
    {
        extensions.push_back(extension_type_pair.first);
    }

    return extensions;
}

std::string manifest::default_type(const std::string &extension) const
{
    if (default_content_types_.find(extension) == default_content_types_.end())
    {
        throw key_not_found();
    }

    return default_content_types_.at(extension);
}

void manifest::register_default_type(const std::string &extension, const std::string &content_type)
{
    default_content_types_[extension] = content_type;
}

void manifest::unregister_default_type(const std::string &extension)
{
    default_content_types_.erase(extension);
}

std::string manifest::next_relationship_id(const path &part) const
{
    if (relationships_.find(part) == relationships_.end()) return "rId1";

    std::size_t index = 1;
    const auto &part_rels = relationships_.at(part);

    while (part_rels.find("rId" + std::to_string(index)) != part_rels.end())
    {
        ++index;
    }

    return "rId" + std::to_string(index);
}

bool manifest::has_override_type(const xlnt::path &part) const
{
    return override_content_types_.find(part) != override_content_types_.end();
}

std::string manifest::override_type(const xlnt::path &part) const
{
    if (!has_override_type(part))
    {
        throw key_not_found();
    }

    return override_content_types_.at(part);
}

bool manifest::operator==(const manifest &other) const
{
    return default_content_types_ == other.default_content_types_
        && override_content_types_ == other.override_content_types_
        && relationships_ == other.relationships_;
}

} // namespace xlnt
