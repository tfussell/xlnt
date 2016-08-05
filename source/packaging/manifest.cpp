// Copyright (c) 2014-2016 Thomas Fussell
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

#include <xlnt/packaging/manifest.hpp>
#include <xlnt/utils/exceptions.hpp>

namespace {

std::string to_string(xlnt::content_type type)
{
	return "";
}

xlnt::content_type from_string(const std::string &type_string)
{
	return xlnt::content_type::unknown;
}

}

namespace xlnt {

void manifest::clear()
{
	clear_types();
	clear_relationships();
	part_infos_.clear();
}

void manifest::clear_types()
{
	clear_default_types();
	clear_override_types();
}

void manifest::clear_default_types()
{
	extension_content_type_map_.clear();
}

void manifest::clear_override_types()
{
	for (auto &info : part_infos_)
	{
		info.second.content_type.clear();
	}
}

void manifest::clear_relationships()
{
	clear_package_relationships();
	clear_part_relationships();
}

void manifest::clear_package_relationships()
{
	package_relationships_.clear();
}

void manifest::clear_part_relationships()
{
	for (auto &info : part_infos_)
	{
		info.second.relationships.clear();
	}
}

bool manifest::has_package_relationship(relationship::type type) const
{
	for (const auto &rel : package_relationships_)
	{
		if (rel.get_type() == type)
		{
			return true;
		}
	}

	return false;
}

bool manifest::has_part_relationship(const path &part, relationship::type type) const
{
	for (const auto &rel : part_infos_.at(part).relationships)
	{
		if (rel.get_type() == type)
		{
			return true;
		}
	}

	return false;
}

relationship manifest::get_package_relationship(relationship::type type) const
{
	for (const auto &rel : package_relationships_)
	{
		if (rel.get_type() == type)
		{
			return rel;
		}
	}

	throw key_not_found();
}

std::vector<relationship> manifest::get_package_relationships(relationship::type type) const
{
	std::vector<relationship> matches;

	for (const auto &rel : package_relationships_)
	{
		if (rel.get_type() == type)
		{
			matches.push_back(rel);
		}
	}

	return matches;
}

relationship manifest::get_part_relationship(const path &part, relationship::type type) const
{
	for (const auto &rel : part_infos_.at(part).relationships)
	{
		if (rel.get_type() == type)
		{
			return rel;
		}
	}

	throw key_not_found();
}

std::vector<relationship> manifest::get_part_relationships(const path &part, relationship::type type) const
{
	std::vector<relationship> matches;

	for (const auto &rel : part_infos_.at(part).relationships)
	{
		if (rel.get_type() == type)
		{
			matches.push_back(rel);
		}
	}

	return matches;
}

content_type manifest::get_part_content_type(const path &part) const
{
	return from_string(get_part_content_type_string(part));
}

std::string manifest::get_part_content_type_string(const path &part) const
{
	if (part_infos_.find(part) == part_infos_.end())
	{
		throw key_not_found();
	}

	return part_infos_.at(part).content_type;
}

void manifest::register_part(const path &part, const path &parent, const std::string &content_type, relationship::type relation)
{
	part_infos_[part] = { content_type, {} };

	relationship rel(next_package_relationship_id(), relation, part, target_mode::internal);
	part_infos_[parent].relationships.push_back(rel);
}

void manifest::register_part(const path &parent, const relationship &rel, const std::string &content_type)
{
	part_infos_[rel.get_target_uri()] = { content_type,{} };
	part_infos_[parent].relationships.push_back(rel);
}

void manifest::register_package_part(const path &part, const std::string &content_type, relationship::type relation)
{
	part_infos_[part] = { content_type, {} };

	relationship rel(next_package_relationship_id(), relation, part, target_mode::internal);
	package_relationships_.push_back(rel);
}

void manifest::unregister_part(const path &part)
{
	if (part_infos_.find(part) != part_infos_.end())
	{
		part_infos_.erase(part);
	}
	else
	{
		throw key_not_found();
	}

	auto package_rels_iter = package_relationships_.begin();

	while (package_rels_iter != package_relationships_.end())
	{
		if (package_rels_iter->get_target_uri() == part)
		{
			package_rels_iter = package_relationships_.erase(package_rels_iter);
			continue;
		}

		++package_rels_iter;
	}

	for (auto &current_part : part_infos_)
	{
		auto rels_iter = current_part.second.relationships.begin();

		while (rels_iter != current_part.second.relationships.end())
		{
			if (rels_iter->get_target_uri() == part)
			{
				rels_iter = current_part.second.relationships.erase(rels_iter);
				continue;
			}

			++rels_iter;
		}
	}
}

std::vector<path> manifest::get_parts() const
{
	std::vector<path> parts;

	for (const auto &part : part_infos_)
	{
		parts.push_back(part.first);
	}

	return parts;
}

bool manifest::has_part(const path &part) const
{
	return part_infos_.find(part) != part_infos_.end();
}

std::vector<relationship> manifest::get_package_relationships() const
{
	return package_relationships_;
}

relationship manifest::get_package_relationship(const std::string &rel_id)
{
	for (const auto rel : package_relationships_)
	{
		if (rel.get_id() == rel_id)
		{
			return rel;
		}
	}

	throw key_not_found();
}

std::vector<relationship> manifest::get_part_relationships(const path &part) const
{
	std::vector<relationship> matches;

	for (const auto &rel : part_infos_.at(part).relationships)
	{
		matches.push_back(rel);
	}

	return matches;
}

relationship manifest::get_part_relationship(const path &part, const std::string &rel_id) const
{
	if (part_infos_.find(part) == part_infos_.end())
	{
		throw key_not_found();
	}

	for (const auto &rel : part_infos_.at(part).relationships)
	{
		if (rel.get_id() == rel_id)
		{
			return rel;
		}
	}

	throw key_not_found();
}

std::string manifest::register_external_package_relationship(relationship::type type, const std::string &target_uri)
{
	return register_external_package_relationship(type, target_uri, next_package_relationship_id());
}

std::string manifest::register_external_package_relationship(relationship::type type, const std::string &target_uri, const std::string &rel_id)
{
	relationship rel(rel_id, type, path(target_uri), target_mode::external);
	package_relationships_.push_back(rel);
	return rel_id;
}

std::string manifest::register_external_relationship(const path &source_part, relationship::type type, const std::string &target_uri)
{
	return register_external_relationship(source_part, type, target_uri, next_relationship_id(source_part));
}

std::string manifest::register_external_relationship(const path &source_part, relationship::type type, const std::string &target_uri, const std::string &rel_id)
{
	relationship rel(rel_id, type, path(target_uri), target_mode::external);
	part_infos_.at(source_part).relationships.push_back(rel);
	return rel_id;
}

bool manifest::has_default_type(const std::string &extension) const
{
	return extension_content_type_map_.find(extension) != extension_content_type_map_.end();
}

std::vector<std::string> manifest::get_default_type_extensions() const
{
	std::vector<std::string> extensions;

	for (const auto &extension_type_pair : extension_content_type_map_)
	{
		extensions.push_back(extension_type_pair.first);
	}

	return extensions;
}

std::string manifest::get_default_type(const std::string &extension) const
{
	if (extension_content_type_map_.find(extension) == extension_content_type_map_.end())
	{
		throw key_not_found();
	}

	return extension_content_type_map_.at(extension);
}

void manifest::register_default_type(const std::string &extension, const std::string &content_type)
{
	extension_content_type_map_[extension] = content_type;
}

void manifest::unregister_default_type(const std::string &extension)
{
	extension_content_type_map_.erase(extension);
}

std::string manifest::next_package_relationship_id() const
{
	std::string r_id("rId");
	std::size_t index = 1;

	while (std::find_if(package_relationships_.begin(), package_relationships_.end(),
		[&](const relationship &r) { return r.get_id() == r_id + std::to_string(index); })
		!= package_relationships_.end())
	{
		++index;
	}

	return r_id + std::to_string(index);
}

std::string manifest::next_relationship_id(const path &part) const
{
	std::string r_id("rId");
	std::size_t index = 1;

	const auto &part_rels = get_part_relationships(part);

	while (std::find_if(part_rels.begin(), part_rels.end(), 
		[&](const relationship &r) { return r.get_id() == r_id + std::to_string(index); }) 
		!= part_rels.end())
	{
		++index;
	}

	return r_id + std::to_string(index);
}

} // namespace xlnt
