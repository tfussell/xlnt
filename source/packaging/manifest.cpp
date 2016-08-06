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

content_type manifest::get_content_type(const path &part) const
{
	return from_string(get_content_type_string(part));
}

std::string manifest::get_content_type_string(const path &part) const
{
	if (part_infos_.find(part) == part_infos_.end())
	{
		auto extension = part.extension();

		if (has_default_type(extension))
		{
			return get_default_type(extension);
		}

		throw key_not_found();
	}

	return part_infos_.at(part).content_type;
}

void manifest::register_override_type(const path &part, const std::string &content_type)
{
	part_infos_[part] = { content_type, {} };
}

std::string manifest::register_package_relationship(relationship::type type, const path &target_uri, target_mode mode)
{
	return register_package_relationship(type, target_uri, mode, next_package_relationship_id());
}

std::string manifest::register_package_relationship(relationship::type type, const path &target_uri, target_mode mode, const std::string &rel_id)
{
	if (has_package_relationship(rel_id))
	{
		throw invalid_parameter();
	}

	relationship rel(rel_id, type, target_uri, mode);
	package_relationships_.push_back(rel);

	return rel_id;
}

bool manifest::has_package_relationship(const std::string &rel_id) const
{
	for (const auto &rel : package_relationships_)
	{
		if (rel.get_id() == rel_id)
		{
			return true;
		}
	}

	return false;
}

std::vector<path> manifest::get_overriden_parts() const
{
	std::vector<path> overriden;

	for (const auto &part : part_infos_)
	{
		if (!part.second.content_type.empty())
		{
			overriden.push_back(part.first);
		}
	}

	return overriden;
}

std::vector<relationship> manifest::get_package_relationships() const
{
	return package_relationships_;
}

relationship manifest::get_package_relationship(const std::string &rel_id) const
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
	if (part_infos_.find(part) == part_infos_.end())
	{
		throw key_not_found();
	}

	return part_infos_.at(part).relationships;
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

std::vector<path> manifest::get_parts_with_relationships() const
{
	std::vector<path> with_relationships;

	for (const auto &info : part_infos_)
	{
		if (!info.second.relationships.empty())
		{
			with_relationships.push_back(info.first);
		}
	}

	return with_relationships;
}

std::string manifest::register_part_relationship(const path &source_uri, relationship::type type, const path &target_uri, target_mode mode)
{
	return register_part_relationship(source_uri, type, target_uri, mode, next_relationship_id(source_uri));
}

std::string manifest::register_part_relationship(const path &source_uri, relationship::type type, const path &target_uri, target_mode mode, const std::string &rel_id)
{
	relationship rel(rel_id, type, path(target_uri), mode);
	part_infos_.at(source_uri).relationships.push_back(rel);
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
