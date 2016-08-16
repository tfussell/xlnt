#pragma once

#include <pugixml.hpp>
#include <sstream>

#include "path_helper.hpp"

class xml_helper
{
public:
    enum class difference_type
    {
        names_differ,
        missing_attribute,
        attribute_values_differ,
        missing_text,
        text_values_differ,
        missing_child,
        child_order_differs,
        equivalent,
    };
    
    struct comparison_result
    {
        difference_type difference;
        std::string value_left;
        std::string value_right;

        operator bool() const
        {
            return difference == difference_type::equivalent;
        }
    };
    
    static bool documents_match(const pugi::xml_document &expected, 
		const pugi::xml_document &observed)
    {
        auto result = compare_xml_nodes(expected.root(), observed.root());

		if (!result)
		{
			std::cout << "documents don't match" << std::endl;

			std::cout << "expected:" << std::endl;
			expected.save(std::cout);
			std::cout << std::endl;

			std::cout << "observed:" << std::endl;
			observed.save(std::cout);
			std::cout << std::endl;

			return false;
		}

		return true;
    }

	static bool string_matches_workbook_part(const std::string &expected,
		xlnt::workbook &wb, const xlnt::path &part)
	{
		std::vector<std::uint8_t> bytes;
		wb.save(bytes);
		xlnt::zip_file archive;
		archive.load(bytes);

		return string_matches_archive_member(expected, archive, part);
	}

	static bool file_matches_workbook_part(const xlnt::path &expected,
		xlnt::workbook &wb, const xlnt::path &part)
	{
		std::vector<std::uint8_t> bytes;
		wb.save(bytes);
		xlnt::zip_file archive;
		archive.load(bytes);

		return file_matches_archive_member(expected, archive, part);
	}

	static bool string_matches_archive_member(const std::string &expected,
		xlnt::zip_file &archive,
		const xlnt::path &member)
	{
		pugi::xml_document expected_document;
		expected_document.load(expected.c_str());

		pugi::xml_document observed_document;
		expected_document.load(archive.read(member).c_str());

		return documents_match(expected_document, observed_document);
	}

	static bool file_matches_archive_member(const xlnt::path &file,
		xlnt::zip_file &archive,
		const xlnt::path &member)
	{
		pugi::xml_document expected_document;
		expected_document.load(file.read_contents().c_str());

		pugi::xml_document observed_document;
		if (!archive.has_file(member)) return false;
		expected_document.load(archive.read(member).c_str());

		return documents_match(expected_document, observed_document);
	}

	static bool file_matches_document(const xlnt::path &expected, 
		const pugi::xml_document &observed)
	{
		return string_matches_document(expected.read_contents(), observed);
	}

    static bool string_matches_document(const std::string &expected_string,
		const pugi::xml_document &observed_document)
    {        
		pugi::xml_document expected_document;
		expected_document.load(expected_string.c_str());

        return documents_match(expected_document, observed_document);
    }
    
    static bool strings_match(const std::string &expected_string, const std::string &observed_string)
    {
		pugi::xml_document left_xml;
		left_xml.load(expected_string.c_str());

		pugi::xml_document right_xml;
		right_xml.load(observed_string.c_str());

        return documents_match(left_xml, right_xml);
    }

	static bool archives_match(xlnt::zip_file &left_archive, xlnt::zip_file &right_archive)
	{
		auto left_info = left_archive.infolist();
		auto right_info = right_archive.infolist();

		if (left_info.size() != right_info.size())
        {
            std::cout << "left has a different number of files than right" << std::endl;

            std::cout << "left has: ";
            for (auto &info : left_info)
            {
                std::cout << info.filename.string() << ", ";
            }
            std::cout << std::endl;

            std::cout << "right has: ";
            for (auto &info : right_info)
            {
                std::cout << info.filename.string() << ", ";
            }
            std::cout << std::endl;
        }
        
        bool match = true;

		for (auto left_member : left_info)
		{
			if (!right_archive.has_file(left_member))
            {
                match = false;
                std::cout << "right is missing file: " << left_member.filename.string() << std::endl;
                continue;
            }

			auto left_member_contents = left_archive.read(left_member);
			auto right_member_contents = right_archive.read(left_member.filename);

			if (!strings_match(left_member_contents, right_member_contents))
			{
				std::cout << left_member.filename.string() << std::endl;
                match = false;
			}
		}

		return match;
	}

	static bool workbooks_match(const xlnt::workbook &left, const xlnt::workbook &right)
	{
		std::vector<std::uint8_t> buffer;

		left.save(buffer);
		xlnt::zip_file left_archive(buffer);

		buffer.clear();

		right.save(buffer);
		xlnt::zip_file right_archive(buffer);

		return archives_match(left_archive, right_archive);
	}
    
    static comparison_result compare_xml_nodes(const pugi::xml_node &left, const pugi::xml_node &right)
    {
        std::string left_temp = left.name();
        std::string right_temp = right.name();
        
        if(left_temp != right_temp)
        {
            return {difference_type::names_differ, left_temp, right_temp};
        }
        
        for(auto &left_attribute : left.attributes())
        {
            left_temp = left_attribute.name();

            if(!right.attribute(left_attribute.name()))
            {
                return {difference_type::missing_attribute, left_temp, "((empty))"};
            }
            
            left_temp = left_attribute.name();
            right_temp = right.attribute(left_attribute.name()).name();
            
            if(left_temp != right_temp)
            {
                return {difference_type::attribute_values_differ, left_temp, right_temp};
            }
        }

		for (auto &right_attribute : right.attributes())
		{
			right_temp = right_attribute.name();

			if (!left.attribute(right_attribute.name()))
			{
				return{ difference_type::missing_attribute, "((empty))", right_temp };
			}

			right_temp = right_attribute.name();
			left_temp = left.attribute(right_attribute.name()).name();

			if (left_temp != right_temp)
			{
				return{ difference_type::attribute_values_differ, left_temp, right_temp };
			}
		}
        
        if(left.text())
        {
            left_temp = left.text().get();
            
            if(!right.text())
            {
                return {difference_type::missing_text, left_temp, "((empty))"};
            }
            
            right_temp = right.text().get();
            
            if(left_temp != right_temp)
            {
                return {difference_type::text_values_differ, left_temp, right_temp};
            }
        }
        else if(right.text())
        {
            right_temp = right.text().get();
            return {difference_type::text_values_differ, "((empty))", right_temp};
        }
        
        auto right_children = right.children();
        auto right_child_iter = right_children.begin();
        
        for(auto left_child : left.children())
        {
            left_temp = left_child.name();
            
            if(right_child_iter == right_children.end())
            {
                return {difference_type::child_order_differs, left_temp, "((end))"};
            }
            
            auto right_child = *right_child_iter;
            right_child_iter++;
            
            auto child_comparison_result = compare_xml_nodes(left_child, right_child);
            
            if(!child_comparison_result)
            {
                return child_comparison_result;
            }
        }
        
        if(right_child_iter != right_children.end())
        {
            right_temp = right_child_iter->name();
            return {difference_type::child_order_differs, "((end))", right_temp};
        }
        
        return {difference_type::equivalent, "", ""};
    }
};
