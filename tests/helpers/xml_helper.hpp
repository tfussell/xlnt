#pragma once

#include <sstream>

#include <detail/include_libstudxml.hpp>
#include <detail/vector_streambuf.hpp>
#include <detail/zstream.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/workbook/workbook.hpp>

class xml_helper
{
public:
    static bool compare_files(const std::string &left,
		const std::string &right, const std::string &content_type)
    {
        // content types are stored in unordered maps, too complicated to compare
        if (content_type == "[Content_Types].xml") return true;

        auto is_xml = (content_type.substr(0, 12) == "application/"
            && content_type.substr(content_type.size() - 4) == "+xml")
            || content_type == "application/xml"
            || content_type == "[Content_Types].xml"
            || content_type == "application/vnd.openxmlformats-officedocument.vmlDrawing";

        return is_xml ? compare_xml_exact(left, right) : left == right;
    }

    static bool compare_xml_exact(const std::string &left,
        const std::string &right, bool suppress_debug_info = false)
    {
        xml::parser left_parser(left.data(), left.size(), "left");
        xml::parser right_parser(right.data(), right.size(), "right");

        bool difference = false;
        auto right_iter = right_parser.begin();

        auto is_whitespace = [](const std::string &v)
        {
            return v.find_first_not_of("\n\r\t ") == std::string::npos;
        };

        for (auto left_event : left_parser)
        {
            if (left_event == xml::parser::event_type::characters
                && is_whitespace(left_parser.value())) continue;

            if (right_iter == right_parser.end())
            {
                difference = true;
                break;
            }

            auto right_event = *right_iter;

            while (right_iter != right_parser.end()
                && right_event == xml::parser::event_type::characters
                && is_whitespace(right_parser.value()))
            {
                ++right_iter;
                right_event = *right_iter;
            }

            if (left_event != right_event)
            {
                difference = true;
                break;
            }

            if (left_event == xml::parser::event_type::start_element)
            {
                auto left_attr_map = left_parser.attribute_map();
                auto right_attr_map = right_parser.attribute_map();

                for (auto attr : left_attr_map)
                {
                    if (right_attr_map.find(attr.first) == right_attr_map.end())
                    {
                        difference = true;
                        break;
                    }

                    if (attr.second.value != right_attr_map.at(attr.first).value)
                    {
                        difference = true;
                        break;
                    }
                }

                for (auto attr : right_attr_map)
                {
                    if (left_attr_map.find(attr.first) == left_attr_map.end())
                    {
                        difference = true;
                        break;
                    }

                    if (attr.second.value != left_attr_map.at(attr.first).value)
                    {
                        difference = true;
                        break;
                    }
                }

                if (difference)
                {
                    break;
                }
                
                if (left_parser.qname() != right_parser.qname())
                {
                    difference = true;
                    break;
                }
            }
            else if (left_event == xml::parser::event_type::characters)
            {
                if (left_parser.value() != right_parser.value())
                {
                    difference = true;
                    break;
                }
            }

            ++right_iter;
        }

		if (difference && !suppress_debug_info)
		{
			std::cout << "documents don't match" << std::endl;

			std::cout << "left:" << std::endl;
            for (auto c : left)
            {
                std::cout << c << std::flush;
            }
			std::cout << std::endl;

			std::cout << "right:" << std::endl;
            for (auto c : right)
            {
                std::cout << c << std::flush;
            }
			std::cout << std::endl;
		}

		return !difference;
    }

	static bool xlsx_archives_match(const std::vector<std::uint8_t> &left,
        const std::vector<std::uint8_t> &right)
	{
        xlnt::detail::vector_istreambuf left_buffer(left);
        std::istream left_stream(&left_buffer);
        xlnt::detail::izstream left_archive(left_stream);

		const auto left_info = left_archive.files();

        xlnt::detail::vector_istreambuf right_buffer(right);
        std::istream right_stream(&right_buffer);
        xlnt::detail::izstream right_archive(right_stream);

		const auto right_info = right_archive.files();

		if (left_info.size() != right_info.size())
        {
            std::cout << "left has a different number of files than right" << std::endl;

            std::cout << "left has: ";
            for (auto &info : left_info)
            {
                std::cout << info.string() << ", ";
            }
            std::cout << std::endl;

            std::cout << "right has: ";
            for (auto &info : right_info)
            {
                std::cout << info.string() << ", ";
            }
            std::cout << std::endl;
        }
        
        bool match = true;

        xlnt::workbook left_workbook;
        left_workbook.load(left);

        xlnt::workbook right_workbook;
        right_workbook.load(right);
        
        auto &left_manifest = left_workbook.manifest();
        auto &right_manifest = right_workbook.manifest();

		for (auto left_member : left_info)
		{
			if (!right_archive.has_file(left_member))
            {
                match = false;
                std::cout << "right is missing file: " << left_member.string() << std::endl;
                break;
            }

            auto left_content_type = left_member.string() == "[Content_Types].xml"
                ? "[Content_Types].xml" : left_manifest.content_type(left_member);
            auto right_content_type = left_member.string() == "[Content_Types].xml"
                ? "[Content_Types].xml" : right_manifest.content_type(left_member);

            if (left_content_type != right_content_type)
            {
                std::cout << "content types differ: "
                    << left_member.string()
                    << " "
                    << left_content_type
                    << " "
                    << right_content_type
                    << std::endl;
                match = false;
                break;
            }

            if (!compare_files(left_archive.read(left_member),
                right_archive.read(left_member), left_content_type))
			{
				std::cout << left_member.string() << std::endl;
                match = false;
                break;
			}
		}

		return match;
	}
};
