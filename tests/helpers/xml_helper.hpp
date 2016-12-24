#pragma once

#include <sstream>

#include <detail/include_libstudxml.hpp>
#include <detail/vector_streambuf.hpp>
#include <detail/zip.hpp>
#include <helpers/path_helper.hpp>
#include <xlnt/packaging/manifest.hpp>

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

    static bool compare_files(const std::string &left,
		const std::string &right, const std::string &content_type)
    {
        auto is_xml = (content_type.substr(0, 12) == "application/"
            && content_type.substr(content_type.size() - 4) == "+xml")
            || content_type == "application/xml"
            || content_type == "[Content_Types].xml"
            || content_type == "application/vnd.openxmlformats-officedocument.vmlDrawing";
        
        if (is_xml)
        {
            return compare_xml_exact(left, right);
        }
        
        return left == right;
    }

    static bool compare_xml_exact(const std::string &left, const std::string &right, bool suppress_debug_info = false)
    {
        xml::parser left_parser(left.data(), left.size(), "left");
        xml::parser right_parser(right.data(), right.size(), "right");

        bool difference = false;
        auto right_iter = right_parser.begin();

        auto is_whitespace = [](const std::string &v)
        {
            return v.find_first_not_of("\n\t ") == std::string::npos;
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

	static bool string_matches_workbook_part(const std::string &expected,
		xlnt::workbook &wb, const xlnt::path &part, const std::string &content_type)
	{
		std::vector<std::uint8_t> bytes;
		wb.save(bytes);
		std::istringstream file_stream(std::string(bytes.begin(), bytes.end()));
		xlnt::detail::zip_file_reader archive(file_stream);

		return string_matches_archive_member(expected, archive, part, content_type);
	}

	static bool file_matches_workbook_part(const xlnt::path &expected,
		xlnt::workbook &wb, const xlnt::path &part, const std::string &content_type)
	{
		std::vector<std::uint8_t> bytes;
		wb.save(bytes);
		std::istringstream file_stream(std::string(bytes.begin(), bytes.end()));
		xlnt::detail::zip_file_reader archive(file_stream);

		return file_matches_archive_member(expected, archive, part, content_type);
	}

	static bool string_matches_archive_member(const std::string &expected,
		xlnt::detail::zip_file_reader &archive,
		const xlnt::path &member,
        const std::string &content_type)
	{
		auto streambuf = archive.open(member);
        std::istream stream(streambuf.get());
		std::string contents((std::istreambuf_iterator<char>(stream)), (std::istreambuf_iterator<char>()));
		return compare_files(expected, contents, content_type);
	}
    
    static bool file_matches_archive_member(const xlnt::path &file,
		xlnt::detail::zip_file_reader &archive,
		const xlnt::path &member,
        const std::string &content_type)
	{
		if (!archive.has_file(member)) return false;
		std::vector<std::uint8_t> member_data;
		xlnt::detail::vector_ostreambuf member_data_buffer(member_data);
		std::ostream member_data_stream(&member_data_buffer);
        auto member_streambuf = archive.open(member);
        std::ostream member_stream(member_streambuf.get());
		member_data_stream << member_stream.rdbuf();
		std::string contents(member_data.begin(), member_data.end());
		return compare_files(file.read_contents(), contents, content_type);
	}

	static bool xlsx_archives_match(const std::vector<std::uint8_t> &left, const std::vector<std::uint8_t> &right)
	{
        xlnt::detail::vector_istreambuf left_buffer(left);
        std::istream left_stream(&left_buffer);
        xlnt::detail::zip_file_reader left_archive(left_stream);

		const auto left_info = left_archive.files();

        xlnt::detail::vector_istreambuf right_buffer(right);
        std::istream right_stream(&right_buffer);
        xlnt::detail::zip_file_reader right_archive(right_stream);

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
                continue;
            }

            auto left_member_streambuf = left_archive.open(left_member);
            std::istream left_member_stream(left_member_streambuf.get());
            std::vector<std::uint8_t> left_contents_raw;
            xlnt::detail::vector_ostreambuf left_contents_buffer(left_contents_raw);
            std::ostream left_contents_stream(&left_contents_buffer);
            left_contents_stream << left_member_stream.rdbuf();
            std::string left_member_contents(left_contents_raw.begin(), left_contents_raw.end());

            auto right_member_streambuf = right_archive.open(left_member);
            std::istream right_member_stream(right_member_streambuf.get());
            std::vector<std::uint8_t> right_contents_raw;
            xlnt::detail::vector_ostreambuf right_contents_buffer(right_contents_raw);
            std::ostream right_contents_stream(&right_contents_buffer);
            right_contents_stream << right_member_stream.rdbuf();
            std::string right_member_contents(right_contents_raw.begin(), right_contents_raw.end());

            std::string left_content_type, right_content_type;
            
            if (left_member.string() != "[Content_Types].xml")
            {
                left_content_type = left_manifest.content_type(left_member);
                right_content_type = right_manifest.content_type(left_member);
            }
            else
            {
                left_content_type = right_content_type = "[Content_Types].xml";
            }
            
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
            }
            else if (!compare_files(left_member_contents, right_member_contents, left_content_type))
			{
				std::cout << left_member.string() << std::endl;
                match = false;
			}
		}

		return match;
	}
};
