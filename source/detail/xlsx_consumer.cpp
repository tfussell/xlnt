#include <cctype>

#include <detail/xlsx_consumer.hpp>

#include <detail/constants.hpp>
#include <detail/workbook_impl.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

bool is_true(const std::string &bool_string)
{
	return bool_string == "1" || bool_string == "true";
}

std::size_t string_to_size_t(const std::string &s)
{
#if ULLONG_MAX == SIZE_MAX
	return std::stoull(s);
#else
	return std::stoul(s);
#endif
}

xlnt::datetime w3cdtf_to_datetime(const std::string &string)
{
	xlnt::datetime result(1900, 1, 1);
	auto separator_index = string.find('-');
	result.year = std::stoi(string.substr(0, separator_index));
	result.month = std::stoi(string.substr(separator_index + 1, string.find('-', separator_index + 1)));
	separator_index = string.find('-', separator_index + 1);
	result.day = std::stoi(string.substr(separator_index + 1, string.find('T', separator_index + 1)));
	separator_index = string.find('T', separator_index + 1);
	result.hour = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
	separator_index = string.find(':', separator_index + 1);
	result.minute = std::stoi(string.substr(separator_index + 1, string.find(':', separator_index + 1)));
	separator_index = string.find(':', separator_index + 1);
	result.second = std::stoi(string.substr(separator_index + 1, string.find('Z', separator_index + 1)));
	return result;
}

struct EnumClassHash
{
	template <typename T>
	std::size_t operator()(T t) const
	{
		return static_cast<std::size_t>(t);
	}
};

template<typename T>
T from_string(const std::string &string);

template<>
xlnt::font::underline_style from_string(const std::string &underline_string)
{
	if (underline_string == "double") return xlnt::font::underline_style::double_;
	if (underline_string == "doubleAccounting") return xlnt::font::underline_style::double_accounting;
	if (underline_string == "single") return xlnt::font::underline_style::single;
	if (underline_string == "singleAccounting") return xlnt::font::underline_style::single_accounting;
	return xlnt::font::underline_style::none;
}

template<>
xlnt::pattern_fill_type from_string(const std::string &fill_type)
{
	if (fill_type == "darkdown") return xlnt::pattern_fill_type::darkdown;
	if (fill_type == "darkgray") return xlnt::pattern_fill_type::darkgray;
	if (fill_type == "darkgrid") return xlnt::pattern_fill_type::darkgrid;
	if (fill_type == "darkhorizontal") return xlnt::pattern_fill_type::darkhorizontal;
	if (fill_type == "darktrellis") return xlnt::pattern_fill_type::darktrellis;
	if (fill_type == "darkup") return xlnt::pattern_fill_type::darkup;
	if (fill_type == "darkvertical") return xlnt::pattern_fill_type::darkvertical;
	if (fill_type == "gray0625") return xlnt::pattern_fill_type::gray0625;
	if (fill_type == "gray125") return xlnt::pattern_fill_type::gray125;
	if (fill_type == "lightdown") return xlnt::pattern_fill_type::lightdown;
	if (fill_type == "lightgray") return xlnt::pattern_fill_type::lightgray;
	if (fill_type == "lighthorizontal") return xlnt::pattern_fill_type::lighthorizontal;
	if (fill_type == "lighttrellis") return xlnt::pattern_fill_type::lighttrellis;
	if (fill_type == "lightup") return xlnt::pattern_fill_type::lightup;
	if (fill_type == "lightvertical") return xlnt::pattern_fill_type::lightvertical;
	if (fill_type == "mediumgray") return xlnt::pattern_fill_type::mediumgray;
	if (fill_type == "solid") return xlnt::pattern_fill_type::solid;
	return  xlnt::pattern_fill_type::none;
};

template<>
xlnt::gradient_fill_type from_string(const std::string &fill_type)
{
	//TODO what's the default?
	if (fill_type == "linear") return xlnt::gradient_fill_type::linear;
	return xlnt::gradient_fill_type::path;
};

template<>
xlnt::border_style from_string(const std::string &border_style_string)
{
	if (border_style_string == "dashdot") return xlnt::border_style::dashdot;
	if (border_style_string == "dashdotdot") return xlnt::border_style::dashdotdot;
	if (border_style_string == "dashed") return xlnt::border_style::dashed;
	if (border_style_string == "dotted") return xlnt::border_style::dotted;
	if (border_style_string == "double") return xlnt::border_style::double_;
	if (border_style_string == "hair") return xlnt::border_style::hair;
	if (border_style_string == "medium") return xlnt::border_style::medium;
	if (border_style_string == "mediumdashdot") return xlnt::border_style::mediumdashdot;
	if (border_style_string == "mediumdashdotdot") return xlnt::border_style::mediumdashdotdot;
	if (border_style_string == "mediumdashed") return xlnt::border_style::mediumdashed;
	if (border_style_string == "slantdashdot") return xlnt::border_style::slantdashdot;
	if (border_style_string == "thick") return xlnt::border_style::thick;
	if (border_style_string == "thin") return xlnt::border_style::thin;
	return  xlnt::border_style::none;
}

template<>
xlnt::vertical_alignment from_string(const std::string &vertical_alignment_string)
{
	if (vertical_alignment_string == "bottom") return xlnt::vertical_alignment::bottom;
	if (vertical_alignment_string == "center") return xlnt::vertical_alignment::center;
	if (vertical_alignment_string == "justify") return xlnt::vertical_alignment::justify;
	if (vertical_alignment_string == "top") return xlnt::vertical_alignment::top;
	return  xlnt::vertical_alignment::none;
}

template<>
xlnt::horizontal_alignment from_string(const std::string &horizontal_alignment_string)
{
	if (horizontal_alignment_string == "center") return xlnt::horizontal_alignment::center;
	if (horizontal_alignment_string == "center-continous") return xlnt::horizontal_alignment::center_continuous;
	if (horizontal_alignment_string == "general") return xlnt::horizontal_alignment::general;
	if (horizontal_alignment_string == "justify") return xlnt::horizontal_alignment::justify;
	if (horizontal_alignment_string == "left") return xlnt::horizontal_alignment::left;
	if (horizontal_alignment_string == "right") return xlnt::horizontal_alignment::right;
	return  xlnt::horizontal_alignment::none;
}

template<>
xlnt::relationship::type from_string(const std::string &type_string)
{
	if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
	{
		return xlnt::relationship::type::extended_properties;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
	{
		return xlnt::relationship::type::custom_properties;
	}
	else if (type_string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
	{
		return xlnt::relationship::type::core_properties;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
	{
		return xlnt::relationship::type::office_document;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
	{
		return xlnt::relationship::type::worksheet;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
	{
		return xlnt::relationship::type::shared_string_table;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
	{
		return xlnt::relationship::type::styles;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
	{
		return xlnt::relationship::type::theme;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
	{
		return xlnt::relationship::type::hyperlink;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
	{
		return xlnt::relationship::type::chartsheet;
	}
	else if (type_string == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail")
	{
		return xlnt::relationship::type::thumbnail;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
	{
		return xlnt::relationship::type::calculation_chain;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
	{
		return xlnt::relationship::type::printer_settings;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
	{
		return xlnt::relationship::type::drawings;
	}
	else if (type_string == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
	{
		return xlnt::relationship::type::image;
	}

	return xlnt::relationship::type::unknown;
}

std::string to_string(xlnt::border_side side)
{
	switch (side)
	{
	case xlnt::border_side::bottom: return "bottom";
	case xlnt::border_side::top: return "top";
	case xlnt::border_side::start: return "left";
	case xlnt::border_side::end: return "right";
	case xlnt::border_side::horizontal: return "horizontal";
	case xlnt::border_side::vertical: return "vertical";
	default:
	case xlnt::border_side::diagonal: return "diagonal";
	}
}

xlnt::protection read_protection(xml::parser &parser)
{
	xlnt::protection prot;

    for (auto event : parser)
    {
        if (event == xml::parser::event_type::start_attribute)
        {
            if (parser.name() == "locked")
            {
                prot.locked(is_true(parser.value()));
            }
            else if (parser.name() == "hidden")
            {
                prot.hidden(is_true(parser.value()));
            }
        }
        else
        {
            break;
        }
    }

	return prot;
}

xlnt::alignment read_alignment(xml::parser &parser)
{
	xlnt::alignment align;
/*
	align.wrap(is_true(alignment_node.attribute("wrapText").value()));
	align.shrink(is_true(alignment_node.attribute("shrinkToFit").value()));

	if (alignment_node.attribute("vertical"))
	{
		std::string vertical = alignment_node.attribute("vertical").value();
		align.vertical(from_string<xlnt::vertical_alignment>(vertical));
	}

	if (alignment_node.attribute("horizontal"))
	{
		std::string horizontal = alignment_node.attribute("horizontal").value();
		align.horizontal(from_string<xlnt::horizontal_alignment>(horizontal));
	}
*/
	return align;
}

void read_number_formats(xml::parser &parser, std::vector<xlnt::number_format> &number_formats)
{
	number_formats.clear();
/*
	for (auto num_fmt_node : number_formats_node.children("numFmt"))
	{
		std::string format_string(num_fmt_node.attribute("formatCode").value());

		if (format_string == "GENERAL")
		{
			format_string = "General";
		}

		xlnt::number_format nf;

		nf.set_format_string(format_string);
		nf.set_id(string_to_size_t(num_fmt_node.attribute("numFmtId").value()));

		number_formats.push_back(nf);
	}
*/
}

xlnt::color read_color(xml::parser &parser)
{
	xlnt::color result;
/*
	if (color_node.attribute("auto"))
	{
		return result;
	}

	if (color_node.attribute("rgb"))
	{
		result = xlnt::rgb_color(color_node.attribute("rgb").value());
	}
	else if (color_node.attribute("theme"))
	{
		result = xlnt::theme_color(string_to_size_t(color_node.attribute("theme").value()));
	}
	else if (color_node.attribute("indexed"))
	{
		result = xlnt::indexed_color(string_to_size_t(color_node.attribute("indexed").value()));
	}

	if (color_node.attribute("tint"))
	{
		result.set_tint(color_node.attribute("tint").as_double());
	}
*/
	return result;
}

xlnt::font read_font(xml::parser &parser)
{
	xlnt::font new_font;
/*
	new_font.size(string_to_size_t(font_node.child("sz").attribute("val").value()));
	new_font.name(font_node.child("name").attribute("val").value());

	if (font_node.child("color"))
	{
		new_font.color(read_color(font_node.child("color")));
	}

	if (font_node.child("family"))
	{
		new_font.family(string_to_size_t(font_node.child("family").attribute("val").value()));
	}

	if (font_node.child("scheme"))
	{
		new_font.scheme(font_node.child("scheme").attribute("val").value());
	}

	if (font_node.child("b"))
	{
		if (font_node.child("b").attribute("val"))
		{
			new_font.bold(is_true(font_node.child("b").attribute("val").value()));
		}
		else
		{
			new_font.bold(true);
		}
	}

	if (font_node.child("strike"))
	{
		if (font_node.child("strike").attribute("val"))
		{
			new_font.strikethrough(is_true(font_node.child("strike").attribute("val").value()));
		}
		else
		{
			new_font.strikethrough(true);
		}
	}

	if (font_node.child("i"))
	{
		if (font_node.child("i").attribute("val"))
		{
			new_font.italic(is_true(font_node.child("i").attribute("val").value()));
		}
		else
		{
			new_font.italic(true);
		}
	}

	if (font_node.child("u"))
	{
		if (font_node.child("u").attribute("val"))
		{
			std::string underline_string = font_node.child("u").attribute("val").value();
			new_font.underline(from_string<xlnt::font::underline_style>(underline_string));
		}
		else
		{
			new_font.underline(xlnt::font::underline_style::single);
		}
	}
*/
	return new_font;
}

void read_fonts(xml::parser &parser, std::vector<xlnt::font> &fonts)
{
	fonts.clear();
/*
	for (auto font_node : fonts_node.children())
	{
		fonts.push_back(read_font(font_node));
	}
*/
}

void read_indexed_colors(xml::parser &parser, std::vector<xlnt::color> &colors)
{
/*
	for (auto color_node : indexed_colors_node.children())
	{
		colors.push_back(read_color(color_node));
	}
*/
}

void read_colors(xml::parser &parser, std::vector<xlnt::color> &colors)
{
	colors.clear();
/*
	if (colors_node.child("indexedColors"))
	{
		read_indexed_colors(colors_node.child("indexedColors"), colors);
	}
*/
}

xlnt::fill read_fill(xml::parser &parser)
{
	xlnt::fill new_fill;
/*
	if (fill_node.child("patternFill"))
	{
		auto pattern_fill_node = fill_node.child("patternFill");
		std::string pattern_fill_type_string = pattern_fill_node.attribute("patternType").value();

		xlnt::pattern_fill pattern;

		if (!pattern_fill_type_string.empty())
		{
			pattern.type(from_string<xlnt::pattern_fill_type>(pattern_fill_type_string));

			if (pattern_fill_node.child("fgColor"))
			{
				pattern.foreground(read_color(pattern_fill_node.child("fgColor")));
			}

			if (pattern_fill_node.child("bgColor"))
			{
				pattern.background(read_color(pattern_fill_node.child("bgColor")));
			}
		}

		new_fill = pattern;
	}
	else if (fill_node.child("gradientFill"))
	{
		auto gradient_fill_node = fill_node.child("gradientFill");
		std::string gradient_fill_type_string = gradient_fill_node.attribute("type").value();

		xlnt::gradient_fill gradient;

		if (!gradient_fill_type_string.empty())
		{
			gradient.type(from_string<xlnt::gradient_fill_type>(gradient_fill_type_string));
		}
		else
		{
			gradient.type(xlnt::gradient_fill_type::linear);
		}

		for (auto stop_node : gradient_fill_node.children("stop"))
		{
			auto position = stop_node.attribute("position").as_double();
			auto color = read_color(stop_node.child("color"));

			gradient.add_stop(position, color);
		}

		new_fill = gradient;
	}
*/
	return new_fill;
}

void read_fills(xml::parser &parser, std::vector<xlnt::fill> &fills)
{
	fills.clear();
/*
	for (auto fill_node : fills_node.children())
	{
		fills.emplace_back();
		fills.back() = read_fill(fill_node);
	}
*/
}

xlnt::border::border_property read_side(xml::parser &parser)
{
	xlnt::border::border_property new_side;
/*
	if (side_node.attribute("style"))
	{
		new_side.style(from_string<xlnt::border_style>(side_node.attribute("style").value()));
	}

	if (side_node.child("color"))
	{
		new_side.color(read_color(side_node.child("color")));
	}
*/
	return new_side;
}

xlnt::border read_border(xml::parser &parser)
{
	xlnt::border new_border;
/*
	for (const auto &side : xlnt::border::all_sides())
	{
		auto side_name = to_string(side);

		if (border_node.child(side_name.c_str()))
		{
			new_border.side(side, read_side(border_node.child(side_name.c_str())));
		}
	}
*/
	return new_border;
}

void read_borders(xml::parser &parser, std::vector<xlnt::border> &borders)
{
	borders.clear();
/*
	for (auto border_node : borders_node.children())
	{
		borders.push_back(read_border(border_node));
	}
*/
}

std::vector<xlnt::relationship> read_relationships(const xlnt::path &part, xlnt::zip_file &archive)
{
	std::vector<xlnt::relationship> relationships;
	if (!archive.has_file(part)) return relationships;

    std::istringstream rels_stream(archive.read(part));
    xml::parser parser(rels_stream, part.string());

    xlnt::uri source(part.string());

    const auto xmlns = xlnt::constants::get_namespace("relationships");
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "Relationships");
    parser.content(xml::content::complex);

    while (true)
	{
        if (parser.peek() == xml::parser::event_type::end_element) break;
        
        parser.next_expect(xml::parser::event_type::start_element, xmlns, "Relationship");

        std::string id(parser.attribute("Id"));
        std::string type_string(parser.attribute("Type"));
        xlnt::uri target(parser.attribute("Target"));

		relationships.emplace_back(id, from_string<xlnt::relationship::type>(type_string),
            source, target, xlnt::target_mode::internal);

        parser.next_expect(xml::parser::event_type::end_element,
            xlnt::constants::get_namespace("relationships"), "Relationship");
	}
    
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "Relationships");

	return relationships;
}

void check_document_type(const std::string &document_content_type)
{
	if (document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
		&& document_content_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml ")
	{
		throw xlnt::invalid_file(document_content_type);
	}
}

} // namespace

namespace xlnt {
namespace detail {

xlsx_consumer::xlsx_consumer(workbook &destination) : destination_(destination)
{
}

void xlsx_consumer::read(const path &source)
{
	destination_.clear();
	source_.load(source);
	populate_workbook();
}

void xlsx_consumer::read(std::istream &source)
{
	destination_.clear();
	source_.load(source);
	populate_workbook();
}

void xlsx_consumer::read(const std::vector<std::uint8_t> &source)
{
	destination_.clear();
	source_.load(source);
	populate_workbook();
}

// Part Writing Methods

void xlsx_consumer::populate_workbook()
{
	auto &manifest = destination_.get_manifest();
	read_manifest();

	for (const auto &rel : manifest.get_relationships(path("/")))
	{
        std::istringstream parser_stream(source_.read(rel.get_target().get_path()));
        xml::parser parser(parser_stream, rel.get_target().get_path().string());

		switch (rel.get_type())
		{
		case relationship::type::core_properties:
			read_core_properties(parser);
			break;
		case relationship::type::extended_properties:
			read_extended_properties(parser);
			break;
		case relationship::type::custom_properties:
			read_custom_property(parser);
			break;
		case relationship::type::office_document:
            check_document_type(manifest.get_content_type(rel.get_target().get_path()));
			read_workbook(parser);
			break;
		case relationship::type::calculation_chain:
			read_calculation_chain(parser);
			break;
		case relationship::type::connections:
			read_connections(parser);
			break;
		case relationship::type::custom_xml_mappings:
			read_custom_xml_mappings(parser);
			break;
		case relationship::type::external_workbook_references:
			read_external_workbook_references(parser);
			break;
		case relationship::type::metadata:
			read_metadata(parser);
			break;
		case relationship::type::pivot_table:
			read_pivot_table(parser);
			break;
		case relationship::type::shared_string_table:
			read_shared_string_table(parser);
			break;
		case relationship::type::shared_workbook_revision_headers:
			read_shared_workbook_revision_headers(parser);
			break;
		case relationship::type::styles:
			read_stylesheet(parser);
			break;
		case relationship::type::theme:
			read_theme(parser);
			break;
		case relationship::type::volatile_dependencies:
			read_volatile_dependencies(parser);
			break;
        default:
            break;
		}
	}

    const auto workbook_rel = manifest.get_relationship(path("/"), relationship::type::office_document);

    // First pass of workbook relationship parts which must be read before sheets (e.g. shared strings)
    
	for (const auto &rel : manifest.get_relationships(workbook_rel.get_target().get_path()))
	{
		path part_path(rel.get_source().get_path().parent().append(rel.get_target().get_path()));
        std::istringstream parser_stream(source_.read(part_path));
        xml::parser parser(parser_stream, rel.get_target().get_path().string());

		switch (rel.get_type())
		{
        case relationship::type::shared_string_table:
            read_shared_string_table(parser);
            break;
        case relationship::type::styles:
            read_stylesheet(parser);
            break;
        case relationship::type::theme:
            read_theme(parser);
            break;
        default:
            break;
		}
	}
    
    // Second pass, read sheets themselves

	for (const auto &rel : manifest.get_relationships(workbook_rel.get_target().get_path()))
    {
		path part_path(rel.get_source().get_path().parent().append(rel.get_target().get_path()));
        std::istringstream parser_stream(source_.read(part_path));
        xml::parser parser(parser_stream, rel.get_target().get_path().string());

		switch (rel.get_type())
		{
		case relationship::type::chartsheet:
			read_chartsheet(rel.get_id(), parser);
			break;
		case relationship::type::dialogsheet:
			read_dialogsheet(rel.get_id(), parser);
			break;
		case relationship::type::worksheet:
			read_worksheet(rel.get_id(), parser);
			break;
        default:
            break;
		}
	}

	// Unknown Parts

	void read_unknown_parts();
	void read_unknown_relationships();
}

// Package Parts

void xlsx_consumer::read_manifest()
{
	path package_rels_path("_rels/.rels");
	if (!source_.has_file(package_rels_path)) throw invalid_file("missing package rels");
	auto package_rels = read_relationships(package_rels_path, source_);

    std::istringstream parser_stream(source_.read(path("[Content_Types].xml")));
    xml::parser parser(parser_stream, "[Content_Types].xml");
    
	auto &manifest = destination_.get_manifest();
    
    const std::string xmlns = "http://schemas.openxmlformats.org/package/2006/content-types";
    parser.next_expect(xml::parser::event_type::start_element, xmlns, "Types");
    parser.content(xml::content::complex);

    while (true)
	{
        if (parser.peek() == xml::parser::event_type::end_element) break;
        
        parser.next_expect(xml::parser::event_type::start_element);
        
        if (parser.name() == "Default")
        {
            manifest.register_default_type(parser.attribute("Extension"),
				parser.attribute("ContentType"));
            parser.next_expect(xml::parser::event_type::end_element, xmlns, "Default");
        }
        else if (parser.name() == "Override")
        {
			manifest.register_override_type(path(parser.attribute("PartName")),
                parser.attribute("ContentType"));
            parser.next_expect(xml::parser::event_type::end_element, xmlns, "Override");
        }
	}
    
    parser.next_expect(xml::parser::event_type::end_element, xmlns, "Types");

	for (const auto &package_rel : package_rels)
	{
		manifest.register_relationship(uri("/"),
            package_rel.get_type(),
			package_rel.get_target(),
			package_rel.get_target_mode(), 
			package_rel.get_id());
	}

	for (const auto &relationship_source : source_.infolist())
	{
		if (relationship_source.filename == path("_rels/.rels") 
			|| relationship_source.filename.extension() != "rels") continue;

		path part(relationship_source.filename.parent().parent());
		part = part.append(relationship_source.filename.split_extension().first);
		uri source(part.string());

		path source_directory = part.parent();

		auto part_rels = read_relationships(relationship_source.filename, source_);

		for (const auto part_rel : part_rels)
		{
			path target_path(source_directory.append(part_rel.get_target().get_path()));
			manifest.register_relationship(source, part_rel.get_type(),
				part_rel.get_target(), part_rel.get_target_mode(), part_rel.get_id());
		}
	}
}

void xlsx_consumer::read_extended_properties(xml::parser &parser)
{
/*
	for (auto property_node : root.child("Properties"))
	{
		std::string name(property_node.name());
		std::string value(property_node.text().get());

		if (name == "Application") destination_.set_application(value);
		else if (name == "DocSecurity") destination_.set_doc_security(std::stoi(value));
		else if (name == "ScaleCrop") destination_.set_scale_crop(is_true(value));
		else if (name == "Company") destination_.set_company(value);
		else if (name == "SharedDoc") destination_.set_shared_doc(is_true(value));
		else if (name == "HyperlinksChanged") destination_.set_hyperlinks_changed(is_true(value));
		else if (name == "AppVersion") destination_.set_app_version(value);
		else if (name == "Application") destination_.set_application(value);
	}
    */
}

void xlsx_consumer::read_core_properties(xml::parser &parser)
{
/*
	for (auto property_node : root.child("cp:coreProperties"))
	{
		std::string name(property_node.name());
		std::string value(property_node.text().get());

		if (name == "dc:creator") destination_.set_creator(value);
		else if (name == "cp:lastModifiedBy") destination_.set_last_modified_by(value);
		else if (name == "dcterms:created") destination_.set_created(w3cdtf_to_datetime(value));
		else if (name == "dcterms:modified") destination_.set_modified(w3cdtf_to_datetime(value));
	}
    */
}

void xlsx_consumer::read_custom_file_properties(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("customFileProperties");
    */
}

// Write SpreadsheetML-Specific Package Parts

void xlsx_consumer::read_workbook(xml::parser &parser)
{
/*
	auto workbook_node = root.child("workbook");

	if (workbook_node.attribute("xmlns:x15"))
	{
		destination_.enable_x15();
	}

	if (workbook_node.child("fileVersion"))
	{
		auto file_version_node = workbook_node.child("fileVersion");

		destination_.d_->has_file_version_ = true;
		destination_.d_->file_version_.app_name = file_version_node.attribute("appName").value();
		destination_.d_->file_version_.last_edited = string_to_size_t(file_version_node.attribute("lastEdited").value());
		destination_.d_->file_version_.lowest_edited = string_to_size_t(file_version_node.attribute("lowestEdited").value());
		destination_.d_->file_version_.rup_build = string_to_size_t(file_version_node.attribute("rupBuild").value());
	}

	if (workbook_node.child("mc:AlternateContent"))
	{
		destination_.set_absolute_path(path(workbook_node.child("mc:AlternateContent")
			.child("mc:Choice").child("x15ac:absPath").attribute("url").value()));
	}

	if (workbook_node.child("bookViews"))
	{
		auto book_views_node = workbook_node.child("bookViews");

		if (book_views_node.child("workbookView"))
		{
			auto book_view_node = book_views_node.child("workbookView");
			workbook_view view;
			view.x_window = string_to_size_t(book_view_node.attribute("xWindow").value());
			view.y_window = string_to_size_t(book_view_node.attribute("yWindow").value());
			view.window_width = string_to_size_t(book_view_node.attribute("windowWidth").value());
			view.window_height = string_to_size_t(book_view_node.attribute("windowHeight").value());
			view.tab_ratio = string_to_size_t(book_view_node.attribute("tabRatio").value());
			destination_.set_view(view);
		}
	}

	if (workbook_node.child("workbookPr"))
	{
		destination_.d_->has_properties_ = true;

		auto workbook_pr_node = workbook_node.child("workbookPr");

		if (workbook_pr_node.attribute("date1904"))
		{
			std::string value = workbook_pr_node.attribute("date1904").value();

			if (value == "1" || value == "true")
			{
				destination_.set_base_date(xlnt::calendar::mac_1904);
			}
		}
	}

	std::size_t index = 0;
	for (auto sheet_node : workbook_node.child("sheets").children("sheet"))
	{
		std::string rel_id(sheet_node.attribute("r:id").value());
		std::string title(sheet_node.attribute("name").value());
		auto id = string_to_size_t(sheet_node.attribute("sheetId").value());

		sheet_title_id_map_[title] = id;
		sheet_title_index_map_[title] = index++;
		destination_.d_->sheet_title_rel_id_map_[title] = rel_id;
	}

	if (workbook_node.child("calcPr"))
	{
		destination_.d_->has_calculation_properties_ = true;
	}

	if (workbook_node.child("extLst"))
	{
		destination_.d_->has_arch_id_ = true;
	}
    */
}

// Write Workbook Relationship Target Parts

void xlsx_consumer::read_calculation_chain(xml::parser &parser)
{
}

void xlsx_consumer::read_chartsheet(const std::string &title, xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("chartsheet");
    */
}

void xlsx_consumer::read_connections(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("connections");
    */
}

void xlsx_consumer::read_custom_property(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("customProperty");
    */
}

void xlsx_consumer::read_custom_xml_mappings(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("connections");
    */
}

void xlsx_consumer::read_dialogsheet(const std::string &title, xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("dialogsheet");
    */
}

void xlsx_consumer::read_external_workbook_references(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("externalLink");
    */
}

void xlsx_consumer::read_metadata(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("metadata");
    */
}

void xlsx_consumer::read_pivot_table(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("pivotTableDefinition");
    */
}

void xlsx_consumer::read_shared_string_table(xml::parser &parser)
{
/*
	auto sst_node = root.child("sst");
	std::size_t unique_count = 0;

	if (sst_node.attribute("uniqueCount"))
	{
		unique_count = string_to_size_t(sst_node.attribute("uniqueCount").value());
	}

	auto &strings = destination_.get_shared_strings();

	for (const auto string_item_node : sst_node.children("si"))
	{
		if (string_item_node.child("t"))
		{
			text t;
			t.set_plain_string(string_item_node.child("t").text().get());
			strings.push_back(t);
		}
		else if (string_item_node.child("r")) // possible multiple text entities.
		{
			text t;

			for (const auto& rich_text_run_node : string_item_node.children("r"))
			{
				if (rich_text_run_node.child("t"))
				{
					text_run run;

					run.set_string(rich_text_run_node.child("t").text().get());

					if (rich_text_run_node.child("rPr"))
					{
						auto run_properties_node = rich_text_run_node.child("rPr");

						if (run_properties_node.child("sz"))
						{
							run.set_size(string_to_size_t(run_properties_node.child("sz").attribute("val").value()));
						}

						if (run_properties_node.child("rFont"))
						{
							run.set_font(run_properties_node.child("rFont").attribute("val").value());
						}

						if (run_properties_node.child("color"))
						{
							run.set_color(run_properties_node.child("color").attribute("rgb").value());
						}

						if (run_properties_node.child("family"))
						{
							run.set_family(string_to_size_t(run_properties_node.child("family").attribute("val").value()));
						}

						if (run_properties_node.child("scheme"))
						{
							run.set_scheme(run_properties_node.child("scheme").attribute("val").value());
						}
					}

					t.add_run(run);
				}
			}

			strings.push_back(t);
		}
	}

	if (unique_count != strings.size())
	{
		throw invalid_file("sizes don't match");
	}
    */
}

void xlsx_consumer::read_shared_workbook_revision_headers(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("headers");
    */
}

void xlsx_consumer::read_shared_workbook(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("revisions");
    */
}

void xlsx_consumer::read_shared_workbook_user_data(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("users");
    */
}

void xlsx_consumer::read_stylesheet(xml::parser &parser)
{
/*
	auto stylesheet_node = root.child("styleSheet");
	auto &stylesheet = destination_.impl().stylesheet_;

	read_borders(stylesheet_node.child("borders"), stylesheet.borders);
	read_fills(stylesheet_node.child("fills"), stylesheet.fills);
	read_fonts(stylesheet_node.child("fonts"), stylesheet.fonts);
	read_number_formats(stylesheet_node.child("numFmts"), stylesheet.number_formats);
	read_colors(stylesheet_node.child("colors"), stylesheet.colors);

	auto cell_styles_node = stylesheet_node.child("cellStyles");
	auto cell_style_range = cell_styles_node.children("cellStyle");
	auto cell_style_iter = cell_style_range.begin();
	auto cell_style_xfs_node = stylesheet_node.child("cellStyleXfs");

	for (auto xf_node : cell_style_xfs_node.children("xf"))
	{
		auto cell_style_node = *(cell_style_iter++);
		auto &style = destination_.create_style(cell_style_node.attribute("name").value());
        
        style.builtin_id(string_to_size_t(cell_style_node.attribute("builtinId").value()));

		// Alignment
		auto alignment_node = xf_node.child("alignment");
		auto apply_alignment_attr = xf_node.attribute("applyAlignment");
		auto alignment_applied = apply_alignment_attr ? is_true(xf_node.attribute("applyAlignment").value()) : alignment_node != nullptr;
		auto alignment = alignment_node ? read_alignment(alignment_node) : xlnt::alignment();
		style.alignment(alignment, alignment_applied);

		// Border
		auto border_applied = is_true(xf_node.attribute("applyBorder").value());
		auto border_index = xf_node.attribute("borderId") ? string_to_size_t(xf_node.attribute("borderId").value()) : 0;
		auto border = stylesheet.borders.at(border_index);
		style.border(stylesheet.borders.at(border_index), border_applied);

		// Fill
		auto fill_applied = is_true(xf_node.attribute("applyFill").value());
		auto fill_index = xf_node.attribute("fillId") ? string_to_size_t(xf_node.attribute("fillId").value()) : 0;
		auto fill = stylesheet.fills.at(fill_index);
		style.fill(fill, fill_applied);

		// Font
		auto font_applied = is_true(xf_node.attribute("applyFont").value());
		auto font_index = xf_node.attribute("fontId") ? string_to_size_t(xf_node.attribute("fontId").value()) : 0;
		auto font = stylesheet.fonts.at(font_index);
		style.font(font, font_applied);

		// Number Format
		auto number_format_applied = is_true(xf_node.attribute("applyNumberFormat").value());
		auto number_format_id = xf_node.attribute("numFmtId") ? string_to_size_t(xf_node.attribute("numFmtId").value()) : 0;
        auto number_format = number_format::general();
        bool is_custom_number_format = false;
        
        for (const auto &nf : stylesheet.number_formats)
        {
            if (nf.get_id() == number_format_id)
            {
                number_format = nf;
                is_custom_number_format = true;
                break;
            }
        }
        if (number_format_id < 164 && !is_custom_number_format)
        {
            number_format = number_format::from_builtin_id(number_format_id);
        }
        
		style.number_format(number_format, number_format_applied);

		// Protection
		auto protection_node = xf_node.child("protection");
		auto apply_protection_attr = xf_node.attribute("applyProtection");
		auto protection_applied = apply_protection_attr ? is_true(apply_protection_attr.value()) : protection_node != nullptr;
		auto protection = protection_node ? read_protection(protection_node) : xlnt::protection();
		style.protection(protection, protection_applied);
	}

	for (auto xf_node : stylesheet_node.child("cellXfs").children("xf"))
	{
		auto &format = destination_.create_format();

		// Alignment
		auto alignment_node = xf_node.child("alignment");
		auto apply_alignment_attr = xf_node.attribute("applyAlignment");
		auto alignment_applied = apply_alignment_attr ? is_true(xf_node.attribute("applyAlignment").value()) : alignment_node != nullptr;
		auto alignment = alignment_node ? read_alignment(alignment_node) : xlnt::alignment();
		format.alignment(alignment, alignment_applied);

		// Border
		auto border_applied = is_true(xf_node.attribute("applyBorder").value());
		auto border_index = xf_node.attribute("borderId") ? string_to_size_t(xf_node.attribute("borderId").value()) : 0;
		auto border = stylesheet.borders.at(border_index);
		format.border(stylesheet.borders.at(border_index), border_applied);

		// Fill
		auto fill_applied = is_true(xf_node.attribute("applyFill").value());
		auto fill_index = xf_node.attribute("fillId") ? string_to_size_t(xf_node.attribute("fillId").value()) : 0;
		auto fill = stylesheet.fills.at(fill_index);
		format.fill(fill, fill_applied);

		// Font
		auto font_applied = is_true(xf_node.attribute("applyFont").value());
		auto font_index = xf_node.attribute("fontId") ? string_to_size_t(xf_node.attribute("fontId").value()) : 0;
		auto font = stylesheet.fonts.at(font_index);
		format.font(font, font_applied);

		// Number Format
		auto number_format_applied = is_true(xf_node.attribute("applyNumberFormat").value());
		auto number_format_id = xf_node.attribute("numFmtId") ? string_to_size_t(xf_node.attribute("numFmtId").value()) : 0;
        auto number_format = number_format::general();
        bool is_custom_number_format = false;
        
        for (const auto &nf : stylesheet.number_formats)
        {
            if (nf.get_id() == number_format_id)
            {
                number_format = nf;
                is_custom_number_format = true;
                break;
            }
        }
        if (number_format_id < 164 && !is_custom_number_format)
        {
            number_format = number_format::from_builtin_id(number_format_id);
        }
        
		format.number_format(number_format, number_format_applied);

		// Protection
		auto protection_node = xf_node.child("protection");
		auto apply_protection_attr = xf_node.attribute("applyProtection");
		auto protection_applied = apply_protection_attr ? is_true(apply_protection_attr.value()) : protection_node != nullptr;
		auto protection = protection_node ? read_protection(protection_node) : xlnt::protection();
		format.protection(protection, protection_applied);
		
		// Style
        if (xf_node.attribute("xfId"))
        {
            auto style_index = string_to_size_t(xf_node.attribute("xfId").value());
            auto style_name = stylesheet.styles.at(style_index).name();
            format.style(stylesheet.get_style(style_name));
        }
	}
    */
}

void xlsx_consumer::read_theme(xml::parser &parser)
{
/*
	root.child("theme");
	destination_.set_theme(theme());
    */
}

void xlsx_consumer::read_volatile_dependencies(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("volTypes");
    */
}

void xlsx_consumer::read_worksheet(const std::string &title, xml::parser &parser)
{
/*
	auto title = std::find_if(destination_.d_->sheet_title_rel_id_map_.begin(),
		destination_.d_->sheet_title_rel_id_map_.end(),
		[&](const std::pair<std::string, std::string> &p)
	{
		return p.second == rel_id;
	})->first;

	auto id = sheet_title_id_map_[title];
	auto index = sheet_title_index_map_[title];

	auto insertion_iter = destination_.d_->worksheets_.begin();
	while (insertion_iter != destination_.d_->worksheets_.end()
		&& sheet_title_index_map_[insertion_iter->title_] < index)
	{
		++insertion_iter;
	}

	destination_.d_->worksheets_.emplace(insertion_iter, &destination_, id, title);

	auto ws = destination_.get_sheet_by_id(id);

	auto worksheet_node = root.child("worksheet");

	if (worksheet_node.attribute("xmlns:x14ac"))
	{
		ws.enable_x14ac();
	}

	xlnt::range_reference full_range;

	if (worksheet_node.child("dimension"))
	{
		std::string dimension(worksheet_node.child("dimension").attribute("ref").value());
		full_range = xlnt::range_reference(dimension);
		ws.d_->has_dimension_ = true;
	}

	if (worksheet_node.child("sheetViews"))
	{
		ws.d_->has_view_ = true;
	}

	if (worksheet_node.child("sheetFormatPr"))
	{
		ws.d_->has_format_properties_ = true;
	}

	if (worksheet_node.child("mergeCells"))
	{
		auto merge_cells_node = worksheet_node.child("mergeCells");
		auto count = std::stoull(merge_cells_node.attribute("count").value());

		for (auto merge_cell_node : merge_cells_node.children("mergeCell"))
		{
			ws.merge_cells(range_reference(merge_cell_node.attribute("ref").value()));
			count--;
		}

		if (count != 0)
		{
			throw invalid_file("sizes don't match");
		}
	}

	auto &shared_strings = destination_.get_shared_strings();
	auto sheet_data_node = worksheet_node.child("sheetData");

	for (auto row_node : sheet_data_node.children("row"))
	{
		auto row_index = static_cast<row_t>(std::stoull(row_node.attribute("r").value()));

		if (row_node.attribute("ht"))
		{
			ws.get_row_properties(row_index).height = std::stold(row_node.attribute("ht").value());
		}

		std::string span_string = row_node.attribute("spans").value();
		auto colon_index = span_string.find(':');

		column_t min_column = 0;
		column_t max_column = 0;

		if (colon_index != std::string::npos)
		{
			min_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(0, colon_index)));
			max_column = static_cast<column_t::index_t>(std::stoll(span_string.substr(colon_index + 1)));
		}
		else
		{
			min_column = full_range.get_top_left().get_column_index();
			max_column = full_range.get_bottom_right().get_column_index();
		}

		for (column_t i = min_column; i <= max_column; i++)
		{
			std::string address = i.column_string() + std::to_string(row_index);

			pugi::xml_node cell_node;
			bool cell_found = false;

			for (auto &check_cell_node : row_node.children("c"))
			{
				if (check_cell_node.attribute("r") && check_cell_node.attribute("r").value() == address)
				{
					cell_node = check_cell_node;
					cell_found = true;
					break;
				}
			}

			if (cell_found)
			{
				bool has_value = cell_node.child("v") != nullptr;
				std::string value_string = has_value ? cell_node.child("v").text().get() : "";

				bool has_type = cell_node.attribute("t") != nullptr;
				std::string type = has_type ? cell_node.attribute("t").value() : "";

				bool has_format = cell_node.attribute("s") != nullptr;
				auto format_id = static_cast<std::size_t>(has_format ? std::stoull(cell_node.attribute("s").value()) : 0LL);

				bool has_formula = cell_node.child("f") != nullptr;
				bool has_shared_formula = has_formula && cell_node.child("f").attribute("t") != nullptr
					&& cell_node.child("f").attribute("t").value() == std::string("shared");

				auto cell = ws.get_cell(address);

				if (has_formula && !has_shared_formula && !ws.get_workbook().get_data_only())
				{
					std::string formula = cell_node.child("f").text().get();
					cell.set_formula(formula);
				}

				if (has_type && type == "inlineStr") // inline string
				{
					std::string inline_string = cell_node.child("is").child("t").text().get();
					cell.set_value(inline_string);
				}
				else if (has_type && type == "s" && !has_formula) // shared string
				{
					auto shared_string_index = static_cast<std::size_t>(std::stoull(value_string));
					auto shared_string = shared_strings.at(shared_string_index);
					cell.set_value(shared_string);
				}
				else if (has_type && type == "b") // boolean
				{
					cell.set_value(value_string != "0");
				}
				else if (has_type && type == "str")
				{
					cell.set_value(value_string);
				}
				else if (has_value && !value_string.empty())
				{
					if (!value_string.empty() && value_string[0] == '#')
					{
						cell.set_error(value_string);
					}
					else
					{
						cell.set_value(std::stold(value_string));
					}
				}

				if (has_format)
				{
					cell.set_format(destination_.get_format(format_id));
				}
			}
		}
	}

	auto cols_node = worksheet_node.child("cols");

	for (auto col_node : cols_node.children("col"))
	{
		auto min = static_cast<column_t::index_t>(std::stoull(col_node.attribute("min").value()));
		auto max = static_cast<column_t::index_t>(std::stoull(col_node.attribute("max").value()));
		auto width = std::stold(col_node.attribute("width").value());
		bool custom = col_node.attribute("customWidth").value() == std::string("1");
		auto column_style = static_cast<std::size_t>(col_node.attribute("style") ? std::stoull(col_node.attribute("style").value()) : 0);

		for (auto column = min; column <= max; column++)
		{
			if (!ws.has_column_properties(column))
			{
				ws.add_column_properties(column, column_properties());
			}

			ws.get_column_properties(min).width = width;
			ws.get_column_properties(min).style = column_style;
			ws.get_column_properties(min).custom = custom;
		}
	}

	if (worksheet_node.child("autoFilter"))
	{
		auto auto_filter_node = worksheet_node.child("autoFilter");
		xlnt::range_reference ref(auto_filter_node.attribute("ref").value());
		ws.auto_filter(ref);
	}

	if (worksheet_node.child("pageMargins"))
	{
		auto page_margins_node = worksheet_node.child("pageMargins");

		page_margins margins;
		margins.set_top(page_margins_node.attribute("top").as_double());
		margins.set_bottom(page_margins_node.attribute("bottom").as_double());
		margins.set_left(page_margins_node.attribute("left").as_double());
		margins.set_right(page_margins_node.attribute("right").as_double());
		margins.set_header(page_margins_node.attribute("header").as_double());
		margins.set_footer(page_margins_node.attribute("footer").as_double());

		ws.set_page_margins(margins);
	}
    */
}

// Sheet Relationship Target Parts

void xlsx_consumer::read_comments(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("comments");
    */
}

void xlsx_consumer::read_drawings(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("wsDr");
    */
}

// Unknown Parts

void xlsx_consumer::read_unknown_parts(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("Hmm");
    */
}

void xlsx_consumer::read_unknown_relationships(xml::parser &parser)
{
/*
	pugi::xml_document document;
	document.load_string(source_.read(path("[Content Types].xml")).c_str());
	document.append_child("Relationships");
    */
}

} // namespace detail
} // namepsace xlnt
