#include <cmath>

#include <detail/xlsx_producer.hpp>
#include <detail/constants.hpp>
#include <detail/include_pugixml.hpp>
#include <detail/workbook_impl.hpp>
#include <xlnt/cell/cell.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/workbook/workbook_view.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

/// <summary>
/// Returns true if d is exactly equal to an integer.
/// </summary>
bool is_integral(long double d)
{
	return d == static_cast<long long int>(d);
}

/// <summary>
/// Serializes document and writes to to the archive at archive_path.
/// </summary>
void write_document_to_archive(const pugi::xml_document &document,
	const xlnt::path &archive_path, xlnt::zip_file &archive)
{
	std::ostringstream out_stream;
	document.save(out_stream);
	archive.write_string(out_stream.str(), archive_path);
}

std::string fill(const std::string &string, std::size_t length = 2)
{
	if (string.size() >= length)
	{
		return string;
	}

	return std::string(length - string.size(), '0') + string;
}

std::string datetime_to_w3cdtf(const xlnt::datetime &dt)
{
	return std::to_string(dt.year) + "-" + fill(std::to_string(dt.month)) + "-" + fill(std::to_string(dt.day)) + "T" +
		fill(std::to_string(dt.hour)) + ":" + fill(std::to_string(dt.minute)) + ":" +
		fill(std::to_string(dt.second)) + "Z";
}

/// <summary>
/// Returns the string representation of the relationship type.
/// </summary>
std::string to_string(xlnt::relationship::type t)
{
	switch (t)
	{
	case xlnt::relationship::type::office_document:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
	case xlnt::relationship::type::thumbnail:
		return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
	case xlnt::relationship::type::calculation_chain:
		return "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain";
	case xlnt::relationship::type::extended_properties:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
	case xlnt::relationship::type::core_properties:
		return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
	case xlnt::relationship::type::worksheet:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
	case xlnt::relationship::type::shared_string_table:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
	case xlnt::relationship::type::styles:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
	case xlnt::relationship::type::theme:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
	case xlnt::relationship::type::hyperlink:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
	case xlnt::relationship::type::chartsheet:
		return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
	default:
		return "??";
	}
}

std::string to_string(xlnt::font::underline_style underline_style)
{
	switch (underline_style)
	{
	case xlnt::font::underline_style::double_: return "double";
	case xlnt::font::underline_style::double_accounting: return "doubleAccounting";
	case xlnt::font::underline_style::single: return "single";
	case xlnt::font::underline_style::single_accounting: return "singleAccounting";
	default:
	case xlnt::font::underline_style::none: return "none";
	}
}

std::string to_string(xlnt::pattern_fill_type fill_type)
{
	switch (fill_type)
	{
	case xlnt::pattern_fill_type::darkdown: return "darkdown";
	case xlnt::pattern_fill_type::darkgray: return "darkgray";
	case xlnt::pattern_fill_type::darkgrid: return "darkgrid";
	case xlnt::pattern_fill_type::darkhorizontal: return "darkhorizontal";
	case xlnt::pattern_fill_type::darktrellis: return "darkhorizontal";
	case xlnt::pattern_fill_type::darkup: return "darkup";
	case xlnt::pattern_fill_type::darkvertical: return "darkvertical";
	case xlnt::pattern_fill_type::gray0625: return "gray0625";
	case xlnt::pattern_fill_type::gray125: return "gray125";
	case xlnt::pattern_fill_type::lightdown: return "lightdown";
	case xlnt::pattern_fill_type::lightgray: return "lightgray";
	case xlnt::pattern_fill_type::lightgrid: return "lightgrid";
	case xlnt::pattern_fill_type::lighthorizontal: return "lighthorizontal";
	case xlnt::pattern_fill_type::lighttrellis: return "lighttrellis";
	case xlnt::pattern_fill_type::lightup: return "lightup";
	case xlnt::pattern_fill_type::lightvertical: return "lightvertical";
	case xlnt::pattern_fill_type::mediumgray: return "mediumgray";
	case xlnt::pattern_fill_type::solid: return "solid";
	default:
	case xlnt::pattern_fill_type::none: return "none";
	}
}

std::string to_string(xlnt::gradient_fill_type fill_type)
{
	return fill_type == xlnt::gradient_fill_type::linear ? "linear" : "path";
}

std::string to_string(xlnt::border_style border_style)
{
	switch (border_style)
	{
	case xlnt::border_style::dashdot: return "dashdot";
	case xlnt::border_style::dashdotdot: return "dashdotdot";
	case xlnt::border_style::dashed: return "dashed";
	case xlnt::border_style::dotted: return "dotted";
	case xlnt::border_style::double_: return "double";
	case xlnt::border_style::hair: return "hair";
	case xlnt::border_style::medium: return "medium";
	case xlnt::border_style::mediumdashdot: return "mediumdashdot";
	case xlnt::border_style::mediumdashdotdot: return "mediumdashdotdot";
	case xlnt::border_style::mediumdashed: return "mediumdashed";
	case xlnt::border_style::slantdashdot: return "slantdashdot";
	case xlnt::border_style::thick: return "thick";
	case xlnt::border_style::thin: return "thin";
	default:
	case xlnt::border_style::none: return "none";
	}
}

std::string to_string(xlnt::vertical_alignment vertical_alignment)
{
	switch (vertical_alignment)
	{
	case xlnt::vertical_alignment::bottom: return "bottom";
	case xlnt::vertical_alignment::center: return "center";
	case xlnt::vertical_alignment::justify: return "justify";
	case xlnt::vertical_alignment::top: return "top";
	default:
	case xlnt::vertical_alignment::none: return "none";
	}
}

std::string to_string(xlnt::horizontal_alignment horizontal_alignment)
{
	switch (horizontal_alignment)
	{
	case xlnt::horizontal_alignment::center: return "center";
	case xlnt::horizontal_alignment::center_continuous: return "center-continous";
	case xlnt::horizontal_alignment::general: return "general";
	case xlnt::horizontal_alignment::justify: return "justify";
	case xlnt::horizontal_alignment::left: return "left";
	case xlnt::horizontal_alignment::right: return "right";
	default:
	case xlnt::horizontal_alignment::none: return "none";
	}
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

void write_relationships(const std::vector<xlnt::relationship> &relationships, pugi::xml_node root)
{
	auto relationships_node = root.append_child("Relationships");
	relationships_node.append_attribute("xmlns").set_value(xlnt::constants::get_namespace("relationships").c_str());

	for (const auto &relationship : relationships)
	{
		auto relationship_node = relationships_node.append_child("Relationship");

		relationship_node.append_attribute("Id").set_value(relationship.get_id().c_str());
		relationship_node.append_attribute("Type").set_value(to_string(relationship.get_type()).c_str());
		relationship_node.append_attribute("Target").set_value(relationship.get_target().get_path().string().c_str());

		if (relationship.get_target_mode() == xlnt::target_mode::external)
		{
			relationship_node.append_attribute("TargetMode").set_value("External");
		}
	}
}


bool write_color(const xlnt::color &color, pugi::xml_node color_node)
{
	switch (color.get_type())
	{
	case xlnt::color::type::theme:
		color_node.append_attribute("theme")
			.set_value(std::to_string(color.get_theme().get_index()).c_str());
		break;

	case xlnt::color::type::indexed:
		color_node.append_attribute("indexed")
			.set_value(std::to_string(color.get_indexed().get_index()).c_str());
		break;

	case xlnt::color::type::rgb:
	default:
		color_node.append_attribute("rgb")
			.set_value(color.get_rgb().get_hex_string().c_str());
		break;
	}

	return true;
}

bool write_dxfs(pugi::xml_node &dxfs_node)
{
	dxfs_node.append_attribute("count").set_value("0");
	return true;
}

bool write_table_styles(pugi::xml_node &table_styles_node)
{
	table_styles_node.append_attribute("count").set_value("0");
	table_styles_node.append_attribute("defaultTableStyle").set_value("TableStyleMedium9");
	table_styles_node.append_attribute("defaultPivotStyle").set_value("PivotStyleMedium7");

	return true;
}

bool write_colors(const std::vector<xlnt::color> &colors, pugi::xml_node &colors_node)
{
	auto indexed_colors_node = colors_node.append_child("indexedColors");

	for (auto &c : colors)
	{
		auto rgb_color_node = indexed_colors_node.append_child("rgbColor");
		auto rgb_attribute = rgb_color_node.append_attribute("rgb");
		rgb_attribute.set_value(c.get_rgb().get_hex_string().c_str());
	}

	return true;
}

} // namespace

namespace xlnt {
namespace detail {

xlsx_producer::xlsx_producer(const workbook &target) : source_(target)
{
}

void xlsx_producer::write(const path &destination)
{
	populate_archive();
	destination_.save(destination);
}

void xlsx_producer::write(std::ostream &destination)
{
	populate_archive();
	destination_.save(destination);
}

void xlsx_producer::write(std::vector<std::uint8_t> &destination)
{
	populate_archive();
	destination_.save(destination);
}

// Part Writing Methods

void xlsx_producer::populate_archive()
{
	write_manifest();

	for (auto &rel : source_.impl().manifest_.get_relationships(path("/")))
	{
		pugi::xml_document document;
        bool write_document = true;

		switch (rel.get_type())
		{
		case relationship::type::core_properties:
			write_core_properties(rel, document.root());
			break;
            
		case relationship::type::extended_properties:
			write_extended_properties(rel, document.root());
			break;
            
		case relationship::type::custom_properties:
			write_custom_properties(rel, document.root());
			break;
            
		case relationship::type::office_document:
			write_workbook(rel, document.root());
			break;
            
		case relationship::type::thumbnail:
            write_thumbnail(rel);
            write_document = false;
            break;
        
        default:
            break;
        }

        if (write_document)
        {
            write_document_to_archive(document, rel.get_target().get_path(), destination_);
        }
	}

	// Unknown Parts

	void write_unknown_parts();
	void write_unknown_relationships();
}

// Package Parts

void xlsx_producer::write_manifest()
{
	pugi::xml_document content_types_document;

	auto types_node = content_types_document.append_child("Types");
	types_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/content-types");

	for (const auto &extension : source_.get_manifest().get_extensions_with_default_types())
	{
		auto default_node = types_node.append_child("Default");
		default_node.append_attribute("Extension").set_value(extension.c_str());
		auto content_type = source_.get_manifest().get_default_type(extension);
		default_node.append_attribute("ContentType").set_value(content_type.c_str());
	}

	for (const auto &part : source_.get_manifest().get_parts_with_overriden_types())
	{
		auto override_node = types_node.append_child("Override");
		override_node.append_attribute("PartName").set_value(part.resolve(path("/")).string().c_str());
		auto content_type = source_.get_manifest().get_override_type(part);
		override_node.append_attribute("ContentType").set_value(content_type.c_str());
	}

	for (const auto &part : source_.get_manifest().get_parts())
	{
		auto part_rels = source_.get_manifest().get_relationships(part);
        
        if (part_rels.empty()) continue;

		pugi::xml_document part_rels_document;
		write_relationships(part_rels, part_rels_document.root());
		path parent = part.parent();

		if (parent.is_absolute())
		{
			parent = path(parent.string().substr(1));
		}

		path rels_path(parent.append("_rels").append(part.filename() + ".rels").string());
		write_document_to_archive(part_rels_document, rels_path, destination_);
	}

	write_document_to_archive(content_types_document, path("[Content_Types].xml"), destination_);
}

void xlsx_producer::write_extended_properties(const relationship &rel, pugi::xml_node root)
{
	auto properties_node = root.append_child("Properties");

	properties_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
	properties_node.append_attribute("xmlns:vt").set_value("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

	properties_node.append_child("Application").text().set(source_.get_application().c_str());
	properties_node.append_child("DocSecurity").text().set(std::to_string(source_.get_doc_security()).c_str());
	properties_node.append_child("ScaleCrop").text().set(source_.get_scale_crop() ? "true" : "false");

	auto heading_pairs_node = properties_node.append_child("HeadingPairs");
	auto heading_pairs_vector_node = heading_pairs_node.append_child("vt:vector");
	heading_pairs_vector_node.append_attribute("size").set_value("2");
	heading_pairs_vector_node.append_attribute("baseType").set_value("variant");
	heading_pairs_vector_node.append_child("vt:variant").append_child("vt:lpstr").text().set("Worksheets");
	heading_pairs_vector_node.append_child("vt:variant")
		.append_child("vt:i4")
		.text().set(std::to_string(source_.get_sheet_titles().size()).c_str());

	auto titles_of_parts_node = properties_node.append_child("TitlesOfParts");
	auto titles_of_parts_vector_node = titles_of_parts_node.append_child("vt:vector");
	titles_of_parts_vector_node.append_attribute("size").set_value(std::to_string(source_.get_sheet_titles().size()).c_str());
	titles_of_parts_vector_node.append_attribute("baseType").set_value("lpstr");

	for (auto ws : source_)
	{
		titles_of_parts_vector_node.append_child("vt:lpstr").text().set(ws.get_title().c_str());
	}

	auto company_node = properties_node.append_child("Company");

	if (!source_.get_company().empty())
	{
		company_node.text().set(source_.get_company().c_str());
	}

	properties_node.append_child("LinksUpToDate").text().set(source_.links_up_to_date() ? "true" : "false");
	properties_node.append_child("SharedDoc").text().set(source_.is_shared_doc() ? "true" : "false");
	properties_node.append_child("HyperlinksChanged").text().set(source_.hyperlinks_changed() ? "true" : "false");
	properties_node.append_child("AppVersion").text().set(source_.get_app_version().c_str());
}

void xlsx_producer::write_core_properties(const relationship &rel, pugi::xml_node root)
{
	auto core_properties_node = root.append_child("cp:coreProperties");

	core_properties_node.append_attribute("xmlns:cp").set_value("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
	core_properties_node.append_attribute("xmlns:dc").set_value("http://purl.org/dc/elements/1.1/");
	core_properties_node.append_attribute("xmlns:dcterms").set_value("http://purl.org/dc/terms/");
	core_properties_node.append_attribute("xmlns:dcmitype").set_value("http://purl.org/dc/dcmitype/");
	core_properties_node.append_attribute("xmlns:xsi").set_value("http://www.w3.org/2001/XMLSchema-instance");

	core_properties_node.append_child("dc:creator").text().set(source_.get_creator().c_str());
	core_properties_node.append_child("cp:lastModifiedBy").text().set(source_.get_last_modified_by().c_str());
	core_properties_node.append_child("dcterms:created").text().set(datetime_to_w3cdtf(source_.get_created()).c_str());
	core_properties_node.child("dcterms:created").append_attribute("xsi:type").set_value("dcterms:W3CDTF");
	core_properties_node.append_child("dcterms:modified").text().set(datetime_to_w3cdtf(source_.get_modified()).c_str());
	core_properties_node.child("dcterms:modified").append_attribute("xsi:type").set_value("dcterms:W3CDTF");

	if (!source_.get_title().empty())
	{
		core_properties_node.append_child("dc:title").text().set(source_.get_title().c_str());
	}

	/*
	core_properties_node.append_child("dc:description");
	core_properties_node.append_child("dc:subject");
	core_properties_node.append_child("cp:keywords");
	core_properties_node.append_child("cp:category");
	*/
}

void xlsx_producer::write_custom_properties(const relationship &rel, pugi::xml_node root)
{
	/*auto properties_node = */root.append_child("Properties");
}

// Write SpreadsheetML-Specific Package Parts

void xlsx_producer::write_workbook(const relationship &rel, pugi::xml_node root)
{
	std::size_t num_visible = 0;
	bool any_defined_names = false;

	for (auto ws : source_)
	{
		if (!ws.has_page_setup() || ws.get_page_setup().get_sheet_state() == sheet_state::visible)
		{
			num_visible++;
		}

		if (ws.has_auto_filter())
		{
			any_defined_names = true;
		}
	}

	if (num_visible == 0)
	{
		throw no_visible_worksheets();
	}

	auto workbook_node = root.append_child("workbook");

	workbook_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	workbook_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

	if (source_.x15_enabled())
	{
		workbook_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
		workbook_node.append_attribute("mc:Ignorable").set_value("x15");
		workbook_node.append_attribute("xmlns:x15").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
	}

	if (source_.has_file_version())
	{
		auto file_version_node = workbook_node.append_child("fileVersion");

		file_version_node.append_attribute("appName").set_value(source_.get_app_name().c_str());
		file_version_node.append_attribute("lastEdited").set_value(std::to_string(source_.get_last_edited()).c_str());
		file_version_node.append_attribute("lowestEdited").set_value(std::to_string(source_.get_lowest_edited()).c_str());
		file_version_node.append_attribute("rupBuild").set_value(std::to_string(source_.get_rup_build()).c_str());
	}

	if (source_.has_properties())
	{
		auto workbook_pr_node = workbook_node.append_child("workbookPr");

		if (source_.has_code_name())
		{
			workbook_pr_node.append_attribute("codeName").set_value(source_.get_code_name().c_str());
		}

		/*
		workbook_pr_node.append_attribute("defaultThemeVersion").set_value("124226");
		workbook_pr_node.append_attribute("date1904").set_value(source_.get_base_date() == calendar::mac_1904 ? "1" : "0");
		*/
	}

	if (source_.has_absolute_path())
	{
		auto alternate_content_node = workbook_node.append_child("mc:AlternateContent");
		alternate_content_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");

		auto choice_node = alternate_content_node.append_child("mc:Choice");
		choice_node.append_attribute("Requires").set_value("x15");

		auto abs_path_node = choice_node.append_child("x15ac:absPath");

		std::string absolute_path = source_.get_absolute_path().string();
		abs_path_node.append_attribute("url").set_value(absolute_path.c_str());
		abs_path_node.append_attribute("xmlns:x15ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");
	}

	if (source_.has_view())
	{
		auto book_views_node = workbook_node.append_child("bookViews");
		auto workbook_view_node = book_views_node.append_child("workbookView");

		const auto &view = source_.get_view();

		//workbook_view_node.append_attribute("activeTab").set_value("0");
		//workbook_view_node.append_attribute("autoFilterDateGrouping").set_value("1");
		//workbook_view_node.append_attribute("firstSheet").set_value("0");
		//workbook_view_node.append_attribute("minimized").set_value("0");
		//workbook_view_node.append_attribute("showHorizontalScroll").set_value("1");
		//workbook_view_node.append_attribute("showSheetTabs").set_value("1");
		//workbook_view_node.append_attribute("showVerticalScroll").set_value("1");
		//workbook_view_node.append_attribute("tabRatio").set_value("600");
		//workbook_view_node.append_attribute("visibility").set_value("visible");
		workbook_view_node.append_attribute("xWindow").set_value(std::to_string(view.x_window).c_str());
		workbook_view_node.append_attribute("yWindow").set_value(std::to_string(view.y_window).c_str());
		workbook_view_node.append_attribute("windowWidth").set_value(std::to_string(view.window_width).c_str());
		workbook_view_node.append_attribute("windowHeight").set_value(std::to_string(view.window_height).c_str());
		workbook_view_node.append_attribute("tabRatio").set_value(std::to_string(view.tab_ratio).c_str());
	}

	auto sheets_node = workbook_node.append_child("sheets");
	pugi::xml_node defined_names_node;
	
	if (any_defined_names)
	{
		defined_names_node = workbook_node.append_child("definedNames");
	}

	for (const auto ws : source_)
	{
		auto sheet_rel_id = source_.d_->sheet_title_rel_id_map_[ws.get_title()];
		auto sheet_rel = source_.d_->manifest_.get_relationship(rel.get_source().get_path(), sheet_rel_id);

		auto sheet_node = sheets_node.append_child("sheet");
		sheet_node.append_attribute("name").set_value(ws.get_title().c_str());
		sheet_node.append_attribute("sheetId").set_value(std::to_string(ws.get_id()).c_str());

		if (ws.has_page_setup() && ws.get_sheet_state() == xlnt::sheet_state::hidden)
		{
			sheet_node.append_attribute("state").set_value("hidden");
		}

		sheet_node.append_attribute("r:id").set_value(sheet_rel_id.c_str());

		if (ws.has_auto_filter())
		{
			auto defined_name_node = defined_names_node.append_child("definedName");
			defined_name_node.append_attribute("name").set_value("_xlnm._FilterDatabase");
			defined_name_node.append_attribute("hidden").set_value(write_bool(true).c_str());
			defined_name_node.append_attribute("localSheetId").set_value("0");
			std::string name =
				"'" + ws.get_title() + "'!" + range_reference::make_absolute(ws.get_auto_filter()).to_string();
			defined_name_node.text().set(name.c_str());
		}
	}

	if (source_.has_calculation_properties())
	{
		auto calc_pr_node = workbook_node.append_child("calcPr");
		calc_pr_node.append_attribute("calcId").set_value("150000");
		//calc_pr_node.append_attribute("calcMode").set_value("auto");
		//calc_pr_node.append_attribute("fullCalcOnLoad").set_value("1");
		calc_pr_node.append_attribute("concurrentCalc").set_value("0");
	}

	if (!source_.get_named_ranges().empty())
	{
		auto defined_names_node = workbook_node.append_child("definedNames");

		for (auto &named_range : source_.get_named_ranges())
		{
			auto defined_name_node = defined_names_node.append_child("s:definedName");
			defined_name_node.append_attribute("xmlns:s").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			defined_name_node.append_attribute("name").set_value(named_range.get_name().c_str());
			const auto &target = named_range.get_targets().front();
			std::string target_string = "'" + target.first.get_title();
			target_string.push_back('\'');
			target_string.push_back('!');
			target_string.append(target.second.to_string());
			defined_name_node.text().set(target_string.c_str());
		}
	}

	if (source_.has_arch_id())
	{
		auto ext_lst_node = workbook_node.append_child("extLst");
		auto ext_node = ext_lst_node.append_child("ext");
		ext_node.append_attribute("uri").set_value("{7523E5D3-25F3-A5E0-1632-64F254C22452}");
		ext_node.append_attribute("xmlns:mx").set_value("http://schemas.microsoft.com/office/mac/excel/2008/main");
		auto arch_id_node = ext_node.append_child("mx:ArchID");
		arch_id_node.append_attribute("Flags").set_value("2");
	}
    
    for (const auto &child_rel : source_.get_manifest().get_relationships(rel.get_target().get_path()))
    {
        pugi::xml_document document;
        
        switch (child_rel.get_type())
        {
        case relationship::type::calculation_chain:
			write_calculation_chain(child_rel, document.root());
			break;
            
		case relationship::type::chartsheet:
			write_chartsheet(child_rel, document.root());
			break;
            
		case relationship::type::connections:
			write_connections(child_rel, document.root());
			break;
            
		case relationship::type::custom_xml_mappings:
			write_custom_xml_mappings(child_rel, document.root());
			break;
            
		case relationship::type::dialogsheet:
			write_dialogsheet(child_rel, document.root());
			break;
            
		case relationship::type::external_workbook_references:
			write_external_workbook_references(child_rel, document.root());
			break;
            
		case relationship::type::metadata:
			write_metadata(child_rel, document.root());
			break;
            
		case relationship::type::pivot_table:
			write_pivot_table(child_rel, document.root());
			break;

		case relationship::type::shared_string_table:
			write_shared_string_table(child_rel, document.root());
			break;

		case relationship::type::shared_workbook_revision_headers:
			write_shared_workbook_revision_headers(child_rel, document.root());
			break;
            
		case relationship::type::styles:
			write_styles(child_rel, document.root());
			break;
            
		case relationship::type::theme:
			write_theme(child_rel, document.root());
			break;
            
		case relationship::type::volatile_dependencies:
			write_volatile_dependencies(child_rel, document.root());
			break;
            
		case relationship::type::worksheet:
			write_worksheet(child_rel, document.root());
			break;
            
        default:
            break;
		}
        
		path archive_path(child_rel.get_source().get_path().parent().append(child_rel.get_target().get_path()));
        write_document_to_archive(document, archive_path, destination_);
    }
}

// Write Workbook Relationship Target Parts

void xlsx_producer::write_calculation_chain(const relationship &rel, pugi::xml_node root)
{
	/*auto calc_chain_node = */root.append_child("calcChain");
}

void xlsx_producer::write_chartsheet(const relationship &rel, pugi::xml_node root)
{
	/*auto chartsheet_node = */root.append_child("chartsheet");
}

void xlsx_producer::write_connections(const relationship &rel, pugi::xml_node root)
{
	/*auto connections_node = */root.append_child("connections");
}

void xlsx_producer::write_custom_xml_mappings(const relationship &rel, pugi::xml_node root)
{
	/*auto map_info_node = */root.append_child("MapInfo");
}

void xlsx_producer::write_dialogsheet(const relationship &rel, pugi::xml_node root)
{
	/*auto dialogsheet_node = */root.append_child("dialogsheet");
}

void xlsx_producer::write_external_workbook_references(const relationship &rel, pugi::xml_node root)
{
	/*auto external_link_node = */root.append_child("externalLink");
}

void xlsx_producer::write_metadata(const relationship &rel, pugi::xml_node root)
{
	/*auto metadata_node = */root.append_child("metadata");
}

void xlsx_producer::write_pivot_table(const relationship &rel, pugi::xml_node root)
{
	/*auto pivot_table_definition_node = */root.append_child("pivotTableDefinition");
}

void xlsx_producer::write_shared_string_table(const relationship &rel, pugi::xml_node root)
{
	auto sst_node = root.append_child("sst");

	sst_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

	sst_node.append_attribute("count").set_value(std::to_string(source_.get_shared_strings().size()).c_str());
	sst_node.append_attribute("uniqueCount").set_value(std::to_string(source_.get_shared_strings().size()).c_str());

	for (const auto &string : source_.get_shared_strings())
	{
		if (string.get_runs().size() == 1 && !string.get_runs().at(0).has_formatting())
		{
			sst_node.append_child("si").append_child("t").text().set(string.get_plain_string().c_str());
		}
		else
		{
			auto string_item_node = sst_node.append_child("si");

			for (const auto &run : string.get_runs())
			{
				auto rich_text_run_node = string_item_node.append_child("r");

				if (run.has_formatting())
				{
					auto run_properties_node = rich_text_run_node.append_child("rPr");

					if (run.has_size())
					{
						run_properties_node.append_child("sz").append_attribute("val").set_value(std::to_string(run.get_size()).c_str());
					}

					if (run.has_color())
					{
						run_properties_node.append_child("color").append_attribute("rgb").set_value(run.get_color().c_str());
					}

					if (run.has_font())
					{
						run_properties_node.append_child("rFont").append_attribute("val").set_value(run.get_font().c_str());
					}

					if (run.has_family())
					{
						run_properties_node.append_child("family").append_attribute("val").set_value(std::to_string(run.get_family()).c_str());
					}

					if (run.has_scheme())
					{
						run_properties_node.append_child("scheme").append_attribute("val").set_value(run.get_scheme().c_str());
					}
				}

				auto text_node = rich_text_run_node.append_child("t");
				text_node.text().set(run.get_string().c_str());
			}
		}
	}
}

void xlsx_producer::write_shared_workbook_revision_headers(const relationship &rel, pugi::xml_node root)
{
	/*auto headers_node = */root.append_child("headers");
}

void xlsx_producer::write_shared_workbook(const relationship &rel, pugi::xml_node root)
{
	/*auto revisions_node = */root.append_child("revisions");
}

void xlsx_producer::write_shared_workbook_user_data(const relationship &rel, pugi::xml_node root)
{
	/*auto users_node = */root.append_child("users");
}

void xlsx_producer::write_styles(const relationship &rel, pugi::xml_node root)
{
	auto stylesheet_node = root.append_child("styleSheet");

	// Namespaces

	stylesheet_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

	if (source_.x15_enabled())
	{
		stylesheet_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
		stylesheet_node.append_attribute("mc:Ignorable").set_value("x14ac");
		stylesheet_node.append_attribute("xmlns:x14ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
	}

	const auto &stylesheet = source_.impl().stylesheet_;

	// Number Formats

	if (!stylesheet.number_formats.empty())
	{
		auto number_formats_node = stylesheet_node.append_child("numFmts");
		const auto &number_formats = stylesheet.number_formats;

		number_formats_node.append_attribute("count").set_value(std::to_string(number_formats.size()).c_str());

		for (const auto &num_fmt : number_formats)
		{
			auto num_fmt_node = number_formats_node.append_child("numFmt");
			num_fmt_node.append_attribute("numFmtId").set_value(std::to_string(num_fmt.get_id()).c_str());
			num_fmt_node.append_attribute("formatCode").set_value(num_fmt.get_format_string().c_str());
		}
	}

	// Fonts

	if (!stylesheet.fonts.empty())
	{
		auto fonts_node = stylesheet_node.append_child("fonts");
		const auto &fonts = stylesheet.fonts;

		fonts_node.append_attribute("count").set_value(std::to_string(fonts.size()).c_str());

		if (source_.x15_enabled())
		{
			fonts_node.append_attribute("x14ac:knownFonts").set_value(std::to_string(fonts.size()).c_str());
		}

		for (const auto &current_font : fonts)
		{
			auto font_node = fonts_node.append_child("font");

			if (current_font.bold())
			{
				auto bold_node = font_node.append_child("b");
				bold_node.append_attribute("val").set_value(write_bool(true).c_str());
			}

			if (current_font.italic())
			{
				auto italic_node = font_node.append_child("i");
				italic_node.append_attribute("val").set_value(write_bool(true).c_str());
			}

			if (current_font.underlined())
			{
				auto underline_node = font_node.append_child("u");
				underline_node.append_attribute("val").set_value(to_string(current_font.underline()).c_str());
			}

			if (current_font.strikethrough())
			{
				auto strike_node = font_node.append_child("strike");
				strike_node.append_attribute("val").set_value(write_bool(true).c_str());
			}

			auto size_node = font_node.append_child("sz");
			size_node.append_attribute("val").set_value(std::to_string(current_font.size()).c_str());

			if (current_font.color())
			{
				auto color_node = font_node.append_child("color");
				write_color(*current_font.color(), color_node);
			}

			auto name_node = font_node.append_child("name");
			name_node.append_attribute("val").set_value(current_font.name().c_str());

			if (current_font.family())
			{
				auto family_node = font_node.append_child("family");
				family_node.append_attribute("val").set_value(std::to_string(*current_font.family()).c_str());
			}

			if (current_font.scheme())
			{
				auto scheme_node = font_node.append_child("scheme");
				scheme_node.append_attribute("val").set_value(current_font.scheme()->c_str());
			}
		}
	}

	// Fills

	if (!stylesheet.fills.empty())
	{
		auto fills_node = stylesheet_node.append_child("fills");
		const auto &fills = stylesheet.fills;

		fills_node.append_attribute("count").set_value(std::to_string(fills.size()).c_str());

		for (auto &fill_ : fills)
		{
			auto fill_node = fills_node.append_child("fill");

			if (fill_.type() == xlnt::fill_type::pattern)
			{
				const auto &pattern = fill_.pattern_fill();

				auto pattern_fill_node = fill_node.append_child("patternFill");
				pattern_fill_node.append_attribute("patternType").set_value(to_string(pattern.type()).c_str());

				if (pattern.foreground())
				{
					write_color(*pattern.foreground(), pattern_fill_node.append_child("fgColor"));
				}

				if (pattern.background())
				{
					write_color(*pattern.background(), pattern_fill_node.append_child("bgColor"));
				}
			}
			else if (fill_.type() == xlnt::fill_type::gradient)
			{
				const auto &gradient = fill_.gradient_fill();

				auto gradient_fill_node = fill_node.append_child("gradientFill");
				auto gradient_fill_type_string = to_string(gradient.type());
				gradient_fill_node.append_attribute("gradientType").set_value(gradient_fill_type_string.c_str());

				if (gradient.degree() != 0)
				{
					gradient_fill_node.append_attribute("degree").set_value(gradient.degree());
				}

				if (gradient.left() != 0)
				{
					gradient_fill_node.append_attribute("left").set_value(gradient.left());
				}

				if (gradient.right() != 0)
				{
					gradient_fill_node.append_attribute("right").set_value(gradient.right());
				}

				if (gradient.top() != 0)
				{
					gradient_fill_node.append_attribute("top").set_value(gradient.top());
				}

				if (gradient.bottom() != 0)
				{
					gradient_fill_node.append_attribute("bottom").set_value(gradient.bottom());
				}

				for (const auto &stop : gradient.stops())
				{
					auto stop_node = gradient_fill_node.append_child("stop");
					stop_node.append_attribute("position").set_value(stop.first);
					write_color(stop.second, stop_node.append_child("color"));
				}
			}
		}
	}

	// Borders

	if (!stylesheet.borders.empty())
	{
		auto borders_node = stylesheet_node.append_child("borders");
		const auto &borders = stylesheet.borders;

		borders_node.append_attribute("count").set_value(std::to_string(borders.size()).c_str());

		for (const auto &current_border : borders)
		{
			auto border_node = borders_node.append_child("border");

			if (current_border.diagonal())
			{
				auto up = *current_border.diagonal() == diagonal_direction::both
					|| *current_border.diagonal() == diagonal_direction::up;
				border_node.append_attribute("diagonalUp").set_value(up ? "true" : "false");

				auto down = *current_border.diagonal() == diagonal_direction::both
					|| *current_border.diagonal() == diagonal_direction::down;
				border_node.append_attribute("diagonalDown").set_value(down ? "true" : "false");
			}

			for (const auto &side : xlnt::border::all_sides())
			{
				if (current_border.side(side))
				{
					const auto current_side = *current_border.side(side);

					auto side_name = to_string(side);
					auto side_node = border_node.append_child(side_name.c_str());

					if (current_side.style())
					{
						auto style_string = to_string(*current_side.style());
						side_node.append_attribute("style").set_value(style_string.c_str());
					}

					if (current_side.color())
					{
						auto color_node = side_node.append_child("color");
						write_color(*current_side.color(), color_node);
					}
				}
			}
		}
	}

	// Formats & Styles

	auto cell_style_xfs_node = stylesheet_node.append_child("cellStyleXfs");
	auto cell_xfs_node = stylesheet_node.append_child("cellXfs");
	cell_xfs_node.append_attribute("count").set_value(std::to_string(stylesheet.formats.size()).c_str());

	// Formats

	for (auto &current_format : stylesheet.formats)
	{
		auto xf_node = cell_xfs_node.append_child("xf");

		xf_node.append_attribute("numFmtId").set_value(std::to_string(current_format.number_format().get_id()).c_str());

        auto font_iter = std::find(stylesheet.fonts.begin(), stylesheet.fonts.end(), current_format.font());
        if (font_iter == stylesheet.fonts.end())
        {
            throw xlnt::exception("font not found");
        }
		auto font_id = std::distance(stylesheet.fonts.begin(), font_iter);
		xf_node.append_attribute("fontId").set_value(std::to_string(font_id).c_str());

        auto fill_iter = std::find(stylesheet.fills.begin(), stylesheet.fills.end(), current_format.fill());
        if (fill_iter == stylesheet.fills.end())
        {
            throw xlnt::exception("fill not found");
        }
		auto fill_id = std::distance(stylesheet.fills.begin(), std::find(stylesheet.fills.begin(), stylesheet.fills.end(), current_format.fill()));
		xf_node.append_attribute("fillId").set_value(std::to_string(fill_id).c_str());

        auto border_iter = std::find(stylesheet.borders.begin(), stylesheet.borders.end(), current_format.border());
        if (border_iter == stylesheet.borders.end())
        {
            throw xlnt::exception("border not found");
        }
		auto border_id = std::distance(stylesheet.borders.begin(), std::find(stylesheet.borders.begin(), stylesheet.borders.end(), current_format.border()));
		xf_node.append_attribute("borderId").set_value(std::to_string(border_id).c_str());

		if (current_format.number_format_applied()) xf_node.append_attribute("applyNumberFormat").set_value(write_bool(true).c_str());
		if (current_format.fill_applied()) xf_node.append_attribute("applyFill").set_value(write_bool(true).c_str());
		if (current_format.font_applied()) xf_node.append_attribute("applyFont").set_value(write_bool(true).c_str());
		if (current_format.border_applied()) xf_node.append_attribute("applyBorder").set_value(write_bool(true).c_str());

		if (current_format.alignment_applied())
		{
			xf_node.append_attribute("applyAlignment").set_value(write_bool(true).c_str());
			auto alignment_node = xf_node.append_child("alignment");
			auto current_alignment = current_format.alignment();

			if (current_alignment.vertical())
			{
				auto vertical = to_string(*current_alignment.vertical());
				alignment_node.append_attribute("vertical").set_value(vertical.c_str());
			}

			if (current_alignment.horizontal())
			{
				auto horizontal = to_string(*current_alignment.horizontal());
				alignment_node.append_attribute("horizontal").set_value(horizontal.c_str());
			}

			if (current_alignment.rotation())
			{
				alignment_node.append_attribute("textRotation").set_value(std::to_string(*current_alignment.rotation()).c_str());
			}

			if (current_alignment.wrap())
			{
				alignment_node.append_attribute("wrapText").set_value(*current_alignment.wrap() ? "true" : "false");
			}

			if (current_alignment.indent())
			{
				alignment_node.append_attribute("indent").set_value(std::to_string(*current_alignment.indent()).c_str());
			}

			if (current_alignment.shrink())
			{
				alignment_node.append_attribute("shrinkToFit").set_value(*current_alignment.shrink() ? "true" : "false");
			}
		}

		if (current_format.protection_applied())
		{
			xf_node.append_attribute("applyProtection").set_value(write_bool(true).c_str());
			auto protection_node = xf_node.append_child("protection");
			auto current_protection = current_format.protection();

			protection_node.append_attribute("locked").set_value(current_protection.locked() ? "true" : "false");
			protection_node.append_attribute("hidden").set_value(current_protection.hidden() ? "true" : "false");
		}

		if (current_format.has_style())
		{
			auto style_iter = std::find_if(stylesheet.styles.begin(), stylesheet.styles.end(),
				[&](const xlnt::style &s) { return s.name() == current_format.style().name(); });
			auto style_index = std::distance(stylesheet.styles.begin(), style_iter);
			xf_node.append_attribute("xfId").set_value(std::to_string(style_index).c_str());
		}
	}

	// Styles

	auto cell_styles_node = stylesheet_node.append_child("cellStyles");
	cell_style_xfs_node.append_attribute("count").set_value(std::to_string(stylesheet.styles.size()).c_str());
	cell_styles_node.append_attribute("count").set_value(std::to_string(stylesheet.styles.size()).c_str());
	std::size_t style_index = 0;

	for (auto &current_style : stylesheet.styles)
	{
		auto cell_style_node = cell_styles_node.append_child("cellStyle");

		cell_style_node.append_attribute("name").set_value(current_style.name().c_str());
		cell_style_node.append_attribute("xfId").set_value(std::to_string(style_index++).c_str());

		if (current_style.builtin_id())
		{
			cell_style_node.append_attribute("builtinId").set_value(std::to_string(*current_style.builtin_id()).c_str());
		}

		if (current_style.hidden())
		{
			cell_style_node.append_attribute("hidden").set_value(write_bool(true).c_str());
		}

		if (current_style.custom())
		{
			cell_style_node.append_attribute("customBuiltin").set_value(write_bool(*current_style.custom()).c_str());
		}

		auto xf_node = cell_style_xfs_node.append_child("xf");

		xf_node.append_attribute("numFmtId").set_value(std::to_string(current_style.number_format().get_id()).c_str());

		auto font_id = std::distance(stylesheet.fonts.begin(), std::find(stylesheet.fonts.begin(), stylesheet.fonts.end(), current_style.font()));
		xf_node.append_attribute("fontId").set_value(std::to_string(font_id).c_str());

		auto fill_id = std::distance(stylesheet.fills.begin(), std::find(stylesheet.fills.begin(), stylesheet.fills.end(), current_style.fill()));
		xf_node.append_attribute("fillId").set_value(std::to_string(fill_id).c_str());

		auto border_id = std::distance(stylesheet.borders.begin(), std::find(stylesheet.borders.begin(), stylesheet.borders.end(), current_style.border()));
		xf_node.append_attribute("borderId").set_value(std::to_string(border_id).c_str());

		if (current_style.number_format_applied()) xf_node.append_attribute("applyNumberFormat").set_value(write_bool(true).c_str());
		if (current_style.fill_applied()) xf_node.append_attribute("applyFill").set_value(write_bool(true).c_str());
		if (current_style.font_applied()) xf_node.append_attribute("applyFont").set_value(write_bool(true).c_str());
		if (current_style.border_applied()) xf_node.append_attribute("applyBorder").set_value(write_bool(true).c_str());

		if (current_style.alignment_applied())
		{
			xf_node.append_attribute("applyAlignment").set_value(write_bool(true).c_str());
			auto alignment_node = xf_node.append_child("alignment");
			auto current_alignment = current_style.alignment();

			if (current_alignment.vertical())
			{
				auto vertical = to_string(*current_alignment.vertical());
				alignment_node.append_attribute("vertical").set_value(vertical.c_str());
			}

			if (current_alignment.horizontal())
			{
				auto horizontal = to_string(*current_alignment.horizontal());
				alignment_node.append_attribute("horizontal").set_value(horizontal.c_str());
			}

			if (current_alignment.rotation())
			{
				alignment_node.append_attribute("textRotation").set_value(std::to_string(*current_alignment.rotation()).c_str());
			}

			if (current_alignment.wrap())
			{
				alignment_node.append_attribute("wrapText").set_value(write_bool(*current_alignment.wrap()).c_str());
			}

			if (current_alignment.indent())
			{
				alignment_node.append_attribute("indent").set_value(std::to_string(*current_alignment.indent()).c_str());
			}

			if (current_alignment.shrink())
			{
				alignment_node.append_attribute("shrinkToFit").set_value(write_bool(*current_alignment.shrink()).c_str());
			}
		}

		if (current_style.protection_applied())
		{
			xf_node.append_attribute("applyProtection").set_value(write_bool(true).c_str());
			auto protection_node = xf_node.append_child("protection");
			auto current_protection = current_style.protection();

			protection_node.append_attribute("locked").set_value(write_bool(current_protection.locked()).c_str());
			protection_node.append_attribute("hidden").set_value(write_bool(current_protection.hidden()).c_str());
		}
	}

	// Dxfs

	auto dxfs_node = stylesheet_node.append_child("dxfs");
	write_dxfs(dxfs_node);

	// Table Styles

	auto table_styles_node = stylesheet_node.append_child("tableStyles");
	write_table_styles(table_styles_node);

	// Colors

	if (!stylesheet.colors.empty())
	{
		auto colors_node = stylesheet_node.append_child("colors");
		write_colors(stylesheet.colors, colors_node);
	}

	// Ext Lst

	auto ext_list_node = stylesheet_node.append_child("extLst");
	auto ext_node = ext_list_node.append_child("ext");
	ext_node.append_attribute("uri").set_value("{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
	ext_node.append_attribute("xmlns:x14").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
	auto slicer_styles_node = ext_node.append_child("x14:slicerStyles");
	slicer_styles_node.append_attribute("defaultSlicerStyle").set_value("SlicerStyleLight1");
}

void xlsx_producer::write_theme(const relationship &rel, pugi::xml_node root)
{
	auto theme_node = root.append_child("a:theme");
	theme_node.append_attribute("xmlns:a").set_value(constants::get_namespace("drawingml").c_str());
	theme_node.append_attribute("name").set_value("Office Theme");

	auto theme_elements_node = theme_node.append_child("a:themeElements");
	auto clr_scheme_node = theme_elements_node.append_child("a:clrScheme");
	clr_scheme_node.append_attribute("name").set_value("Office");

	struct scheme_element
	{
		std::string name;
		std::string sub_element_name;
		std::string val;
	};

	std::vector<scheme_element> scheme_elements = {
		{ "a:dk1", "a:sysClr", "windowText" },{ "a:lt1", "a:sysClr", "window" },
		{ "a:dk2", "a:srgbClr", "44546A" },{ "a:lt2", "a:srgbClr", "E7E6E6" },
		{ "a:accent1", "a:srgbClr", "5B9BD5" },{ "a:accent2", "a:srgbClr", "ED7D31" },
		{ "a:accent3", "a:srgbClr", "A5A5A5" },{ "a:accent4", "a:srgbClr", "FFC000" },
		{ "a:accent5", "a:srgbClr", "4472C4" },{ "a:accent6", "a:srgbClr", "70AD47" },
		{ "a:hlink", "a:srgbClr", "0563C1" },{ "a:folHlink", "a:srgbClr", "954F72" },
	};

	for (auto element : scheme_elements)
	{
		auto element_node = clr_scheme_node.append_child(element.name.c_str());
		element_node.append_child(element.sub_element_name.c_str()).append_attribute("val").set_value(element.val.c_str());

		if (element.name == "a:dk1")
		{
			element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("000000");
		}
		else if (element.name == "a:lt1")
		{
			element_node.child(element.sub_element_name.c_str()).append_attribute("lastClr").set_value("FFFFFF");
		}
	}

	struct font_scheme
	{
		bool typeface;
		std::string script;
		std::string major;
		std::string minor;
	};

	std::vector<font_scheme> font_schemes = {
		{ true, "a:latin", "Calibri Light", "Calibri" },
		{ true, "a:ea", "", "" },
		{ true, "a:cs", "", "" },
		{ false, "Jpan", "Yu Gothic Light", "Yu Gothic" },
		{ false, "Hang", "\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95",
		"\xeb\xa7\x91\xec\x9d\x80 \xea\xb3\xa0\xeb\x94\x95" },
		{ false, "Hans", "DengXian Light", "DengXian" },
		{ false, "Hant", "\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94",
		"\xe6\x96\xb0\xe7\xb4\xb0\xe6\x98\x8e\xe9\xab\x94" },
		{ false, "Arab", "Times New Roman", "Arial" },
		{ false, "Hebr", "Times New Roman", "Arial" },
		{ false, "Thai", "Tahoma", "Tahoma" },
		{ false, "Ethi", "Nyala", "Nyala" },
		{ false, "Beng", "Vrinda", "Vrinda" },
		{ false, "Gujr", "Shruti", "Shruti" },
		{ false, "Khmr", "MoolBoran", "DaunPenh" },
		{ false, "Knda", "Tunga", "Tunga" },
		{ false, "Guru", "Raavi", "Raavi" },
		{ false, "Cans", "Euphemia", "Euphemia" },
		{ false, "Cher", "Plantagenet Cherokee", "Plantagenet Cherokee" },
		{ false, "Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti" },
		{ false, "Tibt", "Microsoft Himalaya", "Microsoft Himalaya" },
		{ false, "Thaa", "MV Boli", "MV Boli" },
		{ false, "Deva", "Mangal", "Mangal" },
		{ false, "Telu", "Gautami", "Gautami" },
		{ false, "Taml", "Latha", "Latha" },
		{ false, "Syrc", "Estrangelo Edessa", "Estrangelo Edessa" },
		{ false, "Orya", "Kalinga", "Kalinga" },
		{ false, "Mlym", "Kartika", "Kartika" },
		{ false, "Laoo", "DokChampa", "DokChampa" },
		{ false, "Sinh", "Iskoola Pota", "Iskoola Pota" },
		{ false, "Mong", "Mongolian Baiti", "Mongolian Baiti" },
		{ false, "Viet", "Times New Roman", "Arial" },
		{ false, "Uigh", "Microsoft Uighur", "Microsoft Uighur" },
		{ false, "Geor", "Sylfaen", "Sylfaen" }
	};

	auto font_scheme_node = theme_elements_node.append_child("a:fontScheme");
	font_scheme_node.append_attribute("name").set_value("Office");

	auto major_fonts_node = font_scheme_node.append_child("a:majorFont");
	auto minor_fonts_node = font_scheme_node.append_child("a:minorFont");

	for (auto scheme : font_schemes)
	{
		if (scheme.typeface)
		{
			auto major_font_node = major_fonts_node.append_child(scheme.script.c_str());
			major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());

			if (scheme.major == "Calibri Light")
			{
				major_font_node.append_attribute("panose").set_value("020F0302020204030204");
			}

			auto minor_font_node = minor_fonts_node.append_child(scheme.script.c_str());
			minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());

			if (scheme.minor == "Calibri")
			{
				minor_font_node.append_attribute("panose").set_value("020F0502020204030204");
			}
		}
		else
		{
			auto major_font_node = major_fonts_node.append_child("a:font");
			major_font_node.append_attribute("script").set_value(scheme.script.c_str());
			major_font_node.append_attribute("typeface").set_value(scheme.major.c_str());

			auto minor_font_node = minor_fonts_node.append_child("a:font");
			minor_font_node.append_attribute("script").set_value(scheme.script.c_str());
			minor_font_node.append_attribute("typeface").set_value(scheme.minor.c_str());
		}
	}

	auto format_scheme_node = theme_elements_node.append_child("a:fmtScheme");
	format_scheme_node.append_attribute("name").set_value("Office");

	auto fill_style_list_node = format_scheme_node.append_child("a:fillStyleLst");
	fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

	auto grad_fill_node = fill_style_list_node.append_child("a:gradFill");
	grad_fill_node.append_attribute("rotWithShape").set_value("1");

	auto grad_fill_list = grad_fill_node.append_child("a:gsLst");
	auto gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("0");
	auto scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("110000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("105000");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("67000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("50000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("105000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("103000");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("73000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("100000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("105000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("109000");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("81000");

	auto lin_node = grad_fill_node.append_child("a:lin");
	lin_node.append_attribute("ang").set_value("5400000");
	lin_node.append_attribute("scaled").set_value("0");

	grad_fill_node = fill_style_list_node.append_child("a:gradFill");
	grad_fill_node.append_attribute("rotWithShape").set_value("1");

	grad_fill_list = grad_fill_node.append_child("a:gsLst");
	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("0");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("103000");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("102000");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("94000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("50000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("110000");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("100000");
	scheme_color_node.append_child("a:shade").append_attribute("val").set_value("100000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("100000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("99000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("120000");
	scheme_color_node.append_child("a:shade").append_attribute("val").set_value("78000");

	lin_node = grad_fill_node.append_child("a:lin");
	lin_node.append_attribute("ang").set_value("5400000");
	lin_node.append_attribute("scaled").set_value("0");

	auto line_style_list_node = format_scheme_node.append_child("a:lnStyleLst");

	auto ln_node = line_style_list_node.append_child("a:ln");
	ln_node.append_attribute("w").set_value("6350");
	ln_node.append_attribute("cap").set_value("flat");
	ln_node.append_attribute("cmpd").set_value("sng");
	ln_node.append_attribute("algn").set_value("ctr");

	auto solid_fill_node = ln_node.append_child("a:solidFill");
	scheme_color_node = solid_fill_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");
	ln_node.append_child("a:miter").append_attribute("lim").set_value("800000");

	ln_node = line_style_list_node.append_child("a:ln");
	ln_node.append_attribute("w").set_value("12700");
	ln_node.append_attribute("cap").set_value("flat");
	ln_node.append_attribute("cmpd").set_value("sng");
	ln_node.append_attribute("algn").set_value("ctr");

	solid_fill_node = ln_node.append_child("a:solidFill");
	scheme_color_node = solid_fill_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");
	ln_node.append_child("a:miter").append_attribute("lim").set_value("800000");

	ln_node = line_style_list_node.append_child("a:ln");
	ln_node.append_attribute("w").set_value("19050");
	ln_node.append_attribute("cap").set_value("flat");
	ln_node.append_attribute("cmpd").set_value("sng");
	ln_node.append_attribute("algn").set_value("ctr");

	solid_fill_node = ln_node.append_child("a:solidFill");
	scheme_color_node = solid_fill_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	ln_node.append_child("a:prstDash").append_attribute("val").set_value("solid");
	ln_node.append_child("a:miter").append_attribute("lim").set_value("800000");

	auto effect_style_list_node = format_scheme_node.append_child("a:effectStyleLst");
	auto effect_style_node = effect_style_list_node.append_child("a:effectStyle");
	auto effect_list_node = effect_style_node.append_child("a:effectLst");

	effect_style_node = effect_style_list_node.append_child("a:effectStyle");
	effect_list_node = effect_style_node.append_child("a:effectLst");

	effect_style_node = effect_style_list_node.append_child("a:effectStyle");
	effect_list_node = effect_style_node.append_child("a:effectLst");
	auto outer_shadow_node = effect_list_node.append_child("a:outerShdw");
	outer_shadow_node.append_attribute("blurRad").set_value("57150");
	outer_shadow_node.append_attribute("dist").set_value("19050");
	outer_shadow_node.append_attribute("dir").set_value("5400000");
	outer_shadow_node.append_attribute("algn").set_value("ctr");
	outer_shadow_node.append_attribute("rotWithShape").set_value("0");
	auto srgb_clr_node = outer_shadow_node.append_child("a:srgbClr");
	srgb_clr_node.append_attribute("val").set_value("000000");
	srgb_clr_node.append_child("a:alpha").append_attribute("val").set_value("63000");

	auto bg_fill_style_list_node = format_scheme_node.append_child("a:bgFillStyleLst");

	bg_fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr").append_attribute("val").set_value("phClr");

	scheme_color_node = bg_fill_style_list_node.append_child("a:solidFill").append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("95000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("170000");

	grad_fill_node = bg_fill_style_list_node.append_child("a:gradFill");
	grad_fill_node.append_attribute("rotWithShape").set_value("1");

	grad_fill_list = grad_fill_node.append_child("a:gsLst");
	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("0");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("93000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("150000");
	scheme_color_node.append_child("a:shade").append_attribute("val").set_value("98000");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("102000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("50000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:tint").append_attribute("val").set_value("98000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("130000");
	scheme_color_node.append_child("a:shade").append_attribute("val").set_value("90000");
	scheme_color_node.append_child("a:lumMod").append_attribute("val").set_value("103000");

	gs_node = grad_fill_list.append_child("a:gs");
	gs_node.append_attribute("pos").set_value("100000");
	scheme_color_node = gs_node.append_child("a:schemeClr");
	scheme_color_node.append_attribute("val").set_value("phClr");
	scheme_color_node.append_child("a:shade").append_attribute("val").set_value("63000");
	scheme_color_node.append_child("a:satMod").append_attribute("val").set_value("120000");

	lin_node = grad_fill_node.append_child("a:lin");
	lin_node.append_attribute("ang").set_value("5400000");
	lin_node.append_attribute("scaled").set_value("0");

	theme_node.append_child("a:objectDefaults");
	theme_node.append_child("a:extraClrSchemeLst");

	auto ext_lst_node = theme_node.append_child("a:extLst");
	auto ext_node = ext_lst_node.append_child("a:ext");
	ext_node.append_attribute("uri").set_value("{05A4C25C-085E-4340-85A3-A5531E510DB2}");
	auto theme_family_node = ext_node.append_child("thm15:themeFamily");
	theme_family_node.append_attribute("xmlns:thm15").set_value("http://schemas.microsoft.com/office/thememl/2012/main");
	theme_family_node.append_attribute("name").set_value("Office Theme");
	theme_family_node.append_attribute("id").set_value("{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}");
	theme_family_node.append_attribute("vid").set_value("{4A3C46E8-61CC-4603-A589-7422A47A8E4A}");
}

void xlsx_producer::write_volatile_dependencies(const relationship &rel, pugi::xml_node root)
{
	/*auto vol_types_node = */root.append_child("volTypes");
}

void xlsx_producer::write_worksheet(const relationship &rel, pugi::xml_node root)
{
	auto title = std::find_if(source_.d_->sheet_title_rel_id_map_.begin(),
		source_.d_->sheet_title_rel_id_map_.end(),
		[&](const std::pair<std::string, std::string> &p)
	{
		return p.second == rel.get_id();
	})->first;

	auto ws = source_.get_sheet_by_title(title);

	auto worksheet_node = root.append_child("worksheet");
	worksheet_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	worksheet_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

	if (ws.x14ac_enabled())
	{
		worksheet_node.append_attribute("xmlns:mc").set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
		worksheet_node.append_attribute("mc:Ignorable").set_value("x14ac");
		worksheet_node.append_attribute("xmlns:x14ac").set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
	}

	if (ws.has_page_setup())
	{
		auto sheet_pr_node = worksheet_node.append_child("sheetPr");
		auto outline_pr_node = sheet_pr_node.append_child("outlinePr");

		outline_pr_node.append_attribute("summaryBelow").set_value("1");
		outline_pr_node.append_attribute("summaryRight").set_value("1");

		auto page_set_up_pr_node = sheet_pr_node.append_child("pageSetUpPr");
		page_set_up_pr_node.append_attribute("fitToPage").set_value(ws.get_page_setup().fit_to_page() ? "1" : "0");
	}

	if (ws.has_dimension())
	{
		auto dimension_node = worksheet_node.append_child("dimension");
		auto dimension = ws.calculate_dimension();
		auto dimension_string = dimension.is_single_cell() ? dimension.get_top_left().to_string() : dimension.to_string();
		dimension_node.append_attribute("ref").set_value(dimension_string.c_str());
	}

	if (ws.has_view())
	{
		auto sheet_views_node = worksheet_node.append_child("sheetViews");
		auto sheet_view_node = sheet_views_node.append_child("sheetView");

		const auto view = ws.get_view();

		sheet_view_node.append_attribute("tabSelected").set_value("1");
		sheet_view_node.append_attribute("workbookViewId").set_value("0");

		if (!view.get_selections().empty() && !ws.has_frozen_panes())
		{
			auto selection_node = sheet_view_node.child("selection") ?
				sheet_view_node.child("selection")
				: sheet_view_node.append_child("selection");

			const auto &first_selection = view.get_selections().front();

			if (first_selection.has_active_cell())
			{
				auto active_cell = first_selection.get_active_cell();
				selection_node.append_attribute("activeCell").set_value(active_cell.to_string().c_str());
				selection_node.append_attribute("sqref").set_value(active_cell.to_string().c_str());
			}
		}

		std::string active_pane = "bottomRight";

		if (ws.has_frozen_panes())
		{
			auto pane_node = sheet_view_node.append_child("pane");

			if (ws.get_frozen_panes().get_column_index() > 1)
			{
				pane_node.append_attribute("xSplit").set_value(std::to_string(ws.get_frozen_panes().get_column_index().index - 1).c_str());
				active_pane = "topRight";
			}

			if (ws.get_frozen_panes().get_row() > 1)
			{
				pane_node.append_attribute("ySplit").set_value(std::to_string(ws.get_frozen_panes().get_row() - 1).c_str());
				active_pane = "bottomLeft";
			}

			if (ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
			{
				auto top_right_node = sheet_view_node.append_child("selection");
				top_right_node.append_attribute("pane").set_value("topRight");
				auto bottom_left_node = sheet_view_node.append_child("selection");
				bottom_left_node.append_attribute("pane").set_value("bottomLeft");
				active_pane = "bottomRight";
			}

			pane_node.append_attribute("topLeftCell").set_value(ws.get_frozen_panes().to_string().c_str());
			pane_node.append_attribute("activePane").set_value(active_pane.c_str());
			pane_node.append_attribute("state").set_value("frozen");

			auto selection_node = sheet_view_node.append_child("selection");

			if (ws.get_frozen_panes().get_row() > 1 && ws.get_frozen_panes().get_column_index() > 1)
			{
				selection_node.append_attribute("pane").set_value("bottomRight");
			}
			else if (ws.get_frozen_panes().get_row() > 1)
			{
				selection_node.append_attribute("pane").set_value("bottomLeft");
			}
			else if (ws.get_frozen_panes().get_column_index() > 1)
			{
				selection_node.append_attribute("pane").set_value("topRight");
			}
		}
	}

	if (ws.has_format_properties())
	{
		auto sheet_format_pr_node = worksheet_node.append_child("sheetFormatPr");
		sheet_format_pr_node.append_attribute("baseColWidth").set_value("10");
		sheet_format_pr_node.append_attribute("defaultRowHeight").set_value("16");
		sheet_format_pr_node.append_attribute("x14ac:dyDescent").set_value("0.2");
	}

	bool has_column_properties = false;

	for (auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
	{
		if (ws.has_column_properties(column))
		{
			has_column_properties = true;
			break;
		}
	}

	if (has_column_properties)
	{
		auto cols_node = worksheet_node.append_child("cols");

		for (auto column = ws.get_lowest_column(); column <= ws.get_highest_column(); column++)
		{
			if (!ws.has_column_properties(column))
			{
				continue;
			}

			const auto &props = ws.get_column_properties(column);

			auto col_node = cols_node.append_child("col");

			col_node.append_attribute("min").set_value(std::to_string(column.index).c_str());
			col_node.append_attribute("max").set_value(std::to_string(column.index).c_str());
			col_node.append_attribute("width").set_value(std::to_string(props.width).c_str());
			col_node.append_attribute("style").set_value(std::to_string(props.style).c_str());
			col_node.append_attribute("customWidth").set_value(props.custom ? "1" : "0");
		}
	}

	std::unordered_map<std::string, std::string> hyperlink_references;

	auto sheet_data_node = worksheet_node.append_child("sheetData");
	const auto &shared_strings = ws.get_workbook().get_shared_strings();

	for (auto row : ws.rows())
	{
		auto min = static_cast<xlnt::row_t>(row.num_cells());
		xlnt::row_t max = 0;
		bool any_non_null = false;

		for (auto cell : row)
		{
			min = std::min(min, cell.get_column().index);
			max = std::max(max, cell.get_column().index);

			if (!cell.garbage_collectible())
			{
				any_non_null = true;
			}
		}

		if (!any_non_null)
		{
			continue;
		}

		auto row_node = sheet_data_node.append_child("row");

		row_node.append_attribute("r").set_value(std::to_string(row.front().get_row()).c_str());
		row_node.append_attribute("spans").set_value((std::to_string(min) + ":" + std::to_string(max)).c_str());

		if (ws.has_row_properties(row.front().get_row()))
		{
			row_node.append_attribute("customHeight").set_value("1");
			auto height = ws.get_row_properties(row.front().get_row()).height;

			if (height == std::floor(height))
			{
				row_node.append_attribute("ht").set_value((std::to_string(static_cast<long long int>(height)) + ".0").c_str());
			}
			else
			{
				row_node.append_attribute("ht").set_value(std::to_string(height).c_str());
			}
		}

		// row_node.append_attribute("x14ac:dyDescent", 0.25);

		for (auto cell : row)
		{
			if (!cell.garbage_collectible())
			{
				auto cell_node = row_node.append_child("c");
				cell_node.append_attribute("r").set_value(cell.get_reference().to_string().c_str());

				if (cell.get_data_type() == cell::type::string)
				{
					if (cell.has_formula())
					{
						cell_node.append_attribute("t").set_value("str");
						cell_node.append_child("f").text().set(cell.get_formula().c_str());
						cell_node.append_child("v").text().set(cell.to_string().c_str());

						continue;
					}

					int match_index = -1;

					for (std::size_t i = 0; i < shared_strings.size(); i++)
					{
						if (shared_strings[i] == cell.get_value<text>())
						{
							match_index = static_cast<int>(i);
							break;
						}
					}

					if (match_index == -1)
					{
						if (cell.get_value<std::string>().empty())
						{
							cell_node.append_attribute("t").set_value("s");
						}
						else
						{
							cell_node.append_attribute("t").set_value("inlineStr");
							auto inline_string_node = cell_node.append_child("is");
							inline_string_node.append_child("t").text().set(cell.get_value<std::string>().c_str());
						}
					}
					else
					{
						cell_node.append_attribute("t").set_value("s");
						auto value_node = cell_node.append_child("v");
						value_node.text().set(std::to_string(match_index).c_str());
					}
				}
				else
				{
					if (cell.get_data_type() != cell::type::null)
					{
						if (cell.get_data_type() == cell::type::boolean)
						{
							cell_node.append_attribute("t").set_value("b");
							auto value_node = cell_node.append_child("v");
							value_node.text().set(cell.get_value<bool>() ? "1" : "0");
						}
						else if (cell.get_data_type() == cell::type::numeric)
						{
							if (cell.has_formula())
							{
								cell_node.append_child("f").text().set(cell.get_formula().c_str());
								cell_node.append_child("v").text().set(cell.to_string().c_str());
								continue;
							}

							cell_node.append_attribute("t").set_value("n");
							auto value_node = cell_node.append_child("v");

							if (is_integral(cell.get_value<long double>()))
							{
								value_node.text().set(std::to_string(cell.get_value<long long>()).c_str());
							}
							else
							{
								std::stringstream ss;
								ss.precision(20);
								ss << cell.get_value<long double>();
								ss.str();
								value_node.text().set(ss.str().c_str());
							}
						}
					}
					else if (cell.has_formula())
					{
						cell_node.append_child("f").text().set(cell.get_formula().c_str());
						cell_node.append_child("v");
						continue;
					}
				}

				if (cell.has_format())
				{
					cell_node.append_attribute("s").set_value(std::to_string(cell.get_format().id()).c_str());
				}
			}
		}
	}

	if (ws.has_auto_filter())
	{
		auto auto_filter_node = worksheet_node.append_child("autoFilter");
		auto_filter_node.append_attribute("ref").set_value(ws.get_auto_filter().to_string().c_str());
	}

	if (!ws.get_merged_ranges().empty())
	{
		auto merge_cells_node = worksheet_node.append_child("mergeCells");
		merge_cells_node.append_attribute("count").set_value(std::to_string(ws.get_merged_ranges().size()).c_str());

		for (auto merged_range : ws.get_merged_ranges())
		{
			auto merge_cell_node = merge_cells_node.append_child("mergeCell");
			merge_cell_node.append_attribute("ref").set_value(merged_range.to_string().c_str());
		}
	}

	const auto sheet_rels = source_.get_manifest().get_relationships(rel.get_target().get_path());

	if (!sheet_rels.empty())
	{
		std::vector<relationship> hyperlink_sheet_rels;

		for (const auto &sheet_rel : sheet_rels)
		{
			if (sheet_rel.get_type() == relationship::type::hyperlink)
			{
				hyperlink_sheet_rels.push_back(sheet_rel);
			}
		}

		if (!hyperlink_sheet_rels.empty())
		{
			auto hyperlinks_node = worksheet_node.append_child("hyperlinks");

			for (const auto &hyperlink_rel : hyperlink_sheet_rels)
			{
				auto hyperlink_node = hyperlinks_node.append_child("hyperlink");
				hyperlink_node.append_attribute("display").set_value(hyperlink_rel.get_target().get_path().string().c_str());
				hyperlink_node.append_attribute("ref").set_value(hyperlink_references.at(hyperlink_rel.get_id()).c_str());
				hyperlink_node.append_attribute("r:id").set_value(hyperlink_rel.get_id().c_str());
			}
		}
	}

	if (ws.has_page_setup())
	{
		auto print_options_node = worksheet_node.append_child("printOptions");
		print_options_node.append_attribute("horizontalCentered").set_value(
			ws.get_page_setup().get_horizontal_centered() ? "1" : "0");
		print_options_node.append_attribute("verticalCentered").set_value(
			ws.get_page_setup().get_vertical_centered() ? "1" : "0");
	}

	if (ws.has_page_margins())
	{
		auto page_margins_node = worksheet_node.append_child("pageMargins");

		//TODO: there must be a better way to do this
		auto remove_trailing_zeros = [](const std::string &n)
		{
			auto decimal = n.find('.');

			if (decimal == std::string::npos) return n;

			auto index = n.size() - 1;

			while (index >= decimal && n[index] == '0')
			{
				index--;
			}

			if (index == decimal)
			{
				return n.substr(0, decimal);
			}

			return n.substr(0, index + 1);
		};

		page_margins_node.append_attribute("left").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_left())).c_str());
		page_margins_node.append_attribute("right").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_right())).c_str());
		page_margins_node.append_attribute("top").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_top())).c_str());
		page_margins_node.append_attribute("bottom").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_bottom())).c_str());
		page_margins_node.append_attribute("header").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_header())).c_str());
		page_margins_node.append_attribute("footer").set_value(remove_trailing_zeros(std::to_string(ws.get_page_margins().get_footer())).c_str());
	}

	if (ws.has_page_setup())
	{
		auto page_setup_node = worksheet_node.append_child("pageSetup");

		std::string orientation_string = ws.get_page_setup().get_orientation()
			== xlnt::orientation::landscape ? "landscape" : "portrait";
		page_setup_node.append_attribute("orientation").set_value(orientation_string.c_str());
		page_setup_node.append_attribute("paperSize").set_value(
			std::to_string(static_cast<int>(ws.get_page_setup().get_paper_size())).c_str());
		page_setup_node.append_attribute("fitToHeight").set_value(ws.get_page_setup().fit_to_height() ? "1" : "0");
		page_setup_node.append_attribute("fitToWidth").set_value(ws.get_page_setup().fit_to_width() ? "1" : "0");
	}

	if (!ws.get_header_footer().is_default())
	{
		auto header_footer_node = worksheet_node.append_child("headerFooter");
		auto odd_header_node = header_footer_node.append_child("oddHeader");
		std::string header_text =
			"&L&\"Calibri,Regular\"&K000000Left Header Text&C&\"Arial,Regular\"&6&K445566Center Header "
			"Text&R&\"Arial,Bold\"&8&K112233Right Header Text";
		odd_header_node.text().set(header_text.c_str());
		auto odd_footer_node = header_footer_node.append_child("oddFooter");
		std::string footer_text =
			"&L&\"Times New Roman,Regular\"&10&K445566Left Footer Text_x000D_And &D and &T&C&\"Times New "
			"Roman,Bold\"&12&K778899Center Footer Text &Z&F on &A&R&\"Times New Roman,Italic\"&14&KAABBCCRight Footer "
			"Text &P of &N";
		odd_footer_node.text().set(footer_text.c_str());
	}
}

// Sheet Relationship Target Parts

void xlsx_producer::write_comments(const relationship &rel, pugi::xml_node root)
{
	/*auto comments_node = */root.append_child("comments");
}

void xlsx_producer::write_drawings(const relationship &rel, pugi::xml_node root)
{
	/*auto ws_dr_node = */root.append_child("wsDr");
}

// Other Parts

void xlsx_producer::write_custom_property()
{
}

void xlsx_producer::write_unknown_parts()
{
}

void xlsx_producer::write_unknown_relationships()
{
}

void xlsx_producer::write_thumbnail(const relationship &rel)
{
    const auto &thumbnail = source_.get_thumbnail();
    std::string thumbnail_string(thumbnail.begin(), thumbnail.end());
    destination_.write_string(thumbnail_string, rel.get_target().get_path());
}

std::string xlsx_producer::write_bool(bool boolean) const
{
	if (source_.d_->short_bools_)
	{
		return boolean ? "1" : "0";
	}

	return boolean ? "true" : "false";
}

} // namespace detail
} // namepsace xlnt
