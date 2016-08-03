#include <detail/xlsx_writer.hpp>

#include <detail/constants.hpp>
#include <detail/include_pugixml.hpp>
#include <xlnt/utils/path.hpp>
#include <xlnt/packaging/manifest.hpp>
#include <xlnt/packaging/zip_file.hpp>
#include <xlnt/workbook/const_worksheet_iterator.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/worksheet.hpp>

namespace {

void write_document_to_archive(const pugi::xml_document &document, 
	const xlnt::path &archive_path, xlnt::zip_file &archive)
{
	std::ostringstream out_stream;
	document.save(out_stream);
	archive.write_string(out_stream.str(), archive_path);
}

// Package Parts

void write_package_relationships(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;

	auto relationships_node = document.append_child("Relationships");
	relationships_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/relationships");

	for (const auto &relationship : target.get_root_relationships())
	{
		auto relationship_node = relationships_node.append_child("Relationship");

		relationship_node.append_attribute("Id").set_value(relationship.get_id().c_str());
		relationship_node.append_attribute("Type").set_value(relationship.get_type_string().c_str());
		relationship_node.append_attribute("Target").set_value(relationship.get_target_uri().c_str());
	}

	write_document_to_archive(document, xlnt::constants::part_root_relationships(), archive);
}

void write_content_types(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;

	auto types_node = document.append_child("Types");
	types_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/content-types");

	for (const auto &default_type : target.get_manifest().get_default_types())
	{
		auto default_node = types_node.append_child("Default");
		default_node.append_attribute("Extension").set_value(default_type.second.get_extension().c_str());
		default_node.append_attribute("ContentType").set_value(default_type.second.get_content_type().c_str());
	}

	for (const auto &override_type : target.get_manifest().get_override_types())
	{
		auto override_node = types_node.append_child("Override");
		override_node.append_attribute("PartName").set_value(override_type.second.get_part().to_string('/').c_str());
		override_node.append_attribute("ContentType").set_value(override_type.second.get_content_type().c_str());
	}

	write_document_to_archive(document, xlnt::constants::part_content_types(), archive);
}

void write_app_properties(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("appProperties");
	write_document_to_archive(document, xlnt::constants::part_app(), archive);
}

void write_core_properties(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("coreProperties");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_custom_file_properties(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("customFileProperties");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

// SpreadsheetML Package Parts

void write_workbook(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;

	auto workbook_node = document.append_child("workbook");
	workbook_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	workbook_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

	auto sheets_node = workbook_node.append_child("sheets");
	auto sheet_node = sheets_node.append_child("sheet");
	sheet_node.append_attribute("name").set_value(1);
	sheet_node.append_attribute("sheetId").set_value(1);
	sheet_node.append_attribute("r:id").set_value("rId1");

	write_document_to_archive(document, xlnt::path("workbook.xml"), archive);
}

void write_workbook_relationships(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;

	auto relationships_node = document.append_child("Relationships");
	relationships_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/package/2006/relationships");

	auto relationship_node = relationships_node.append_child("Relationship");
	relationship_node.append_attribute("Id").set_value("rId1");
	relationship_node.append_attribute("Type").set_value("http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet");
	relationship_node.append_attribute("Target").set_value("sheet1.xml");

	write_document_to_archive(document, xlnt::path("_rels/workbook.xml.rels"), archive);
}

// Workbook Relationship Target Parts

void write_calculation_chain(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("calcChain");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_chartsheet(const xlnt::worksheet &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("chartsheet");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_connections(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("connections");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_custom_property(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("customProperty");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_custom_xml_mappings(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("connections");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_dialogsheet(const xlnt::worksheet &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("dialogsheet");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_external_workbook_references(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("externalLink");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_metadata(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("metadata");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_pivot_table(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("pivotTableDefinition");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_shared_string_table(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("sst");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_shared_workbook_revision_headers(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("headers");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_shared_workbook(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("revisions");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_shared_workbook_user_data(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("users");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_styles(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("styleSheet");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_theme(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("theme");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_volatile_dependencies(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("volTypes");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_worksheet(const xlnt::worksheet &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;

	auto worksheet_node = document.append_child("worksheet");
	worksheet_node.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	worksheet_node.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/package/2006/relationships");

	worksheet_node.append_child("sheetData");

	write_document_to_archive(document, xlnt::path("sheet1.xml"), archive);
}

// Sheet Relationship Target Parts

void write_comments(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("comments");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_drawings(const xlnt::worksheet &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("wsDr");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

// Unknown Parts

void write_unknown_parts(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("relationships");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

void write_unknown_relationships(const xlnt::workbook &target, xlnt::zip_file &archive)
{
	pugi::xml_document document;
	auto root_node = document.append_child("Relationships");
	write_document_to_archive(document, xlnt::constants::part_core(), archive);
}

} // namespace

namespace xlnt
{

xlsx_writer::xlsx_writer(const workbook &target) : target_(target)
{
}

void xlsx_writer::write(const path &destination)
{
	xlnt::zip_file archive;
	populate_archive(archive);
	archive.save(destination);
}

void xlsx_writer::write(std::ostream &destination)
{
	xlnt::zip_file archive;
	populate_archive(archive);
	archive.save(destination);
}

void xlsx_writer::write(std::vector<std::uint8_t> &destination)
{
	xlnt::zip_file archive;
	populate_archive(archive);
	archive.save(destination);
}

void xlsx_writer::populate_archive(zip_file &archive)
{
	write_package_relationships(target_, archive);
	write_content_types(target_, archive);

	write_workbook(target_, archive);
	write_workbook_relationships(target_, archive);

	for (auto ws : target_)
	{
		write_worksheet(ws, archive);
	}
}

} // namepsace xlnt
