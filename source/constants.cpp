#include <limits>

#include "constants.h"
#include "config.h"

namespace xlnt {

const row_t constants::MinRow = 1;
const row_t constants::MaxRow = LimitStyle == limit_style::excel ? (1u << 20) : UINT32_MAX;
const column_t constants::MinColumn = 1;
const column_t constants::MaxColumn = LimitStyle == limit_style::excel ? (1u << 14) : LimitStyle == limit_style::openpyxl ? 18278 : UINT32_MAX;

// constants
const std::string constants::PackageProps = "docProps";
const std::string constants::PackageXl = "xl";
const std::string constants::PackageRels = "_rels";
const std::string constants::PackageTheme = PackageXl + "/" + "theme";
const std::string constants::PackageWorksheets = PackageXl + "/" + "worksheets";
const std::string constants::PackageDrawings = PackageXl + "/" + "drawings";
const std::string constants::PackageCharts = PackageXl + "/" + "charts";

const std::string constants::ArcContentTypes = "[Content_Types].xml";
const std::string constants::ArcRootRels = PackageRels + "/.rels";
const std::string constants::ArcWorkbookRels = PackageXl + "/" + PackageRels + "/workbook.xml.rels";
const std::string constants::ArcCore = PackageProps + "/core.xml";
const std::string constants::ArcApp = PackageProps + "/app.xml";
const std::string constants::ArcWorkbook = PackageXl + "/workbook.xml";
const std::string constants::ArcStyles = PackageXl + "/styles.xml";
const std::string constants::ArcTheme = PackageTheme + "/theme1.xml";
const std::string constants::ArcSharedString = PackageXl + "/sharedStrings.xml";

const std::unordered_map<std::string, std::string> constants::Namespaces =
{
    {"spreadsheetml", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
    {"content-types", "http://schemas.openxmlformats.org/package/2006/content-types"},
    {"relationships", "http://schemas.openxmlformats.org/package/2006/relationships"},
    {"drawingml", "http://schemas.openxmlformats.org/drawingml/2006/main"},
    {"r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
    {"cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
    {"dc", "http://purl.org/dc/elements/1.1/"},
    {"dcterms", "http://purl.org/dc/terms/"},
    {"dcmitype", "http://purl.org/dc/dcmitype/"},
    {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
    {"vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"},
    {"xml", "http://www.w3.org/XML/1998/namespace"}
};

}