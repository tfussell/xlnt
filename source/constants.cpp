#include "constants.h"

namespace xlnt {

const int constants::MinRow = 0;
const int constants::MinColumn = 0;
const int constants::MaxColumn = 16384;
const int constants::MaxRow = 1048576;

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
    {"cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
    {"dc", "http://purl.org/dc/elements/1.1/"},
    {"dcterms", "http://purl.org/dc/terms/"},
    {"dcmitype", "http://purl.org/dc/dcmitype/"},
    {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
    {"vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"},
    {"xml", "http://www.w3.org/XML/1998/namespace"}
};

}