#pragma once

#include <string>
#include <unordered_map>

namespace xlnt {

struct constants
{
    static const int MinRow;
    static const int MinColumn;
    static const int MaxColumn;
    static const int MaxRow;

    // constants
    static const std::string PackageProps;
    static const std::string PackageXl;
    static const std::string PackageRels;
    static const std::string PackageTheme;
    static const std::string PackageWorksheets;
    static const std::string PackageDrawings;
    static const std::string PackageCharts;

    static const std::string ArcContentTypes;
    static const std::string ArcRootRels;
    static const std::string ArcWorkbookRels;
    static const std::string ArcCore;
    static const std::string ArcApp;
    static const std::string ArcWorkbook;
    static const std::string ArcStyles;
    static const std::string ArcTheme;
    static const std::string ArcSharedString;

    static const std::unordered_map<std::string, std::string> Namespaces;
};
    
} // namespace xlnt
