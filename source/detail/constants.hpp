#pragma once

#include <string>
#include <unordered_map>

#include <xlnt/cell/types.hpp>

namespace xlnt {

struct constants
{
    static const row_t MinRow();
    static const row_t MaxRow();
    static const column_t MinColumn();
    static const column_t MaxColumn();

    // constants
    static const string PackageProps();
    static const string PackageXl();
    static const string PackageRels();
    static const string PackageTheme();
    static const string PackageWorksheets();
    static const string PackageDrawings();
    static const string PackageCharts();

    static const string ArcContentTypes();
    static const string ArcRootRels();
    static const string ArcWorkbookRels();
    static const string ArcCore();
    static const string ArcApp();
    static const string ArcWorkbook();
    static const string ArcStyles();
    static const string ArcTheme();
    static const string ArcSharedString();

    static const std::unordered_map<string, string> Namespaces();
    
    static const string Namespace(const string &id);
};

} // namespace xlnt
