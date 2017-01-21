<img height="100" src="https://raw.githubusercontent.com/tfussell/xlnt/gh-pages/images/logo.png" alt="xlnt"><br/>
====

[![Travis Build Status](https://travis-ci.org/tfussell/xlnt.svg)](https://travis-ci.org/tfussell/xlnt)
[![AppVeyor Build status](https://ci.appveyor.com/api/projects/status/2hs79a1xoxy16sol?svg=true)](https://ci.appveyor.com/project/tfussell/xlnt)
[![Coverage Status](https://coveralls.io/repos/github/tfussell/xlnt/badge.svg?branch=master)](https://coveralls.io/github/tfussell/xlnt?branch=master)
[![ReadTheDocs Documentation Status](https://readthedocs.org/projects/xlnt/badge/?version=latest)](http://xlnt.readthedocs.org/en/latest/?badge=latest)
[![License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](http://opensource.org/licenses/MIT)

## Introduction
xlnt is a modern C++ library for manipulating spreadsheets in memory and reading/writing them from/to XLSX files as described in [ECMA 376 4th edition](http://www.ecma-international.org/publications/standards/Ecma-376.htm). xlnt is currently under active feature development and is on track for the version 1.0 release in the next few weeks. Until then, the API could have significant changes. For a high-level summary of what you can do with this library, see [here](https://thomas.fussell.io/xlnt/#features).

## Example

Including xlnt in your project, creating a new spreadsheet, and saving it as "example.xlsx"

```c++
#include <xlnt/xlnt.hpp>

int main()
{
    xlnt::workbook wb;
    xlnt::worksheet ws = wb.active_sheet();
    ws.cell("A1").value(5);
    ws.cell("B2").value("string data");
    ws.cell("C3").formula("=RAND()");
    ws.merge_cells("C3:C4");
    ws.freeze_panes("B2");
    wb.save("example.xlsx");
    return 0;
}
// compile with -std=c++14 -Ixlnt/include -Lxlnt/lib -lxlnt
```

## Documentation

Documentation for the current release of xlnt is available [here](https://thomas.fussell.io/xlnt).

## License
xlnt is released to the public for free under the terms of the MIT License. See [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENCE.md) for the full text of the license and the licenses of xlnt's third-party dependencies. [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENCE.md) should be distributed alongside any assemblies that use xlnt in source or compiled form.

