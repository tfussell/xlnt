<img height="100" src="https://user-images.githubusercontent.com/1735211/29433390-f37fa28e-836c-11e7-8a60-f8df4c30b424.png" alt="xlnt logo"><br/>
====

[![Travis Build Status](https://travis-ci.org/tfussell/xlnt.svg?branch=master)](https://travis-ci.org/tfussell/xlnt)
[![AppVeyor Build status](https://ci.appveyor.com/api/projects/status/2hs79a1xoxy16sol?svg=true)](https://ci.appveyor.com/project/tfussell/xlnt)
[![Coverage Status](https://coveralls.io/repos/github/tfussell/xlnt/badge.svg?branch=master)](https://coveralls.io/github/tfussell/xlnt?branch=master)
[![Documentation Status](https://legacy.gitbook.com/button/status/book/tfussell/xlnt)](https://tfussell.gitbooks.io/xlnt/content/)
[![License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](http://opensource.org/licenses/MIT)

## Introduction
xlnt is a modern C++ library for manipulating spreadsheets in memory and reading/writing them from/to XLSX files as described in [ECMA 376 4th edition](http://www.ecma-international.org/publications/standards/Ecma-376.htm). The first public release of xlnt version 1.0 was on May 10th, 2017. Current work is focused on increasing compatibility, improving performance, and brainstorming future development goals. For a high-level summary of what you can do with this library, see [the feature list](https://tfussell.gitbooks.io/xlnt/content/docs/introduction/Features.html). Contributions are welcome in the form of pull requests or discussions on [the repository's Issues page](https://github.com/tfussell/xlnt/issues).

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
// compile with -std=c++14 -Ixlnt/include -lxlnt
```

## Documentation

Documentation for the current release of xlnt is available [here](https://tfussell.gitbooks.io/xlnt/content/).

## License
xlnt is released to the public for free under the terms of the MIT License. See [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENSE.md) for the full text of the license and the licenses of xlnt's third-party dependencies. [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENSE.md) should be distributed alongside any assemblies that use xlnt in source or compiled form.
