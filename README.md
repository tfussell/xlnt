xlnt
====

## Introduction
xlnt is a C++11 library for reading, writing, and modifying xlsx files. The API is generally based on [openpyxl](https://bitbucket.org/openpyxl/openpyxl), a python library for reading and writing xlsx/xlsm files. It is still very much a work in progress, but the core development work is complete.

## Usage
Including xlnt in your project
```c++
// with -Ixlnt/include
#include <xlnt/xlnt.hpp>
```

Creating a new spreadsheet and saving it
```c++
xlnt::workbook wb;
xlnt::worksheet ws = wb.get_active_sheet();
ws.cell("A1").set_value(5);
ws.cell("B2").set_value("string data");
ws.cell("C3").set_formula("=RAND()");
ws.merge_cells("C3:C4");
ws.freeze_panes("B2");
wb.save("book1.xlsx");
```

Opening an existing spreadsheet and printing all rows
```c++
xlnt::workbook wb2;
wb2.load("book2.xlsx");

// no need to use references, iterators are only wrappers around pointers to memory in the workbook
for(auto row : wb2["sheet2"].rows())
{
    for(auto cell : row)
    {
        std::cout << cell.get_value().to_string() << std::endl;
    }
}
```

## Building
xlnt is regularly built and passes all 200+ tests in GCC 4.8.2, MSVC 12, and Clang 3.4.

Workspaces for Visual Studio 2013, and GNU Make can be created using the development version of premake (currently available [here](https://bitbucket.org/premake/premake-dev), binaries [here](http://sourceforge.net/projects/premake/files/Premake/nightlies/)) and the premake5.lua script in the build directory. XCode workspaces can be generated using premake4.lua and premake4 (binaries [here](http://sourceforge.net/projects/premake/files/Premake/4.3/)).

In Windows, with Visual Studio 2013:
```batch
cd build
premake5 vs2013
start vs2013/xlnt.sln
```

In Linux, with GCC 4.8:
```bash
cd build
premake5 gmake
cd gmake
make
```

In OSX, with Clang 3.4 (can also use gmake as described above):
```bash
cd build
premake4 xcode4
open xcode4/xlnt.xcworkspace
```

## Dependencies
xlnt uses the following libraries, which are included in the source tree for convenience:
- [miniz v1.15_r4](https://code.google.com/p/miniz/) (public domain/unlicense)
- [pugixml v1.4](http://pugixml.org/) (MIT license)

## License
xlnt is currently released to the public for free under the terms of the MIT License:

Copyright (c) 2014 Thomas Fussell

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
