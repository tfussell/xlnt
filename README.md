xlnt
====

[![Build Status](https://travis-ci.org/tfussell/xlnt.svg)](https://travis-ci.org/tfussell/xlnt)

## Introduction
xlnt is a C++14 library for reading, writing, and modifying XLSX files as described in [ECMA 376](http://www.ecma-international.org/publications/standards/Ecma-376.htm). The API is roughly based on [openpyxl](https://bitbucket.org/openpyxl/openpyxl), a python library for reading and writing xlsx/xlsm files. This is still very much a work in progress, but the core development work is complete.

## Usage
Including xlnt in your project
```c++
// with -std=c++14 -Ixlnt/include -Lxlnt/lib -lxlnt
#include <xlnt/xlnt.hpp>
```

Creating a new spreadsheet and saving it
```c++
xlnt::workbook wb;
xlnt::worksheet ws = wb.get_active_sheet();
ws.get_cell("A1").set_value(5);
ws.get_cell("B2").set_value("string data");
ws.get_cell("C3").set_formula("=RAND()");
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
        std::cout << cell << std::endl;
    }
}
```

## Building
xlnt is regularly built and passes all 200+ tests in GCC 4.8.2, MSVC 14, and Clang (using Apple LLVM 7.0).

Build configurations for Visual Studio 2015, GNU Make, and Xcode can be created using cmake and the cmake scripts in the project directory, cmake.

In OSX, with Xcode:
```bash
mkdir build
cd build
cmake -G Xcode ../cmake
open xlnt.xcproject
```

In Windows, with Visual Studio 2015:
```batch
mkdir build
cd build
cmake -G "Visual Studio 14 2015" ../cmake
start xlnt.sln
```

In Linux or OSX, with GNU Make:
```bash
mkdir build
cd build
cmake -G "Unix Makefiles" ../cmake
make
```

## Dependencies
xlnt uses the following libraries, which are included in the source tree (pugixml and cxxtest as [git submodules](https://git-scm.com/book/en/v2/Git-Tools-Submodules#Cloning-a-Project-with-Submodules)) for convenience:
- [miniz v1.15_r4](https://code.google.com/p/miniz/) (public domain/unlicense)
- [pugixml v1.6](http://pugixml.org/) (MIT license)
- [cxxtest v4.4](http://cxxtest.com/) (LGPLv3 license [only used for testing, separate from main library assembly])

## License
xlnt is currently released to the public for free under the terms of the MIT License:

Copyright (c) 2015 Thomas Fussell

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
