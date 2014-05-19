xlnt
====

Introduction
----
xlnt is a c++ library that reads and write XLSX files. The API is roughly based on openpyxl, a python XLSX library. It is still very much a work in progress, but I expect the basic functionality to be working in the near future.

Usage
----
Using xlnt in your project
```c++
#include <xlnt.h>
```

Create a new spreadsheet
```c++
xlnt::workbook wb;
xlnt::worksheet ws = wb.get_active_sheet();
ws.cell("A1") = 5;
ws.cell("B2") = "string data";
ws.cell("C3") = "=RAND()";
ws.merge_cells("C3:C4");
ws.freeze_panes("B2");
wb.save("book1.xlsx");
```

Open an existing spreadsheet
```c++
xlnt::workbook wb2;
wb2.load("book2.xlsx");
wb2["sheet1"].get_dimensions();
for(auto &row : wb2["sheet2"])
{
    for(auto cell : row)
    {
        std::cout << cell.to_string() << std::endl;
    }
}
```

Building
----
It compiles in all of the major compilers. Currently it is being built in GCC 4.8.2, MSVC 12, and Clang 3.3.

Workspaces for Visual Studio, XCode, and GNU Make can be created using premake and the premake5.lua file in the build directory.

Dependencies
----
xlnt requires the following libraries:
- [zlib v1.2.8](http://zlib.net/) (zlib/libpng license)
- [pugixml v1.4](http://pugixml.org/) (MIT license)

License
----
xlnt is currently released under the terms of the MIT License:

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
