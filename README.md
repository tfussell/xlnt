<img height="100" src="https://raw.githubusercontent.com/tfussell/xlnt/gh-pages/images/xlnt.png" alt="xlnt"><br/>
====

[![Travis Build Status](https://travis-ci.org/tfussell/xlnt.svg)](https://travis-ci.org/tfussell/xlnt)
[![AppVeyor Build status](https://ci.appveyor.com/api/projects/status/2hs79a1xoxy16sol?svg=true)](https://ci.appveyor.com/project/tfussell/xlnt)
[![Coverage Status](https://coveralls.io/repos/github/tfussell/xlnt/badge.svg?branch=master)](https://coveralls.io/github/tfussell/xlnt?branch=master)
[![ReadTheDocs Documentation Status](https://readthedocs.org/projects/xlnt/badge/?version=latest)](http://xlnt.readthedocs.org/en/latest/?badge=latest)
[![License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](http://opensource.org/licenses/MIT)

## Introduction
xlnt is a C++14 library for reading, writing, and modifying xlsx files as described in [ECMA 376 4th edition](http://www.ecma-international.org/publications/standards/Ecma-376.htm). The API was initially based on [openpyxl](https://bitbucket.org/openpyxl/openpyxl), a Python library to read/write Excel 2007 xlsx/xlsm files, which was itself based on [PHPExcel](https://github.com/PHPOffice/PHPExcel), pure PHP library for reading and writing spreadsheet files. xlnt is still very much a work in progress, but the core development work is complete. For a high-level summary of what you can do with this library, see [here](http://xlnt.readthedocs.io/en/latest/#summary-of-features).

## Usage
Including xlnt in your project
```c++
// with -std=c++14 -Ixlnt/include -Lxlnt/lib -lxlnt
#include <xlnt/xlnt.hpp>
```

Creating a new spreadsheet and saving it as "example.xlsx"
```c++
xlnt::workbook wb;
xlnt::worksheet ws = wb.get_active_sheet();
ws.get_cell("A1").set_value(5);
ws.get_cell("B2").set_value("string data");
ws.get_cell("C3").set_formula("=RAND()");
ws.merge_cells("C3:C4");
ws.freeze_panes("B2");
wb.save("example.xlsx");
```

## Building
xlnt uses continous integration (thanks [Travis CI](https://travis-ci.org/) and [AppVeyor](https://www.appveyor.com/)!) and passes all 270+ tests in GCC 4, GCC 5, VS2015 U3, and Clang (using Apple LLVM 8.0). Build configurations for Visual Studio 2015, GNU Make, and Xcode can be created using [cmake](https://cmake.org/) v3.1+. A full list of cmake generators can be found [here](https://cmake.org/cmake/help/v3.0/manual/cmake-generators.7.html). A basic build would look like (starting in the root xlnt directory):

```bash
mkdir build
cd build
cmake ..
make -j8
```

The resulting shared (e.g. libxlnt.dylib) library would be found in the build/lib directory. Other cmake configuration options for xlnt can be found using "cmake -LH". These options include building a static library instead of shared and whether to build sample executables or not. An example of building a static library with an Xcode project:

```bash
mkdir build
cd build
cmake -D STATIC=ON -G Xcode ..
cmake --build .
cd bin && ./xlnt.test
```
*Note for Windows: cmake defaults to building a 32-bit library project. To build a 64-bit library, use the Win64 generator*
```bash
cmake -G "Visual Studio 14 2015 Win64" ..
```

## Dependencies
xlnt requires the following libraries which are all distributed under permissive open source licenses. All libraries are included in the source tree as [git submodules](https://git-scm.com/book/en/v2/Git-Tools-Submodules#Cloning-a-Project-with-Submodules) for convenience except for pole and partio's zip component which have been modified and reside in xlnt/source/detail:
- [zlib v1.2.8](http://zlib.net/) (zlib License)
- [libstudxml v1.1.0](http://www.codesynthesis.com/projects/libstudxml/) (MIT license)
- [utfcpp v2.3.4](http://utfcpp.sourceforge.net/) (Boost Software License, Version 1.0)
- [Crypto++ v5.6.5](https://www.cryptopp.com/) (Boost Software License, Version 1.0)
- [pole v0.5](https://github.com/catlan/pole) (BSD 2-Clause License)
- [partio v1.1.0](https://github.com/wdas/partio) (BSD 3-Clause License with specific non-attribution clause)

Additionally, the xlnt test suite (bin/xlnt.test) requires an extra testing library. This test executable is separate from the main xlnt library assembly so the terms of this library's license should not apply to users of solely the xlnt library:
- [cxxtest v4.4](http://cxxtest.com/) (LGPLv3 License)

Initialize the submodules from the HEAD of their respective Git repositories with this command called from the xlnt root directory:
```bash
git submodule update --init --remote
```

## Documentation

More detailed documentation with examples can be found on [Read The Docs](http://xlnt.readthedocs.org/en/latest/) (Warning: As of November 9, 2016 this is very out of date).

## License
xlnt is released to the public for free under the terms of the MIT License. See [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENCE.md) for the full text of the license and the licenses of xlnt's third-party dependencies. [LICENSE.md](https://github.com/tfussell/xlnt/blob/master/LICENCE.md) should be distributed alongside any assemblies that use xlnt in source or compiled form.

