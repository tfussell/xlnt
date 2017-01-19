
# Getting Started

## Getting xlnt

### Binaries

### Homebrew

### Arch

### vcpkg

## Compiling from Source

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

### Dependencies
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

### Static vs. Dynamic
