xlnt
====

Introduction
----
xlnt is a c++ library that reads and write XLSX files. The API is roughly based on openpyxl, a python XLSX library. It is still very much a work in progress, but I expect the basic functionality to be working in the near future.

Building
----
It compiles in all of the major compilers. Currently it is being built in GCC 4.8.2, MSVC 12, and Clang 3.3.

Workspaces for Visual Studio, XCode, and GNU Make can be created using premake and the premake5.lua file in the build directory.

Dependencies
----
xlnt requires the following libraries:
- [libopc v0.0.3 + mce + plib](http://libopc.codeplex.com/)
    - [zlib v1.2.5](https://github.com/madler/zlib)
    - [libxml2 v2.7.7](http://xmlsoft.org/index.html)
    - [python v2.7+](https://www.python.org/)
- [pugixml v1.4](http://pugixml.org/)

License
----
xlnt is currently released under the terms of the MIT License.