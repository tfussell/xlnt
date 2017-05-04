# Getting xlnt

## Binaries

## Homebrew

## Arch

## vcpkg

## Compiling xlnt 1.x.x from Source on Ubuntu 16.04 LTS (Xenial Xerus) 
Time required: Approximately 5 minutes (depending on your internet speed)
```
sudo apt-get update
sudo apt-get upgrade
sudo apt-get install cmake
sudo apt-get install zlibc
```
The following steps update the compiler and set the appropriate environment variables - see note [1] below for the reason why we need to update the standard available compiler
```
sudo add-apt-repository ppa:ubuntu-toolchain-r/test
sudo apt update
sudo apt-get upgrade
sudo apt-get install gcc-6 g++-6
export CC=/usr/bin/gcc-6  
export CXX=/usr/bin/g++-6
```
The following steps will intall xlnt
Download the zip file from the xlnt repository
https://github.com/tfussell/xlnt/archive/master.zip
```
cd ~
unzip Downloads/xlnt-master.zip
cd xlnt-master
cmake .
make -j 2
sudo make install
```
The following step will map the shared library names to the location of the corresponding shared library files
```
sudo ldconfig
```
xlnt will now be ready to use on your Ubuntu instance. 

[1] 
Xlnt requires a minimum of gcc 6.2.0
The most recent gcc version available using the standard APT repositories is gcc 5.4.0 (obtained through build-essential 12.1ubuntu2). If these older versions of gcc are used an error "workbook.cpp error 1502:31 'extended_property' is not a class, namespace or enumeration" will occur during the xlnt make command.

## Compiling from Source

Build configurations for Visual Studio, GNU Make, Ninja, and Xcode can be created using [cmake](https://cmake.org/) v3.2+. A full list of cmake generators can be found [here](https://cmake.org/cmake/help/v3.0/manual/cmake-generators.7.html). A basic build would look like (starting in the root xlnt directory):

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
