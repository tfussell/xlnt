# Getting xlnt

## Binaries

## Homebrew

## Arch

## vcpkg

## Compiling from Source - Ubuntu 16.04 LTS (Xenial Xerus) 
Please run the following commands and read the comments as you go.
You will notice that we update and upgrade quite frequently, this is to ensure that all dependancies are met at all times
```
sudo apt-get update
sudo apt-get upgrade
sudo apt-get install cmake
sudo apt-get install zlibc
sudo apt-get install zlib1g
sudo apt-get install zlib1g-dev
sudo apt-get install libssl-dev
sudo add-apt-repository ppa:ubuntu-toolchain-r/test
sudo apt update
sudo apt-get upgrade
sudo apt-get install gcc-6 g++-6
sudo apt update
sudo apt-get upgrade
```
The following steps will set up some environment variables for this session
```
export CC=/usr/bin/gcc-6  
export CXX=/usr/bin/g++-6
```
The following steps will install the appropriate crypto libraries for c++
Download the zip file from https://github.com/weidai11/cryptopp/archive/master.zip
```
cd ~
unzip Downloads/cryptopp-master.zip
cd cryptopp-master
make -j 2
sudo make install
```
The following steps will install the latest compiler from source
Download gcc 6.3.0 from https://gcc.gnu.org/mirrors.html I used the Japanese mirror as this is the closest location to me http://ftp.tsukuba.wide.ad.jp/software/gcc/releases/gcc-6.3.0/gcc-6.3.0.tar.gz

```
cd ~
tar -zxvf Downloads/gcc-6.3.0.tar.gz
cd gcc-6.3.0
./contrib/download_prerequisites
mkdir build
cd build
../configure --disable-multilib --enable-languages=c,c++
make -j 2
sudo make install
```
If the above make command fails please use "make clean pristine" and then remove and remake the build directory. This will clean up and prepare the environment for another attempt. A common reason for failure on virtual machines is lack of RAM (if you don't have enough RAM you may get an error like this "recipe for target 's-attrtab' failed"). Otherwise, if all goes well in about 30 to 60 minutes the compiler will be ready and we will move on to the next steps.

The following step will install the docbook2x library
```
sudo apt-get install docbook2x
```
The following steps will install the libexpat library
Download the zip file from the following repository https://github.com/libexpat/libexpat 
```
cd ~
unzip Downloads/libexpat-master.zip
cd libexpat-master/expat
cmake .
make -j 2
sudo make install
```
The following step will map the shared library names to the location of the corresponding shared library files
```
sudo ldconfig
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
xlnt will now be ready to use on your Ubuntu instance. 

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
