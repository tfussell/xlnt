sudo update-alternatives --install /usr/bin/gcc gcc /usr/bin/gcc-4.9 90
sudo update-alternatives --install /usr/bin/g++ g++ /usr/bin/g++-4.9 90
sudo update-alternatives --install /usr/bin/gcov gcov /usr/bin/gcov-4.9 90
cmake -D DEBUG=1 -D COVERAGE=1 -D SHARED=0 -D STATIC=1 -D TESTS=1 ..
cmake --build .
./bin/xlnt.test
cd ..
export OLDWD=$(pwd)
cd "build/CMakeFiles/xlnt.static.dir$(pwd)"
coveralls --root $OLDWD --verbose -x ".cpp" --gcov-options '\-p' --exclude include --exclude third-party --exclude tests --exclude samples --exclude benchmarks
