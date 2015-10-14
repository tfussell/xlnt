ACTION=$1

if [ "$ACTION" = "clean" ]; then
    rm -rf build
    exit 0
fi

mkdir build
cd build
cmake ..
make
