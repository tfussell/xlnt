#!/bin/sh

for var in "$@"; do
    split=(${var/=/ })
    export ${split[0]}=${split[1]}
done

if [ "$USE_CMAKE" = 1 ]; then
    cd cmake
else
    cd genie
fi

./build-linux.sh $1
