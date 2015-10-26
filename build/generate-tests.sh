DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd $DIR
../third-party/cxxtest/bin/cxxtestgen --runner=ErrorPrinter -o ../tests/runner-autogen.cpp ../tests/*.hpp