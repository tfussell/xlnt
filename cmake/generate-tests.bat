cd %~dp0
../third-party/cxxtest/bin/cxxtestgen --runner=ErrorPrinter -o "%1"/tests/runner-autogen.cpp ../tests/*.hpp ../source/*/tests/*.hpp
