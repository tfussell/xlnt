cmake_minimum_required(VERSION 3.1)
project(xlnt VERSION 1.2)

# Require C99 and C++11 compilers
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Project metadata
set(PROJECT_VENDOR "Thomas Fussell")
set(PROJECT_CONTACT "thomas.fussell@gmail.com")
set(PROJECT_URL "https://github.com/tfussell/xlnt")
set(PROJECT_DESCRIPTION "cross-platform user-friendly xlsx library for C++11")

# Commonly used project directories
set(XLNT_ROOT_DIR ${CMAKE_CURRENT_SOURCE_DIR}/..)
set(XLNT_INCLUDE_DIR ${XLNT_ROOT_DIR}/include)
set(XLNT_SOURCE_DIR ${XLNT_ROOT_DIR}/source)
set(THIRD_PARTY_DIR ${XLNT_ROOT_DIR}/third-party)

# Include libstudxml library
add_subdirectory(${THIRD_PARTY_DIR}/libstudxml
  ${CMAKE_CURRENT_BINARY_DIR}/third-party/libstudxml)

if(COVERAGE)
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

# Non-target-specific compiler settings
if(MSVC)
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /W4") # level 4 warnings
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /MP") # multi-processor compilation
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall") # all warnings
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wextra") # extra warnings
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unknown-pragmas") # ignore MSVC and Clang pragmas
elseif(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Weverything") # all warnings
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Werror") # warnings are errors
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-c++98-compat") # ignore warnings about C++98 compatibility
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-c++98-compat-pedantic") # ignore pedantic warnings about C++98 compatibility
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-padded") # ignore padding warnings
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-documentation-unknown-command") # ignore unknown commands in Javadoc-style comments
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unknown-pragmas") # ignore Windows and GCC pragmas
endif()

if(STATIC_CRT)
  include(${XLNT_ROOT_DIR}/cmake/ucm.cmake)
  ucm_set_runtime(STATIC)
endif()

if(APPLE)
  # Prevent a warning about deployment target not being set by setting it to current OSX version
  execute_process(COMMAND "sw_vers -productVersion | awk -F'.' '{print $1\".\"$2}'"
    OUTPUT_VARIABLE OSX_VERSION)
  set(CMAKE_OSX_DEPLOYMENT_TARGET ${OSX_VERSION})
endif()

file(GLOB ROOT_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/*.hpp)
file(GLOB CELL_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/cell/*.hpp)
file(GLOB CELL_SOURCES ${XLNT_SOURCE_DIR}/cell/*.cpp)
file(GLOB CHARTS_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/charts/*.hpp)
file(GLOB CHARTS_SOURCES ${XLNT_SOURCE_DIR}/charts/*.cpp)
file(GLOB CHARTSHEET_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/chartsheet/*.hpp)
file(GLOB CHARTSHEET_SOURCES ${XLNT_SOURCE_DIR}/chartsheet/*.cpp)
file(GLOB DRAWING_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/drawing/*.hpp)
file(GLOB DRAWING_SOURCES ${XLNT_SOURCE_DIR}/drawing/*.cpp)
file(GLOB FORMULA_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/formula/*.hpp)
file(GLOB FORMULA_SOURCES ${XLNT_SOURCE_DIR}/formula/*.cpp)
file(GLOB PACKAGING_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/packaging/*.hpp)
file(GLOB PACKAGING_SOURCES ${XLNT_SOURCE_DIR}/packaging/*.cpp)
file(GLOB STYLES_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/styles/*.hpp)
file(GLOB STYLES_SOURCES ${XLNT_SOURCE_DIR}/styles/*.cpp)
file(GLOB UTILS_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/utils/*.hpp)
file(GLOB UTILS_SOURCES ${XLNT_SOURCE_DIR}/utils/*.cpp)
file(GLOB WORKBOOK_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/workbook/*.hpp)
file(GLOB WORKBOOK_SOURCES ${XLNT_SOURCE_DIR}/workbook/*.cpp)
file(GLOB WORKSHEET_HEADERS ${XLNT_INCLUDE_DIR}/xlnt/worksheet/*.hpp)
file(GLOB WORKSHEET_SOURCES ${XLNT_SOURCE_DIR}/worksheet/*.cpp)

file(GLOB DETAIL_ROOT_HEADERS ${XLNT_SOURCE_DIR}/detail/*.hpp)
file(GLOB DETAIL_ROOT_SOURCES ${XLNT_SOURCE_DIR}/detail/*.cpp)
file(GLOB DETAIL_CRYPTOGRAPHY_HEADERS ${XLNT_SOURCE_DIR}/detail/cryptography/*.hpp)
file(GLOB DETAIL_CRYPTOGRAPHY_SOURCES ${XLNT_SOURCE_DIR}/detail/cryptography/*.c*)
file(GLOB DETAIL_EXTERNAL_HEADERS ${XLNT_SOURCE_DIR}/detail/external/*.hpp)
file(GLOB DETAIL_HEADER_FOOTER_HEADERS ${XLNT_SOURCE_DIR}/detail/header_footer/*.hpp)
file(GLOB DETAIL_HEADER_FOOTER_SOURCES ${XLNT_SOURCE_DIR}/detail/header_footer/*.cpp)
file(GLOB DETAIL_IMPLEMENTATIONS_HEADERS ${XLNT_SOURCE_DIR}/detail/implementations/*.hpp)
file(GLOB DETAIL_IMPLEMENTATIONS_SOURCES ${XLNT_SOURCE_DIR}/detail/implementations/*.cpp)
file(GLOB DETAIL_NUMBER_FORMAT_HEADERS ${XLNT_SOURCE_DIR}/detail/number_format/*.hpp)
file(GLOB DETAIL_NUMBER_FORMAT_SOURCES ${XLNT_SOURCE_DIR}/detail/number_format/*.cpp)
file(GLOB DETAIL_SERIALIZATION_HEADERS ${XLNT_SOURCE_DIR}/detail/serialization/*.hpp)
file(GLOB DETAIL_SERIALIZATION_SOURCES ${XLNT_SOURCE_DIR}/detail/serialization/*.cpp)

set(DETAIL_HEADERS ${DETAIL_ROOT_HEADERS} ${DETAIL_CRYPTOGRAPHY_HEADERS}
  ${DETAIL_EXTERNAL_HEADERS} ${DETAIL_HEADER_FOOTER_HEADERS}
  ${DETAIL_IMPLEMENTATIONS_HEADERS} ${DETAIL_NUMBER_FORMAT_HEADERS}
  ${DETAIL_SERIALIZATION_HEADERS})
set(DETAIL_SOURCES ${DETAIL_ROOT_SOURCES} ${DETAIL_CRYPTOGRAPHY_SOURCES}
  ${DETAIL_EXTERNAL_SOURCES} ${DETAIL_HEADER_FOOTER_SOURCES}
  ${DETAIL_IMPLEMENTATIONS_SOURCES} ${DETAIL_NUMBER_FORMAT_SOURCES}
  ${DETAIL_SERIALIZATION_SOURCES})

set(XLNT_HEADERS ${ROOT_HEADERS} ${CELL_HEADERS} ${CHARTS_HEADERS}
  ${CHARTSHEET_HEADERS} ${DRAWING_HEADERS} ${FORMULA_HEADERS}
  ${PACKAGING_HEADERS} ${STYLES_HEADERS} ${UTILS_HEADERS}
  ${WORKBOOK_HEADERS} ${WORKSHEET_HEADERS} ${DETAIL_HEADERS} ${DETAIL_CRYPTO_HEADERS})
set(XLNT_SOURCES ${CELL_SOURCES} ${CHARTS_SOURCES} ${CHARTSHEET_SOURCES}
  ${DRAWING_SOURCES} ${FORMULA_SOURCES} ${PACKAGING_SOURCES}
  ${STYLES_SOURCES} ${UTILS_SOURCES} ${WORKBOOK_SOURCES}
  ${WORKSHEET_SOURCES} ${DETAIL_SOURCES} ${DETAIL_CRYPTO_SOURCES})

if(CMAKE_INSTALL_PREFIX_INITIALIZED_TO_DEFAULT)
  # Set a default CMAKE_INSTALL_PREFIX if one wasn't specified

  if(MSVC)
    # No good place to install on Windows so just install to cmake build directory in /installed
    set(DEFAULT_INSTALL_PREFIX "${CMAKE_BINARY_DIR}/installed")
  else()
    # Install to /usr/local on other platforms
    set(DEFAULT_INSTALL_PREFIX "/usr/local")
  endif()

  # Apply the variable and save to cache
  set(CMAKE_INSTALL_PREFIX ${DEFAULT_INSTALL_PREFIX}
    CACHE PATH "default install path" FORCE)
endif()

# Append "d" to the name of the compiled library
set(CMAKE_DEBUG_POSTFIX "d" CACHE STRING "Set debug library postfix")

if(NOT STATIC)
  # Compile shared library
  add_library(xlnt SHARED
    ${XLNT_HEADERS}
    ${XLNT_SOURCES}
    $<TARGET_OBJECTS:libstudxml>)

  target_compile_definitions(xlnt PRIVATE XLNT_SHARED=1)

  # Set SO version (for symbolic links like libxlnt.so.1.2 and libxlnt.so.1)
  set_target_properties(xlnt
    PROPERTIES
    VERSION ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}.${PROJECT_VERSION_MINOR}
    INSTALL_NAME_DIR "${LIB_DEST_DIR}")
else()
  # Compile static library
  add_library(xlnt STATIC ${XLNT_HEADERS} ${XLNT_SOURCES} $<TARGET_OBJECTS:libstudxml>)
  target_compile_definitions(xlnt PUBLIC XLNT_STATIC=1)
endif()

# Includes
target_include_directories(xlnt PUBLIC ${XLNT_INCLUDE_DIR})
target_include_directories(xlnt PRIVATE ${XLNT_SOURCE_DIR})
target_include_directories(xlnt PRIVATE ${XLNT_SOURCE_DIR}/../third-party/libstudxml)
target_include_directories(xlnt PRIVATE ${XLNT_SOURCE_DIR}/../third-party/utfcpp)

# Platform- and file-specific settings, MSVC
if(MSVC)
  target_compile_definitions(xlnt PRIVATE _CRT_SECURE_NO_WARNINGS=1)

  if(NOT STATIC)
    target_compile_definitions(xlnt PRIVATE XLNT_EXPORT=1)
  endif()

  set_target_properties(xlnt PROPERTIES
    COMPILE_FLAGS "/wd\"4251\" /wd\"4275\" /wd\"4068\"")
  set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/detail/serialization/miniz.cpp
    PROPERTIES
    COMPILE_FLAGS "/wd\"4244\" /wd\"4334\" /wd\"4127\"")
  set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/detail/cryptography/aes.cpp
    PROPERTIES
    COMPILE_FLAGS "/wd\"4996\"")
endif()

# Platform- and file-specific settings, Clang
if(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
  set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/detail/serialization/miniz.cpp
    PROPERTIES
    COMPILE_FLAGS "-Wno-undef")
  set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/detail/serialization/zstream.cpp
    PROPERTIES
    COMPILE_FLAGS "-Wno-undef -Wno-shorten-64-to-32")
endif()

# Platform- and file-specific settings, GCC
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
  set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/detail/serialization/miniz.cpp
    PROPERTIES
    COMPILE_FLAGS "-Wno-strict-aliasing")
endif()

# Group files into pseudo-folders in IDEs
source_group(xlnt FILES ${ROOT_HEADERS})
source_group(cell FILES ${CELL_HEADERS} ${CELL_SOURCES})
source_group(detail FILES ${DETAIL_ROOT_HEADERS} ${DETAIL_ROOT_SOURCES})
source_group(detail\\cryptography FILES ${DETAIL_CRYPTOGRAPHY_HEADERS} ${DETAIL_CRYPTOGRAPHY_SOURCES})
source_group(detail\\external FILES ${DETAIL_EXTERNAL_HEADERS})
source_group(detail\\header_footer FILES ${DETAIL_HEADER_FOOTER_HEADERS} ${DETAIL_HEADER_FOOTER_SOURCES})
source_group(detail\\implementations FILES ${DETAIL_IMPLEMENTATIONS_HEADERS} ${DETAIL_IMPLEMENTATIONS_SOURCES})
source_group(detail\\number_format FILES ${DETAIL_NUMBER_FORMAT_HEADERS} ${DETAIL_NUMBER_FORMAT_SOURCES})
source_group(detail\\serialization FILES ${DETAIL_SERIALIZATION_HEADERS} ${DETAIL_SERIALIZATION_SOURCES})
source_group(packaging FILES ${PACKAGING_HEADERS} ${PACKAGING_SOURCES})
source_group(styles FILES ${STYLES_HEADERS} ${STYLES_SOURCES})
source_group(utils FILES ${UTILS_HEADERS} ${UTILS_SOURCES})
source_group(workbook FILES ${WORKBOOK_HEADERS} ${WORKBOOK_SOURCES})
source_group(worksheet FILES ${WORKSHEET_HEADERS} ${WORKSHEET_SOURCES})

if(NOT INC_DEST_DIR)
  # Default location to install include files
  set(INC_DEST_DIR ${CMAKE_INSTALL_PREFIX}/include)
endif()

if(NOT LIB_DEST_DIR)
  # Default location to install library files
  set(LIB_DEST_DIR ${CMAKE_INSTALL_PREFIX}/lib)
endif()

if(NOT BIN_DEST_DIR)
  # Default location to install runtime files
   set(BIN_DEST_DIR ${CMAKE_INSTALL_PREFIX}/bin)
endif()

if(NOT MAN_DEST_DIR)
  # Default location to install runtime files
   set(MAN_DEST_DIR ${CMAKE_INSTALL_PREFIX}/share/man)
endif()

# Install library
install(TARGETS xlnt
  LIBRARY DESTINATION ${LIB_DEST_DIR}
  ARCHIVE DESTINATION ${LIB_DEST_DIR}
  RUNTIME DESTINATION ${BIN_DEST_DIR})

# Install include directory
install(DIRECTORY ${XLNT_INCLUDE_DIR}/xlnt
  DESTINATION include
  PATTERN ".DS_Store" EXCLUDE)

# Install LICENSE.md
install(FILES ${XLNT_ROOT_DIR}/docs/xlnt.3
DESTINATION ${MAN_DEST_DIR}/man3)

# Configure uninstall
configure_file("${XLNT_ROOT_DIR}/cmake/cmake_uninstall.cmake.in"
  "${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake"
  IMMEDIATE @ONLY)

# Add uninstall target
add_custom_target(uninstall
  COMMAND ${CMAKE_COMMAND} -P
    ${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake)

if(NOT MSVC)
  # Set pkg-config variables
  set(PKG_CONFIG_LIBDIR ${LIB_DEST_DIR})
  set(PKG_CONFIG_INCLUDEDIR ${INC_DEST_DIR})
  set(PKG_CONFIG_LIBS "-L\${libdir} -lxlnt")
  set(PKG_CONFIG_CFLAGS "-I\${includedir}")

  # Replace variables in pkg-config template
  configure_file("${XLNT_ROOT_DIR}/cmake/pkg-config.pc.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/xlnt.pc")

  # pkg-config install
  install(FILES "${CMAKE_CURRENT_BINARY_DIR}/xlnt.pc"
    DESTINATION ${LIB_DEST_DIR}/pkgconfig)
endif()
