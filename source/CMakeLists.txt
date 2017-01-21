cmake_minimum_required(VERSION 3.1)
project(xlnt VERSION 0.9.3)

set(PROJECT_VENDOR "Thomas Fussell")
set(PROJECT_CONTACT "thomas.fussell@gmail.com")
set(PROJECT_URL "https://github.com/tfussell/xlnt")
set(PROJECT_DESCRIPTION "cross-platform user-friendly xlsx library for C++14")

if(APPLE)
    option(FRAMEWORK "Set to ON to package dylib and headers into a .framework, OSX only" OFF)
endif()

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(XLNT_ROOT_DIR ${CMAKE_CURRENT_SOURCE_DIR}/..)

set(XLNT_INCLUDE_DIR ${XLNT_ROOT_DIR}/include)
set(XLNT_SOURCE_DIR ${XLNT_ROOT_DIR}/source)
set(THIRD_PARTY_DIR ${XLNT_ROOT_DIR}/third-party)
add_subdirectory(${THIRD_PARTY_DIR} ${CMAKE_CURRENT_BINARY_DIR}/third-party)

if(COVERAGE)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

if(MSVC)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /W4")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Wno-unknown-pragmas")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Weverything -Wno-c++98-compat -Wno-c++98-compat-pedantic -Wno-padded -Werror")
endif()

if(APPLE)
    # Prevent a warning about deployment target not being set
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
file(GLOB DETAIL_HEADERS ${XLNT_SOURCE_DIR}/detail/*.hpp)
file(GLOB DETAIL_SOURCES ${XLNT_SOURCE_DIR}/detail/*.cpp)

set(XLNT_HEADERS ${ROOT_HEADERS} ${CELL_HEADERS} ${CHARTS_HEADERS} 
    ${CHARTSHEET_HEADERS} ${DRAWING_HEADERS} ${FORMULA_HEADERS}
	${PACKAGING_HEADERS} ${STYLES_HEADERS} ${UTILS_HEADERS} 
	${WORKBOOK_HEADERS} ${WORKSHEET_HEADERS} ${DETAIL_HEADERS})
set(XLNT_SOURCES ${CELL_SOURCES} ${CHARTS_SOURCES} ${CHARTSHEET_SOURCES} 
    ${DRAWING_SOURCES} ${FORMULA_SOURCES} ${PACKAGING_SOURCES} 
	${STYLES_SOURCES} ${UTILS_SOURCES} ${WORKBOOK_SOURCES} 
	${WORKSHEET_SOURCES} ${DETAIL_SOURCES})

if(CMAKE_INSTALL_PREFIX_INITIALIZED_TO_DEFAULT)
    set(CMAKE_INSTALL_PREFIX "${CMAKE_BINARY_DIR}/installed" 
        CACHE PATH "default install path" FORCE)
endif()

if(NOT INC_DEST_DIR)
    set(INC_DEST_DIR ${CMAKE_INSTALL_PREFIX}/include)
endif()

if(NOT LIB_DEST_DIR)
    set(LIB_DEST_DIR ${CMAKE_INSTALL_PREFIX}/lib)
endif()

if(NOT BIN_DEST_DIR)
    set(BIN_DEST_DIR ${CMAKE_INSTALL_PREFIX}/bin)
endif()

if(NOT STATIC)
    add_library(xlnt SHARED ${XLNT_HEADERS} ${XLNT_SOURCES})
    target_compile_definitions(xlnt PRIVATE XLNT_SHARED=1)

    if(MSVC)
        target_compile_definitions(xlnt PRIVATE XLNT_EXPORT=1 _CRT_SECURE_NO_WARNINGS=1)
    endif()

    set_target_properties(xlnt
        PROPERTIES
        VERSION ${PROJECT_VERSION}
        SOVERSION ${PROJECT_VERSION_MAJOR}.${PROJECT_VERSION_MINOR}
        INSTALL_NAME_DIR "${LIB_DEST_DIR}")

    if(FRAMEWORK)
        add_custom_command(TARGET xlnt
            POST_BUILD
            COMMAND mkdir -p "${CMAKE_CURRENT_BINARY_DIR}/xlnt.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp -R ${XLNT_INCLUDE_DIR}/xlnt/* "${CMAKE_CURRENT_BINARY_DIR}/xlnt.framework/Versions/${PROJECT_VERSION_FULL}/Headers"
            COMMAND cp "lib/lib${PROJECT_NAME}.${PROJECT_VERSION_FULL}.dylib" "${CMAKE_CURRENT_BINARY_DIR}/xlnt.framework/Versions/${PROJECT_VERSION_FULL}/xlnt"
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/xlnt.framework/Versions" && ln -s "${PROJECT_VERSION_FULL}" Current
            COMMAND cd "${CMAKE_CURRENT_BINARY_DIR}/xlnt.framework" && ln -s Versions/Current/* ./
        )
    endif()
else()
    add_library(xlnt STATIC ${XLNT_HEADERS} ${XLNT_SOURCES})
    target_compile_definitions(xlnt PUBLIC XLNT_STATIC=1)

    if(MSVC)
        target_compile_definitions(xlnt PRIVATE _CRT_SECURE_NO_WARNINGS=1)
    endif()
endif()

target_link_libraries(xlnt ${CRYPTOPP_LIBRARY} ${ZLIB_LIBRARY_RELEASE} libstudxml)

target_include_directories(xlnt
    PUBLIC ${XLNT_INCLUDE_DIR}
    PRIVATE ${XLNT_SOURCE_DIR}
    PRIVATE ${ZLIB_INCLUDE_DIR}
    PRIVATE ${CRYPTOPP_INCLUDE_DIR})

if(MSVC)
    set_target_properties(xlnt PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\" /wd\"4068\" /MP")
endif()

source_group(xlnt FILES ${ROOT_HEADERS})
source_group(detail FILES ${DETAIL_HEADERS} ${DETAIL_SOURCES})
source_group(cell FILES ${CELL_HEADERS} ${CELL_SOURCES})
source_group(charts FILES ${CHARTS_HEADERS} ${CHARTS_SOURCES})
source_group(chartsheet FILES ${CHARTSHEET_HEADERS} ${CHARTSHEET_SOURCES})
source_group(drawing FILES ${DRAWING_HEADERS} ${DRAWING_SOURCES})
source_group(formula FILES ${FORMULA_HEADERS} ${FORMULA_SOURCES})
source_group(packaging FILES ${PACKAGING_HEADERS} ${PACKAGING_SOURCES})
source_group(serialization FILES ${SERIALIZATION_HEADERS} ${SERIALIZATION_SOURCES})
source_group(styles FILES ${STYLES_HEADERS} ${STYLES_SOURCES})
source_group(utils FILES ${UTILS_HEADERS} ${UTILS_SOURCES})
source_group(workbook FILES ${WORKBOOK_HEADERS} ${WORKBOOK_SOURCES})
source_group(worksheet FILES ${WORKSHEET_HEADERS} ${WORKSHEET_SOURCES})

set(PKG_CONFIG_LIBDIR ${LIB_DEST_DIR})
set(PKG_CONFIG_INCLUDEDIR ${INC_DEST_DIR})
set(PKG_CONFIG_LIBS "-L\${libdir} -lxlnt")
set(PKG_CONFIG_CFLAGS "-I\${includedir}")

install(TARGETS xlnt
    LIBRARY DESTINATION ${LIB_DEST_DIR}
    ARCHIVE DESTINATION ${LIB_DEST_DIR}
    RUNTIME DESTINATION ${BIN_DEST_DIR})

if(NOT MSVC)
    configure_file("${XLNT_ROOT_DIR}/cmake/pkg-config.pc.cmake"
        "${CMAKE_CURRENT_BINARY_DIR}/xlnt.pc")
    install(FILES "${CMAKE_CURRENT_BINARY_DIR}/xlnt.pc"
        DESTINATION ${LIB_DEST_DIR}/pkgconfig)
endif()

install(DIRECTORY ${XLNT_INCLUDE_DIR}/xlnt
    DESTINATION include
    PATTERN ".DS_Store" EXCLUDE)

configure_file("${XLNT_ROOT_DIR}/cmake/cmake_uninstall.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake"
    IMMEDIATE @ONLY)

add_custom_target(uninstall
    COMMAND ${CMAKE_COMMAND} -P 
        ${CMAKE_CURRENT_BINARY_DIR}/cmake_uninstall.cmake)
