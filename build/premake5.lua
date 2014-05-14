solution "xlnt"
    configurations { "Debug", "Release" }
    platforms { "x64" }
    location ("./" .. _ACTION)
    configuration "not windows"
        buildoptions { 
            "-std=c++11",
            "-Wno-unknown-pragmas"
        }
    configuration "Debug"
        flags { "Symbols" }
	optimize "Off"
    configuration "Release"
        optimize "Full"

project "xlnt.test"
    kind "ConsoleApp"
    language "C++"
    targetname "xlnt.test"
    includedirs { 
       "$(opc_prefix)",
       "$(opc_prefix)/config",
       "$(opc_prefix)/third_party/libxml2-2.7.7/include",
       "$(cxxtest_prefix)",
       "../include",
       "../third-party/pugixml/src"
    }
    defines { "WIN32" }
    files { 
       "../source/tests/**.h",
       "../source/tests/runner-autogen.cpp"
    }
    links { 
        "xlnt",
	"mce",
	"opc",
	"plib",
	"xml",
	"zlib"
    }
    prebuildcommands { "cxxtestgen --runner=ErrorPrinter -o ../../source/tests/runner-autogen.cpp ../../source/tests/*.h" }
    flags { 
       "Unicode",
       "NoEditAndContinue",
       "NoManifest",
       "NoPCH"
    }
    configuration "Debug"
	targetdir "../bin/debug"
    configuration "Release"
        flags { "LinkTimeOptimization" }
	targetdir "../bin/release"
    configuration "windows"
        includedirs { "$(opc_prefix)/plib/config/msvc/plib/include" }
        defines { "WIN32" }
	links { "Shlwapi" }
        libdirs { "$(opc_prefix)/win32/x64/Debug" }
    configuration "not windows"
        includedirs { "$(opc_prefix)/build/plib/config/darwin-debug-gcc-i386/plib/include" }
        libdirs { "$(opc_prefix)/build/darwin-debug-gcc-i386/static" }

project "xlnt"
    kind "StaticLib"
    language "C++"
    warnings "Extra"
    targetdir "../lib/"
    includedirs { 
       "$(opc_prefix)",
       "$(opc_prefix)/config",
       "$(opc_prefix)/third_party/libxml2-2.7.7/include",
       "../include/xlnt",
       "../third-party/pugixml/src",
       "../source/",
       "../third-party/zlib/",
       "../third-party/zlib/contrib/minizip"
    }
    files {
       "../source/**.cpp",
       "../source/**.h",
       "../include/**.h",
       "../third-party/pugixml/src/pugixml.cpp",
       "../third-party/zlib/*.c",
       "../third-party/zlib/contrib/minizip/*.c"
    }
    excludes {
       "../source/tests/**.cpp",
       "../source/tests/**.h",
       "../third-party/zlib/contrib/minizip/miniunz.c",
       "../third-party/zlib/contrib/minizip/minizip.c"
    }
    flags { 
       "Unicode",
       "NoEditAndContinue",
       "NoManifest",
       "NoPCH"
    }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        includedirs { "$(opc_prefix)/plib/config/msvc/plib/include" }
        defines { "WIN32" }
    configuration "not windows"
        includedirs { "$(opc_prefix)/build/plib/config/darwin-debug-gcc-i386/plib/include" }
    configuration "**.c"
        flags { "NoWarnings" }
