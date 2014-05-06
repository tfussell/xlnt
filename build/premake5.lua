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
       "$(opc_prefix)/plib/config/msvc/plib/include",
       "$(opc_prefix)/third_party/libxml2-2.7.7/include",
       "$(cxxtest_prefix)"
    }
    defines { "WIN32" }
    files { 
       "../source/tests/**.h",
       "../source/tests/runner-autogen.cpp"
    }
    links { 
        "xlnt",
	"Shlwapi",
	"mce",
	"opc",
	"plib",
	"xml",
	"zlib"
    }
    prebuildcommands { "cxxtestgen --runner=ParenPrinter -o ../../source/tests/runner-autogen.cpp ../../source/tests/packaging/*.h" }
    libdirs { "$(opc_prefix)/win32/x64/Debug" }
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

project "xlnt"
    kind "StaticLib"
    language "C++"
    warnings "Extra"
    targetdir "../lib/"
    includedirs { 
       "$(opc_prefix)",
       "$(opc_prefix)/config",
       "$(opc_prefix)/plib/config/msvc/plib/include",
       "$(opc_prefix)/third_party/libxml2-2.7.7/include",
       "../include/xlnt"
    }
    defines { "WIN32" }
    files {
       "../source/**.cpp",
       "../source/**.h",
       "../include/**.h"
    }
    excludes {
       "../source/tests/**.cpp",
       "../source/tests/**.h"
    }
    flags { 
       "Unicode",
       "NoEditAndContinue",
       "NoManifest",
       "NoPCH"
    }
    configuration "Debug"
        flags { "FatalWarnings" }
