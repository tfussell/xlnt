solution "xlnt"
    configurations { "Debug", "Release" }
    platforms { "x64" }
    location ("./" .. _ACTION)
    configuration "Debug"
        flags { "Symbols" }
	optimize "Off"
    configuration "Release"
        optimize "Full"

project "xlnt.test"
    kind "ConsoleApp"
    language "C++"
    targetname "xlnt.test"
    targetdir "../bin"
    includedirs { 
       "../include",
       "../third-party/pugixml/src",
       "../third-party/miniz",
       "../third-party/cxxtest"
    }
    files { 
       "../tests/*.hpp",
       "../tests/runner-autogen.cpp"
    }
    links { "xlnt" }
    prebuildcommands { "../../third-party/cxxtest/bin/cxxtestgen --runner=ErrorPrinter -o ../../tests/runner-autogen.cpp ../../tests/*.hpp" }
    flags { "Unicode" }
    configuration "Release"
        flags { "LinkTimeOptimization" }
    configuration "windows"
        defines { "WIN32" }
	links { "Shlwapi" }

project "xlnt"
    kind "StaticLib"
    language "C++"
    warnings "Extra"
    targetdir "../lib/"
    includedirs { 
       "../include/xlnt",
       "../third-party/miniz",
       "../third-party/pugixml/src"
    }
    files {
       "../source/**.cpp",
       "../source/**.hpp",
       "../include/xlnt/**.hpp",
       "../third-party/miniz/miniz.c",
       "../third-party/pugixml/src/pugixml.cpp"
    }
    flags { "Unicode" }
    configuration "Debug"
        flags { "FatalWarnings" }
    configuration "windows"
        defines { 
	   "WIN32",
	   "_CRT_SECURE_NO_WARNINGS"
	}
