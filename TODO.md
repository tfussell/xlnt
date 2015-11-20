For release:
Move all method definitions out of headers.
Clean up iterators in workbook, worksheet, cell_vector, range.
One class per header.
Set up Coveralls.
Synchronize tests for workbook, worksheet, cell, core_properties, manifest, relationship, stylesheet, number_format, format_rule, tokenizer.
Synchronize samples.
Summarize what works and what doesn't.
Finish documenting all header files.
Port benchmarks.
Clean up CMake scripts.
Release 0.9.0 for Windows (static 32-bit, dll 32-bit, static 64-bit, dll 32-bit), Ubuntu 14.04 64-bit (static, shared), OSX 10.11 (static, dylib, framework).
XML wrappers aren't very pretty, improve them?
100% test coverage.

Then:
Implement conditional formatting.
Implement charts.
Implement chartsheets.
Implement drawing.
Implement formula evaluation?
Implement pivot tables.
Implement XML validation.