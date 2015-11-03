Simple usage
============

Write a workbook
----------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;

	std::string dest_filename = "empty_book.xlsx";

	auto ws1 = wb.get_active_sheet();
	ws1.set_title("range names");

	for(xlnt::row_t row = 1; row < 40; row++)
	{
	    std::vector<int> to_append(600, 0);
	    std::iota(std::begin(to_append), std::end(to_append), 0);
	    ws1.append(to_append);
	}

	auto ws2 = wb.create_sheet("Pi");

	ws2.get_cell("F5").set_value(3.14);

	for(xlnt::row_t row = 10; row < 20; row++)
	{
	    for(xlnt::column_t column = 27; column < 54; column++)
	    {
	        ws3.get_cell(column, row).set_value(column.column_string());
	    }
	}

	std::cout << ws3.get_cell("AA10") << std::endl;

	wb.save(dest_filename);

	return 0;
    }

Read an existing workbook
-------------------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;
	wb.load("empty_book.xlsx");
	auto sheet_ranges = wb.get_range("range names");

	std::cout << sheet_ranges["D18"] << std::endl;
	// prints: 3

        return 0;
    }

.. note ::

    There are several optional parameters that can be used in xlnt::workbook::load (in order):

    - `guess_types` will enable or disable (default) type inference when
      reading cells.

    - `data_only` controls whether cells with formulae have either the
      formula (default) or the value stored the last time Excel read the sheet.

    - `keep_vba` controls whether any Visual Basic elements are preserved or
      not (default). If they are preserved they are still not editable.


.. warning ::

    xlnt does currently not read all possible items in an Excel file so
    images and charts will be lost from existing files if they are opened and
    saved with the same name.


Using number formats
--------------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;
	wb.guess_types(true);

	auto ws = wb.get_active_sheet();
	ws.get_cell("A1").set_value(xlnt::datetime(2010, 7, 21));
        std::cout << ws.get_cell("A1").get_number_format().get_format_string() << std::endl
        // prints: yyyy-mm-dd h:mm:ss

        // set percentage using a string followed by the percent sign
        ws.get_cell("B1").set_value("3.14%");
	std::cout << cell.get_value<long double>() << std::endl;
	// prints: 0.031400000000000004
	std::cout << cell << std::endl;
	// prints: 3.14%
	std::cout << cell.get_number_format().get_format_string() << std::endl;
	// prints: 0%

	return 0;
    }

Using formulae
--------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;
	auto ws = wb.get_active_sheet();
	ws.get_cell("A1").set_formula("=SUM(1, 1)");
	wb.save("formula.xlsx");
    }

.. warning::
    NB you must use the English name for a function and function arguments *must* be separated by commas and not other punctuation such as semi-colons.

xlsx never evaluates formula but it is possible to check the name of a formula:

.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        bool found = xlnt::formulae::exists("HEX2DEC");
	std::cout << (found ? "True" : "False") << std::endl;
	// prints: True

	return 0;
    }

If you're trying to use a formula that isn't known this could be because you're using a formula that was not included in the initial specification. Such formulae must be prefixed with `xlfn.` to work.

Merge / Unmerge cells
---------------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;
	auto ws = wb.get_active_sheet();

        ws.merge_cells("A1:B1");
        ws.unmerge_cells("A1:B1");

	// or

	ws.merge_cells(1, 2, 4, 2)
	ws.unmerge_cells(1, 2, 4, 2);

	return 0;
    }


Inserting an image
-------------------
.. code-block:: cpp

    #include <xlnt/xlnt.hpp>

    int main()
    {
        xlnt::workbook wb;
	auto ws = wb.get_active_sheet();
	ws.get_cell("A1").set_value("You should see three logos below");

	// create an image
	auto img = xlnt::image("logo.png");

	// add to worksheet and anchor next to cells
	ws.add_image(img, "A1");
	wb.save("logo.xlsx");

        return 0;
    }

Fold columns (outline)
----------------------
.. code-block:: cpp

    int main()
    {
        xlnt::workbook wb;
	auto ws = wb.create_sheet();
	bool hidden = true;
	ws.group_columns("A", "D", hidden);
	wb.save("group.xlsx");

        return 0;
    }
