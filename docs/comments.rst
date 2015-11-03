Comments
========

Adding a comment to a cell
--------------------------

Comments have a text attribute and an author attribute, which must both be set.

.. code-block:: cpp

    xlnt::workbook workbook;
    auto worksheet = workbook.get_active_sheet();
    auto comment = worksheet.get_cell("A1").get_comment();
    comment = xlnt::comment("This is the comment text", "Comment Author");
    std::cout << comment.get_text() << std::endl;
    std::cout << comment.get_author() << std::endl;

You cannot assign the same Comment object to two different cells. Doing so
raises an xlnt::attribute_error.

.. code-block:: cpp

    xlnt::workbook workbook;
    auto worksheet = workbook.get_active_sheet();
    xlnt::comment comment("Text", "Author");
    worksheet.get_cell("A1").set_comment(comment);
    worksheet.get_cell("B1").set_comment(comment);

    // prints: terminate called after throwing an instance of 'xlnt::attribute_error'

Loading and saving comments
----------------------------

Comments present in a workbook when loaded are stored in the comment
attribute of their respective cells automatically. Comments remaining in a workbook when it is saved are automatically saved to
the workbook file.
