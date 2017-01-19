# Properties

```c++
xlnt::workbook wb;

wb.core_property(xlnt::core_property::category, "hors categorie");
wb.core_property(xlnt::core_property::content_status, "good");
wb.core_property(xlnt::core_property::created, xlnt::datetime(2017, 1, 15));
wb.core_property(xlnt::core_property::creator, "me");
wb.core_property(xlnt::core_property::description, "description");
wb.core_property(xlnt::core_property::identifier, "id");
wb.core_property(xlnt::core_property::keywords, { "wow", "such" });
wb.core_property(xlnt::core_property::language, "Esperanto");
wb.core_property(xlnt::core_property::last_modified_by, "someone");
wb.core_property(xlnt::core_property::last_printed, xlnt::datetime(2017, 1, 15));
wb.core_property(xlnt::core_property::modified, xlnt::datetime(2017, 1, 15));
wb.core_property(xlnt::core_property::revision, "3");
wb.core_property(xlnt::core_property::subject, "subject");
wb.core_property(xlnt::core_property::title, "title");
wb.core_property(xlnt::core_property::version, "1.0");

wb.extended_property(xlnt::extended_property::application, "xlnt");
wb.extended_property(xlnt::extended_property::app_version, "0.9.3");
wb.extended_property(xlnt::extended_property::characters, 123);
wb.extended_property(xlnt::extended_property::characters_with_spaces, 124);
wb.extended_property(xlnt::extended_property::company, "Incorporated Inc.");
wb.extended_property(xlnt::extended_property::dig_sig, "?");
wb.extended_property(xlnt::extended_property::doc_security, 0);
wb.extended_property(xlnt::extended_property::heading_pairs, true);
wb.extended_property(xlnt::extended_property::hidden_slides, false);
wb.extended_property(xlnt::extended_property::h_links, 0);
wb.extended_property(xlnt::extended_property::hyperlink_base, 0);
wb.extended_property(xlnt::extended_property::hyperlinks_changed, true);
wb.extended_property(xlnt::extended_property::lines, 42);
wb.extended_property(xlnt::extended_property::links_up_to_date, false);
wb.extended_property(xlnt::extended_property::manager, "johnny");
wb.extended_property(xlnt::extended_property::m_m_clips, "?");
wb.extended_property(xlnt::extended_property::notes, "note");
wb.extended_property(xlnt::extended_property::pages, 19);
wb.extended_property(xlnt::extended_property::paragraphs, 18);
wb.extended_property(xlnt::extended_property::presentation_format, "format");
wb.extended_property(xlnt::extended_property::scale_crop, true);
wb.extended_property(xlnt::extended_property::shared_doc, false);
wb.extended_property(xlnt::extended_property::slides, 17);
wb.extended_property(xlnt::extended_property::template_, "template!");
wb.extended_property(xlnt::extended_property::titles_of_parts, { "title" });
wb.extended_property(xlnt::extended_property::total_time, 16);
wb.extended_property(xlnt::extended_property::words, 101);

wb.custom_property("test", { 1, 2, 3 });
wb.custom_property("Editor", "John Smith");

wb.save("lots_of_properties.xlsx");
```