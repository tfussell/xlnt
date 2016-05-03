#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/serialization/style_serializer.hpp>

class test_style_writer : public CxxTest::TestSuite
{
public:
    void test_write_number_formats()
    {
        xlnt::workbook wb;
        wb.get_active_sheet().get_cell("A1").set_number_format(xlnt::number_format("YYYY"));
        xlnt::style_serializer writer(wb);
        xlnt::xml_document observed;
        auto num_fmts_node = observed.add_child("numFmts");
        writer.write_number_formats(num_fmts_node);
        std::string expected =
        "    <numFmts count=\"1\">"
        "    <numFmt formatCode=\"YYYY\" numFmtId=\"164\"></numFmt>"
        "    </numFmts>";
        auto diff = Helper::compare_xml(expected, observed);
        TS_ASSERT(diff);
    }
    /*
    class TestStyleWriter(object):
    
    void setup(self):
    self.workbook = Workbook()
    self.worksheet = self.workbook.create_sheet()
    
    void _test_no_style(self):
    w = StyleWriter(self.workbook)
    assert len(w.wb._cell_styles) == 1  # there is always the empty (defaul) style
    
    void _test_nb_style(self):
    for i in range(1, 6):
    cell = self.worksheet.cell(row=1, column=i)
    cell.font = Font(size=i)
    _ = cell.style_id
    w = StyleWriter(self.workbook)
    assert len(w.wb._cell_styles) == 6  # 5 + the default
    
    cell = self.worksheet.cell('A10')
    cell.border=Border(top=Side(border_style=borders.BORDER_THIN))
    _ = cell.style_id
    w = StyleWriter(self.workbook)
    assert len(w.wb._cell_styles) == 7
    
    
    void _test_default_xfs(self):
    w = StyleWriter(self.workbook)
    fonts = nft = borders = fills = DummyElement()
    w._write_cell_styles()
    xml = tostring(w._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    </cellXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_xfs_number_format(self):
    for idx, nf in enumerate(["0.0%", "0.00%", "0.000%"], 1):
    cell = self.worksheet.cell(row=idx, column=1)
    cell.number_format = nf
    _ = cell.style_id # add to workbook styles
    w = StyleWriter(self.workbook)
    w._write_cell_styles()
    
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="4">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    <xf borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>
    <xf borderId="0" fillId="0" fontId="0" numFmtId="10" xfId="0"/>
    <xf borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>
    </cellXfs>
    </styleSheet>
    """
    
    xml = tostring(w._root)
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_xfs_fonts(self):
    cell = self.worksheet.cell('A1')
    cell.font = Font(size=12, bold=True)
    _ = cell.style_id # update workbook styles
    w = StyleWriter(self.workbook)
    
    w._write_cell_styles()
    xml = tostring(w._root)
    
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="2">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    <xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0"/>
    </cellXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_xfs_fills(self):
    cell = self.worksheet.cell('A1')
    cell.fill = fill=PatternFill(fill_type='solid',
                                 start_color=Color(colors.DARKYELLOW))
    _ = cell.style_id # update workbook styles
    w = StyleWriter(self.workbook)
    w._write_cell_styles()
    
    xml = tostring(w._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="2">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    <xf borderId="0" fillId="2" fontId="0" numFmtId="0" xfId="0"/>
    </cellXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_xfs_borders(self):
    cell = self.worksheet.cell('A1')
    cell.border=Border(top=Side(border_style=borders.BORDER_THIN,
                                color=Color(colors.DARKYELLOW)))
    _ = cell.style_id # update workbook styles
    
    w = StyleWriter(self.workbook)
    w._write_cell_styles()
    
    xml = tostring(w._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="2">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    <xf borderId="1" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    </cellXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_protection(self):
    cell = self.worksheet.cell('A1')
    cell.protection = Protection(locked=True, hidden=True)
    _ = cell.style_id
    
    w = StyleWriter(self.workbook)
    w._write_cell_styles()
    xml = tostring(w._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellXfs count="2">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    <xf applyProtection="1" borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0">
    <protection hidden="1" locked="1"/>
    </xf>
    </cellXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_named_styles(self):
    writer = StyleWriter(self.workbook)
    writer._write_named_styles()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
    </cellStyleXfs>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_style_names(self):
    writer = StyleWriter(self.workbook)
    writer._write_style_names()
    xml = tostring(writer._root)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0" hidden="0"/>
    </cellStyles>
    </styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    
    void _test_simple_styles(datadir):
    wb = Workbook(guess_types=True)
    ws = wb.active
    now = datetime.datetime.now()
    for idx, v in enumerate(['12.34%', now, 'This is a test', '31.31415', None], 1):
    ws.append([v])
    _ = ws.cell(column=1, row=idx).style_id
    
# set explicit formats
    ws['D9'].number_format = numbers.FORMAT_NUMBER_00
    ws['D9'].protection = Protection(locked=True)
    ws['D9'].style_id
    ws['E1'].protection = Protection(hidden=True)
    ws['E1'].style_id
    
    assert len(wb._cell_styles) == 5
    writer = StyleWriter(wb)
    
    datadir.chdir()
    with open('simple-styles.xml') as reference_file:
    expected = reference_file.read()
    xml = writer.write_table()
    diff = compare_xml(xml, expected)
    assert diff is None, diff
    
    
    void _test_empty_workbook():
    wb = Workbook()
    writer = StyleWriter(wb)
    expected = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="0"/>
    <fonts count="1">
    <font>
    <name val="Calibri"/>
    <family val="2"/>
    <color theme="1"/>
    <sz val="11"/>
    <scheme val="minor"/>
    </font>
    </fonts>
    <fills count="2">
    <fill>
    <patternFill />
    </fill>
    <fill>
    <patternFill patternType="gray125"/>
    </fill>
    </fills>
    <borders count="1">
    <border>
    <left/>
    <right/>
    <top/>
    <bottom/>
    <diagonal/>
    </border>
    </borders>
    <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
    </cellXfs>
    <cellStyles count="1">
    <cellStyle builtinId="0" name="Normal" xfId="0" hidden="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultPivotStyle="PivotStyleLight16" defaultTableStyle="TableStyleMedium9"/>
    </styleSheet>
    """
    xml = writer.write_table()
    diff = compare_xml(xml, expected)
    assert diff is None, diff
     */
};
