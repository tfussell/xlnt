#pragma once

#include <iostream>
#include <cxxtest/TestSuite.h>

#include <xlnt/xlnt.h>

class ChartTestSuite : public CxxTest::TestSuite
{
public:
    ChartTestSuite()
    {

    }

    void setUp()
    {
        /*
        xlnt::workbook wb;
        auto ws = wb.get_active();
        ws.set_title("data");

        for(int i = 0; i < 10; i++)
        {
            ws.cell(i, 0) = i;
            auto chart = BarChart();
            chart.title = "TITLE";
            chart.add_serie(Serie(Reference(ws, (0, 0), (10, 0))));
            chart._series[-1].color = Color.GREEN;
            cw = ChartWriter(chart);
            root = Element("test");
        }
        */
    }

    void test_write_title()
    {
        /*
        cw._write_title(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title></test>");
        */
    }

    void test_write_xaxis()
    {
        /*
        cw._write_axis(root, chart.x_axis, "c:catAx");
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx></test>");
        */
    }

    void test_write_yaxis()
    {
        /*
        cw._write_axis(root, chart.y_axis, "c:valAx");
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></test>");
        */
    }

    void test_write_series()
    {
        //cw._write_series(root);
        //TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>\"data\"!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></test>");
    }

    void test_write_legend()
    {
        /*cw._write_legend(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>");*/
    }

    void test_no_write_legend()
    {
        /*xlnt::workbook wb;
        auto ws = wb.get_active();
        ws.set_title("data");
        for(int i = 0; i < 10; i++)
        {
            ws.cell(i, 0) = i;
            ws.cell(i, 1) = i;
            scatterchart = ScatterChart();
            scatterchart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)), ;
            xvalues = Reference(ws, (0, 1), (10, 1))));
            cw = ChartWriter(scatterchart);
            root = Element("test");
            scatterchart.show_legend = False;
            cw._write_legend(root);
            TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test />");
        }*/
    }

    void test_write_print_settings()
    {
        /*cw._write_print_settings(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>");*/
    }

    void test_write_chart()
    {
        //cw._write_chart(root);
        //// Truncate floats because results differ with Python >= 3.2 and <= 3.1
        //test_xml = sub("([0-9][.][0-9]{4})[0-9]*", "\\1", get_xml(root));
        //TS_ASSERT_EQUALS(test_xml, "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:chart><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>TITLE</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner" /><c:xMode val="edge" /><c:yMode val="edge" /><c:x val="1.2857" /><c:y val="0.2125" /><c:w val="0.6" /><c:h val="0.6" /></c:manualLayout></c:layout><c:barChart><c:barDir val="col" /><c:grouping val="clustered" /><c:ser><c:idx val="0" /><c:order val="0" /><c:spPr><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00" /></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none" /></c:marker><c:val><c:numRef><c:f>\"data\"!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:marker val="1" /><c:axId val="60871424" /><c:axId val="60873344" /></c:barChart><c:catAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:axPos val="b" /><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /></c:catAx><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="between" /><c:majorUnit val="2.0" /></c:valAx></c:plotArea><c:legend><c:legendPos val="r" /><c:layout /></c:legend><c:plotVisOnly val="1" /></c:chart></test>");
    }

    void setUp_scatter()
    {
        /*wb = Workbook();
        ws = wb.get_active_sheet();
        ws.title = "data";
        for i in range(10)
        {
            ws.cell(row = i, column = 0).value = i;
            ws.cell(row = i, column = 1).value = i;
            scatterchart = ScatterChart();
            scatterchart.add_serie(Serie(Reference(ws, (0, 0), (10, 0)), ;
            xvalues = Reference(ws, (0, 1), (10, 1))));
            cw = ChartWriter(scatterchart);
            root = Element("test");
        }*/
    }

    void test_write_xaxis_scatter()
    {
        /*scatterchart.x_axis.title = "test x axis title";
        cw._write_axis(root, scatterchart.x_axis, "c:valAx");
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:valAx><c:axId val="60871424" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="b" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>test x axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60873344" /><c:crosses val="autoZero" /><c:auto val="1" /><c:lblAlgn val="ctr" /><c:lblOffset val="100" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>");*/
    }

    void test_write_yaxis_scatter()
    {
        /*scatterchart.y_axis.title = "test y axis title";
        cw._write_axis(root, scatterchart.y_axis, "c:valAx");
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:valAx><c:axId val="60873344" /><c:scaling><c:orientation val="minMax" /><c:max val="10.0" /><c:min val="0.0" /></c:scaling><c:axPos val="l" /><c:majorGridlines /><c:numFmt formatCode="General" sourceLinked="1" /><c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr /></a:pPr><a:r><a:rPr lang="fr-FR" /><a:t>test y axis title</a:t></a:r></a:p></c:rich></c:tx><c:layout /></c:title><c:tickLblPos val="nextTo" /><c:crossAx val="60871424" /><c:crosses val="autoZero" /><c:crossBetween val="midCat" /><c:majorUnit val="2.0" /></c:valAx></test>");*/
    }

    void test_write_series_scatter()
    {
        /*cw._write_series(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:ser><c:idx val="0" /><c:order val="0" /><c:marker><c:symbol val="none" /></c:marker><c:xVal><c:numRef><c:f>\"data\"!$B$1:$B$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>\"data\"!$A$1:$A$11</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="11" /><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt><c:pt idx="4"><c:v>4</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>7</c:v></c:pt><c:pt idx="8"><c:v>8</c:v></c:pt><c:pt idx="9"><c:v>9</c:v></c:pt><c:pt idx="10"><c:v>None</c:v></c:pt></c:numCache></c:numRef></c:yVal></c:ser></test>");*/
    }

    void test_write_legend_scatter()
    {
        /*cw._write_legend(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:legend><c:legendPos val="r" /><c:layout /></c:legend></test>");*/
    }

    void test_write_print_settings_scatter()
    {
        /*cw._write_print_settings(root);
        TS_ASSERT_EQUALS(get_xml(root), "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?><test><c:printSettings><c:headerFooter /><c:pageMargins b="0.75" footer="0.3" header="0.3" l="0.7" r="0.7" t="0.75" /><c:pageSetup /></c:printSettings></test>");*/
    }

    void test_write_chart_scatter()
    {
        //cw._write_chart(root);
        //// Truncate floats because results differ with Python >= 3.2 and <= 3.1
        //test_xml = sub("([0-9][.][0-9]{4})[0-9]*", "\\1", get_xml(root));
        //TS_ASSERT_EQUALS(test_xml, "<?xml version=\"1.0\" encoding=\"" + utf8_xml_str + "\"?>;
    }
};
