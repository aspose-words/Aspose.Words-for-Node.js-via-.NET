// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

/// <summary>
/// Apply data labels with custom number format and separator to several data points in a series.
/// </summary>
function applyDataLabels(series, labelsCount, numberFormat, separator) {
  for (let i = 0; i < labelsCount; i++) {
    series.hasDataLabels = true;
    expect(series.dataLabels.at(i).isVisible).toEqual(false);

    series.dataLabels.at(i).showCategoryName = true;
    series.dataLabels.at(i).showSeriesName = true;
    series.dataLabels.at(i).showValue = true;
    series.dataLabels.at(i).showLeaderLines = true;
    series.dataLabels.at(i).showLegendKey = true;
    series.dataLabels.at(i).showPercentage = false;
    series.dataLabels.at(i).isHidden = false;

    expect(series.dataLabels.at(i).showDataLabelsRange).toEqual(false);
    series.dataLabels.at(i).numberFormat.formatCode = numberFormat;
    series.dataLabels.at(i).separator = separator;
    expect(series.dataLabels.at(i).showDataLabelsRange).toEqual(false);
    expect(series.dataLabels.at(i).isVisible).toEqual(true);
    expect(series.dataLabels.at(i).isHidden).toEqual(false);
  }
}

/// <summary>
/// Applies a number of data points to a series.
/// </summary>
function applyDataPoints(series, dataPointsCount, markerSymbol, dataPointSize) {
  for (let i = 0; i < dataPointsCount; i++) {
    let point = series.dataPoints.at(i);
    point.marker.symbol = markerSymbol;
    point.marker.size = dataPointSize;
    expect(point.index).toEqual(i);
  }
}

/// <summary>
/// Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
/// </summary>
function appendChart(builder, chartType, width, height) {
  let chartShape = builder.insertChart(chartType, width, height);
  let chart = chartShape.chart;
  chart.series.clear();
  expect(chart.series.count).toEqual(0); 
  return chart;
}

describe("ExCharts", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('ChartTitle', () => {
    //ExStart:ChartTitle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Chart
    //ExFor:aw.Drawing.Charts.Chart.title
    //ExFor:ChartTitle
    //ExFor:aw.Drawing.Charts.ChartTitle.overlay
    //ExFor:aw.Drawing.Charts.ChartTitle.show
    //ExFor:aw.Drawing.Charts.ChartTitle.text
    //ExFor:aw.Drawing.Charts.ChartTitle.font
    //ExSummary:Shows how to insert a chart and set a title.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a chart shape with a document builder and get its chart.
    let chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Bar, 400, 300);
    let chart = chartShape.chart;

    // Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
    let title = chart.title;
    title.text = "My Chart";
    title.font.size = 15;
    title.font.color = "#008000";

    // Set the "Show" property to "true" to make the title visible. 
    title.show = true;

    // Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
    title.overlay = true;

    doc.save(base.artifactsDir + "Charts.chartTitle.docx");
    //ExEnd:ChartTitle

    doc = new aw.Document(base.artifactsDir + "Charts.chartTitle.docx");
    chartShape = doc.getShape(0, true);

    expect(chartShape.shapeType).toEqual(aw.Drawing.ShapeType.NonPrimitive);
    expect(chartShape.hasChart).toEqual(true);

    title = chartShape.chart.title;

    expect(title.text).toEqual("My Chart");
    expect(title.overlay).toEqual(true);
    expect(title.show).toEqual(true);
  });

  test('DataLabelNumberFormat', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.numberFormat
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.font
    //ExFor:aw.Drawing.Charts.ChartNumberFormat.formatCode
    //ExSummary:Shows how to enable and configure data labels for a chart series.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add a line chart, then clear its demo data series to start with a clean chart,
    // and then set a title.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let chart = shape.chart;
    chart.series.clear();
    chart.title.text = "Monthly sales report";

    // Insert a custom chart series with months as categories for the X-axis,
    // and respective decimal amounts for the Y-axis.
    let series = chart.series.add("Revenue",
       [ "January", "February", "March" ],
       [ 25.611, 21.439, 33.750 ]);

    // Enable data labels, and then apply a custom number format for values displayed in the data labels.
    // This format will treat displayed decimal values as millions of US Dollars.
    series.hasDataLabels = true;
    let dataLabels = series.dataLabels;
    dataLabels.showValue = true;
    dataLabels.numberFormat.formatCode = "\"US$\" #,##0.000\"M\"";
    dataLabels.font.size = 12;

    doc.save(base.artifactsDir + "Charts.DataLabelNumberFormat.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.DataLabelNumberFormat.docx");
    series = doc.getShape(0, true).chart.series.at(0);

    expect(series.hasDataLabels).toEqual(true);
    expect(series.dataLabels.showValue).toEqual(true);
    expect(series.dataLabels.numberFormat.formatCode).toEqual("\"US$\" #,##0.000\"M\"");
  });

  test('DataArraysWrongSize', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let chart = shape.chart;

    let seriesColl = chart.series;
    seriesColl.clear();

    let categories = [ "Cat1", null, "Cat3", "Cat4", "Cat5", null ];
    seriesColl.add("AW Series 1", categories, [ 1, 2, NaN, 4, 5, 6 ]);
    seriesColl.add("AW Series 2", categories, [ 2, 3, NaN, 5, 6, 7 ]);

    expect(()=> seriesColl.add("AW Series 3", categories, [ NaN, 4, 5, NaN, NaN ])).toThrow("Data arrays must be of the same size.");
    expect(() => seriesColl.add("AW Series 4", categories, [ NaN, NaN, NaN, NaN, NaN ])).toThrow("Data arrays must be of the same size.");
  });

  test('EmptyValuesInChartData', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let chart = shape.chart;

    let seriesColl = chart.series;
    seriesColl.clear();

    let categories = [ "Cat1", null, "Cat3", "Cat4", "Cat5", null ];
    seriesColl.add("AW Series 1", categories, [ 1, 2, NaN, 4, 5, 6 ]);
    seriesColl.add("AW Series 2", categories, [ 2, 3, NaN, 5, 6, 7 ]);
    seriesColl.add("AW Series 3", categories, [ NaN, 4, 5, NaN, 7, 8 ]);
    seriesColl.add("AW Series 4", categories, [ NaN, NaN, NaN, NaN, NaN, 9 ]);

    doc.save(base.artifactsDir + "Charts.EmptyValuesInChartData.docx");
  });

  test('AxisProperties', () => {
    //ExStart
    //ExFor:ChartAxis
    //ExFor:aw.Drawing.Charts.ChartAxis.categoryType
    //ExFor:aw.Drawing.Charts.ChartAxis.crosses
    //ExFor:aw.Drawing.Charts.ChartAxis.reverseOrder
    //ExFor:aw.Drawing.Charts.ChartAxis.majorTickMark
    //ExFor:aw.Drawing.Charts.ChartAxis.minorTickMark
    //ExFor:aw.Drawing.Charts.ChartAxis.majorUnit
    //ExFor:aw.Drawing.Charts.ChartAxis.minorUnit
    //ExFor:aw.Drawing.Charts.AxisTickLabels.offset
    //ExFor:aw.Drawing.Charts.AxisTickLabels.position
    //ExFor:aw.Drawing.Charts.AxisTickLabels.isAutoSpacing
    //ExFor:aw.Drawing.Charts.ChartAxis.tickMarkSpacing
    //ExFor:AxisCategoryType
    //ExFor:AxisCrosses
    //ExFor:aw.Drawing.Charts.Chart.axisX
    //ExFor:aw.Drawing.Charts.Chart.axisY
    //ExFor:aw.Drawing.Charts.Chart.axisZ
    //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 500, 300);
    let chart = shape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
    chart.series.add("Aspose Test Series",
      ["Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 640, 320, 280, 120, 150 ]);

    // Chart axes have various options that can change their appearance,
    // such as their direction, major/minor unit ticks, and tick marks.
    let xAxis = chart.axisX;
    xAxis.categoryType = aw.Drawing.Charts.AxisCategoryType.Category;
    xAxis.crosses = aw.Drawing.Charts.AxisCrosses.Minimum;
    xAxis.reverseOrder = false;
    xAxis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Inside;
    xAxis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    xAxis.majorUnit = 10.0;
    xAxis.minorUnit = 15.0;
    xAxis.tickLabels.offset = 50;
    xAxis.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.Low;
    xAxis.tickLabels.isAutoSpacing = false;
    xAxis.tickMarkSpacing = 1;

    let yAxis = chart.axisY;
    yAxis.categoryType = aw.Drawing.Charts.AxisCategoryType.Automatic;
    yAxis.crosses = aw.Drawing.Charts.AxisCrosses.Maximum;
    yAxis.reverseOrder = true;
    yAxis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Inside;
    yAxis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    yAxis.majorUnit = 100.0;
    yAxis.minorUnit = 20.0;
    yAxis.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.NextToAxis;

    // Column charts do not have a Z-axis.
    expect(chart.axisZ).toBe(null);

    doc.save(base.artifactsDir + "Charts.AxisProperties.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.AxisProperties.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisX.categoryType).toEqual(aw.Drawing.Charts.AxisCategoryType.Category);
    expect(chart.axisX.crosses).toEqual(aw.Drawing.Charts.AxisCrosses.Minimum);
    expect(chart.axisX.reverseOrder).toEqual(false);
    expect(chart.axisX.majorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Inside);
    expect(chart.axisX.minorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Cross);
    expect(chart.axisX.majorUnit).toEqual(1.0);
    expect(chart.axisX.minorUnit).toEqual(0.5);
    expect(chart.axisX.tickLabels.offset).toEqual(50);
    expect(chart.axisX.tickLabels.position).toEqual(aw.Drawing.Charts.AxisTickLabelPosition.Low);
    expect(chart.axisX.tickLabels.isAutoSpacing).toEqual(false);
    expect(chart.axisX.tickMarkSpacing).toEqual(1);

    expect(chart.axisY.categoryType).toEqual(aw.Drawing.Charts.AxisCategoryType.Category);
    expect(chart.axisY.crosses).toEqual(aw.Drawing.Charts.AxisCrosses.Maximum);
    expect(chart.axisY.reverseOrder).toEqual(true);
    expect(chart.axisY.majorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Inside);
    expect(chart.axisY.minorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Cross);
    expect(chart.axisY.majorUnit).toEqual(100.0);
    expect(chart.axisY.minorUnit).toEqual(20.0);
    expect(chart.axisY.tickLabels.position).toEqual(aw.Drawing.Charts.AxisTickLabelPosition.NextToAxis);
  });

  test('AxisCollection', () => {
    //ExStart
    //ExFor:ChartAxisCollection
    //ExFor:aw.Drawing.Charts.Chart.axes
    //ExSummary:Shows how to work with axes collection.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 500, 300);
    let chart = shape.chart;

    // Hide the major grid lines on the primary and secondary Y axes.
    for (let axis of chart.axes)
    {
      if (axis.type == aw.Drawing.Charts.ChartAxisType.Value)
        axis.hasMajorGridlines = false;
    }

    doc.save(base.artifactsDir + "Charts.AxisCollection.docx");
    //ExEnd
  });

  test.skip('DateTimeValues - TODO: Failed expect(chart.axisX.baseTimeUnit).toEqual.', () => {
    //ExStart
    //ExFor:AxisBound
    //ExFor:AxisBound.#ctor(Double)
    //ExFor:AxisBound.#ctor(DateTime)
    //ExFor:aw.Drawing.Charts.AxisScaling.minimum
    //ExFor:aw.Drawing.Charts.AxisScaling.maximum
    //ExFor:aw.Drawing.Charts.ChartAxis.scaling
    //ExFor:AxisTickMark
    //ExFor:AxisTickLabelPosition
    //ExFor:AxisTimeUnit
    //ExFor:aw.Drawing.Charts.ChartAxis.baseTimeUnit
    //ExFor:aw.Drawing.Charts.ChartAxis.hasMajorGridlines
    //ExFor:aw.Drawing.Charts.ChartAxis.hasMinorGridlines
    //ExSummary:Shows how to insert chart with date/time values.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let chart = shape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
    chart.series.add("Aspose Test Series",
      [
        Date.parse("2017-11-06"), Date.parse("2017-11-09"), Date.parse("2017-11-15"),
        Date.parse("2017-11-21"), Date.parse("2017-11-25"), Date.parse("2017-11-29")
      ],
      [ 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 ]);


    // Set lower and upper bounds for the X-axis.
    let xAxis = chart.axisX;
    xAxis.scaling.minimum = new aw.Drawing.Charts.AxisBound(Date.parse("2017-11-05"));
    xAxis.scaling.maximum = new aw.Drawing.Charts.AxisBound(Date.parse("2017-12-03"));

    // Set the major units of the X-axis to a week, and the minor units to a day.
    xAxis.baseTimeUnit = aw.Drawing.Charts.AxisTimeUnit.Days;
    xAxis.majorUnit = 7.0;
    xAxis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    xAxis.minorUnit = 1.0;
    xAxis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Outside;
    xAxis.hasMajorGridlines = true;
    xAxis.hasMinorGridlines = true;

    // Define Y-axis properties for decimal values.
    let yAxis = chart.axisY;
    yAxis.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.High;
    yAxis.majorUnit = 100.0;
    yAxis.minorUnit = 50.0;
    yAxis.displayUnit.unit = aw.Drawing.Charts.AxisBuiltInUnit.Hundreds;
    yAxis.scaling.minimum = new aw.Drawing.Charts.AxisBound(100);
    yAxis.scaling.maximum = new aw.Drawing.Charts.AxisBound(700);
    yAxis.hasMajorGridlines = true;
    yAxis.hasMinorGridlines = true;

    doc.save(base.artifactsDir + "Charts.DateTimeValues.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.DateTimeValues.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisX.scaling.minimum).toEqual(new aw.Drawing.Charts.AxisBound(Date.parse("2017-11-05")));
    expect(chart.axisX.scaling.maximum).toEqual(new aw.Drawing.Charts.AxisBound(Date.parse("2017-12-03")));
    expect(chart.axisX.baseTimeUnit).toEqual(aw.Drawing.Charts.AxisTimeUnit.Days);
    expect(chart.axisX.majorUnit).toEqual(7.0);
    expect(chart.axisX.minorUnit).toEqual(1.0);
    expect(chart.axisX.majorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Cross);
    expect(chart.axisX.minorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Outside);
    expect(chart.axisX.hasMajorGridlines).toEqual(true);
    expect(chart.axisX.hasMinorGridlines).toEqual(true);

    expect(chart.axisY.tickLabels.position).toEqual(aw.Drawing.Charts.AxisTickLabelPosition.High);
    expect(chart.axisY.majorUnit).toEqual(100.0);
    expect(chart.axisY.minorUnit).toEqual(50.0);
    expect(chart.axisY.displayUnit.unit).toEqual(aw.Drawing.Charts.AxisBuiltInUnit.Hundreds);
    expect(chart.axisY.scaling.minimum).toEqual(new aw.Drawing.Charts.AxisBound(100));
    expect(chart.axisY.scaling.maximum).toEqual(new aw.Drawing.Charts.AxisBound(700));
    expect(chart.axisY.hasMajorGridlines).toEqual(true);
    expect(chart.axisY.hasMinorGridlines).toEqual(true);
  });

  test('HideChartAxis', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartAxis.hidden
    //ExSummary:Shows how to hide chart axes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let chart = shape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
    chart.series.add("AW Series 1",
      [ "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" ],
      [ 1.2, 0.3, 2.1, 2.9, 4.2 ]);

    // Hide the chart axes to simplify the appearance of the chart. 
    chart.axisX.hidden = true;
    chart.axisY.hidden = true;

    doc.save(base.artifactsDir + "Charts.HideChartAxis.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.HideChartAxis.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisX.hidden).toEqual(true);
    expect(chart.axisY.hidden).toEqual(true);
  });

  test('SetNumberFormatToChartAxis', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartAxis.numberFormat
    //ExFor:ChartNumberFormat
    //ExFor:aw.Drawing.Charts.ChartNumberFormat.formatCode
    //ExFor:aw.Drawing.Charts.ChartNumberFormat.isLinkedToSource
    //ExSummary:Shows how to set formatting for chart values.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 500, 300);
    let chart = shape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Add a custom series to the chart with categories for the X-axis,
    // and large respective numeric values for the Y-axis. 
    chart.series.add("Aspose Test Series",
      [ "Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 1900000, 850000, 2100000, 600000, 1500000 ]);

    // Set the number format of the Y-axis tick labels to not group digits with commas. 
    chart.axisY.numberFormat.formatCode = "#,##0";

    // This flag can override the above value and draw the number format from the source cell.
    expect(chart.axisY.numberFormat.isLinkedToSource).toEqual(false);

    doc.save(base.artifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisY.numberFormat.formatCode).toEqual("#,##0");
  });

  test.each([aw.Drawing.Charts.ChartType.Column,
    aw.Drawing.Charts.ChartType.Line,
    aw.Drawing.Charts.ChartType.Pie,
    aw.Drawing.Charts.ChartType.Bar,
    aw.Drawing.Charts.ChartType.Area])('TestDisplayChartsWithConversion', (chartType) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(chartType, 500, 300);
    let chart = shape.chart;
    chart.series.clear();
            
    chart.series.add("Aspose Test Series",
      [ "Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 1900000, 850000, 2100000, 600000, 1500000 ]);

    doc.save(base.artifactsDir + "Charts.TestDisplayChartsWithConversion.docx");
    doc.save(base.artifactsDir + "Charts.TestDisplayChartsWithConversion.pdf");
  });

  test('Surface3DChart', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Surface3D, 500, 300);
    let chart = shape.chart;
    chart.series.clear();

    chart.series.add("Aspose Test Series 1",
      [ "Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 1900000, 850000, 2100000, 600000, 1500000 ]);

    chart.series.add("Aspose Test Series 2",
      [ "Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 900000, 50000, 1100000, 400000, 2500000 ]);

    chart.series.add("Aspose Test Series 3",
      [ "Word", "PDF", "Excel", "GoogleDocs", "Note" ],
      [ 500000, 820000, 1500000, 400000, 100000 ]);

    doc.save(base.artifactsDir + "Charts.surfaceChart.docx");
    doc.save(base.artifactsDir + "Charts.surfaceChart.pdf");
  });

  test('DataLabelsBubbleChart', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.separator
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showBubbleSize
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showCategoryName
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showSeriesName
    //ExSummary:Shows how to work with data labels of a bubble chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chart = builder.insertChart(aw.Drawing.Charts.ChartType.Bubble, 500, 300).chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Add a custom series with X/Y coordinates and diameter of each of the bubbles. 
    let series = chart.series.add("Aspose Test Series",
      [ 2.9, 3.5, 1.1, 4.0, 4.0 ],
      [ 1.9, 8.5, 2.1, 6.0, 1.5 ],
      [ 9.0, 4.5, 2.5, 8.0, 5.0 ]);

    // Enable data labels, and then modify their appearance.
    series.hasDataLabels = true;
    let dataLabels = series.dataLabels;
    dataLabels.showBubbleSize = true;
    dataLabels.showCategoryName = true;
    dataLabels.showSeriesName = true;
    dataLabels.separator = " & ";

    doc.save(base.artifactsDir + "Charts.DataLabelsBubbleChart.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.DataLabelsBubbleChart.docx");
    dataLabels = doc.getShape(0, true).chart.series.at(0).dataLabels;

    expect(dataLabels.showBubbleSize).toEqual(true);
    expect(dataLabels.showCategoryName).toEqual(true);
    expect(dataLabels.showSeriesName).toEqual(true);
    expect(dataLabels.separator).toEqual(" & ");
  });

  test('DataLabelsPieChart', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.separator
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showLeaderLines
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showLegendKey
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showPercentage
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.showValue
    //ExSummary:Shows how to work with data labels of a pie chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chart = builder.insertChart(aw.Drawing.Charts.ChartType.Pie, 500, 300).chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Insert a custom chart series with a category name for each of the sectors, and their frequency table.
    let series = chart.series.add("Aspose Test Series",
      [ "Word", "PDF", "Excel" ],
      [ 2.7, 3.2, 0.8 ]);

    // Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
    series.hasDataLabels = true;
    let dataLabels = series.dataLabels;
    dataLabels.showLeaderLines = true;
    dataLabels.showLegendKey = true;
    dataLabels.showPercentage = true;
    dataLabels.showValue = true;
    dataLabels.separator = "; ";

    doc.save(base.artifactsDir + "Charts.DataLabelsPieChart.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.DataLabelsPieChart.docx");
    dataLabels = doc.getShape(0, true).chart.series.at(0).dataLabels;

    expect(dataLabels.showLeaderLines).toEqual(true);
    expect(dataLabels.showLegendKey).toEqual(true);
    expect(dataLabels.showPercentage).toEqual(true);
    expect(dataLabels.showValue).toEqual(true);
    expect(dataLabels.separator).toEqual("; ");
  });

  //ExStart
  //ExFor:ChartSeries
  //ExFor:ChartSeries.DataLabels
  //ExFor:ChartSeries.DataPoints
  //ExFor:ChartSeries.Name
  //ExFor:ChartDataLabel
  //ExFor:ChartDataLabel.Index
  //ExFor:ChartDataLabel.IsVisible
  //ExFor:ChartDataLabel.NumberFormat
  //ExFor:ChartDataLabel.Separator
  //ExFor:ChartDataLabel.ShowCategoryName
  //ExFor:ChartDataLabel.ShowDataLabelsRange
  //ExFor:ChartDataLabel.ShowLeaderLines
  //ExFor:ChartDataLabel.ShowLegendKey
  //ExFor:ChartDataLabel.ShowPercentage
  //ExFor:ChartDataLabel.ShowSeriesName
  //ExFor:ChartDataLabel.ShowValue
  //ExFor:ChartDataLabel.IsHidden
  //ExFor:ChartDataLabelCollection
  //ExFor:ChartDataLabelCollection.ClearFormat
  //ExFor:ChartDataLabelCollection.Count
  //ExFor:ChartDataLabelCollection.GetEnumerator
  //ExFor:ChartDataLabelCollection.Item(Int32)
  //ExSummary:Shows how to apply labels to data points in a line chart.
  test('DataLabels', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 400, 300);
    let chart = chartShape.chart;

    expect(chart.series.count).toEqual(3);
    expect(chart.series.at(0).name).toEqual("Series 1");
    expect(chart.series.at(1).name).toEqual("Series 2");
    expect(chart.series.at(2).name).toEqual("Series 3");

    // Apply data labels to every series in the chart.
    // These labels will appear next to each data point in the graph and display its value.
    for (let series of chart.series)
    {
      applyDataLabels(series, 4, "000.0", ", ");
      expect(series.dataLabels.count).toEqual(4);
    }

    // Change the separator string for every data label in a series.
    for (let label of chart.series.at(0).dataLabels)
    {
        expect(label.separator).toEqual(", ");
        label.separator = " & ";
    }

    // For a cleaner looking graph, we can remove data labels individually.
    chart.series.at(1).dataLabels.at(2).clearFormat();

    // We can also strip an entire series of its data labels at once.
    chart.series.at(2).dataLabels.clearFormat();

    doc.save(base.artifactsDir + "Charts.dataLabels.docx");
  });
  //ExEnd

  //ExStart
  //ExFor:ChartSeries.Smooth
  //ExFor:ChartDataPoint
  //ExFor:ChartDataPoint.Index
  //ExFor:ChartDataPointCollection
  //ExFor:ChartDataPointCollection.ClearFormat
  //ExFor:ChartDataPointCollection.Count
  //ExFor:ChartDataPointCollection.GetEnumerator
  //ExFor:ChartDataPointCollection.Item(Int32)
  //ExFor:ChartMarker
  //ExFor:ChartMarker.Size
  //ExFor:ChartMarker.Symbol
  //ExFor:IChartDataPoint
  //ExFor:IChartDataPoint.InvertIfNegative
  //ExFor:IChartDataPoint.Marker
  //ExFor:MarkerSymbol
  //ExSummary:Shows how to work with data points on a line chart.
  test('ChartDataPoint', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 350);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(3);
    expect(chart.series.at(0).name).toEqual("Series 1");
    expect(chart.series.at(1).name).toEqual("Series 2");
    expect(chart.series.at(2).name).toEqual("Series 3");

    // Emphasize the chart's data points by making them appear as diamond shapes.
    for (let series of chart.series)
      applyDataPoints(series, 4, aw.Drawing.Charts.MarkerSymbol.Diamond, 15);

    // Smooth out the line that represents the first data series.
    chart.series.at(0).smooth = true;

    // Verify that data points for the first series will not invert their colors if the value is negative.
    for (let p of chart.series.at(0).dataPoints)
    {
      expect(p.invertIfNegative).toEqual(false);
    }

    // For a cleaner looking graph, we can clear format individually.
    chart.series.at(1).dataPoints.at(2).clearFormat();

    // We can also strip an entire series of data points at once.
    chart.series.at(2).dataPoints.clearFormat();

    doc.save(base.artifactsDir + "Charts.ChartDataPoint.docx");
  });
  //ExEnd

  test('PieChartExplosion', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.IChartDataPoint.explosion
    //ExSummary:Shows how to move the slices of a pie chart away from the center.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Pie, 500, 350);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(1);
    expect(chart.series.at(0).name).toEqual("Sales");

    // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
    // Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
    // Aspose.words create data points automatically if them does not exist.
    let dataPoint = chart.series.at(0).dataPoints.at(0);
    dataPoint.explosion = 10;

    // Displace the second portion by a greater distance.
    dataPoint = chart.series.at(0).dataPoints.at(1);
    dataPoint.explosion = 40;

    doc.save(base.artifactsDir + "Charts.PieChartExplosion.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.PieChartExplosion.docx");
    let series = doc.getShape(0, true).chart.series.at(0);

    expect(series.dataPoints.at(0).explosion).toEqual(10);
    expect(series.dataPoints.at(1).explosion).toEqual(40);
  });

  test('Bubble3D', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataLabel.showBubbleSize
    //ExFor:aw.Drawing.Charts.ChartDataLabel.font
    //ExFor:aw.Drawing.Charts.IChartDataPoint.bubble3D
    //ExSummary:Shows how to use 3D effects with bubble charts.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Bubble3D, 500, 350);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(1);
    expect(chart.series.at(0).name).toEqual("Y-Values");
    expect(chart.series.at(0).bubble3D).toEqual(true);

    // Apply a data label to each bubble that displays its diameter.
    for (let i = 0; i < 3; i++)
    {
      chart.series.at(0).hasDataLabels = true;
      chart.series.at(0).dataLabels.at(i).showBubbleSize = true;
      chart.series.at(0).dataLabels.at(i).font.size = 12;
    }

    doc.save(base.artifactsDir + "Charts.bubble3D.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.bubble3D.docx");
    let series = doc.getShape(0, true).chart.series.at(0);

    for (let i = 0; i < 3; i++)
    {
      expect(series.dataLabels.at(i).showBubbleSize).toEqual(true);
    }
  });


  //ExStart
  //ExFor:ChartAxis.Type
  //ExFor:ChartAxisType
  //ExFor:ChartType
  //ExFor:Chart.Series
  //ExFor:ChartSeriesCollection.Add(String,DateTime[],Double[])
  //ExFor:ChartSeriesCollection.Add(String,Double[],Double[])
  //ExFor:ChartSeriesCollection.Add(String,Double[],Double[],Double[])
  //ExFor:ChartSeriesCollection.Add(String,String[],Double[])
  //ExSummary:Shows how to create an appropriate type of chart series for a graph type.
  test('ChartSeriesCollection', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // There are several ways of populating a chart's series collection.
    // Different series schemas are intended for different chart types.
    // 1 -  Column chart with columns grouped and banded along the X-axis by category:
    let chart = appendChart(builder, aw.Drawing.Charts.ChartType.Column, 500, 300);

    let categories = [ "Category 1", "Category 2", "Category 3" ];

    // Insert two series of decimal values containing a value for each respective category.
    // This column chart will have three groups, each with two columns.
    chart.series.add("Series 1", categories, [ 76.6, 82.1, 91.6 ]);
    chart.series.add("Series 2", categories, [ 64.2, 79.5, 94.0 ]);

    // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
    expect(chart.axisX.type).toEqual(aw.Drawing.Charts.ChartAxisType.Category);
    expect(chart.axisY.type).toEqual(aw.Drawing.Charts.ChartAxisType.Value);

    // 2 -  Area chart with dates distributed along the X-axis:
    chart = appendChart(builder, aw.Drawing.Charts.ChartType.Area, 500, 300);

    let dates = [ Date.parse("2014-03-31"),
      Date.parse("2017-01-23"),
      Date.parse("2017-06-18"),
      Date.parse("2019-11-22"),
      Date.parse("2020-09-07")
    ];

    // Insert a series with a decimal value for each respective date.
    // The dates will be distributed along a linear X-axis,
    // and the values added to this series will create data points.
    chart.series.add("Series 1", dates, [ 15.8, 21.5, 22.9, 28.7, 33.1 ]);

    expect(chart.axisX.type).toEqual(aw.Drawing.Charts.ChartAxisType.Category);
    expect(chart.axisY.type).toEqual(aw.Drawing.Charts.ChartAxisType.Value);

    // 3 -  2D scatter plot:
    chart = appendChart(builder, aw.Drawing.Charts.ChartType.Scatter, 500, 300);

    // Each series will need two decimal arrays of equal length.
    // The first array contains X-values, and the second contains corresponding Y-values
    // of data points on the chart's graph.
    chart.series.add("Series 1",
      [ 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 ],
      [ 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 ]);
    chart.series.add("Series 2",
      [ 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 ],
      [ 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 ]);

    expect(chart.axisX.type).toEqual(aw.Drawing.Charts.ChartAxisType.Value);
    expect(chart.axisY.type).toEqual(aw.Drawing.Charts.ChartAxisType.Value);

    // 4 -  Bubble chart:
    chart = appendChart(builder, aw.Drawing.Charts.ChartType.Bubble, 500, 300);

    // Each series will need three decimal arrays of equal length.
    // The first array contains X-values, the second contains corresponding Y-values,
    // and the third contains diameters for each of the graph's data points.
    chart.series.add("Series 1",
      [ 1.1, 5.0, 9.8 ],
      [ 1.2, 4.9, 9.9 ],
      [ 2.0, 4.0, 8.0 ]);

    doc.save(base.artifactsDir + "Charts.ChartSeriesCollection.docx");
  });
  //ExEnd

  test('ChartSeriesCollectionModify', () => {
    //ExStart
    //ExFor:ChartSeriesCollection
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.clear
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.count
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.getEnumerator
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.item(Int32)
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.removeAt(Int32)
    //ExSummary:Shows how to add and remove series data in a chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a column chart that will contain three series of demo data by default.
    let chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 400, 300);
    let chart = chartShape.chart;

    // Each series has four decimal values: one for each of the four categories.
    // Four clusters of three columns will represent this data.
    let chartData = chart.series;

    expect(chartData.count).toEqual(3);

    // Print the name of every series in the chart.
    for (let s of chart.series) {
      console.log(s.name);
    }

    // These are the names of the categories in the chart.
    let categories = [ "Category 1", "Category 2", "Category 3", "Category 4" ];

    // We can add a series with new values for existing categories.
    // This chart will now contain four clusters of four columns.
    chart.series.add("Series 4", categories, [ 4.4, 7.0, 3.5, 2.1 ]);
    expect(chartData.count).toEqual(4);
    expect(chartData.at(3).name).toEqual("Series 4");

    // A chart series can also be removed by index, like this.
    // This will remove one of the three demo series that came with the chart.
    chartData.removeAt(2);

    for (let s in chartData) {
      if (s.name == "Series 3")
        expect(s).toEqual(false);
    }
    expect(chartData.count).toEqual(3);
    expect(chartData.at(2).name).toEqual("Series 4");

    // We can also clear all the chart's data at once with this method.
    // When creating a new chart, this is the way to wipe all the demo data
    // before we can begin working on a blank chart.
    chartData.clear();
    expect(chartData.count).toEqual(0);

    //ExEnd
  });

  test('AxisScaling', () => {
    //ExStart
    //ExFor:AxisScaleType
    //ExFor:AxisScaling
    //ExFor:aw.Drawing.Charts.AxisScaling.logBase
    //ExFor:aw.Drawing.Charts.AxisScaling.type
    //ExSummary:Shows how to apply logarithmic scaling to a chart axis.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 450, 300);
    let chart = chartShape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Insert a series with X/Y coordinates for five points.
    chart.series.add("Series 1",
      [ 1.0, 2.0, 3.0, 4.0, 5.0 ],
      [ 1.0, 20.0, 400.0, 8000.0, 160000.0 ]);

    // The scaling of the X-axis is linear by default,
    // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
    // A linear axis is not ideal for our Y-values
    // since the points with the smaller Y-values will be harder to read.
    // A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
    // will spread the plotted points, allowing us to read their values on the chart more easily.
    chart.axisY.scaling.type = aw.Drawing.Charts.AxisScaleType.Logarithmic;
    chart.axisY.scaling.logBase = 20;

    doc.save(base.artifactsDir + "Charts.AxisScaling.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.AxisScaling.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisX.scaling.type).toEqual(aw.Drawing.Charts.AxisScaleType.Linear);
    expect(chart.axisY.scaling.type).toEqual(aw.Drawing.Charts.AxisScaleType.Logarithmic);
    expect(chart.axisY.scaling.logBase).toEqual(20.0);
  });

  test('AxisBound', () => {
    //ExStart
    //ExFor:AxisBound.#ctor
    //ExFor:aw.Drawing.Charts.AxisBound.isAuto
    //ExFor:aw.Drawing.Charts.AxisBound.value
    //ExFor:aw.Drawing.Charts.AxisBound.valueAsDate
    //ExSummary:Shows how to set custom axis bounds.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 450, 300);
    let chart = chartShape.chart;

    // Clear the chart's demo data series to start with a clean chart.
    chart.series.clear();

    // Add a series with two decimal arrays. The first array contains the X-values,
    // and the second contains corresponding Y-values for points in the scatter chart.
    chart.series.add("Series 1",
      [ 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 ],
      [ 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 ]);

    // By default, default scaling is applied to the graph's X and Y-axes,
    // so that both their ranges are big enough to encompass every X and Y-value of every series.
    expect(chart.axisX.scaling.minimum.isAuto).toEqual(true);

    // We can define our own axis bounds.
    // In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
    chart.axisX.scaling.minimum = new aw.Drawing.Charts.AxisBound(0);
    chart.axisX.scaling.maximum = new aw.Drawing.Charts.AxisBound(10);
    chart.axisY.scaling.minimum = new aw.Drawing.Charts.AxisBound(0);
    chart.axisY.scaling.maximum = new aw.Drawing.Charts.AxisBound(10);

    expect(chart.axisX.scaling.minimum.isAuto).toEqual(false);
    expect(chart.axisY.scaling.minimum.isAuto).toEqual(false);

    // Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
    chartShape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 450, 300);
    chart = chartShape.chart;
    chart.series.clear();

    let dates = [ Date.parse("1973-05-11"),
      Date.parse("1981-02-04"),
      Date.parse("1985-09-23"),
      Date.parse("1989-06-28"),
      Date.parse("1994-12-15")
    ];

    chart.series.add("Series 1", dates, [ 3.0, 4.7, 5.9, 7.1, 8.9 ]);

    // We can set axis bounds in the form of dates as well, limiting the chart to a period.
    // Setting the range to 1980-1990 will omit the two of the series values
    // that are outside of the range from the graph.
    chart.axisX.scaling.minimum = new aw.Drawing.Charts.AxisBound(Date.parse("1980-01-01"));
    chart.axisX.scaling.maximum = new aw.Drawing.Charts.AxisBound(Date.parse("1990-01-01"));

    doc.save(base.artifactsDir + "Charts.AxisBound.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.AxisBound.docx");
    chart = doc.getShape(0, true).chart;

    expect(chart.axisX.scaling.minimum.isAuto).toEqual(false);
    expect(chart.axisX.scaling.minimum.value).toEqual(0.0);
    expect(chart.axisX.scaling.maximum.value).toEqual(10.0);

    expect(chart.axisY.scaling.minimum.isAuto).toEqual(false);
    expect(chart.axisY.scaling.minimum.value).toEqual(0.0);
    expect(chart.axisY.scaling.maximum.value).toEqual(10.0);

    chart = doc.getShape(1, true).chart;

    expect(chart.axisX.scaling.minimum.isAuto).toEqual(false);
    expect(chart.axisX.scaling.minimum).toEqual(new aw.Drawing.Charts.AxisBound(Date.parse("1980-01-01")));
    expect(chart.axisX.scaling.maximum).toEqual(new aw.Drawing.Charts.AxisBound(Date.parse("1990-01-01")));

    expect(chart.axisY.scaling.minimum.isAuto).toEqual(true);
  });

  test('ChartLegend', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.Chart.legend
    //ExFor:ChartLegend
    //ExFor:aw.Drawing.Charts.ChartLegend.overlay
    //ExFor:aw.Drawing.Charts.ChartLegend.position
    //ExFor:LegendPosition
    //ExSummary:Shows how to edit the appearance of a chart's legend.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 450, 300);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(3);
    expect(chart.series.at(0).name).toEqual("Series 1");
    expect(chart.series.at(1).name).toEqual("Series 2");
    expect(chart.series.at(2).name).toEqual("Series 3");

    // Move the chart's legend to the top right corner.
    let legend = chart.legend;
    legend.position = aw.Drawing.Charts.LegendPosition.TopRight;

    // Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
    legend.overlay = true;

    doc.save(base.artifactsDir + "Charts.ChartLegend.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.ChartLegend.docx");

    legend = doc.getShape(0, true).chart.legend;

    expect(legend.overlay).toEqual(true);
    expect(legend.position).toEqual(aw.Drawing.Charts.LegendPosition.TopRight);
  });

  test('AxisCross', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartAxis.axisBetweenCategories
    //ExFor:aw.Drawing.Charts.ChartAxis.crossesAt
    //ExSummary:Shows how to get a graph axis to cross at a custom location.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 450, 250);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(3);
    expect(chart.series.at(0).name).toEqual("Series 1");
    expect(chart.series.at(1).name).toEqual("Series 2");
    expect(chart.series.at(2).name).toEqual("Series 3");

    // For column charts, the Y-axis crosses at zero by default,
    // which means that columns for all values below zero point down to represent negative values.
    // We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
    let axis = chart.axisX;
    axis.crosses = aw.Drawing.Charts.AxisCrosses.Custom;
    axis.crossesAt = 3;
    axis.axisBetweenCategories = true;

    doc.save(base.artifactsDir + "Charts.AxisCross.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.AxisCross.docx");
    axis = doc.getShape(0, true).chart.axisX;

    expect(axis.axisBetweenCategories).toEqual(true);
    expect(axis.crosses).toEqual(aw.Drawing.Charts.AxisCrosses.Custom);
    expect(axis.crossesAt).toEqual(3.0);
  });


  test('AxisDisplayUnit', () => {
    //ExStart
    //ExFor:AxisBuiltInUnit
    //ExFor:aw.Drawing.Charts.ChartAxis.displayUnit
    //ExFor:aw.Drawing.Charts.ChartAxis.majorUnitIsAuto
    //ExFor:aw.Drawing.Charts.ChartAxis.majorUnitScale
    //ExFor:aw.Drawing.Charts.ChartAxis.minorUnitIsAuto
    //ExFor:aw.Drawing.Charts.ChartAxis.minorUnitScale
    //ExFor:aw.Drawing.Charts.ChartAxis.tickLabelSpacing
    //ExFor:aw.Drawing.Charts.ChartAxis.tickLabelAlignment
    //ExFor:AxisDisplayUnit
    //ExFor:aw.Drawing.Charts.AxisDisplayUnit.customUnit
    //ExFor:aw.Drawing.Charts.AxisDisplayUnit.unit
    //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 450, 250);
    let chart = shape.chart;

    expect(chart.series.count).toEqual(1);
    expect(chart.series.at(0).name).toEqual("Y-Values");

    // Set the minor tick marks of the Y-axis to point away from the plot area,
    // and the major tick marks to cross the axis.
    let axis = chart.axisY;
    axis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    axis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Outside;

    // Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
    axis.majorUnit = 10;
    axis.minorUnit = 1;

    // Set the Y-axis bounds to -10 and 20.
    // This Y-axis will now display 4 major tick marks and 27 minor tick marks.
    axis.scaling.minimum = new aw.Drawing.Charts.AxisBound(-10);
    axis.scaling.maximum = new aw.Drawing.Charts.AxisBound(20);

    // For the X-axis, set the major tick marks at every 10 units,
    // every minor tick mark at 2.5 units.
    axis = chart.axisX;
    axis.majorUnit = 10;
    axis.minorUnit = 2.5;

    // Configure both types of tick marks to appear inside the graph plot area.
    axis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Inside;
    axis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Inside;

    // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
    axis.scaling.minimum = new aw.Drawing.Charts.AxisBound(-10);
    axis.scaling.maximum = new aw.Drawing.Charts.AxisBound(30);
    axis.tickLabels.alignment = aw.ParagraphAlignment.Right;

    expect(axis.tickLabels.spacing).toEqual(1);

    // Set the tick labels to display their value in millions.
    axis.displayUnit.unit = aw.Drawing.Charts.AxisBuiltInUnit.Millions;

    // We can set a more specific value by which tick labels will display their values.
    // This statement is equivalent to the one above.
    axis.displayUnit.customUnit = 1000000;
    expect(axis.displayUnit.unit).toEqual(aw.Drawing.Charts.AxisBuiltInUnit.Custom);

    doc.save(base.artifactsDir + "Charts.AxisDisplayUnit.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Charts.AxisDisplayUnit.docx");
    shape = doc.getShape(0, true);

    expect(shape.width).toEqual(450.0);
    expect(shape.height).toEqual(250.0);

    axis = shape.chart.axisX;

    expect(axis.majorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Inside);
    expect(axis.minorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Inside);
    expect(axis.majorUnit).toEqual(10.0);
    expect(axis.scaling.minimum.value).toEqual(-10.0);
    expect(axis.scaling.maximum.value).toEqual(30.0);
    expect(axis.tickLabels.spacing).toEqual(1);
    expect(axis.tickLabels.alignment).toEqual(aw.ParagraphAlignment.Right);
    expect(axis.displayUnit.unit).toEqual(aw.Drawing.Charts.AxisBuiltInUnit.Custom);
    expect(axis.displayUnit.customUnit).toEqual(1000000.0);

    axis = shape.chart.axisY;

    expect(axis.majorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Cross);
    expect(axis.minorTickMark).toEqual(aw.Drawing.Charts.AxisTickMark.Outside);
    expect(axis.majorUnit).toEqual(10.0);
    expect(axis.minorUnit).toEqual(1.0);
    expect(axis.scaling.minimum.value).toEqual(-10.0);
    expect(axis.scaling.maximum.value).toEqual(20.0);
  });


  test('MarkerFormatting', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartMarker.format
    //ExFor:aw.Drawing.Charts.ChartFormat.fill
    //ExFor:aw.Drawing.Charts.ChartFormat.stroke
    //ExFor:aw.Drawing.Stroke.foreColor
    //ExFor:aw.Drawing.Stroke.backColor
    //ExFor:aw.Drawing.Stroke.visible
    //ExFor:aw.Drawing.Stroke.transparency
    //ExFor:aw.Drawing.Fill.presetTextured(PresetTexture)
    //ExSummary:Show how to set marker formatting.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 432, 252);
    let chart = shape.chart;

    // Delete default generated series.
    chart.series.clear();
    let series = chart.series.add("AW Series 1", [ 0.7, 1.8, 2.6, 3.9 ],
      [ 2.7, 3.2, 0.8, 1.7 ]);

    // Set marker formatting.
    series.marker.size = 40;
    series.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Square;
    let dataPoints = series.dataPoints;
    dataPoints.at(0).marker.format.fill.presetTextured(aw.Drawing.PresetTexture.Denim);
    dataPoints.at(0).marker.format.stroke.foreColor = "#FFFF00";
    dataPoints.at(0).marker.format.stroke.backColor = "#FF0000";
    dataPoints.at(1).marker.format.fill.presetTextured(aw.Drawing.PresetTexture.WaterDroplets);
    dataPoints.at(1).marker.format.stroke.foreColor = "#FFFF00";
    dataPoints.at(1).marker.format.stroke.visible = false;
    dataPoints.at(2).marker.format.fill.presetTextured(aw.Drawing.PresetTexture.GreenMarble);
    dataPoints.at(2).marker.format.stroke.foreColor = "#FFFF00";
    dataPoints.at(3).marker.format.fill.presetTextured(aw.Drawing.PresetTexture.Oak);
    dataPoints.at(3).marker.format.stroke.foreColor = "#FFFF00";
    dataPoints.at(3).marker.format.stroke.transparency = 0.5;

    doc.save(base.artifactsDir + "Charts.MarkerFormatting.docx");
    //ExEnd
  });


  test('SeriesColor', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartSeries.format
    //ExSummary:Sows how to set series color.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    let seriesColl = chart.series;

    // Delete default generated series.
    seriesColl.clear();

    // Create category names array.
    let categories = [ "Category 1", "Category 2" ];

    // Adding new series. Value and category arrays must be the same size.
    let series1 = seriesColl.add("Series 1", categories, [ 1, 2 ]);
    let series2 = seriesColl.add("Series 2", categories, [ 3, 4 ]);
    let series3 = seriesColl.add("Series 3", categories, [ 5, 6 ]);

    // Set series color.
    series1.format.fill.foreColor = "#FF0000";
    series2.format.fill.foreColor = "#FFFF00";
    series3.format.fill.foreColor = "#0000FF";

    doc.save(base.artifactsDir + "Charts.SeriesColor.docx");
    //ExEnd
  });


  test('DataPointsFormatting', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataPoint.format
    //ExSummary:Shows how to set individual formatting for categories of a column chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;

    // Delete default generated series.
    chart.series.clear();

    // Adding new series.
    let series = chart.series.add("Series 1",
      [ "Category 1", "Category 2", "Category 3", "Category 4" ],
      [ 1, 2, 3, 4 ]);

    // Set column formatting.
    let dataPoints = series.dataPoints;
    dataPoints.at(0).format.fill.presetTextured(aw.Drawing.PresetTexture.Denim);
    dataPoints.at(1).format.fill.foreColor = "#FF0000";
    dataPoints.at(2).format.fill.foreColor = "#FFFF00"
    dataPoints.at(3).format.fill.foreColor = "#0000FF";

    doc.save(base.artifactsDir + "Charts.DataPointsFormatting.docx");
    //ExEnd
  });


  test('LegendEntries', () => {
    //ExStart
    //ExFor:ChartLegendEntryCollection
    //ExFor:aw.Drawing.Charts.ChartLegend.legendEntries
    //ExFor:aw.Drawing.Charts.ChartLegendEntry.isHidden
    //ExSummary:Shows how to work with a legend entry for chart series.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    let series = chart.series;
    series.clear();

    let categories = [ "AW Category 1", "AW Category 2" ];

    let series1 = series.add("Series 1", categories, [ 1, 2 ]);
    series.add("Series 2", categories, [ 3, 4 ]);
    series.add("Series 3", categories, [ 5, 6 ]);
    series.add("Series 4", categories, [ 0, 0 ]);

    let legendEntries = chart.legend.legendEntries;
    legendEntries.at(3).isHidden = true;

    doc.save(base.artifactsDir + "Charts.legendEntries.docx");
    //ExEnd
  });


  test('LegendFont', () => {
    //ExStart:LegendFont
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:aw.Drawing.Charts.ChartLegendEntry.font
    //ExFor:aw.Drawing.Charts.ChartLegend.font
    //ExSummary:Shows how to work with a legend font.
    let doc = new aw.Document(base.myDir + "Reporting engine template - Chart series.docx");
    let chart = doc.getShape(0, true).chart;

    let chartLegend = chart.legend;
    // Set default font size all legend entries.
    chartLegend.font.size = 14;
    // Change font for specific legend entry.
    chartLegend.legendEntries.at(1).font.italic = true;
    chartLegend.legendEntries.at(1).font.size = 12;

    doc.save(base.artifactsDir + "Charts.LegendFont.docx");
    //ExEnd:LegendFont
  });


  test('RemoveSpecificChartSeries', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartSeries.seriesType
    //ExFor:ChartSeriesType
    //ExSummary:Shows how to remove specific chart serie.
    let doc = new aw.Document(base.myDir + "Reporting engine template - Chart series.docx");
    let chart = doc.getShape(0, true).chart;

    // Remove all series of the Column type.
    for (var i = chart.series.count - 1; i >= 0; i--)
    {
      if (chart.series.at(i).seriesType == aw.Drawing.Charts.ChartSeriesType.Column)
        chart.series.removeAt(i);
    }

    chart.series.add(
      "Aspose Series",
      [ "Category 1", "Category 2", "Category 3", "Category 4" ],
      [ 5.6, 7.1, 2.9, 8.9 ]);

    doc.save(base.artifactsDir + "Charts.RemoveSpecificChartSeries.docx");
    //ExEnd
  });


  test('PopulateChartWithData', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartXValue.fromDouble(Double)
    //ExFor:aw.Drawing.Charts.ChartYValue.fromDouble(Double)
    //ExFor:aw.Drawing.Charts.ChartSeries.add(ChartXValue, ChartYValue)
    //ExSummary:Shows how to populate chart series with data.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder();

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;
    let series1 = chart.series.at(0);

    // Clear X and Y values of the first series.
    series1.clearValues();

    // Populate the series with data.
    series1.add(aw.Drawing.Charts.ChartXValue.fromDouble(3), aw.Drawing.Charts.ChartYValue.fromDouble(10));
    series1.add(aw.Drawing.Charts.ChartXValue.fromDouble(5), aw.Drawing.Charts.ChartYValue.fromDouble(5));
    series1.add(aw.Drawing.Charts.ChartXValue.fromDouble(7), aw.Drawing.Charts.ChartYValue.fromDouble(11));
    series1.add(aw.Drawing.Charts.ChartXValue.fromDouble(9), aw.Drawing.Charts.ChartYValue.fromDouble(17));

    let series2 = chart.series.at(1);

    // Clear X and Y values of the second series.
    series2.clearValues();

    // Populate the series with data.
    series2.add(aw.Drawing.Charts.ChartXValue.fromDouble(2), aw.Drawing.Charts.ChartYValue.fromDouble(4));
    series2.add(aw.Drawing.Charts.ChartXValue.fromDouble(4), aw.Drawing.Charts.ChartYValue.fromDouble(7));
    series2.add(aw.Drawing.Charts.ChartXValue.fromDouble(6), aw.Drawing.Charts.ChartYValue.fromDouble(14));
    series2.add(aw.Drawing.Charts.ChartXValue.fromDouble(8), aw.Drawing.Charts.ChartYValue.fromDouble(7));

    doc.save(base.artifactsDir + "Charts.PopulateChartWithData.docx");
    //ExEnd
  });


  test('GetChartSeriesData', () => {
    //ExStart
    //ExFor:ChartXValueCollection
    //ExFor:ChartYValueCollection
    //ExSummary:Shows how to get chart series data.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder();

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;
    let series = chart.series.at(0);

    const minValue = -1.7976931348623157E+308;
    const minValueIndex = 0;
    const maxValue = 1.7976931348623157E+308;
    const maxValueIndex = 0;

    for (var i = 0; i < series.yvalues.count; i++)
    {
      // Clear individual format of all data points.
      // Data points and data values are one-to-one in column charts.
      series.dataPoints.at(i).clearFormat();

      // Get Y value.
      let yValue = series.yvalues.at(i).doubleValue;

      if (yValue < minValue)
      {
        minValue = yValue;
        minValueIndex = i;
      }

      if (yValue > maxValue)
      {
        maxValue = yValue;
        maxValueIndex = i;
      }
    }

    // Change colors of the max and min values.
    series.dataPoints.at(minValueIndex).format.fill.foreColor = "#FF0000";
    series.dataPoints.at(maxValueIndex).format.fill.foreColor = "#008000";

    doc.save(base.artifactsDir + "Charts.GetChartSeriesData.docx");
    //ExEnd
  });


  test('ChartDataValues', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartXValue.fromString(String)
    //ExFor:aw.Drawing.Charts.ChartSeries.remove(Int32)
    //ExFor:aw.Drawing.Charts.ChartSeries.add(ChartXValue, ChartYValue)
    //ExSummary:Shows how to add/remove chart data values.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder();

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;
    let department1Series = chart.series.at(0);
    let department2Series = chart.series.at(1);

    // Remove the first value in the both series.
    department1Series.remove(0);
    department2Series.remove(0);

    // Add new values to the both series.
    let newXCategory = aw.Drawing.Charts.ChartXValue.fromString("Q1, 2023");
    department1Series.add(newXCategory, aw.Drawing.Charts.ChartYValue.fromDouble(10.3));
    department2Series.add(newXCategory, aw.Drawing.Charts.ChartYValue.fromDouble(5.7));

    doc.save(base.artifactsDir + "Charts.ChartDataValues.docx");
    //ExEnd
  });


  test('FormatDataLables', () => {
    //ExStart
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.format
    //ExFor:aw.Drawing.Charts.ChartFormat.shapeType
    //ExFor:ChartShapeType
    //ExSummary:Shows how to set fill, stroke and callout formatting for chart data labels.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;

    // Delete default generated series.
    chart.series.clear();

    // Add new series.
    let series = chart.series.add("AW Series 1",
      [ "AW Category 1", "AW Category 2", "AW Category 3", "AW Category 4" ],
      [ 100, 200, 300, 400 ]);

    // Show data labels.
    series.hasDataLabels = true;
    series.dataLabels.showValue = true;

    // Format data labels as callouts.
    let format = series.dataLabels.format;
    format.shapeType = aw.Drawing.Charts.ChartShapeType.WedgeRectCallout;
    format.stroke.color = "#006400";
    format.fill.solid("#008000");
    series.dataLabels.font.color = "#FFFF00";

    // Change fill and stroke of an individual data label.
    let labelFormat = series.dataLabels.at(0).format;
    labelFormat.stroke.color = "#00008B";
    labelFormat.fill.solid("#0000FF");

    doc.save(base.artifactsDir + "Charts.FormatDataLables.docx");
    //ExEnd
  });


  test('ChartAxisTitle', () => {
    //ExStart:ChartAxisTitle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ChartAxisTitle
    //ExFor:aw.Drawing.Charts.ChartAxisTitle.text
    //ExFor:aw.Drawing.Charts.ChartAxisTitle.show
    //ExFor:aw.Drawing.Charts.ChartAxisTitle.overlay
    //ExFor:aw.Drawing.Charts.ChartAxisTitle.font
    //ExSummary:Shows how to set chart axis title.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    let seriesColl = chart.series;
    // Delete default generated series.
    seriesColl.clear();

    seriesColl.add("AW Series 1", [ "AW Category 1", "AW Category 2" ], [ 1, 2 ]);

    let chartAxisXTitle = chart.axisX.title;
    chartAxisXTitle.text = "Categories";
    chartAxisXTitle.show = true;
    let chartAxisYTitle = chart.axisY.title;
    chartAxisYTitle.text = "Values";
    chartAxisYTitle.show = true;
    chartAxisYTitle.overlay = true;
    chartAxisYTitle.font.size = 12;
    chartAxisYTitle.font.color = "#0000FF";

    doc.save(base.artifactsDir + "Charts.chartAxisTitle.docx");
    //ExEnd:ChartAxisTitle
  });


  test.each([ [[ 1, 2, NaN, 4, 5, 6], "" ],
    [[ NaN, 4, 5, NaN, 7, 8], "" ],
    [[ NaN, NaN, NaN, NaN, NaN, 9], "" ],
    [[ NaN, 4, 5, NaN, NaN], "Data arrays must be of the same size." ],
    [[ NaN, NaN, NaN, NaN, NaN], "Data arrays must be of the same size." ] ])
    ('DataArraysWrongSize', (seriesValue, exceptionText) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 500, 300);
    let seriesColl = shape.chart.series;
    seriesColl.clear();

    let categories = [ "Word", null, "Excel", "GoogleDocs", "Note", null ];
    if (exceptionText == "")
      seriesColl.add("AW Series", categories, seriesValue);
    else
      expect(() => seriesColl.add("AW Series", categories, seriesValue)).toThrow(exceptionText);
  });


  test('CopyDataPointFormat', () => {
    //ExStart:CopyDataPointFormat
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:aw.Drawing.Charts.ChartSeries.copyFormatFrom(int)
    //ExFor:aw.Drawing.Charts.ChartDataPointCollection.hasDefaultFormat(int)
    //ExFor:aw.Drawing.Charts.ChartDataPointCollection.copyFormat(int, int)
    //ExSummary:Shows how to copy data point format.
    let doc = new aw.Document(base.myDir + "DataPoint format.docx");

    // Get the chart and series to update format.
    let shape = doc.getShape(0, true);
    let series = shape.chart.series.at(0);
    let dataPoints = series.dataPoints;

    expect(dataPoints.hasDefaultFormat(0)).toEqual(true);
    expect(dataPoints.hasDefaultFormat(1)).toEqual(false);

    // Copy format of the data point with index 1 to the data point with index 2
    // so that the data point 2 looks the same as the data point 1.
    dataPoints.copyFormat(0, 1);

    expect(dataPoints.hasDefaultFormat(0)).toEqual(true);
    expect(dataPoints.hasDefaultFormat(1)).toEqual(true);

    // Copy format of the data point with index 0 to the series defaults so that all data points
    // in the series that have the default format look the same as the data point 0.
    series.copyFormatFrom(1);

    expect(dataPoints.hasDefaultFormat(0)).toEqual(true);
    expect(dataPoints.hasDefaultFormat(1)).toEqual(true);

    doc.save(base.artifactsDir + "Charts.CopyDataPointFormat.docx");
    //ExEnd:CopyDataPointFormat
  });


  test('ResetDataPointFill', () => {
    //ExStart:ResetDataPointFill
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:aw.Drawing.Charts.ChartFormat.isDefined
    //ExFor:aw.Drawing.Charts.ChartFormat.setDefaultFill
    //ExSummary:Shows how to reset the fill to the default value defined in the series.
    let doc = new aw.Document(base.myDir + "DataPoint format.docx");

    let shape = doc.getShape(0, true);
    let series = shape.chart.series.at(0);
    let dataPoint = series.dataPoints.at(1);

    expect(dataPoint.format.isDefined).toEqual(true);

    dataPoint.format.setDefaultFill();

    doc.save(base.artifactsDir + "Charts.ResetDataPointFill.docx");
    //ExEnd:ResetDataPointFill
  });


  test('DataTable', () => {
    //ExStart:DataTable
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:ChartDataTable
    //ExFor:aw.Drawing.Charts.ChartDataTable.show
    //ExSummary:Shows how to show data table with chart series data.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;
            
    let series = chart.series;
    series.clear();
    let xValues = [ 2020, 2021, 2022, 2023 ];
    series.add("Series1", xValues, [ 5, 11, 2, 7 ]);
    series.add("Series2", xValues, [ 6, 5.5, 7, 7.8 ]);
    series.add("Series3", xValues, [ 10, 8, 7, 9 ]);

    let dataTable = chart.dataTable;
    dataTable.show = true;

    dataTable.hasLegendKeys = false;
    dataTable.hasHorizontalBorder = false;
    dataTable.hasVerticalBorder = false;

    dataTable.font.italic = true;
    dataTable.format.stroke.weight = 1;
    dataTable.format.stroke.dashStyle = aw.Drawing.DashStyle.ShortDot;
    dataTable.format.stroke.color = "#00008B";

    doc.save(base.artifactsDir + "Charts.dataTable.docx");
    //ExEnd:DataTable
  });


  test('ChartFormat', () => {
    //ExStart:ChartFormat
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:aw.Drawing.Charts.Chart.format
    //ExFor:aw.Drawing.Charts.ChartTitle.format
    //ExFor:aw.Drawing.Charts.ChartAxisTitle.format
    //ExFor:aw.Drawing.Charts.ChartLegend.format
    //ExSummary:Shows how to use chart formating.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;

    // Delete series generated by default.
    let series = chart.series;
    series.clear();

    let categories = [ "Category 1", "Category 2" ];
    series.add("Series 1", categories, [ 1, 2 ]);
    series.add("Series 2", categories, [ 3, 4 ]);

    // Format chart background.
    chart.format.fill.solid("#2F4F4F");

    // Hide axis tick labels.
    chart.axisX.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.None;
    chart.axisY.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.None;

    // Format chart title.
    chart.title.format.fill.solid("#FFFACD");

    // Format axis title.
    chart.axisX.title.show = true;
    chart.axisX.title.format.fill.solid("#FFFACD");

    // Format legend.
    chart.legend.format.fill.solid("#FFFACD");

    doc.save(base.artifactsDir + "Charts.ChartFormat.docx");
    //ExEnd:ChartFormat

    doc = new aw.Document(base.artifactsDir + "Charts.ChartFormat.docx");

    shape = doc.getShape(0, true);
    chart = shape.chart;

    expect(chart.format.fill.color).toEqual("#2F4F4F");
    expect(chart.title.format.fill.color).toEqual("#FFFACD");
    expect(chart.axisX.title.format.fill.color).toEqual("#FFFACD");
    expect(chart.legend.format.fill.color).toEqual("#FFFACD");
  });


  test('SecondaryAxis', () => {
    //ExStart:SecondaryAxis
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ChartSeriesGroup
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.axisGroup
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.axisX
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.axisY
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.series
    //ExFor:aw.Drawing.Charts.ChartSeriesGroupCollection.add(ChartSeriesType)
    //ExFor:AxisGroup
    //ExSummary:Shows how to work with the secondary axis of chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 450, 250);
    let chart = shape.chart;
    let series = chart.series;

    // Delete default generated series.
    series.clear();

    let categories = [ "Category 1", "Category 2", "Category 3" ];
    series.add("Series 1 of primary series group", categories, [ 2, 3, 4 ]);
    series.add("Series 2 of primary series group", categories, [ 5, 2, 3 ]);

    // Create an additional series group, also of the line type.
    let newSeriesGroup = chart.seriesGroups.add(aw.Drawing.Charts.ChartSeriesType.Line);
    // Specify the use of secondary axes for the new series group.
    newSeriesGroup.axisGroup = aw.Drawing.Charts.AxisGroup.Secondary;
    // Hide the secondary X axis.
    newSeriesGroup.axisX.hidden = true;
    // Define title of the secondary Y axis.
    newSeriesGroup.axisY.title.show = true;
    newSeriesGroup.axisY.title.text = "Secondary Y axis";

    // Add a series to the new series group.
    let series3 =
      newSeriesGroup.series.add("Series of secondary series group", categories, [ 13, 11, 16 ]);
    series3.format.stroke.weight = 3.5;

    doc.save(base.artifactsDir + "Charts.SecondaryAxis.docx");
    //ExEnd:SecondaryAxis
  });


  test('ConfigureGapOverlap', () => {
    //ExStart:ConfigureGapOverlap
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.gapWidth
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.overlap
    //ExSummary:Show how to configure gap width and overlap.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 450, 250);
    let seriesGroup = shape.chart.seriesGroups.at(0);

    // Set column gap width and overlap.
    seriesGroup.gapWidth = 450;
    seriesGroup.overlap = -75;

    doc.save(base.artifactsDir + "Charts.ConfigureGapOverlap.docx");
    //ExEnd:ConfigureGapOverlap
  });


  test('BubbleScale', () => {
    //ExStart:BubbleScale
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:aw.Drawing.Charts.ChartSeriesGroup.bubbleScale
    //ExSummary:Show how to set size of the bubbles.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a bubble 3D chart.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Bubble3D, 450, 250);
    let seriesGroup = shape.chart.seriesGroups.at(0);

    // Set bubble scale to 200%.
    seriesGroup.bubbleScale = 200;

    doc.save(base.artifactsDir + "Charts.bubbleScale.docx");
    //ExEnd:BubbleScale
  });


  test('RemoveSecondaryAxis', () => {
    //ExStart:RemoveSecondaryAxis
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:aw.Drawing.Charts.ChartSeriesGroupCollection.count
    //ExFor:aw.Drawing.Charts.ChartSeriesGroupCollection.item(Int32)
    //ExFor:aw.Drawing.Charts.ChartSeriesGroupCollection.removeAt(Int32)
    //ExSummary:Show how to remove secondary axis.
    let doc = new aw.Document(base.myDir + "Combo chart.docx");

    let shape = doc.getShape(0, true);
    let chart = shape.chart;
    let seriesGroups = chart.seriesGroups;

    // Find secondary axis and remove from the collection.
    for (let i = 0; i < seriesGroups.count; i++)
      if (seriesGroups.at(i).axisGroup == aw.Drawing.Charts.AxisGroup.Secondary)
        seriesGroups.removeAt(i);
    //ExEnd:RemoveSecondaryAxis
  });


  test('TreemapChart', () => {
    //ExStart:TreemapChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.add(String, ChartMultilevelValue.at(], double[))
    //ExFor:ChartMultilevelValue.#ctor(String, String)
    //ExSummary:Shows how to create treemap chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a Treemap chart.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Treemap, 450, 280);
    let chart = shape.chart;
    chart.title.text = "World Population";

    // Delete default generated series.
    chart.series.clear();

    // Add a series.
    let series = chart.series.add(
      "Population by Region",
      [
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "China"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "India"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Indonesia"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Pakistan"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Bangladesh"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Japan"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Philippines"),
        new aw.Drawing.Charts.ChartMultilevelValue("Asia", "Other"),
        new aw.Drawing.Charts.ChartMultilevelValue("Africa", "Nigeria"),
        new aw.Drawing.Charts.ChartMultilevelValue("Africa", "Ethiopia"),
        new aw.Drawing.Charts.ChartMultilevelValue("Africa", "Egypt"),
        new aw.Drawing.Charts.ChartMultilevelValue("Africa", "Other"),
        new aw.Drawing.Charts.ChartMultilevelValue("Europe", "Russia"),
        new aw.Drawing.Charts.ChartMultilevelValue("Europe", "Germany"),
        new aw.Drawing.Charts.ChartMultilevelValue("Europe", "Other"),
        new aw.Drawing.Charts.ChartMultilevelValue("Latin America", "Brazil"),
        new aw.Drawing.Charts.ChartMultilevelValue("Latin America", "Mexico"),
        new aw.Drawing.Charts.ChartMultilevelValue("Latin America", "Other"),
        new aw.Drawing.Charts.ChartMultilevelValue("Northern America", "United States"),
        new aw.Drawing.Charts.ChartMultilevelValue("Northern America", "Other"),
        new aw.Drawing.Charts.ChartMultilevelValue("Oceania")
      ],
      [
        1409670000, 1400744000, 279118866, 241499431, 169828911, 123930000, 112892781, 764000000,
        223800000, 107334000, 105914499, 903000000,
        146150789, 84607016, 516000000,
        203080756, 129713690, 310000000,
        335893238, 35000000,
        42000000
      ]);

    // Show data labels.
    series.hasDataLabels = true;
    series.dataLabels.showValue = true;
    series.dataLabels.showCategoryName = true;
    let thousandSeparator = ".";
    series.dataLabels.numberFormat.formatCode = `#${thousandSeparator}0`;

    doc.save(base.artifactsDir + "Charts.treemap.docx");
    //ExEnd:TreemapChart
  });


  test('SunburstChart', () => {
    //ExStart:SunburstChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.add(String, ChartMultilevelValue.at(], double[))
    //ExSummary:Shows how to create sunburst chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a Sunburst chart.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Sunburst, 450, 450);
    let chart = shape.chart;
    chart.title.text = "Sales";

    // Delete default generated series.
    chart.series.clear();

    // Add a series.
    let series = chart.series.add(
      "Sales",
      [
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Europe", "UK", "London Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Europe", "UK", "Liverpool Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Europe", "UK", "Manchester Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Europe", "France", "Paris Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Europe", "France", "Lyon Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "USA", "Denver Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "USA", "Seattle Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "USA", "Detroit Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "USA", "Houston Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "Canada", "Toronto Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - NA", "Canada", "Montreal Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Oceania", "Australia", "Sydney Dep."),
        new aw.Drawing.Charts.ChartMultilevelValue("Sales - Oceania", "New Zealand", "Auckland Dep.")
      ],
      [ 1236, 851, 536, 468, 179, 527, 799, 1148, 921, 457, 482, 761, 694 ]);

    // Show data labels.
    series.hasDataLabels = true;
    series.dataLabels.showValue = false;
    series.dataLabels.showCategoryName = true;

    doc.save(base.artifactsDir + "Charts.sunburst.docx");
    //ExEnd:SunburstChart
  });


    function histogramChart() {
      //ExStart:HistogramChart
      //GistId:65919861586e42e24f61a3ccb65f8f4e
      //ExFor:ChartSeriesCollection.Add(String, double[])
      //ExSummary:Shows how to create histogram chart.
      let doc = new aw.Document();
      let builder = new aw.DocumentBuilder(doc);

      // Insert a Histogram chart.
      let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Histogram, 450, 450);
      let chart = shape.chart;
      chart.title.text = "Avg Temperature since 1991";

      // Delete default generated series.
      chart.series.clear();

      // Add a series.
      chart.series.add(
        "Avg Temperature",
        [
          51.8, 53.6, 50.3, 54.7, 53.9, 54.3, 53.4, 52.9, 53.3, 53.7, 53.8, 52.0, 55.0, 52.1, 53.4,
          53.8, 53.8, 51.9, 52.1, 52.7, 51.8, 56.6, 53.3, 55.6, 56.3, 56.2, 56.1, 56.2, 53.6, 55.7,
          56.3, 55.9, 55.6
        ]);

      doc.save(base.artifactsDir + "Charts.Histogram.docx");
      //ExEnd:HistogramChart
    }

    function paretoChart() {
      //ExStart:ParetoChart
      //GistId:65919861586e42e24f61a3ccb65f8f4e
      //ExFor:ChartSeriesCollection.Add(String, String[], double[])
      //ExSummary:Shows how to create pareto chart.
      let doc = new aw.Document();
      let builder = new aw.DocumentBuilder(doc);

      // Insert a Pareto chart.
      let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Pareto, 450, 450);
      let chart = shape.chart;
      chart.title.text = "Best-Selling Car";

      // Delete default generated series.
      chart.series.clear();

      // Add a series.
      chart.series.add(
        "Best-Selling Car",
        [ "Tesla Model Y", "Toyota Corolla", "Toyota RAV4", "Ford F-Series", "Honda CR-V" ],
        [ 1.43, 0.91, 1.17, 0.98, 0.85 ]);

      doc.save(base.artifactsDir + "Charts.Pareto.docx");
      //ExEnd:ParetoChart
    }

    function boxAndWhiskerChart() {
      //ExStart:BoxAndWhiskerChart
      //GistId:65919861586e42e24f61a3ccb65f8f4e
      //ExFor:ChartSeriesCollection.Add(String, String[], double[])
      //ExSummary:Shows how to create box and whisker chart.
      let doc = new aw.Document();
      let builder = new aw.DocumentBuilder(doc);

      // Insert a Box & Whisker chart.
      let shape = builder.insertChart(aw.Drawing.Charts.ChartType.BoxAndWhisker, 450, 450);
      let chart = shape.chart;
      chart.title.text = "Points by Years";

      // Delete default generated series.
      chart.series.clear();

      // Add a series.
      let series = chart.series.add(
        "Points by Years",
        [
          "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC",
          "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR",
          "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA"
        ],
        [
          91, 80, 100, 77, 90, 104, 105, 118, 120, 101,
          114, 107, 110, 60, 79, 78, 77, 102, 101, 113,
          94, 93, 84, 71, 80, 103, 80, 94, 100, 101
        ]);

      // Show data labels.
      series.hasDataLabels = true;

      doc.save(base.artifactsDir + "Charts.BoxAndWhisker.docx");
      //ExEnd:BoxAndWhiskerChart
    }

  test('WaterfallChart', () => {
    //ExStart:WaterfallChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.add(String, String.at(], double[), bool[])
    //ExSummary:Shows how to create waterfall chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a Waterfall chart.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Waterfall, 450, 450);
    let chart = shape.chart;
    chart.title.text = "New Zealand GDP";

    // Delete default generated series.
    chart.series.clear();

    // Add a series.
    let series = chart.series.add(
      "New Zealand GDP",
      [ "2018", "2019 growth", "2020 growth", "2020", "2021 growth", "2022 growth", "2022" ],
      [ 100, 0.57, -0.25, 100.32, 20.22, -2.92, 117.62 ],
      [ true, false, false, true, false, false, true ]);

    // Show data labels.
    series.hasDataLabels = true;

    doc.save(base.artifactsDir + "Charts.waterfall.docx");
    //ExEnd:WaterfallChart
  });


  test('FunnelChart', () => {
    //ExStart:FunnelChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:aw.Drawing.Charts.ChartSeriesCollection.add(String, String.at(], double[))
    //ExSummary:Shows how to create funnel chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a Funnel chart.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Funnel, 450, 450);
    let chart = shape.chart;
    chart.title.text = "Population by Age Group";

    // Delete default generated series.
    chart.series.clear();

    // Add a series.
    let series = chart.series.add(
      "Population by Age Group",
      [ "0-9", "10-19", "20-29", "30-39", "40-49", "50-59", "60-69", "70-79", "80-89", "90-" ],
      [ 0.121, 0.128, 0.132, 0.146, 0.124, 0.124, 0.111, 0.075, 0.032, 0.007 ]);

    // Show data labels.
    series.hasDataLabels = true;
    let decimalSeparator = ".";
    series.dataLabels.numberFormat.formatCode = `0${decimalSeparator}0%`;

    doc.save(base.artifactsDir + "Charts.funnel.docx");
    //ExEnd:FunnelChart
  });


  test('LabelOrientationRotation', () => {
    //ExStart:LabelOrientationRotation
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.orientation
    //ExFor:aw.Drawing.Charts.ChartDataLabelCollection.rotation
    //ExFor:aw.Drawing.Charts.ChartDataLabel.rotation
    //ExFor:aw.Drawing.Charts.ChartDataLabel.orientation
    //ExFor:ShapeTextOrientation
    //ExSummary:Shows how to change orientation and rotation for data labels.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let series = shape.chart.series.at(0);
    let dataLabels = series.dataLabels;

    // Show data labels.
    series.hasDataLabels = true;
    dataLabels.showValue = true;
    dataLabels.showCategoryName = true;

    // Define data label shape.
    dataLabels.format.shapeType = aw.Drawing.Charts.ChartShapeType.UpArrow;
    dataLabels.format.stroke.fill.solid("#00008B");

    // Set data label orientation and rotation for the entire series.
    dataLabels.orientation = aw.Drawing.ShapeTextOrientation.VerticalFarEast;
    dataLabels.rotation = -45;

    // Change orientation and rotation of the first data label.
    dataLabels.at(0).orientation = aw.Drawing.ShapeTextOrientation.Horizontal;
    dataLabels.at(0).rotation = 45;

    doc.save(base.artifactsDir + "Charts.LabelOrientationRotation.docx");
    //ExEnd:LabelOrientationRotation
  });
});
