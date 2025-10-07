// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithCharts", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('FormatNumberOfDataLabel', () => {
    //ExStart:FormatNumberOfDataLabel
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 432, 252);

    let chart = shape.chart;
    chart.title.text = "Data Labels With Different Number Format";

    // Delete default generated series.
    chart.series.clear();

    let series1 = chart.series.add("Aspose Series 1",
        ["Category 1", "Category 2", "Category 3"],
        [2.5, 1.5, 3.5]);

    series1.hasDataLabels = true;
    series1.dataLabels.showValue = true;
    series1.dataLabels.at(0).numberFormat.formatCode = "\"$\"#,##0.00";
    series1.dataLabels.at(1).numberFormat.formatCode = "dd/mm/yyyy";
    series1.dataLabels.at(2).numberFormat.formatCode = "0.00%";

    // Or you can set format code to be linked to a source cell,
    // in this case NumberFormat will be reset to general and inherited from a source cell.
    series1.dataLabels.at(2).numberFormat.isLinkedToSource = true;

    doc.save(base.artifactsDir + "WorkingWithCharts.FormatNumberOfDataLabel.docx");
    //ExEnd:FormatNumberOfDataLabel
  });

  test('CreateChartUsingShape', () => {
    //ExStart:CreateChartUsingShape
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 432, 252);

    let chart = shape.chart;
    chart.title.show = true;
    chart.title.text = "Line Chart Title";
    chart.title.overlay = false;

    // Please note if null or empty value is specified as title text, auto generated title will be shown.

    chart.legend.position = aw.Drawing.Charts.LegendPosition.Left;
    chart.legend.overlay = true;

    doc.save(base.artifactsDir + "WorkingWithCharts.CreateChartUsingShape.docx");
    //ExEnd:CreateChartUsingShape
  });

  test('InsertSimpleColumnChart', () => {
    //ExStart:InsertSimpleColumnChart
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // You can specify different chart types and sizes.
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    //ExStart:ChartSeriesCollection
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let seriesColl = chart.series;

    console.log(seriesColl.count);
    //ExEnd:ChartSeriesCollection

    // Delete default generated series.
    seriesColl.clear();

    // Create category names array, in this example we have two categories.
    let categories = ["Category 1", "Category 2"];

    // Please note, data arrays must not be empty and arrays must be the same size.
    seriesColl.add("Aspose Series 1", categories, [1, 2]);
    seriesColl.add("Aspose Series 2", categories, [3, 4]);
    seriesColl.add("Aspose Series 3", categories, [5, 6]);
    seriesColl.add("Aspose Series 4", categories, [7, 8]);
    seriesColl.add("Aspose Series 5", categories, [9, 10]);

    doc.save(base.artifactsDir + "WorkingWithCharts.InsertSimpleColumnChart.docx");
    //ExEnd:InsertSimpleColumnChart
  });

  test('InsertColumnChart', () => {
    //ExStart:InsertColumnChart
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    chart.series.add("Aspose Series 1", ["Category 1", "Category 2"], [1, 2]);

    doc.save(base.artifactsDir + "WorkingWithCharts.InsertColumnChart.docx");
    //ExEnd:InsertColumnChart
  });

  test('InsertAreaChart', () => {
    //ExStart:InsertAreaChart
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Area, 432, 252);

    let chart = shape.chart;
    chart.series.add("Aspose Series 1", [
            new Date(2002, 4, 1),
            new Date(2002, 5, 1),
            new Date(2002, 6, 1),
            new Date(2002, 7, 1),
            new Date(2002, 8, 1)
        ],
        [32, 32, 28, 12, 15]);

    doc.save(base.artifactsDir + "WorkingWithCharts.InsertAreaChart.docx");
    //ExEnd:InsertAreaChart
  });

  test('InsertBubbleChart', () => {
    //ExStart:InsertBubbleChart
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Bubble, 432, 252);

    let chart = shape.chart;
    chart.series.add("Aspose Series 1", [0.7, 1.8, 2.6], [2.7, 3.2, 0.8],
        [10, 4, 8]);

    doc.save(base.artifactsDir + "WorkingWithCharts.InsertBubbleChart.docx");
    //ExEnd:InsertBubbleChart
  });

  test('InsertScatterChart', () => {
    //ExStart:InsertScatterChart
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 432, 252);

    let chart = shape.chart;
    chart.series.add("Aspose Series 1", [0.7, 1.8, 2.6], [2.7, 3.2, 0.8]);

    doc.save(base.artifactsDir + "WorkingWithCharts.InsertScatterChart.docx");
    //ExEnd:InsertScatterChart
  });

  test('DefineAxisProperties', () => {
    //ExStart:DefineAxisProperties
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert chart
    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Area, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        [
            new Date(2002, 0, 1), new Date(2002, 5, 1), new Date(2002, 6, 1),
            new Date(2002, 7, 1), new Date(2002, 8, 1)
        ],
        [640, 320, 280, 120, 150]);

    let xAxis = chart.axisX;
    let yAxis = chart.axisY;

    // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
    xAxis.categoryType = aw.Drawing.Charts.AxisCategoryType.Category;
    xAxis.crosses = aw.Drawing.Charts.AxisCrosses.Custom;
    xAxis.crossesAt = 3; // Measured in display units of the Y axis (hundreds).
    xAxis.reverseOrder = true;
    xAxis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    xAxis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Outside;
    xAxis.tickLabels.offset = 200;

    yAxis.tickLabels.position = aw.Drawing.Charts.AxisTickLabelPosition.High;
    yAxis.majorUnit = 100;
    yAxis.minorUnit = 50;
    yAxis.displayUnit.unit = aw.Drawing.Charts.AxisBuiltInUnit.Hundreds;
    yAxis.scaling.minimum = new aw.Drawing.Charts.AxisBound(100);
    yAxis.scaling.maximum = new aw.Drawing.Charts.AxisBound(700);

    doc.save(base.artifactsDir + "WorkingWithCharts.DefineAxisProperties.docx");
    //ExEnd:DefineAxisProperties
  });

  test('DateTimeValuesToAxis', () => {
    //ExStart:DateTimeValuesToAxis
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);
    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        [
            new Date(2017, 10, 6), new Date(2017, 10, 9), new Date(2017, 10, 15),
            new Date(2017, 10, 21), new Date(2017, 10, 25), new Date(2017, 10, 29)
        ],
        [1.2, 0.3, 2.1, 2.9, 4.2, 5.3]);

    let xAxis = chart.axisX;
    xAxis.scaling.minimum = new aw.Drawing.Charts.AxisBound(new Date(2017, 10, 5));
    xAxis.scaling.maximum = new aw.Drawing.Charts.AxisBound(new Date(2017, 11, 3));

    // Set major units to a week and minor units to a day.
    xAxis.majorUnit = 7;
    xAxis.minorUnit = 1;
    xAxis.majorTickMark = aw.Drawing.Charts.AxisTickMark.Cross;
    xAxis.minorTickMark = aw.Drawing.Charts.AxisTickMark.Outside;

    doc.save(base.artifactsDir + "WorkingWithCharts.DateTimeValuesToAxis.docx");
    //ExEnd:DateTimeValuesToAxis
  });

  test('NumberFormatForAxis', () => {
    //ExStart:NumberFormatForAxis
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
        [1900000, 850000, 2100000, 600000, 1500000]);

    chart.axisY.numberFormat.formatCode = "#,##0";

    doc.save(base.artifactsDir + "WorkingWithCharts.NumberFormatForAxis.docx");
    //ExEnd:NumberFormatForAxis
  });

  test('BoundsOfAxis', () => {
    //ExStart:BoundsOfAxis
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
        [1.2, 0.3, 2.1, 2.9, 4.2]);

    chart.axisY.scaling.minimum = new aw.Drawing.Charts.AxisBound(0);
    chart.axisY.scaling.maximum = new aw.Drawing.Charts.AxisBound(6);

    doc.save(base.artifactsDir + "WorkingWithCharts.BoundsOfAxis.docx");
    //ExEnd:BoundsOfAxis
  });

  test('IntervalUnitBetweenLabelsOnAxis', () => {
    //ExStart:IntervalUnitBetweenLabelsOnAxis
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
        [1.2, 0.3, 2.1, 2.9, 4.2]);

    chart.axisX.tickLabels.spacing = 2;

    doc.save(base.artifactsDir + "WorkingWithCharts.IntervalUnitBetweenLabelsOnAxis.docx");
    //ExEnd:IntervalUnitBetweenLabelsOnAxis
  });

  test('HideChartAxis', () => {
    //ExStart:HideChartAxis
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    chart.series.add("Aspose Series 1",
        ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5"],
        [1.2, 0.3, 2.1, 2.9, 4.2]);

    chart.axisY.hidden = true;

    doc.save(base.artifactsDir + "WorkingWithCharts.HideChartAxis.docx");
    //ExEnd:HideChartAxis
  });

  test('TickMultiLineLabelAlignment', () => {
    //ExStart:TickMultiLineLabelAlignment
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Scatter, 450, 250);

    let axis = shape.chart.axisX;
    // This property has effect only for multi-line labels.
    axis.tickLabels.alignment = aw.ParagraphAlignment.Right;

    doc.save(base.artifactsDir + "WorkingWithCharts.TickMultiLineLabelAlignment.docx");
    //ExEnd:TickMultiLineLabelAlignment
  });

  test('ChartDataLabel', () => {
    //ExStart:WorkWithChartDataLabel
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Bar, 432, 252);

    let chart = shape.chart;
    let series0 = shape.chart.series.at(0);

    let labels = series0.dataLabels;
    labels.showLegendKey = true;
    // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
    // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its
    // corresponding data point.
    labels.showLeaderLines = true;
    labels.showCategoryName = false;
    labels.showPercentage = false;
    labels.showSeriesName = true;
    labels.showValue = true;
    labels.separator = "/";
    labels.showValue = true;

    doc.save(base.artifactsDir + "WorkingWithCharts.ChartDataLabel.docx");
    //ExEnd:WorkWithChartDataLabel
  });

  test('DefaultOptionsForDataLabels', () => {
    //ExStart:DefaultOptionsForDataLabels
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Pie, 432, 252);

    let chart = shape.chart;

    chart.series.clear();

    let series = chart.series.add("Aspose Series 1",
        ["Category 1", "Category 2", "Category 3"],
        [2.7, 3.2, 0.8]);

    let labels = series.dataLabels;
    labels.showPercentage = true;
    labels.showValue = true;
    labels.showLeaderLines = false;
    labels.separator = " - ";

    doc.save(base.artifactsDir + "WorkingWithCharts.DefaultOptionsForDataLabels.docx");
    //ExEnd:DefaultOptionsForDataLabels
  });

  test('SingleChartDataPoint', () => {
    //ExStart:WorkWithSingleChartDataPoint
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 432, 252);

    let chart = shape.chart;
    let series0 = chart.series.at(0);
    let series1 = chart.series.at(1);

    let dataPointCollection = series0.dataPoints;
    let dataPoint00 = dataPointCollection.at(0);
    let dataPoint01 = dataPointCollection.at(1);

    dataPoint00.explosion = 50;
    dataPoint00.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Circle;
    dataPoint00.marker.size = 15;

    dataPoint01.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Diamond;
    dataPoint01.marker.size = 20;

    let dataPoint12 = series1.dataPoints.at(2);
    dataPoint12.invertIfNegative = true;
    dataPoint12.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Star;
    dataPoint12.marker.size = 20;

    doc.save(base.artifactsDir + "WorkingWithCharts.SingleChartDataPoint.docx");
    //ExEnd:WorkWithSingleChartDataPoint
  });

  test('SingleChartSeries', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 432, 252);

    let chart = shape.chart;

    //ExStart:WorkWithSingleChartSeries
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let series0 = chart.series.at(0);
    let series1 = chart.series.at(1);

    series0.name = "Chart Series Name 1";
    series1.name = "Chart Series Name 2";

    // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
    series0.smooth = true;
    series1.smooth = true;
    //ExEnd:WorkWithSingleChartSeries

    //ExStart:ChartDataPoint
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    // Specifies whether by default the parent element shall inverts its colors if the value is negative.
    series0.invertIfNegative = true;

    series0.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Circle;
    series0.marker.size = 15;

    series1.marker.symbol = aw.Drawing.Charts.MarkerSymbol.Star;
    series1.marker.size = 10;
    //ExEnd:ChartDataPoint

    doc.save(base.artifactsDir + "WorkingWithCharts.SingleChartSeries.docx");

  });

  test('FillFormatting', () => {
    //ExStart:FillFormatting
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Column, 432, 252);

    let chart = shape.chart;
    let seriesColl = chart.series;

    // Delete default generated series.
    seriesColl.clear();

    // Create category names array.
    let categories = ["AW Category 1", "AW Category 2"];

    // Adding new series. Value and category arrays must be the same size.
    let series1 = seriesColl.add("AW Series 1", categories, [1, 2]);
    let series2 = seriesColl.add("AW Series 2", categories, [3, 4]);
    let series3 = seriesColl.add("AW Series 3", categories, [5, 6]);

    // Set series color.
    series1.format.fill.foreColor = "#FF0000";
    series2.format.fill.foreColor = "#FFFF00";
    series3.format.fill.foreColor = "#0000FF";

    doc.save(base.artifactsDir + "WorkingWithCharts.FillFormatting.docx");
    //ExEnd:FillFormatting
  });

  test('StrokeFormatting', () => {
    //ExStart:StrokeFormatting
    //GistId:7ce46b3fa44be2f120f85d4e070329db
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertChart(aw.Drawing.Charts.ChartType.Line, 432, 252);

    let chart = shape.chart;
    let seriesColl = chart.series;

    // Delete default generated series.
    seriesColl.clear();

    // Adding new series.
    let series1 = seriesColl.add("AW Series 1", [0.7, 1.8, 2.6],
        [2.7, 3.2, 0.8]);
    let series2 = seriesColl.add("AW Series 2", [0.5, 1.5, 2.5],
        [3, 1, 2]);

    // Set series color.
    series1.format.stroke.foreColor = "#FF0000";
    series1.format.stroke.weight = 5;
    series2.format.stroke.foreColor = "#90EE90";
    series2.format.stroke.weight = 5;

    doc.save(base.artifactsDir + "WorkingWithCharts.StrokeFormatting.docx");
    //ExEnd:StrokeFormatting
  });
});
