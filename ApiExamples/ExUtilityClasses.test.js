// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;


describe("ExUtilityClasses", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('PointsAndInches', () => {
    //ExStart
    //ExFor:ConvertUtil
    //ExFor:ConvertUtil.pointToInch
    //ExFor:ConvertUtil.inchToPoint
    //ExSummary:Shows how to specify page properties in inches.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
    // such as inches when defining boundaries.
    let pageSetup = builder.pageSetup;
    pageSetup.topMargin = aw.ConvertUtil.inchToPoint(1.0);
    pageSetup.bottomMargin = aw.ConvertUtil.inchToPoint(2.0);
    pageSetup.leftMargin = aw.ConvertUtil.inchToPoint(2.5);
    pageSetup.rightMargin = aw.ConvertUtil.inchToPoint(1.5);

    // An inch is 72 points.
    expect(aw.ConvertUtil.inchToPoint(1)).toEqual(72.0);
    expect(aw.ConvertUtil.pointToInch(72)).toEqual(1.0);

    // Add content to demonstrate the new margins.
    builder.writeln(`This Text is ${pageSetup.leftMargin} points/${aw.ConvertUtil.pointToInch(pageSetup.leftMargin)} inches from the left, ` +
            `${pageSetup.rightMargin} points/${aw.ConvertUtil.pointToInch(pageSetup.rightMargin)} inches from the right, ` +
            `${pageSetup.topMargin} points/${aw.ConvertUtil.pointToInch(pageSetup.topMargin)} inches from the top, ` +
            `and ${pageSetup.bottomMargin} points/${aw.ConvertUtil.pointToInch(pageSetup.bottomMargin)} inches from the bottom of the page.`);

    doc.save(base.artifactsDir + "UtilityClasses.PointsAndInches.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "UtilityClasses.PointsAndInches.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.topMargin).toBeCloseTo(72.0, 2);
    expect(aw.ConvertUtil.pointToInch(pageSetup.topMargin)).toBeCloseTo(1.0, 2);
    expect(pageSetup.bottomMargin).toBeCloseTo(144.0, 2);
    expect(aw.ConvertUtil.pointToInch(pageSetup.bottomMargin)).toBeCloseTo(2.0, 2);
    expect(pageSetup.leftMargin).toBeCloseTo(180.0, 2);
    expect(aw.ConvertUtil.pointToInch(pageSetup.leftMargin)).toBeCloseTo(2.5, 2);
    expect(pageSetup.rightMargin).toBeCloseTo(108.0, 2);
    expect(aw.ConvertUtil.pointToInch(pageSetup.rightMargin)).toBeCloseTo(1.5, 2);
  });


  test('PointsAndMillimeters', () => {
    //ExStart
    //ExFor:ConvertUtil.millimeterToPoint
    //ExSummary:Shows how to specify page properties in millimeters.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a more familiar measurement unit,
    // such as millimeters when defining boundaries.
    let pageSetup = builder.pageSetup;
    pageSetup.topMargin = aw.ConvertUtil.millimeterToPoint(30);
    pageSetup.bottomMargin = aw.ConvertUtil.millimeterToPoint(50);
    pageSetup.leftMargin = aw.ConvertUtil.millimeterToPoint(80);
    pageSetup.rightMargin = aw.ConvertUtil.millimeterToPoint(40);

    // A centimeter is approximately 28.3 points.
    expect(aw.ConvertUtil.millimeterToPoint(10)).toBeCloseTo(28.34, 1);

    // Add content to demonstrate the new margins.
    builder.writeln(`This Text is ${pageSetup.leftMargin} points from the left, ` +
            `${pageSetup.rightMargin} points from the right, ` +
            `${pageSetup.topMargin} points from the top, ` +
            `and ${pageSetup.bottomMargin} points from the bottom of the page.`);

    doc.save(base.artifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "UtilityClasses.PointsAndMillimeters.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.topMargin).toBeCloseTo(85.05, 2);
    expect(pageSetup.bottomMargin).toBeCloseTo(141.75, 2);
    expect(pageSetup.leftMargin).toBeCloseTo(226.75, 2);
    expect(pageSetup.rightMargin).toBeCloseTo(113.4, 2);
  });


  test('PointsAndPixels', () => {
    //ExStart
    //ExFor:ConvertUtil.pixelToPoint(double)
    //ExFor:ConvertUtil.pointToPixel(double)
    //ExSummary:Shows how to specify page properties in pixels.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A section's "Page Setup" defines the size of the page margins in points.
    // We can also use the "ConvertUtil" class to use a different measurement unit,
    // such as pixels when defining boundaries.
    let pageSetup = builder.pageSetup;
    pageSetup.topMargin = aw.ConvertUtil.pixelToPoint(100);
    pageSetup.bottomMargin = aw.ConvertUtil.pixelToPoint(200);
    pageSetup.leftMargin = aw.ConvertUtil.pixelToPoint(225);
    pageSetup.rightMargin = aw.ConvertUtil.pixelToPoint(125);

    // A pixel is 0.75 points.
    expect(aw.ConvertUtil.pixelToPoint(1)).toEqual(0.75);
    expect(aw.ConvertUtil.pointToPixel(0.75)).toEqual(1.0);

    // The default DPI value used is 96.
    expect(aw.ConvertUtil.pixelToPoint(1, 96)).toEqual(0.75);

    // Add content to demonstrate the new margins.
    builder.writeln(`This Text is ${pageSetup.leftMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.leftMargin)} pixels from the left, ` +
            `${pageSetup.rightMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.rightMargin)} pixels from the right, ` +
            `${pageSetup.topMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.topMargin)} pixels from the top, ` +
            `and ${pageSetup.bottomMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.bottomMargin)} pixels from the bottom of the page.`);

    doc.save(base.artifactsDir + "UtilityClasses.PointsAndPixels.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "UtilityClasses.PointsAndPixels.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.topMargin).toBeCloseTo(75.0, 2);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.topMargin)).toBeCloseTo(100.0, 2);
    expect(pageSetup.bottomMargin).toBeCloseTo(150.0, 2);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.bottomMargin)).toBeCloseTo(200.0, 2);
    expect(pageSetup.leftMargin).toBeCloseTo(168.75, 2);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.leftMargin)).toBeCloseTo(225.0, 2);
    expect(pageSetup.rightMargin).toBeCloseTo(93.75, 2);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.rightMargin)).toBeCloseTo(125.0, 2);
  });


  test('PointsAndPixelsDpi', () => {
    //ExStart
    //ExFor:ConvertUtil.pixelToNewDpi
    //ExFor:ConvertUtil.pixelToPoint(double, double)
    //ExFor:ConvertUtil.pointToPixel(double, double)
    //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Define the size of the top margin of this section in pixels, according to a custom DPI.
    const myDpi = 192;

    let pageSetup = builder.pageSetup;
    pageSetup.topMargin = aw.ConvertUtil.pixelToPoint(100, myDpi);

    expect(pageSetup.topMargin).toBeCloseTo(37.5, 2);

    // At the default DPI of 96, a pixel is 0.75 points.
    expect(aw.ConvertUtil.pixelToPoint(1)).toEqual(0.75);

    builder.writeln(`This Text is ${pageSetup.topMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.topMargin, myDpi)} ` +
            `pixels (at a DPI of ${myDpi}) from the top of the page.`);

    // Set a new DPI and adjust the top margin value accordingly.
    const newDpi = 300;
    pageSetup.topMargin = aw.ConvertUtil.pixelToNewDpi(pageSetup.topMargin, myDpi, newDpi);
    expect(pageSetup.topMargin).toBeCloseTo(59.0, 2);

    builder.writeln(`At a DPI of ${newDpi}, the text is now ${pageSetup.topMargin} points/${aw.ConvertUtil.pointToPixel(pageSetup.topMargin, myDpi)} ` +
            "pixels from the top of the page.");

    doc.save(base.artifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "UtilityClasses.PointsAndPixelsDpi.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.topMargin).toBeCloseTo(59.0, 2);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.topMargin)).toBeCloseTo(78.66, 1);
    expect(aw.ConvertUtil.pointToPixel(pageSetup.topMargin, myDpi)).toBeCloseTo(157.33, 2);
    expect(aw.ConvertUtil.pointToPixel(100)).toBeCloseTo(133.33, 2);
    expect(aw.ConvertUtil.pointToPixel(100, myDpi)).toBeCloseTo(266.66, 1);
  });

});
