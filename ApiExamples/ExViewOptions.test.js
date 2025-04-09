// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const fs = require('fs');


describe("ExViewOptions", () => {
  test('SetZoomPercentage', () => {
    //ExStart
    //ExFor:Document.viewOptions
    //ExFor:ViewOptions
    //ExFor:ViewOptions.viewType
    //ExFor:ViewOptions.zoomPercent
    //ExFor:ViewOptions.zoomType
    //ExFor:ZoomType
    //ExFor:ViewType
    //ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    doc.viewOptions.viewType = aw.Settings.ViewType.PageLayout;
    doc.viewOptions.zoomPercent = 50;

    expect(doc.viewOptions.zoomType).toEqual(aw.Settings.ZoomType.Custom);
    expect(doc.viewOptions.zoomType).toEqual(aw.Settings.ZoomType.None);

    doc.save(base.artifactsDir + "ViewOptions.SetZoomPercentage.doc");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ViewOptions.SetZoomPercentage.doc");

    expect(doc.viewOptions.viewType).toEqual(aw.Settings.ViewType.PageLayout);
    expect(doc.viewOptions.zoomPercent).toEqual(50.0);
    expect(doc.viewOptions.zoomType).toEqual(aw.Settings.ZoomType.None);
  });


  test.each([aw.Settings.ZoomType.PageWidth,
    aw.Settings.ZoomType.FullPage,
    aw.Settings.ZoomType.TextFit])('SetZoomType', (zoomType) => {
    //ExStart
    //ExFor:Document.viewOptions
    //ExFor:ViewOptions
    //ExFor:ViewOptions.zoomType
    //ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Set the "ZoomType" property to "ZoomType.PageWidth" to get Microsoft Word
    // to automatically zoom the document to fit the width of the page.
    // Set the "ZoomType" property to "ZoomType.FullPage" to get Microsoft Word
    // to automatically zoom the document to make the entire first page visible.
    // Set the "ZoomType" property to "ZoomType.TextFit" to get Microsoft Word
    // to automatically zoom the document to fit the inner text margins of the first page.
    doc.viewOptions.zoomType = zoomType;

    doc.save(base.artifactsDir + "ViewOptions.SetZoomType.doc");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ViewOptions.SetZoomType.doc");

    expect(doc.viewOptions.zoomType).toEqual(zoomType);
  });


  test.each([false,
    true])('DisplayBackgroundShape', (displayBackgroundShape) => {
    //ExStart
    //ExFor:ViewOptions.displayBackgroundShape
    //ExSummary:Shows how to hide/display document background images in view options.
    // Use an HTML string to create a new document with a flat background color.
    const html = 
    `<html>
      <body style='background-color: blue'>
        <p>Hello world!</p>
      </body>
    </html>`;

    let doc = new aw.Document(Buffer.from(html));

    // The source for the document has a flat color background,
    // the presence of which will set the "DisplayBackgroundShape" flag to "true".
    expect(doc.viewOptions.displayBackgroundShape).toEqual(true);

    // Keep the "DisplayBackgroundShape" as "true" to get the document to display the background color.
    // This may affect some text colors to improve visibility.
    // Set the "DisplayBackgroundShape" to "false" to not display the background color.
    doc.viewOptions.displayBackgroundShape = displayBackgroundShape;

    doc.save(base.artifactsDir + "ViewOptions.displayBackgroundShape.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ViewOptions.displayBackgroundShape.docx");

    expect(doc.viewOptions.displayBackgroundShape).toEqual(displayBackgroundShape);
  });


  test.each([false,
    true])('DisplayPageBoundaries', (doNotDisplayPageBoundaries) => {
    //ExStart
    //ExFor:ViewOptions.doNotDisplayPageBoundaries
    //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert content that spans across 3 pages.
    builder.writeln("Paragraph 1, Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Paragraph 2, Page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Paragraph 3, Page 3.");

    // Insert a header and a footer.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.writeln("This is the header.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.writeln("This is the footer.");

    // This document contains a small amount of content that takes up a few full pages worth of space.
    // Set the "DoNotDisplayPageBoundaries" flag to "true" to get older versions of Microsoft Word to omit headers,
    // footers, and much of the vertical whitespace when displaying our document.
    // Set the "DoNotDisplayPageBoundaries" flag to "false" to get older versions of Microsoft Word
    // to normally display our document.
    doc.viewOptions.doNotDisplayPageBoundaries = doNotDisplayPageBoundaries;

    doc.save(base.artifactsDir + "ViewOptions.DisplayPageBoundaries.doc");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ViewOptions.DisplayPageBoundaries.doc");

    expect(doc.viewOptions.doNotDisplayPageBoundaries).toEqual(doNotDisplayPageBoundaries);
  });


  test.each([false,
    true])('FormsDesign', (useFormsDesign) => {
    //ExStart
    //ExFor:ViewOptions.formsDesign
    //ExSummary:Shows how to enable/disable forms design mode.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Set the "FormsDesign" property to "false" to keep forms design mode disabled.
    // Set the "FormsDesign" property to "true" to enable forms design mode.
    doc.viewOptions.formsDesign = useFormsDesign;

    doc.save(base.artifactsDir + "ViewOptions.formsDesign.xml");

    expect(fs.readFileSync(base.artifactsDir + "ViewOptions.formsDesign.xml").toString().includes("<w:formsDesign />")).toEqual(useFormsDesign);
    //ExEnd
  });
});
