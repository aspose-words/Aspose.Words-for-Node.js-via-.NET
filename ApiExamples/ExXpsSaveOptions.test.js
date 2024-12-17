// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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


describe("ExXpsSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('OutlineLevels', () => {
    //ExStart
    //ExFor:XpsSaveOptions
    //ExFor:XpsSaveOptions.#ctor
    //ExFor:aw.Saving.XpsSaveOptions.outlineOptions
    //ExFor:aw.Saving.XpsSaveOptions.saveFormat
    //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    expect(builder.paragraphFormat.isHeading).toEqual(true);

    builder.writeln("Heading 1");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;

    builder.writeln("Heading 1.1");
    builder.writeln("Heading 1.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading3;

    builder.writeln("Heading 1.2.1");
    builder.writeln("Heading 1.2.2");

    // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    let saveOptions = new aw.Saving.XpsSaveOptions();

    expect(saveOptions.saveFormat).toEqual(aw.SaveFormat.Xps);

    // The output XPS document will contain an outline, a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
    // The last two headings we have inserted above will not appear.
    saveOptions.outlineOptions.headingsOutlineLevels = 2;

    doc.save(base.artifactsDir + "XpsSaveOptions.OutlineLevels.xps", saveOptions);
    //ExEnd
  });


  test.each([false,
    true])('BookFold', (renderTextAsBookFold) => {
    //ExStart
    //ExFor:XpsSaveOptions.#ctor(SaveFormat)
    //ExFor:aw.Saving.XpsSaveOptions.useBookFoldPrintingSettings
    //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    let xpsOptions = new aw.Saving.XpsSaveOptions(aw.SaveFormat.Xps);

    // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
    // in the output XPS in a way that helps us use it to make a booklet.
    // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
    xpsOptions.useBookFoldPrintingSettings = renderTextAsBookFold;

    // If we are rendering the document as a booklet, we must set the "MultiplePages"
    // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
    if (renderTextAsBookFold)
      for (let s of doc.sections.toArray())
      {
        s.pageSetup.multiplePages = aw.Settings.MultiplePagesType.BookFoldPrinting;
      }

    // Once we print this document, we can turn it into a booklet by stacking the pages
    // to come out of the printer and folding down the middle.
    doc.save(base.artifactsDir + "XpsSaveOptions.BookFold.xps", xpsOptions);
    //ExEnd
  });


  test.each([false,
    true])('OptimizeOutput', (optimizeOutput) => {
    //ExStart
    //ExFor:aw.Saving.FixedPageSaveOptions.optimizeOutput
    //ExSummary:Shows how to optimize document objects while saving to xps.
    let doc = new aw.Document(base.myDir + "Unoptimized document.docx");

    // Create an "XpsSaveOptions" object to pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    let saveOptions = new aw.Saving.XpsSaveOptions();
    // Set the "OptimizeOutput" property to "true" to take measures such as removing nested or empty canvases
    // and concatenating adjacent runs with identical formatting to optimize the output document's content.
    // This may affect the appearance of the document.
    // Set the "OptimizeOutput" property to "false" to save the document normally.
    saveOptions.optimizeOutput = optimizeOutput;

    doc.save(base.artifactsDir + "XpsSaveOptions.optimizeOutput.xps", saveOptions);
    //ExEnd

    var testedFileLength = fs.statSync(base.artifactsDir + "XpsSaveOptions.optimizeOutput.xps").size;
    if (optimizeOutput)
      expect(testedFileLength).toBeLessThan(44000);
    else
      expect(testedFileLength).toBeLessThan(65000);

    TestUtil.docPackageFileContainsString(
      optimizeOutput
        ? "Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" " +
        "UnicodeString=\"This document contains complex content which can be optimized to save space when \""
        : "<Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" UnicodeString=\"This\"",
      base.artifactsDir + "XpsSaveOptions.optimizeOutput.xps", "1.fpage");
  });


  test('ExportExactPages', () => {
    //ExStart
    //ExFor:aw.Saving.FixedPageSaveOptions.pageSet
    //ExFor:PageSet.#ctor(int[])
    //ExSummary:Shows how to extract pages based on exact page indices.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add five pages to the document.
    for (let i = 1; i < 6; i++)
    {
      builder.write("Page " + i);
      builder.insertBreak(aw.BreakType.PageBreak);
    }

    // Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how that method converts the document to .XPS.
    let xpsOptions = new aw.Saving.XpsSaveOptions();

    // Use the "PageSet" property to select a set of the document's pages to save to output XPS.
    // In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
    xpsOptions.pageSet = new aw.Saving.PageSet([0, 1, 3]);

    doc.save(base.artifactsDir + "XpsSaveOptions.ExportExactPages.xps", xpsOptions);
    //ExEnd
  });
});
