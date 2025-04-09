// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExPclSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('RasterizeElements', () => {
    //ExStart
    //ExFor:PclSaveOptions
    //ExFor:PclSaveOptions.saveFormat
    //ExFor:PclSaveOptions.rasterizeTransformedElements
    //ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PclSaveOptions();
    saveOptions.saveFormat = aw.SaveFormat.Pcl;
    saveOptions.rasterizeTransformedElements = true

    doc.save(base.artifactsDir + "PclSaveOptions.RasterizeElements.pcl", saveOptions);
    //ExEnd
  });


  test('FallbackFontName', () => {
    //ExStart
    //ExFor:PclSaveOptions.fallbackFontName
    //ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Non-existent font";
    builder.write("Hello world!");

    let saveOptions = new aw.Saving.PclSaveOptions();
    saveOptions.fallbackFontName = "Times New Roman";

    // This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
    // Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
    doc.save(base.artifactsDir + "PclSaveOptions.SetPrinterFont.pcl", saveOptions);
    //ExEnd
  });


  test('AddPrinterFont', () => {
    //ExStart
    //ExFor:PclSaveOptions.addPrinterFont(string, string)
    //ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Courier";
    builder.write("Hello world!");

    let saveOptions = new aw.Saving.PclSaveOptions();
    saveOptions.addPrinterFont("Courier New", "Courier");

    // When printing this document, the printer will use the "Courier New" font
    // to access places where our document used the "Courier" font.
    doc.save(base.artifactsDir + "PclSaveOptions.addPrinterFont.pcl", saveOptions);
    //ExEnd
  });


  test('GetPreservedPaperTrayInformation', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Paper tray information is now preserved when saving document to PCL format.
    // Following information is transferred from document's model to PCL file.
    for (let s of doc.sections)
    {
      let section = s.asSection();
      section.pageSetup.firstPageTray = 15;
      section.pageSetup.otherPagesTray = 12;
    }

    doc.save(base.artifactsDir + "PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
  });

});
