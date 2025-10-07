// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const MemoryStream = require('memorystream');


describe("WorkingWithTxtSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('AddBidiMarks', () => {
    //ExStart:AddBidiMarks
    //GistId:ee038b97a80cf17ce52665651e81d832
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    builder.paragraphFormat.bidi = true;
    builder.writeln("שלום עולם!");
    builder.writeln("مرحبا بالعالم!");

    let saveOptions = new aw.Saving.TxtSaveOptions();
    saveOptions.addBidiMarks = true;

    doc.save(base.artifactsDir + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
    //ExEnd:AddBidiMarks
  });

  test('UseTabForListIndentation', () => {
    //ExStart:UseTabForListIndentation
    //GistId:ee038b97a80cf17ce52665651e81d832
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a list with three levels of indentation.
    builder.listFormat.applyNumberDefault();
    builder.writeln("Item 1");
    builder.listFormat.listIndent();
    builder.writeln("Item 2");
    builder.listFormat.listIndent();
    builder.write("Item 3");

    let saveOptions = new aw.Saving.TxtSaveOptions();
    saveOptions.listIndentation.count = 1;
    saveOptions.listIndentation.character = '\t';

    doc.save(base.artifactsDir + "WorkingWithTxtSaveOptions.UseTabForListIndentation.txt", saveOptions);
    //ExEnd:UseTabForListIndentation
  });

  test('UseSpaceForListIndentation', () => {
    //ExStart:UseSpaceForListIndentation
    //GistId:ee038b97a80cf17ce52665651e81d832
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a list with three levels of indentation.
    builder.listFormat.applyNumberDefault();
    builder.writeln("Item 1");
    builder.listFormat.listIndent();
    builder.writeln("Item 2");
    builder.listFormat.listIndent();
    builder.write("Item 3");

    let saveOptions = new aw.Saving.TxtSaveOptions();
    saveOptions.listIndentation.count = 3;
    saveOptions.listIndentation.character = ' ';

    doc.save(base.artifactsDir + "WorkingWithTxtSaveOptions.UseSpaceForListIndentation.txt", saveOptions);
    //ExEnd:UseSpaceForListIndentation
  });

  test('ExportHeadersFootersMode', () => {
    //ExStart:ExportHeadersFootersMode
    //GistId:ee038b97a80cf17ce52665651e81d832
    let doc = new aw.Document();

    // Insert even and primary headers/footers into the document.
    // The primary header/footers will override the even headers/footers.
    doc.firstSection.headersFooters.add(new aw.HeaderFooter(doc, aw.HeaderFooterType.HeaderEven));
    doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderEven).appendParagraph("Even header");
    doc.firstSection.headersFooters.add(new aw.HeaderFooter(doc, aw.HeaderFooterType.FooterEven));
    doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterEven).appendParagraph("Even footer");
    doc.firstSection.headersFooters.add(new aw.HeaderFooter(doc, aw.HeaderFooterType.HeaderPrimary));
    doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).appendParagraph("Primary header");
    doc.firstSection.headersFooters.add(new aw.HeaderFooter(doc, aw.HeaderFooterType.FooterPrimary));
    doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary).appendParagraph("Primary footer");

    // Insert pages to display these headers and footers.
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Page 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.write("Page 3");

    let options = new aw.Saving.TxtSaveOptions();
    options.saveFormat = aw.SaveFormat.Text;

    // All headers and footers are placed at the very end of the output document.
    options.exportHeadersFootersMode = aw.Saving.TxtExportHeadersFootersMode.AllAtEnd;
    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.AllAtEnd.txt", options);

    // Only primary headers and footers are exported at the beginning and end of each section.
    options.exportHeadersFootersMode = aw.Saving.TxtExportHeadersFootersMode.PrimaryOnly;
    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.PrimaryOnly.txt", options);

    // No headers and footers are exported.
    options.exportHeadersFootersMode = aw.Saving.TxtExportHeadersFootersMode.None;
    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.None.txt", options);
    //ExEnd:ExportHeadersFootersMode
  });

});