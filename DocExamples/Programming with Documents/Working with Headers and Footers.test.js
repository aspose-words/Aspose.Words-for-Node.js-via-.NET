// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithHeadersAndFooters", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CreateHeaderFooter', () => {
    //ExStart:CreateHeaderFooter
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use HeaderPrimary and FooterPrimary
    // if you want to set header/footer for all document.
    // This header/footer type also responsible for odd pages.
    //ExStart:HeaderFooterType
    //GistId:af238004afed43ffa79beb305a41e642
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("Header for page.");
    //ExEnd:HeaderFooterType

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.write("Footer for page.");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
    //ExEnd:CreateHeaderFooter
  });

  test('DifferentFirstPage', () => {
    //ExStart:DifferentFirstPage
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify that we want different headers and footers for first page.
    builder.pageSetup.differentFirstPageHeaderFooter = true;

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderFirst);
    builder.write("Header for the first page.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterFirst);
    builder.write("Footer for the first page.");

    builder.moveToSection(0);
    builder.writeln("Page 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.DifferentFirstPage.docx");
    //ExEnd:DifferentFirstPage
  });

  test('OddEvenPages', () => {
    //ExStart:OddEvenPages
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify that we want different headers and footers for even and odd pages.
    builder.pageSetup.oddAndEvenPagesHeaderFooter = true;

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderEven);
    builder.write("Header for even pages.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("Header for odd pages.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterEven);
    builder.write("Footer for even pages.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.write("Footer for odd pages.");

    builder.moveToSection(0);
    builder.writeln("Page 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.OddEvenPages.docx");
    //ExEnd:OddEvenPages
  });

  test('InsertImage', () => {
    //ExStart:InsertImage
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.insertImage(base.imagesDir + "Logo.jpg", aw.Drawing.RelativeHorizontalPosition.RightMargin, 10,
        aw.Drawing.RelativeVerticalPosition.Page, 10, 50, 50, aw.Drawing.WrapType.Through);

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.InsertImage.docx");
    //ExEnd:InsertImage
  });

  test('FontProps', () => {
    //ExStart:FontProps
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.font.name = "Arial";
    builder.font.bold = true;
    builder.font.size = 14;
    builder.write("Header for page.");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.FontProps.docx");
    //ExEnd:FontProps
  });

  test('PageNumbers', () => {
    //ExStart:PageNumbers
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Right;
    builder.write("Page ");
    builder.insertField("PAGE", "");
    builder.write(" of ");
    builder.insertField("NUMPAGES", "");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.PageNumbers.docx");
    //ExEnd:PageNumbers
  });

  test('LinkToPreviousHeaderFooter', () => {
    //ExStart:LinkToPreviousHeaderFooter
    //GistId:af238004afed43ffa79beb305a41e642
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.pageSetup.differentFirstPageHeaderFooter = true;

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderFirst);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.font.name = "Arial";
    builder.font.bold = true;
    builder.font.size = 14;
    builder.write("Header for the first page.");

    builder.moveToDocumentEnd();
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    let currentSection = builder.currentSection;
    let pageSetup = currentSection.pageSetup;
    pageSetup.orientation = aw.Orientation.Landscape;
    // This section does not need a different first-page header/footer we need only one title page in the document,
    // and the header/footer for this page has already been defined in the previous section.
    pageSetup.differentFirstPageHeaderFooter = false;

    // This section displays headers/footers from the previous section
    // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
    // is different for the new section.
    currentSection.headersFooters.linkToPrevious(false);
    currentSection.headersFooters.clear();

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.font.name = "Arial";
    builder.font.size = 12;
    builder.write("New Header for the first page.");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.LinkToPreviousHeaderFooter.docx");
    //ExEnd:LinkToPreviousHeaderFooter
  });

  test('SectionsWithDifferentHeaders', () => {
    //ExStart:SectionsWithDifferentHeaders
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let pageSetup = builder.currentSection.pageSetup;
    pageSetup.differentFirstPageHeaderFooter = true;
    pageSetup.headerDistance = 20;

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderFirst);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.font.name = "Arial";
    builder.font.bold = true;
    builder.font.size = 14;
    builder.write("Header for the first page.");

    builder.moveToDocumentEnd();
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    // Insert a positioned image into the top/left corner of the header.
    // Distance from the top/left edges of the page is set to 10 points.
    builder.insertImage(base.imagesDir + "Logo.jpg", aw.Drawing.RelativeHorizontalPosition.Page, 10,
        aw.Drawing.RelativeVerticalPosition.Page, 10, 50, 50, aw.Drawing.WrapType.Through);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Right;
    builder.write("Header for odd page.");

    doc.save(base.artifactsDir + "WorkingWithHeadersAndFooters.SectionsWithDifferentHeaders.docx");
    //ExEnd:SectionsWithDifferentHeaders
  });

  //ExStart:CopyHeadersFootersFromPreviousSection
  //GistId:af238004afed43ffa79beb305a41e642
  /// <summary>
  /// Clones and copies headers/footers form the previous section to the specified section.
  /// </summary>
  function copyHeadersFootersFromPreviousSection(section) {
    let previousSection = section.previousSibling.asSection();

    if (previousSection == null)
      return;

    section.headersFooters.clear();

    for (let headerFooter of previousSection.headersFooters)
      section.headersFooters.add(headerFooter.clone(true));
  }
  //ExEnd:CopyHeadersFootersFromPreviousSection

});