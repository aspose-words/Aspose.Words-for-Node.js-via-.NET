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


describe("ExPdfSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('OnePage', async () => {
    //ExStart
    //ExFor:FixedPageSaveOptions.pageSet
    //ExFor:Document.save(Stream, SaveOptions)
    //ExSummary:Shows how to convert only some of the pages in a document to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3.");

    var stream = fs.createWriteStream(base.artifactsDir + "PdfSaveOptions.OnePage.pdf");
    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
    options.pageSet = new aw.Saving.PageSet(1);

    // This document will contain one page starting from page two, which will only contain the second page.
    doc.save(stream, options);
    await new Promise(resolve => stream.on("finish", resolve));
    //ExEnd
  });


  test('HeadingsOutlineLevels', () => {
    //ExStart
    //ExFor:ParagraphFormat.isHeading
    //ExFor:PdfSaveOptions.outlineOptions
    //ExFor:PdfSaveOptions.saveFormat
    //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
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

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.saveFormat = aw.SaveFormat.Pdf;

    // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
    // The last two headings we have inserted above will not appear.
    saveOptions.outlineOptions.headingsOutlineLevels = 2;

    doc.save(base.artifactsDir + "PdfSaveOptions.headingsOutlineLevels.pdf", saveOptions);
    //ExEnd
  });


  test.each([false,
    true])('CreateMissingOutlineLevels', (createMissingOutlineLevels) => {
    //ExStart
    //ExFor:OutlineOptions.createMissingOutlineLevels
    //ExFor:PdfSaveOptions.outlineOptions
    //ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert headings that can serve as TOC entries of levels 1 and 5.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    expect(builder.paragraphFormat.isHeading).toEqual(true);

    builder.writeln("Heading 1");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading5;

    builder.writeln("Heading 1.1.1.1.1");
    builder.writeln("Heading 1.1.1.1.2");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();

    // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
    saveOptions.outlineOptions.headingsOutlineLevels = 5;

    // This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4.
    // The output PDF document will treat outline levels 2, 3, and 4 as "missing".
    // Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
    // leaving blank outline entries since there are no usable headings.
    // Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
    // and treat the outline level 5 headings as level 2.
    saveOptions.outlineOptions.createMissingOutlineLevels = createMissingOutlineLevels;

    doc.save(base.artifactsDir + "PdfSaveOptions.createMissingOutlineLevels.pdf", saveOptions);
    //ExEnd
  });


  test.each([false,
    true])('TableHeadingOutlines', (createOutlinesForHeadingsInTables) => {
    //ExStart
    //ExFor:OutlineOptions.createOutlinesForHeadingsInTables
    //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a table with three rows. The first row,
    // whose text we will format in a heading-type style, will serve as the column header.
    builder.startTable();
    builder.insertCell();
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.write("Customers");
    builder.endRow();
    builder.insertCell();
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Normal;
    builder.write("John Doe");
    builder.endRow();
    builder.insertCell();
    builder.write("Jane Doe");
    builder.endTable();

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let pdfSaveOptions = new aw.Saving.PdfSaveOptions();

    // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "1" to get the outline
    // to only register headings with heading levels that are no larger than 1.
    pdfSaveOptions.outlineOptions.headingsOutlineLevels = 1;

    // Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
    // such as the one we have created above from the outline.
    // Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
    // in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
    pdfSaveOptions.outlineOptions.createOutlinesForHeadingsInTables = createOutlinesForHeadingsInTables;

    doc.save(base.artifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
    //ExEnd
  });


  test('ExpandedOutlineLevels', () => {
    //ExStart
    //ExFor:Document.save(String, SaveOptions)
    //ExFor:PdfSaveOptions
    //ExFor:OutlineOptions.headingsOutlineLevels
    //ExFor:OutlineOptions.expandedOutlineLevels
    //ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert headings of levels 1 to 5.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    expect(builder.paragraphFormat.isHeading).toEqual(true);

    builder.writeln("Heading 1");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;

    builder.writeln("Heading 1.1");
    builder.writeln("Heading 1.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading3;

    builder.writeln("Heading 1.2.1");
    builder.writeln("Heading 1.2.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading4;

    builder.writeln("Heading 1.2.2.1");
    builder.writeln("Heading 1.2.2.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading5;

    builder.writeln("Heading 1.2.2.2.1");
    builder.writeln("Heading 1.2.2.2.2");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
    // Clicking on an entry in this outline will take us to the location of its respective heading.
    // Set the "HeadingsOutlineLevels" property to "4" to exclude all headings whose levels are above 4 from the outline.
    options.outlineOptions.headingsOutlineLevels = 4;

    // If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
    // an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
    // In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
    // the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on.
    // In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
    // Set the "ExpandedOutlineLevels" property to "2" to automatically expand all heading level 2 and lower outline entries
    // and collapse all level and 3 and higher entries when we open the document.
    options.outlineOptions.expandedOutlineLevels = 2;

    doc.save(base.artifactsDir + "PdfSaveOptions.expandedOutlineLevels.pdf", options);
    //ExEnd
  });


  test.each([false,
    true])('UpdateFields', (updateFields) => {
    //ExStart
    //ExFor:PdfSaveOptions.clone
    //ExFor:SaveOptions.updateFields
    //ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
    // We will need to manually update them using updating methods such as "Field.update()", and "Document.updateFields()"
    // each time we need them to display accurate values.
    builder.write("Page ");
    builder.insertField("PAGE", "");
    builder.write(" of ");
    builder.insertField("NUMPAGES", "");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Hello World!");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
    // This is the preferable option if we know that all our fields will be up to date before saving.
    // Set the "UpdateFields" property to "true" to iterate through all the document
    // fields and update them before we save it as a PDF. This will make sure that all the fields will display
    // the most accurate values in the PDF.
    options.updateFields = updateFields;

    // We can clone PdfSaveOptions objects.
    expect(options).not.toBe(options.clone());

    doc.save(base.artifactsDir + "PdfSaveOptions.updateFields.pdf", options);
    //ExEnd
  });


  test.each([false,
    true])('PreserveFormFields', (preserveFormFields) => {
    //ExStart
    //ExFor:PdfSaveOptions.preserveFormFields
    //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Please select a fruit: ");

    // Insert a combo box which will allow a user to choose an option from a collection of strings.
    builder.insertComboBox("MyComboBox", [ "Apple", "Banana", "Cherry" ], 0);

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let pdfOptions = new aw.Saving.PdfSaveOptions();

    // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
    // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
    // their current values and display them as plain text in the output PDF.
    pdfOptions.preserveFormFields = preserveFormFields;

    doc.save(base.artifactsDir + "PdfSaveOptions.preserveFormFields.pdf", pdfOptions);
    //ExEnd
  });


  test.each([aw.Saving.PdfCompliance.PdfA2u,
    aw.Saving.PdfCompliance.Pdf17,
    aw.Saving.PdfCompliance.PdfA2a,
    aw.Saving.PdfCompliance.PdfUa1,
    aw.Saving.PdfCompliance.Pdf20,
    aw.Saving.PdfCompliance.PdfA4,
    aw.Saving.PdfCompliance.PdfA4Ua2,
    aw.Saving.PdfCompliance.PdfUa2])('Compliance', (pdfCompliance) => {
    //ExStart
    //ExFor:aw.Saving.PdfSaveOptions.compliance
    //ExFor:PdfCompliance
    //ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
    let doc = new aw.Document(base.myDir + "Images.docx");

      // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
      // to modify how that method converts the document to .PDF.
      // Note that some PdfSaveOptions are prohibited when saving to one of the standards and automatically fixed.
      // Use IWarningCallback to know which options are automatically fixed.
    let saveOptions = new aw.Saving.PdfSaveOptions();

      // Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
      // which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
      // Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
      // Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
      // which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
      // Set the "Compliance" property to "PdfCompliance.PdfUa1" to comply with the "PDF/UA-1" (ISO 14289-1) standard,
      // which aims to define represent electronic documents in PDF that allow the file to be accessible.
      // Set the "Compliance" property to "PdfCompliance.Pdf20" to comply with the "PDF 2.0" (ISO 32000-2) standard.
      // Set the "Compliance" property to "PdfCompliance.PdfA4" to comply with the "PDF/A-4" (ISO 19004:2020) standard,
      // which preserving document static visual appearance over time.
      // Set the "Compliance" property to "PdfCompliance.PdfA4Ua2" to comply with both PDF/A-4 (ISO 19005-4:2020)
      // and PDF/UA-2 (ISO 14289-2:2024) standards.
      // Set the "Compliance" property to "PdfCompliance.PdfUa2" to comply with the PDF/UA-2 (ISO 14289-2:2024) standard.
      // This helps with making documents searchable but may significantly increase the size of already large documents.
    saveOptions.compliance = pdfCompliance;

    doc.save(base.artifactsDir + "PdfSaveOptions.compliance.pdf", saveOptions);
    //ExEnd
  });


  test.each([aw.Saving.PdfTextCompression.None,
    aw.Saving.PdfTextCompression.Flate])('TextCompression', (pdfTextCompression) => {
    //ExStart
    //ExFor:PdfSaveOptions
    //ExFor:PdfSaveOptions.textCompression
    //ExFor:PdfTextCompression
    //ExSummary:Shows how to apply text compression when saving a document to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    for (let i = 0; i < 100; i++)
      builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
              "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
    // compression to text when we save the document to PDF.
    // Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
    // to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
    options.textCompression = pdfTextCompression;

    doc.save(base.artifactsDir + "PdfSaveOptions.textCompression.pdf", options);
    //ExEnd

    var filePath = base.artifactsDir + "PdfSaveOptions.textCompression.pdf";
    var testedFileLength = fs.statSync(base.artifactsDir + "PdfSaveOptions.textCompression.pdf").size;

    switch (pdfTextCompression)
    {
      case aw.Saving.PdfTextCompression.None:
        expect(testedFileLength < 69000).toEqual(true);
        TestUtil.fileContainsString("<</Length 11 0 R>>stream", filePath);
        break;
      case aw.Saving.PdfTextCompression.Flate:
        expect(testedFileLength < 27000).toEqual(true);
        TestUtil.fileContainsString("<</Length 11 0 R/Filter/FlateDecode>>stream", filePath);
        break;
    }
  });


  test.each([aw.Saving.PdfImageCompression.Auto,
    aw.Saving.PdfImageCompression.Jpeg])('ImageCompression', (pdfImageCompression) => {
    //ExStart
    //ExFor:PdfSaveOptions.imageCompression
    //ExFor:PdfSaveOptions.jpegQuality
    //ExFor:PdfImageCompression
    //ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Jpeg image:");
    builder.insertImage(base.imageDir + "Logo.jpg");
    builder.insertParagraph();
    builder.writeln("Png image:");
    builder.insertImage(base.imageDir + "Transparent background logo.png");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let pdfSaveOptions = new aw.Saving.PdfSaveOptions();
    // Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
    // "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
    // Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
    // "ImageCompression" property to control the quality of all images that end up in the output PDF.
    pdfSaveOptions.imageCompression = pdfImageCompression;
    // Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
    pdfSaveOptions.jpegQuality = 10;

    doc.save(base.artifactsDir + "PdfSaveOptions.imageCompression.pdf", pdfSaveOptions);
    //ExEnd
  });


  test.each([aw.Saving.PdfImageColorSpaceExportMode.Auto,
    aw.Saving.PdfImageColorSpaceExportMode.SimpleCmyk])('ImageColorSpaceExportMode', (pdfImageColorSpaceExportMode) => {
    //ExStart
    //ExFor:PdfImageColorSpaceExportMode
    //ExFor:PdfSaveOptions.imageColorSpaceExportMode
    //ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Jpeg image:");
    builder.insertImage(base.imageDir + "Logo.jpg");
    builder.insertParagraph();
    builder.writeln("Png image:");
    builder.insertImage(base.imageDir + "Transparent background logo.png");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let pdfSaveOptions = new aw.Saving.PdfSaveOptions();

    // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.words to
    // automatically select the color space for images in the document that it converts to PDF.
    // In most cases, the color space will be RGB.
    // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk"
    // to use the CMYK color space for all images in the saved PDF.
    // Aspose.words will also apply Flate compression to all images and ignore the "ImageCompression" property's value.
    pdfSaveOptions.imageColorSpaceExportMode = pdfImageColorSpaceExportMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.imageColorSpaceExportMode.pdf", pdfSaveOptions);
    //ExEnd
  });


  test('DownsampleOptions', () => {
    //ExStart
    //ExFor:DownsampleOptions
    //ExFor:DownsampleOptions.downsampleImages
    //ExFor:DownsampleOptions.resolution
    //ExFor:DownsampleOptions.resolutionThreshold
    //ExFor:PdfSaveOptions.downsampleOptions
    //ExSummary:Shows how to change the resolution of images in the PDF document.
    let doc = new aw.Document(base.myDir + "Images.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // By default, Aspose.words downsample all images in a document that we save to PDF to 220 ppi.
    expect(options.downsampleOptions.downsampleImages).toEqual(true);
    expect(options.downsampleOptions.resolution).toEqual(220);
    expect(options.downsampleOptions.resolutionThreshold).toEqual(0);

    doc.save(base.artifactsDir + "PdfSaveOptions.downsampleOptions.default.pdf", options);

    // Set the "Resolution" property to "36" to downsample all images to 36 ppi.
    options.downsampleOptions.resolution = 36;

    // Set the "ResolutionThreshold" property to only apply the downsampling to
    // images with a resolution that is above 128 ppi.
    options.downsampleOptions.resolutionThreshold = 128;

    // Only the first two images from the document will be downsampled at this stage.
    doc.save(base.artifactsDir + "PdfSaveOptions.downsampleOptions.LowerResolution.pdf", options);
    //ExEnd
  });


  test.each([aw.Saving.ColorMode.Grayscale,
    aw.Saving.ColorMode.Normal])('ColorRendering', (colorMode) => {
    //ExStart
    //ExFor:PdfSaveOptions
    //ExFor:ColorMode
    //ExFor:FixedPageSaveOptions.colorMode
    //ExSummary:Shows how to change image color with saving options property.
    let doc = new aw.Document(base.myDir + "Images.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    // Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
    // The size of the output document may be larger with this setting.
    // Set the "ColorMode" property to "Normal" to render all images in color.
    let pdfSaveOptions = new aw.Saving.PdfSaveOptions();
    pdfSaveOptions.colorMode = colorMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
    //ExEnd
  });


  test.each([false,
    true])('DocTitle', (displayDocTitle) => {
    //ExStart
    //ExFor:PdfSaveOptions.displayDocTitle
    //ExSummary:Shows how to display the title of the document as the title bar.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    doc.builtInDocumentProperties.title = "Windows bar pdf title";

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    // Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
    // to display the value of the document's "Title" built-in property in the tab that belongs to this document.
    // Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
    let pdfSaveOptions = new aw.Saving.PdfSaveOptions();
    pdfSaveOptions.displayDocTitle = displayDocTitle;;

    doc.save(base.artifactsDir + "PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
    //ExEnd
  });


  test.each([false,
    true])('MemoryOptimization', (memoryOptimization) => {
    //ExStart
    //ExFor:SaveOptions.createSaveOptions(SaveFormat)
    //ExFor:SaveOptions.memoryOptimization
    //ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = aw.Saving.SaveOptions.createSaveOptions(aw.SaveFormat.Pdf);

    // Set the "MemoryOptimization" property to "true" to lower the memory footprint of large documents' saving operations
    // at the cost of increasing the duration of the operation.
    // Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
    saveOptions.memoryOptimization = memoryOptimization;

    doc.save(base.artifactsDir + "PdfSaveOptions.memoryOptimization.pdf", saveOptions);
    //ExEnd
  });


  test.each([ ["https://www.google.com/search?q= aspose", "https://www.google.com/search?q=%20aspose"],
    ["https://www.google.com/search?q=%20aspose", "https://www.google.com/search?q=%20aspose"] ])('EscapeUri', (uri, result) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.insertHyperlink("Testlink", uri, false);

    doc.save(base.artifactsDir + "PdfSaveOptions.EscapedUri.pdf");
  });


  test.each([false,
    true])('OpenHyperlinksInNewWindow', (openHyperlinksInNewWindow) => {
    //ExStart
    //ExFor:PdfSaveOptions.openHyperlinksInNewWindow
    //ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.insertHyperlink("Testlink", "https://www.google.com/search?q=%20aspose", false);

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "OpenHyperlinksInNewWindow" property to "true" to save all hyperlinks using Javascript code
    // that forces readers to open these links in new windows/browser tabs.
    // Set the "OpenHyperlinksInNewWindow" property to "false" to save all hyperlinks normally.
    options.openHyperlinksInNewWindow = openHyperlinksInNewWindow;

    doc.save(base.artifactsDir + "PdfSaveOptions.openHyperlinksInNewWindow.pdf", options);
    //ExEnd

    if (openHyperlinksInNewWindow)
      TestUtil.fileContainsString(
        "<</Type/Annot/Subtype/Link/Rect[70.84999847 707.35101318 110.17799377 721.15002441]/BS" +
        "<</Type/Border/S/S/W 0>>/A<</Type/Action/S/JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
        base.artifactsDir + "PdfSaveOptions.openHyperlinksInNewWindow.pdf");
    else
      TestUtil.fileContainsString(
        "<</Type/Annot/Subtype/Link/Rect[70.84999847 707.35101318 110.17799377 721.15002441]/BS" +
        "<</Type/Border/S/S/W 0>>/A<</Type/Action/S/URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
        base.artifactsDir + "PdfSaveOptions.openHyperlinksInNewWindow.pdf");
  });


  //ExStart
  //ExFor:MetafileRenderingMode
  //ExFor:MetafileRenderingOptions
  //ExFor:MetafileRenderingOptions.EmulateRasterOperations
  //ExFor:MetafileRenderingOptions.RenderingMode
  //ExFor:IWarningCallback
  //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
  //ExSummary:Shows added a fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
  test.skip('HandleBinaryRasterWarnings - TODO: WORDSNODEJS-108 - Add support of IWarningCallback', () => {
    let doc = new aw.Document(base.myDir + "WMF with image.docx");

    let metafileRenderingOptions = new aw.Saving.MetafileRenderingOptions();

    // Set the "EmulateRasterOperations" property to "false" to fall back to bitmap when
    // it encounters a metafile, which will require raster operations to render in the output PDF.
    metafileRenderingOptions.emulateRasterOperations = false;

    // Set the "RenderingMode" property to "VectorWithFallback" to try to render every metafile using vector graphics.
    metafileRenderingOptions.renderingMode = aw.Saving.MetafileRenderingMode.VectorWithFallback;

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF and applies the configuration
    // in our MetafileRenderingOptions object to the saving operation.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.metafileRenderingOptions = metafileRenderingOptions;

    let callback = new HandleDocumentWarnings();
    doc.warningCallback = callback;

    doc.save(base.artifactsDir + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

    expect(callback.Warnings.count).toEqual(1);
    expect(callback.Warnings.at(0).description).toEqual("'R2_XORPEN' binary raster operation is not supported.");
  });

/*
    /// <summary>
    /// Prints and collects formatting loss-related warnings that occur upon saving a document.
    /// </summary>
  public class HandleDocumentWarnings : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
      if (info.warningType == aw.WarningType.MinorFormattingLoss)
      {
        console.log("Unsupported operation: " + info.description);
        Warnings.warning(info);
      }
    }

    public WarningInfoCollection Warnings = new aw.WarningInfoCollection();
  }
    //ExEnd
*/

  test.each([aw.Saving.HeaderFooterBookmarksExportMode.None,
    aw.Saving.HeaderFooterBookmarksExportMode.First,
    aw.Saving.HeaderFooterBookmarksExportMode.All])('HeaderFooterBookmarksExportMode', (headerFooterBookmarksExportMode) => {
    //ExStart
    //ExFor:HeaderFooterBookmarksExportMode
    //ExFor:OutlineOptions
    //ExFor:OutlineOptions.defaultBookmarksOutlineLevel
    //ExFor:PdfSaveOptions.headerFooterBookmarksExportMode
    //ExFor:PdfSaveOptions.pageMode
    //ExFor:PdfPageMode
    //ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
    let doc = new aw.Document(base.myDir + "Bookmarks in headers and footers.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();

    // Set the "PageMode" property to "PdfPageMode.UseOutlines" to display the outline navigation pane in the output PDF.
    saveOptions.pageMode = aw.Saving.PdfPageMode.UseOutlines;

    // Set the "DefaultBookmarksOutlineLevel" property to "1" to display all
    // bookmarks at the first level of the outline in the output PDF.
    saveOptions.outlineOptions.defaultBookmarksOutlineLevel = 1;

    // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.None" to
    // not export any bookmarks that are inside headers/footers.
    // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.First" to
    // only export bookmarks in the first section's header/footers.
    // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.All" to
    // export bookmarks that are in all headers/footers.
    saveOptions.headerFooterBookmarksExportMode = headerFooterBookmarksExportMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.headerFooterBookmarksExportMode.pdf", saveOptions);
    //ExEnd
  });


  test.skip('UnsupportedImageFormatWarning - TODO: WORDSNODEJS-108 - Add support of IWarningCallback', () => {
    let doc = new aw.Document(base.myDir + "Corrupted image.docx");

    let saveWarningCallback = new SaveWarningCallback();
    doc.warningCallback = saveWarningCallback;

    doc.save(base.artifactsDir + "PdfSaveOption.UnsupportedImageFormatWarning.pdf", aw.SaveFormat.Pdf);

    expect(saveWarningCallback.SaveWarnings.at(0).description).toEqual("Image can not be processed. Possibly unsupported image format.");
  });


  /*
  public class SaveWarningCallback : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
      if (info.warningType == aw.WarningType.MinorFormattingLoss)
      {
        console.log(`${info.warningType}: ${info.description}.`);
        SaveWarnings.warning(info);
      }
    }

    internal WarningInfoCollection SaveWarnings = new aw.WarningInfoCollection();
  }
*/


  test.each([false,
    true])('EmulateRenderingToSizeOnPage', (renderToSize) => {
    //ExStart
    //ExFor:MetafileRenderingOptions.emulateRenderingToSizeOnPage
    //ExFor:MetafileRenderingOptions.emulateRenderingToSizeOnPageResolution
    //ExSummary:Shows how to display of the metafile according to the size on page.
    let doc = new aw.Document(base.myDir + "WMF with text.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();


    // Set the "EmulateRenderingToSizeOnPage" property to "true"
    // to emulate rendering according to the metafile size on page.
    // Set the "EmulateRenderingToSizeOnPage" property to "false"
    // to emulate metafile rendering to its default size in pixels.
    saveOptions.metafileRenderingOptions.emulateRenderingToSizeOnPage = renderToSize;
    saveOptions.metafileRenderingOptions.emulateRenderingToSizeOnPageResolution = 50;

    doc.save(base.artifactsDir + "PdfSaveOptions.emulateRenderingToSizeOnPage.pdf", saveOptions);
    //ExEnd
  });


  test.skip.each([false,
    true])('EmbedFullFonts - TODO: WORDSNODEJS-110 - Add marshalling of IList<T> results to nodewgen.', (embedFullFonts) => {
    //ExStart
    //ExFor:PdfSaveOptions.#ctor
    //ExFor:PdfSaveOptions.embedFullFonts
    //ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Arvo";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    // Configure our font sources to ensure that we have access to both the fonts in this document.
    let originalFontsSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();
    let folderFontSource =  new aw.Fonts.FolderFontSource(base.fontsDir, true);
    aw.Fonts.FontSettings.defaultInstance.setFontsSources([ originalFontsSources.at(0), folderFontSource ]);

    let fontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();
    expect(fontSources[0].getAvailableFonts().some(f => f.fullFontName == "Arial")).toEqual(true);
    expect(fontSources[1].getAvailableFonts().some(f => f.fullFontName == "Arvo")).toEqual(true);

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Since our document contains a custom font, embedding in the output document may be desirable.
    // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
    // The document's size may become very large, but we will have full use of all fonts if we edit the PDF.
    // Set the "EmbedFullFonts" property to "false" to apply subsetting to fonts, saving only the glyphs
    // that the document is using. The file will be considerably smaller,
    // but we may need access to any custom fonts if we edit the document.
    options.embedFullFonts = embedFullFonts;

    doc.save(base.artifactsDir + "PdfSaveOptions.embedFullFonts.pdf", options);

    // Restore the original font sources.
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontsSources);
    //ExEnd

    var testedFileLength = fs.statSync(base.artifactsDir + "PdfSaveOptions.embedFullFonts.pdf").size;
    if (embedFullFonts)
      expect(testedFileLength < 571000).toEqual(true);
    else
      expect(testedFileLength < 24000).toEqual(true);
  });



  test.each([aw.Saving.PdfFontEmbeddingMode.EmbedAll,
    aw.Saving.PdfFontEmbeddingMode.EmbedNone,
    aw.Saving.PdfFontEmbeddingMode.EmbedNonstandard])('EmbedWindowsFonts', (pdfFontEmbeddingMode) => {
    //ExStart
    //ExFor:PdfSaveOptions.fontEmbeddingMode
    //ExFor:PdfFontEmbeddingMode
    //ExSummary:Shows how to set Aspose.words to skip embedding Arial and Times New Roman fonts into a PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // "Arial" is a standard font, and "Courier New" is a nonstandard font.
    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Courier New";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();
    // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
    options.embedFullFonts = true;
    // Set the "FontEmbeddingMode" property to "EmbedAll" to embed all fonts in the output PDF.
    // Set the "FontEmbeddingMode" property to "EmbedNonstandard" to only allow nonstandard fonts' embedding in the output PDF.
    // Set the "FontEmbeddingMode" property to "EmbedNone" to not embed any fonts in the output PDF.
    options.fontEmbeddingMode = pdfFontEmbeddingMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf", options);
    //ExEnd

    var testedFileLength = fs.statSync(base.artifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf").size;
    switch (pdfFontEmbeddingMode)
    {
      case aw.Saving.PdfFontEmbeddingMode.EmbedAll:
        expect(testedFileLength < 1040000).toEqual(true);
        break;
      case aw.Saving.PdfFontEmbeddingMode.EmbedNonstandard:
        expect(testedFileLength < 492000).toEqual(true);
        break;
      case aw.Saving.PdfFontEmbeddingMode.EmbedNone:
        expect(testedFileLength < 4300).toEqual(true);
        break;
    }
  });


  test.each([false,
    true])('EmbedCoreFonts', (useCoreFonts) => {
    //ExStart
    //ExFor:PdfSaveOptions.useCoreFonts
    //ExSummary:Shows how enable/disable PDF Type 1 font substitution.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Courier New";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();
    // Set the "UseCoreFonts" property to "true" to replace some fonts,
    // including the two fonts in our document, with their PDF Type 1 equivalents.
    // Set the "UseCoreFonts" property to "false" to not apply PDF Type 1 fonts.
    options.useCoreFonts = useCoreFonts;

    doc.save(base.artifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf", options);
    //ExEnd

    var testedFileLength = fs.statSync(base.artifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf").size;
    if (useCoreFonts)
      expect(testedFileLength < 2000).toEqual(true);
    else
      expect(testedFileLength < 33500).toEqual(true);
  });


  test.each([false,
    true])('AdditionalTextPositioning', (applyAdditionalTextPositioning) => {
    //ExStart
    //ExFor:PdfSaveOptions.additionalTextPositioning
    //ExSummary:Show how to write additional text positioning operators.
    let doc = new aw.Document(base.myDir + "Text positioning operators.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.textCompression = aw.Saving.PdfTextCompression.None;
    // Set the "AdditionalTextPositioning" property to "true" to attempt to fix incorrect
    // element positioning in the output PDF, should there be any, at the cost of increased file size.
    // Set the "AdditionalTextPositioning" property to "false" to render the document as usual.
    saveOptions.additionalTextPositioning = applyAdditionalTextPositioning;

    doc.save(base.artifactsDir + "PdfSaveOptions.additionalTextPositioning.pdf", saveOptions);
    //ExEnd
  });


  test.each([false, 
    true])('SaveAsPdfBookFold', (renderTextAsBookfold) => {
    //ExStart
    //ExFor:PdfSaveOptions.useBookFoldPrintingSettings
    //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
    // in the output PDF in a way that helps us use it to make a booklet.
    // Set the "UseBookFoldPrintingSettings" property to "false" to render the PDF normally.
    options.useBookFoldPrintingSettings = renderTextAsBookfold;

    // If we are rendering the document as a booklet, we must set the "MultiplePages"
    // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
    if (renderTextAsBookfold)
      for (let s of doc.sections.toArray())
      {
        s.pageSetup.multiplePages = aw.Settings.MultiplePagesType.BookFoldPrinting;
      }

    // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
    // and the contents will line up in a way that creates a booklet.
    doc.save(base.artifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
    //ExEnd
  });


  test('ZoomBehaviour', () => {
    //ExStart
    //ExFor:PdfSaveOptions.zoomBehavior
    //ExFor:PdfSaveOptions.zoomFactor
    //ExFor:PdfZoomBehavior
    //ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    // Set the "ZoomBehavior" property to "PdfZoomBehavior.ZoomFactor" to get a PDF reader to
    // apply a percentage-based zoom factor when we open the document with it.
    // Set the "ZoomFactor" property to "25" to give the zoom factor a value of 25%.
    let options = new aw.Saving.PdfSaveOptions();
    options.zoomBehavior = aw.Saving.PdfZoomBehavior.ZoomFactor;
    options.zoomFactor = 25;

    // When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
    doc.save(base.artifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf", options);
    //ExEnd
  });


  test.each([aw.Saving.PdfPageMode.FullScreen,
    aw.Saving.PdfPageMode.UseThumbs,
    aw.Saving.PdfPageMode.UseOC,
    aw.Saving.PdfPageMode.UseOutlines,
    aw.Saving.PdfPageMode.UseNone,
    aw.Saving.PdfPageMode.UseAttachments])('PageMode', (pageMode) => {
    //ExStart
    //ExFor:PdfSaveOptions.pageMode
    //ExFor:PdfPageMode
    //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "PageMode" property to "PdfPageMode.FullScreen" to get the PDF reader to open the saved
    // document in full-screen mode, which takes over the monitor's display and has no controls visible.
    // Set the "PageMode" property to "PdfPageMode.UseThumbs" to get the PDF reader to display a separate panel
    // with a thumbnail for each page in the document.
    // Set the "PageMode" property to "PdfPageMode.UseOC" to get the PDF reader to display a separate panel
    // that allows us to work with any layers present in the document.
    // Set the "PageMode" property to "PdfPageMode.UseOutlines" to get the PDF reader
    // also to display the outline, if possible.
    // Set the "PageMode" property to "PdfPageMode.UseNone" to get the PDF reader to display just the document itself.
    // Set the "PageMode" property to "PdfPageMode.UseAttachments" to make visible attachments panel.
    options.pageMode = pageMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.pageMode.pdf", options);
    //ExEnd

    const docLocaleName = "en-US";

    switch (pageMode)
    {
      case aw.Saving.PdfPageMode.FullScreen:
        TestUtil.fileContainsString(
          `<</Type/Catalog/Pages 3 0 R/PageMode/FullScreen/Lang(${docLocaleName})/Metadata 4 0 R>>`,
          base.artifactsDir + "PdfSaveOptions.pageMode.pdf");
        break;
      case aw.Saving.PdfPageMode.UseThumbs:
        TestUtil.fileContainsString(
          `<</Type/Catalog/Pages 3 0 R/PageMode/UseThumbs/Lang(${docLocaleName})/Metadata 4 0 R>>`,
          base.artifactsDir + "PdfSaveOptions.pageMode.pdf");
        break;
      case aw.Saving.PdfPageMode.UseOC:
        TestUtil.fileContainsString(
          `<</Type/Catalog/Pages 3 0 R/PageMode/UseOC/Lang(${docLocaleName})/Metadata 4 0 R>>`,
          base.artifactsDir + "PdfSaveOptions.pageMode.pdf");
        break;
      case aw.Saving.PdfPageMode.UseOutlines:
      case aw.Saving.PdfPageMode.UseNone:
        TestUtil.fileContainsString(`<</Type/Catalog/Pages 3 0 R/Lang(${docLocaleName})/Metadata 4 0 R>>`,
          base.artifactsDir + "PdfSaveOptions.pageMode.pdf");
        break;
      case aw.Saving.PdfPageMode.UseAttachments:
        TestUtil.fileContainsString(
          `<</Type/Catalog/Pages 3 0 R/PageMode/UseAttachments/Lang(${docLocaleName})/Metadata 4 0 R>>`,
          base.artifactsDir + "PdfSaveOptions.pageMode.pdf");
        break;
    }
  });


  test('NoteHyperlinks', () => {
    //ExStart
    //ExFor:PdfSaveOptions.createNoteHyperlinks
    //ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
    let doc = new aw.Document(base.myDir + "Footnotes and endnotes.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "CreateNoteHyperlinks" property to "true" to turn all footnote/endnote symbols
    // in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
    // Set the "CreateNoteHyperlinks" property to "false" not to have footnote/endnote symbols link to anything.
    options.createNoteHyperlinks = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf", options);
    //ExEnd

    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[157.80099487 720.90106201 159.35600281 733.55004883]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 677 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[202.16900635 720.90106201 206.06201172 733.55004883]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 79 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[212.23199463 699.2510376 215.34199524 711.90002441]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 654 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[258.15499878 699.2510376 262.04800415 711.90002441]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 85 68 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[85.05000305 68.19904327 88.66500092 79.69804382]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 202 733 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[85.05000305 56.70004272 88.66500092 68.19904327]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 258 711 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[85.05000305 666.10205078 86.4940033 677.60107422]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 157 733 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
    TestUtil.fileContainsString(
      "<</Type/Annot/Subtype/Link/Rect[85.05000305 643.10406494 87.93800354 654.60308838]/BS<</Type/Border/S/S/W 0>>/Dest[5 0 R /XYZ 212 711 0]>>",
      base.artifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
  });


  test.each([aw.Saving.PdfCustomPropertiesExport.None,
    aw.Saving.PdfCustomPropertiesExport.Standard,
    aw.Saving.PdfCustomPropertiesExport.Metadata])('CustomPropertiesExport(%o)', (pdfCustomPropertiesExportMode) => {
    //ExStart
    //ExFor:PdfCustomPropertiesExport
    //ExFor:PdfSaveOptions.customPropertiesExport
    //ExSummary:Shows how to export custom properties while converting a document to PDF.
    let doc = new aw.Document();

    doc.customDocumentProperties.add("Company", "My value");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.None" to discard
    // custom document properties as we save the document to .PDF.
    // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Standard"
    // to preserve custom properties within the output PDF document.
    // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Metadata"
    // to preserve custom properties in an XMP packet.
    options.customPropertiesExport = pdfCustomPropertiesExportMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf", options);
    //ExEnd

    switch (pdfCustomPropertiesExportMode)
    {
      case aw.Saving.PdfCustomPropertiesExport.None:
        // TestUtil.fileNotContainString(
        //   doc.customDocumentProperties.at(0).name,
        //   base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf");
        // TestUtil.fileNotContainString(
        //   "<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
        //   base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf");
        break;
      case aw.Saving.PdfCustomPropertiesExport.Standard:
        TestUtil.fileContainsString(
          "<</Creator(",
          base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf");
        TestUtil.fileContainsString("/Company(",
          base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf");
        break;
      case aw.Saving.PdfCustomPropertiesExport.Metadata:
        TestUtil.fileContainsString("<</Type/Metadata/Subtype/XML/Length 8 0 R/Filter/FlateDecode>>",
          base.artifactsDir + "PdfSaveOptions.customPropertiesExport.pdf");
        break;
    }
  });


  test.each([aw.Saving.DmlEffectsRenderingMode.None,
    aw.Saving.DmlEffectsRenderingMode.Simplified,
    aw.Saving.DmlEffectsRenderingMode.Fine])('DrawingMLEffects', (effectsRenderingMode) => {
    //ExStart
    //ExFor:DmlRenderingMode
    //ExFor:DmlEffectsRenderingMode
    //ExFor:PdfSaveOptions.dmlEffectsRenderingMode
    //ExFor:SaveOptions.dmlEffectsRenderingMode
    //ExFor:SaveOptions.dmlRenderingMode
    //ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
    let doc = new aw.Document(base.myDir + "DrawingML shape effects.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.None" to discard all DrawingML effects.
    // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Simplified"
    // to render a simplified version of DrawingML effects.
    // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Fine" to
    // render DrawingML effects with more accuracy and also with more processing cost.
    options.dmlEffectsRenderingMode = effectsRenderingMode;

    expect(options.dmlRenderingMode).toEqual(aw.Saving.DmlRenderingMode.DrawingML);

    doc.save(base.artifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf", options);
    //ExEnd
  });


  test.each([aw.Saving.DmlRenderingMode.Fallback,
    aw.Saving.DmlRenderingMode.DrawingML])('DrawingMLFallback', (dmlRenderingMode) => {
    //ExStart
    //ExFor:DmlRenderingMode
    //ExFor:SaveOptions.dmlRenderingMode
    //ExSummary:Shows how to render fallback shapes when saving to PDF.
    let doc = new aw.Document(base.myDir + "DrawingML shape fallbacks.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "DmlRenderingMode" property to "DmlRenderingMode.Fallback"
    // to substitute DML shapes with their fallback shapes.
    // Set the "DmlRenderingMode" property to "DmlRenderingMode.DrawingML"
    // to render the DML shapes themselves.
    options.dmlRenderingMode = dmlRenderingMode;

    doc.save(base.artifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf", options);
    //ExEnd

    switch (dmlRenderingMode)
    {
      case aw.Saving.DmlRenderingMode.DrawingML:
        TestUtil.fileContainsString(
          "<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
          base.artifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
        break;
      case aw.Saving.DmlRenderingMode.Fallback:
        TestUtil.fileContainsString(
          "<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABE 14 0 R>>/ExtGState<</GS1 11 0 R/GS2 12 0 R/GS3 17 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
          base.artifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
        break;
    }
  });


  test.each([false,
    true])('ExportDocumentStructure', (exportDocumentStructure) => {
    //ExStart
    //ExFor:PdfSaveOptions.exportDocumentStructure
    //ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.paragraphFormat.style = doc.styles.at("Normal");
    builder.write(
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();
    // Set the "ExportDocumentStructure" property to "true" to make the document structure, such tags, available via the
    // "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
    // Set the "ExportDocumentStructure" property to "false" to not export the document structure.
    options.exportDocumentStructure = exportDocumentStructure;

    // Suppose we export document structure while saving this document. In that case,
    // we can open it using Adobe Acrobat and find tags for elements such as the heading
    // and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
    doc.save(base.artifactsDir + "PdfSaveOptions.exportDocumentStructure.pdf", options);
    //ExEnd

    if (exportDocumentStructure)
    {
      TestUtil.fileContainsString("<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABD 13 0 R>>/ExtGState<</GS1 11 0 R/GS2 16 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs/S>>",
        base.artifactsDir + "PdfSaveOptions.exportDocumentStructure.pdf");
    }
    else
    {
      TestUtil.fileContainsString("<</Type/Page/Parent 3 0 R/Contents 6 0 R/MediaBox[0 0 612 792]/Resources<</Font<</FAAAAI 8 0 R/FAAABC 12 0 R>>>>/Group<</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
        base.artifactsDir + "PdfSaveOptions.exportDocumentStructure.pdf");
    }
  });


  test.each([false,
    true])('InterpolateImages', (interpolateImages) => {
    //ExStart
    //ExFor:PdfSaveOptions.interpolateImages
    //ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imageDir + "Transparent background logo.png");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    // Set the "InterpolateImages" property to "true" to get the reader that opens this document to interpolate images.
    // Their resolution should be lower than that of the device that is displaying the document.
    // Set the "InterpolateImages" property to "false" to make it so that the reader does not apply any interpolation.
    saveOptions.interpolateImages = interpolateImages;

    // When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
    // to see the interpolation effect if we saved the document with it enabled.
    doc.save(base.artifactsDir + "PdfSaveOptions.interpolateImages.pdf", saveOptions);
    //ExEnd

    if (interpolateImages)
    {
      TestUtil.fileContainsString("<</Type/XObject/Subtype/Image/Width 400/Height 400/ColorSpace/DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Interpolate true/Length 11 0 R/Filter/FlateDecode>>",
        base.artifactsDir + "PdfSaveOptions.interpolateImages.pdf");
    }
    else
    {
      TestUtil.fileContainsString("<</Type/XObject/Subtype/Image/Width 400/Height 400/ColorSpace/DeviceRGB/BitsPerComponent 8/SMask 10 0 R/Length 11 0 R/Filter/FlateDecode>>",
        base.artifactsDir + "PdfSaveOptions.interpolateImages.pdf");
    }
  });


  test.skip('Dml3DEffectsRenderingModeTest - TODO: WORDSNODEJS-108 - Add support of IWarningCallback', () => {
    let doc = new aw.Document(base.myDir + "DrawingML shape 3D effects.docx");

    let warningCallback = new RenderCallback();
    doc.warningCallback = warningCallback;

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.dml3DEffectsRenderingMode = aw.Saving.Dml3DEffectsRenderingMode.Advanced;

    doc.save(base.artifactsDir + "PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

    expect(warningCallback.count).toEqual(48);
  });

/*
  public class RenderCallback : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
      console.log(`${info.warningType}: ${info.description}.`);
      mWarnings.add(info);
    }

    public WarningInfo this.at(int i) => mWarnings.at(i);

      /// <summary>
      /// Clears warning collection.
      /// </summary>
    public void Clear()
    {
      mWarnings.clear();
    }

    public int Count => mWarnings.count;

      /// <summary>
      /// Returns true if a warning with the specified properties has been generated.
      /// </summary>
    public bool Contains(WarningSource source, WarningType type, string description)
    {
      return mWarnings.any(warning =>
        warning.source == source && warning.warningType == type && warning.description == description);
    }

    private readonly List<WarningInfo> mWarnings = new aw.Lists.List<WarningInfo>();
  }*/

    
  test('PdfDigitalSignature', () => {
    //ExStart
    //ExFor:PdfDigitalSignatureDetails
    //ExFor:PdfDigitalSignatureDetails.#ctor
    //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
    //ExFor:PdfDigitalSignatureDetails.hashAlgorithm
    //ExFor:PdfDigitalSignatureDetails.location
    //ExFor:PdfDigitalSignatureDetails.reason
    //ExFor:PdfDigitalSignatureDetails.signatureDate
    //ExFor:PdfDigitalSignatureHashAlgorithm
    //ExFor:PdfSaveOptions.digitalSignatureDetails
    //ExFor:PdfDigitalSignatureDetails.certificateHolder
    //ExSummary:Shows how to sign a generated PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Contents of signed PDF.");

    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Configure the "DigitalSignatureDetails" object of the "SaveOptions" object to
    // digitally sign the document as we render it with the "Save" method.
    let signingTime = new Date(2015, 7, 20);
    options.digitalSignatureDetails =
      new aw.Saving.PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "My Office", signingTime);
    options.digitalSignatureDetails.hashAlgorithm = aw.Saving.PdfDigitalSignatureHashAlgorithm.RipeMD160;

    expect(options.digitalSignatureDetails.reason).toEqual("Test Signing");
    expect(options.digitalSignatureDetails.location).toEqual("My Office");
    //expect(options.digitalSignatureDetails.signatureDate).toEqual(signingTime);

    doc.save(base.artifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
    //ExEnd

    TestUtil.fileContainsString("<</Type/Annot/Subtype/Widget/Rect[0 0 0 0]/FT/Sig/T",
      base.artifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf");

    expect(aw.FileFormatUtil.detectFileFormat(
      base.artifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf").hasDigitalSignature).toBe(false);
  });


  test('PdfDigitalSignatureTimestamp', () => {
    //ExStart
    //ExFor:PdfDigitalSignatureDetails.timestampSettings
    //ExFor:PdfDigitalSignatureTimestampSettings
    //ExFor:PdfDigitalSignatureTimestampSettings.#ctor
    //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
    //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
    //ExFor:PdfDigitalSignatureTimestampSettings.password
    //ExFor:PdfDigitalSignatureTimestampSettings.serverUrl
    //ExFor:PdfDigitalSignatureTimestampSettings.timeout
    //ExFor:PdfDigitalSignatureTimestampSettings.userName
    //ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Signed PDF contents.");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF.
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");
    options.digitalSignatureDetails = new aw.Saving.PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", Date.now());

    // Create a timestamp authority-verified timestamp.
    options.digitalSignatureDetails.timestampSettings =
      new aw.Saving.PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword");

    // The default lifespan of the timestamp is 100 seconds.
    expect(options.digitalSignatureDetails.timestampSettings.timeout.totalSeconds).toEqual(100.0);

    const timeout = {
      get minutes() { return 30; }
    }

    // We can set our timeout period via the constructor.
    options.digitalSignatureDetails.timestampSettings =
      new aw.Saving.PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", timeout);

    expect(options.digitalSignatureDetails.timestampSettings.timeout.totalSeconds).toEqual(1800);
    expect(options.digitalSignatureDetails.timestampSettings.serverUrl).toEqual("https://freetsa.org/tsr");
    expect(options.digitalSignatureDetails.timestampSettings.userName).toEqual("JohnDoe");
    expect(options.digitalSignatureDetails.timestampSettings.password).toEqual("MyPassword");

    // The "Save" method will apply our signature to the output document at this time.
    doc.save(base.artifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
    //ExEnd

    expect(aw.FileFormatUtil.detectFileFormat(base.artifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").hasDigitalSignature).toEqual(false);
    TestUtil.fileContainsString("<</Type/Annot/Subtype/Widget/Rect[0 0 0 0]/FT/Sig/T",
      base.artifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");
  });


  test.each([aw.Saving.EmfPlusDualRenderingMode.Emf,
    aw.Saving.EmfPlusDualRenderingMode.EmfPlus,
    aw.Saving.EmfPlusDualRenderingMode.EmfPlusWithFallback])('RenderMetafile', (renderingMode) => {
    //ExStart
    //ExFor:EmfPlusDualRenderingMode
    //ExFor:MetafileRenderingOptions.emfPlusDualRenderingMode
    //ExFor:MetafileRenderingOptions.useEmfEmbeddedToWmf
    //ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
    let doc = new aw.Document(base.myDir + "EMF.docx");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();

    // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.Emf"
    // to only render the EMF part of an EMF+ dual metafile.
    // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlus" to
    // to render the EMF+ part of an EMF+ dual metafile.
    // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlusWithFallback"
    // to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
    // Otherwise, Aspose.words will render the EMF part.
    saveOptions.metafileRenderingOptions.emfPlusDualRenderingMode = renderingMode;

    // Set the "UseEmfEmbeddedToWmf" property to "true" to render embedded EMF data
    // for metafiles that we can render as vector graphics.
    saveOptions.metafileRenderingOptions.useEmfEmbeddedToWmf = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
    //ExEnd
  });


  test('EncryptionPermissions', () => {
    //ExStart
    //ExFor:PdfEncryptionDetails.#ctor(String,String,PdfPermissions)
    //ExFor:PdfSaveOptions.encryptionDetails
    //ExFor:PdfEncryptionDetails.permissions
    //ExFor:PdfEncryptionDetails.ownerPassword
    //ExFor:PdfEncryptionDetails.userPassword
    //ExFor:PdfPermissions
    //ExFor:PdfEncryptionDetails
    //ExSummary:Shows how to set permissions on a saved PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    // Extend permissions to allow the editing of annotations.
    let encryptionDetails =
      new aw.Saving.PdfEncryptionDetails("password", '', aw.Saving.PdfPermissions.ModifyAnnotations | aw.Saving.PdfPermissions.DocumentAssembly);

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    // Enable encryption via the "EncryptionDetails" property.
    saveOptions.encryptionDetails = encryptionDetails;

    // When we open this document, we will need to provide the password before accessing its contents.
    doc.save(base.artifactsDir + "PdfSaveOptions.EncryptionPermissions.pdf", saveOptions);
    //ExEnd
  });


  test.each([aw.Saving.NumeralFormat.ArabicIndic,
    aw.Saving.NumeralFormat.Context,
    aw.Saving.NumeralFormat.EasternArabicIndic,
    aw.Saving.NumeralFormat.European,
    aw.Saving.NumeralFormat.System])('SetNumeralFormat', (numeralFormat) => {
    //ExStart
    //ExFor:FixedPageSaveOptions.numeralFormat
    //ExFor:NumeralFormat
    //ExSummary:Shows how to set the numeral format used when saving to PDF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.localeId = 5121;//new CultureInfo("ar-AR").LCID;
    builder.writeln("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Set the "NumeralFormat" property to "NumeralFormat.ArabicIndic" to
    // use glyphs from the U+0660 to U+0669 range as numbers.
    // Set the "NumeralFormat" property to "NumeralFormat.Context" to
    // look up the locale to determine what number of glyphs to use.
    // Set the "NumeralFormat" property to "NumeralFormat.EasternArabicIndic" to
    // use glyphs from the U+06F0 to U+06F9 range as numbers.
    // Set the "NumeralFormat" property to "NumeralFormat.European" to use european numerals.
    // Set the "NumeralFormat" property to "NumeralFormat.System" to determine the symbol set from regional settings.
    options.numeralFormat = numeralFormat;

    doc.save(base.artifactsDir + "PdfSaveOptions.SetNumeralFormat.pdf", options);
    //ExEnd
  });


  test('ExportPageSet', () => {
    //ExStart
    //ExFor:FixedPageSaveOptions.pageSet
    //ExFor:PageSet.all
    //ExFor:PageSet.even
    //ExFor:PageSet.odd
    //ExSummary:Shows how to export Odd pages from the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    for (let i = 0; i < 5; i++)
    {
      builder.writeln(`Page ${i + 1} (${(i % 2 == 0 ? "odd" : "even")})`);
      if (i < 4)
        builder.insertBreak(aw.BreakType.PageBreak);
    }

    // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to .PDF.
    let options = new aw.Saving.PdfSaveOptions();

    // Below are three PageSet properties that we can use to filter out a set of pages from
    // our document to save in an output PDF document based on the parity of their page numbers.
    // 1 -  Save only the even-numbered pages:
    options.pageSet = aw.Saving.PageSet.even;

    doc.save(base.artifactsDir + "PdfSaveOptions.ExportPageSet.even.pdf", options);

    // 2 -  Save only the odd-numbered pages:
    options.pageSet = aw.Saving.PageSet.odd;

    doc.save(base.artifactsDir + "PdfSaveOptions.ExportPageSet.odd.pdf", options);

    // 3 -  Save every page:
    options.pageSet = aw.Saving.PageSet.all;

    doc.save(base.artifactsDir + "PdfSaveOptions.ExportPageSet.all.pdf", options);
    //ExEnd
  });


  test('ExportLanguageToSpanTag', () => {
    //ExStart
    //ExFor:PdfSaveOptions.exportLanguageToSpanTag
    //ExSummary:Shows how to create a "Span" tag in the document structure to export the text language.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    builder.writeln("Hola mundo!");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    // Note, when "ExportDocumentStructure" is false, "ExportLanguageToSpanTag" is ignored.
    saveOptions.exportDocumentStructure = true;
    saveOptions.exportLanguageToSpanTag = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.exportLanguageToSpanTag.pdf", saveOptions);
    //ExEnd
  });


  test('AttachmentsEmbeddingMode', () => {
    //ExStart:AttachmentsEmbeddingMode
    //GistId:1a265b92fa0019b26277ecfef3c20330
    //ExFor:PdfSaveOptions.attachmentsEmbeddingMode
    //ExSummary:Shows how to add embed attachments to the PDF document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertOleObject(base.myDir + "Spreadsheet.xlsx", "Excel.Sheet", false, true, null);

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.attachmentsEmbeddingMode = aw.Saving.PdfAttachmentsEmbeddingMode.Annotations;

    doc.save(base.artifactsDir + "PdfSaveOptions.PdfEmbedAttachments.pdf", saveOptions);
    //ExEnd:AttachmentsEmbeddingMode
  });


  test('CacheBackgroundGraphics', () => {
    //ExStart
    //ExFor:PdfSaveOptions.cacheBackgroundGraphics
    //ExSummary:Shows how to cache graphics placed in document's background.
    let doc = new aw.Document(base.myDir + "Background images.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.cacheBackgroundGraphics = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.cacheBackgroundGraphics.pdf", saveOptions);

    let asposeToPdfSize = fs.statSync(base.artifactsDir + "PdfSaveOptions.cacheBackgroundGraphics.pdf").size;
    let wordToPdfSize = fs.statSync(base.myDir + "Background images (word to pdf).pdf").size;

    expect(asposeToPdfSize <= wordToPdfSize).toBe(true);
    //ExEnd
  });


  test('ExportParagraphGraphicsToArtifact', () => {
    //ExStart
    //ExFor:PdfSaveOptions.exportParagraphGraphicsToArtifact
    //ExSummary:Shows how to export paragraph graphics as artifact (underlines, text emphasis, etc.).
    let doc = new aw.Document(base.myDir + "PDF artifacts.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.exportDocumentStructure = true;
    saveOptions.exportParagraphGraphicsToArtifact = true;
    saveOptions.textCompression = aw.Saving.PdfTextCompression.None;

    doc.save(base.artifactsDir + "PdfSaveOptions.exportParagraphGraphicsToArtifact.pdf", saveOptions);
    //ExEnd
  });


  test('PageLayout', () => {
    //ExStart:PageLayout
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:PdfSaveOptions.pageLayout
    //ExFor:PdfPageLayout
    //ExSummary:Shows how to display pages when opened in a PDF reader.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // Display the pages two at a time, with odd-numbered pages on the left.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.pageLayout = aw.Saving.PdfPageLayout.TwoPageLeft;

    doc.save(base.artifactsDir + "PdfSaveOptions.pageLayout.pdf", saveOptions);
    //ExEnd:PageLayout
  });


  test('SdtTagAsFormFieldName', () => {
    //ExStart:SdtTagAsFormFieldName
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:PdfSaveOptions.useSdtTagAsFormFieldName
    //ExSummary:Shows how to use SDT control Tag or Id property as a name of form field in PDF.
    let doc = new aw.Document(base.myDir + "Form fields.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.preserveFormFields = true;
    // When set to 'false', SDT control Id property is used as a name of form field in PDF.
    // When set to 'true', SDT control Tag property is used as a name of form field in PDF.
    saveOptions.useSdtTagAsFormFieldName = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.SdtTagAsFormFieldName.pdf", saveOptions);
    //ExEnd:SdtTagAsFormFieldName
  });


  test('RenderChoiceFormFieldBorder', () => {
    //ExStart:RenderChoiceFormFieldBorder
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:PdfSaveOptions.renderChoiceFormFieldBorder
    //ExSummary:Shows how to render PDF choice form field border.
    let doc = new aw.Document(base.myDir + "Legacy drop-down.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.preserveFormFields = true;
    saveOptions.renderChoiceFormFieldBorder = true;

    doc.save(base.artifactsDir + "PdfSaveOptions.renderChoiceFormFieldBorder.pdf", saveOptions);
    //ExEnd:RenderChoiceFormFieldBorder
  });


});
