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


describe("ExTxtSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  function readTextFile(filename, removebom = true, encoding = "utf8") {
    // Read text and remove BOM.
    let text = fs.readFileSync(filename, encoding).toString();
    if (removebom) {
      text = text.trimStart("\uFEFF");
    }
    return text;
  }


  test.each([false,
    true])('PageBreaks', (forcePageBreaks) => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptionsBase.forcePageBreaks
    //ExSummary:Shows how to specify whether to preserve page breaks when exporting a document to plaintext.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Page 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save"
    // method to modify how we save the document to plaintext.
    let saveOptions = new aw.Saving.TxtSaveOptions();

    // The Aspose.words "Document" objects have page breaks, just like Microsoft Word documents.
    // Save formats such as ".txt" are one continuous body of text without page breaks.
    // Set the "ForcePageBreaks" property to "true" to preserve all page breaks in the form of '\f' characters.
    // Set the "ForcePageBreaks" property to "false" to discard all page breaks.
    saveOptions.forcePageBreaks = forcePageBreaks;

    doc.save(base.artifactsDir + "TxtSaveOptions.PageBreaks.txt", saveOptions);
            
    // If we load a plaintext document with page breaks,
    // the "Document" object will use them to split the body into pages.
    doc = new aw.Document(base.artifactsDir + "TxtSaveOptions.PageBreaks.txt");

    expect(doc.pageCount).toEqual(forcePageBreaks ? 3 : 1);
    //ExEnd

    TestUtil.fileContainsString(
      forcePageBreaks ? "Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n" : "Page 1\r\nPage 2\r\nPage 3\r\n\r\n",
      base.artifactsDir + "TxtSaveOptions.PageBreaks.txt");
  });


  test.each([false,
    true])('AddBidiMarks', (addBidiMarks) => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptions.addBidiMarks
    //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    builder.paragraphFormat.bidi = true;
    builder.writeln("שלום עולם!");
    builder.writeln("مرحبا بالعالم!");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let saveOptions = new aw.Saving.TxtSaveOptions();
    saveOptions.encoding = "UTF-16";

    // Set the "AddBidiMarks" property to "true" to add marks before runs
    // with right-to-left text to indicate the fact.
    // Set the "AddBidiMarks" property to "false" to write all left-to-right
    // and right-to-left run equally with nothing to indicate which is which.
    saveOptions.addBidiMarks = addBidiMarks;

    doc.save(base.artifactsDir + "TxtSaveOptions.addBidiMarks.txt", saveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.addBidiMarks.txt", false, "utf16le"); 

    if (addBidiMarks)
    {
      expect(docText).toEqual("\uFEFFHello world!‎\r\nשלום עולם!‏\r\nمرحبا بالعالم!‏\r\n\r\n");
      expect(docText.includes("\u200f")).toEqual(true);
    }
    else
    {
      expect(docText).toEqual("\uFEFFHello world!\r\nשלום עולם!\r\nمرحبا بالعالم!\r\n\r\n");
      expect(docText.includes("\u200f")).toEqual(false);
    }
    //ExEnd
  });


  test.each([aw.Saving.TxtExportHeadersFootersMode.AllAtEnd,
    aw.Saving.TxtExportHeadersFootersMode.PrimaryOnly,
    aw.Saving.TxtExportHeadersFootersMode.None])('ExportHeadersFooters', (txtExportHeadersFootersMode) => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptionsBase.exportHeadersFootersMode
    //ExFor:TxtExportHeadersFootersMode
    //ExSummary:Shows how to specify how to export headers and footers to plain text format.
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

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let saveOptions = new aw.Saving.TxtSaveOptions();

    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.None"
    // to not export any headers/footers.
    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.PrimaryOnly"
    // to only export primary headers/footers.
    // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.AllAtEnd"
    // to place all headers and footers for all section bodies at the end of the document.
    saveOptions.exportHeadersFootersMode = txtExportHeadersFootersMode;

    doc.save(base.artifactsDir + "TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.ExportHeadersFooters.txt");

    switch (txtExportHeadersFootersMode)
    {
      case aw.Saving.TxtExportHeadersFootersMode.AllAtEnd:
        expect(docText).toEqual("Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n" +
                                    "Even header\r\n\r\n" +
                                    "Primary header\r\n\r\n" +
                                    "Even footer\r\n\r\n" +
                                    "Primary footer\r\n\r\n");
        break;
      case aw.Saving.TxtExportHeadersFootersMode.PrimaryOnly:
        expect(docText).toEqual("Primary header\r\n" +
                                    "Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n" +
                                    "Primary footer\r\n");
        break;
      case aw.Saving.TxtExportHeadersFootersMode.None:
        expect(docText).toEqual("Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n");
        break;
    }
    //ExEnd
  });


  test('TxtListIndentation', () => {
    //ExStart
    //ExFor:TxtListIndentation
    //ExFor:aw.Saving.TxtListIndentation.count
    //ExFor:aw.Saving.TxtListIndentation.character
    //ExFor:aw.Saving.TxtSaveOptions.listIndentation
    //ExSummary:Shows how to configure list indenting when saving a document to plaintext.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a list with three levels of indentation.
    builder.listFormat.applyNumberDefault();
    builder.writeln("Item 1");
    builder.listFormat.listIndent();
    builder.writeln("Item 2");
    builder.listFormat.listIndent(); 
    builder.write("Item 3");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let txtSaveOptions = new aw.Saving.TxtSaveOptions();

    // Set the "Character" property to assign a character to use
    // for padding that simulates list indentation in plaintext.
    txtSaveOptions.listIndentation.character = ' ';

    // Set the "Count" property to specify the number of times
    // to place the padding character for each list indent level.
    txtSaveOptions.listIndentation.count = 3;

    doc.save(base.artifactsDir + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.TxtListIndentation.txt");

    expect(docText).toEqual("1. Item 1\r\n" +
                            "   a. Item 2\r\n" +
                            "      i. Item 3\r\n");
    //ExEnd
  });


  test.each([false,
    true])('SimplifyListLabels', (simplifyListLabels) => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptions.simplifyListLabels
    //ExSummary:Shows how to change the appearance of lists when saving a document to plaintext.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a bulleted list with five levels of indentation.
    builder.listFormat.applyBulletDefault();
    builder.writeln("Item 1");
    builder.listFormat.listIndent();
    builder.writeln("Item 2");
    builder.listFormat.listIndent();
    builder.writeln("Item 3");
    builder.listFormat.listIndent();
    builder.writeln("Item 4");
    builder.listFormat.listIndent();
    builder.write("Item 5");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let txtSaveOptions = new aw.Saving.TxtSaveOptions();

    // Set the "SimplifyListLabels" property to "true" to convert some list
    // symbols into simpler ASCII characters, such as '*', 'o', '+', '>', etc.
    // Set the "SimplifyListLabels" property to "false" to preserve as many original list symbols as possible.
    txtSaveOptions.simplifyListLabels = simplifyListLabels;

    doc.save(base.artifactsDir + "TxtSaveOptions.simplifyListLabels.txt", txtSaveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.simplifyListLabels.txt");

    if (simplifyListLabels)
      expect(docText).toEqual("* Item 1\r\n" +
                                "  > Item 2\r\n" +
                                "    + Item 3\r\n" +
                                "      - Item 4\r\n" +
                                "        o Item 5\r\n");
    else
      expect(docText).toEqual("· Item 1\r\n" +
                                "o Item 2\r\n" +
                                "§ Item 3\r\n" +
                                "· Item 4\r\n" +
                                "o Item 5\r\n");
    //ExEnd
  });

  
  test('ParagraphBreak', () => {
    //ExStart
    //ExFor:TxtSaveOptions
    //ExFor:aw.Saving.TxtSaveOptions.saveFormat
    //ExFor:TxtSaveOptionsBase
    //ExFor:aw.Saving.TxtSaveOptionsBase.paragraphBreak
    //ExSummary:Shows how to save a .txt document with a custom paragraph break.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Paragraph 1.");
    builder.writeln("Paragraph 2.");
    builder.write("Paragraph 3.");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let txtSaveOptions = new aw.Saving.TxtSaveOptions();

    expect(txtSaveOptions.saveFormat).toEqual(aw.SaveFormat.Text);

    // Set the "ParagraphBreak" to a custom value that we wish to put at the end of every paragraph.
    txtSaveOptions.paragraphBreak = " End of paragraph.\n\n\t";

    doc.save(base.artifactsDir + "TxtSaveOptions.paragraphBreak.txt", txtSaveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.paragraphBreak.txt");

    expect(docText).toEqual("Paragraph 1. End of paragraph.\n\n\t" +
                            "Paragraph 2. End of paragraph.\n\n\t" +
                            "Paragraph 3. End of paragraph.\n\n\t");
    //ExEnd
  });


  test('Encoding', () => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptionsBase.encoding
    //ExSummary:Shows how to set encoding for a .txt output document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add some text with characters from outside the ASCII character set.
    builder.write("À È Ì Ò Ù.");

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let txtSaveOptions = new aw.Saving.TxtSaveOptions();

    // Verify that the "Encoding" property contains the appropriate encoding for our document's contents.
    expect(txtSaveOptions.encoding).toEqual("utf-8");

    doc.save(base.artifactsDir + "TxtSaveOptions.encoding.UTF8.txt", txtSaveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.encoding.UTF8.txt", false, "utf8");

    expect(docText).toEqual("\uFEFFÀ È Ì Ò Ù.\r\n");

    // Using an unsuitable encoding may result in a loss of document contents.
    txtSaveOptions.encoding = "us-ascii";
    doc.save(base.artifactsDir + "TxtSaveOptions.encoding.ASCII.txt", txtSaveOptions);
    docText = readTextFile(base.artifactsDir + "TxtSaveOptions.encoding.ASCII.txt", false, "ascii");

    expect(docText).toEqual("? ? ? ? ?.\r\n");
    //ExEnd
  });


  test.each([false,
    true])('PreserveTableLayout(%o)', (preserveTableLayout) => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptions.preserveTableLayout
    //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1");
    builder.insertCell();
    builder.write("Row 1, cell 2");
    builder.endRow();
    builder.insertCell();
    builder.write("Row 2, cell 1");
    builder.insertCell();
    builder.write("Row 2, cell 2");
    builder.endTable();

    // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
    // to modify how we save the document to plaintext.
    let txtSaveOptions = new aw.Saving.TxtSaveOptions();

    // Set the "PreserveTableLayout" property to "true" to apply whitespace padding to the contents
    // of the output plaintext document to preserve as much of the table's layout as possible.
    // Set the "PreserveTableLayout" property to "false" to save all tables' contents
    // as a continuous body of text, with just a new line for each row.
    txtSaveOptions.preserveTableLayout = preserveTableLayout;

    doc.save(base.artifactsDir + "TxtSaveOptions.preserveTableLayout.txt", txtSaveOptions);

    let docText = readTextFile(base.artifactsDir + "TxtSaveOptions.preserveTableLayout.txt");
    console.log(docText);

    if (preserveTableLayout)
      expect(docText).toEqual("Row 1, cell 1                                           Row 1, cell 2\r\n" +
                              "Row 2, cell 1                                           Row 2, cell 2\r\n\r\n");
    else
      expect(docText).toEqual("Row 1, cell 1\r" +
                              "Row 1, cell 2\r" +
                              "Row 2, cell 1\r" +
                              "Row 2, cell 2\r\r\n");
    //ExEnd
  });


  test('MaxCharactersPerLine', () => {
    //ExStart
    //ExFor:aw.Saving.TxtSaveOptions.maxCharactersPerLine
    //ExSummary:Shows how to set maximum number of characters per line.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    // Set 30 characters as maximum allowed per one line.
    let saveOptions = new aw.Saving.TxtSaveOptions();
    saveOptions.maxCharactersPerLine = 30;

    doc.save(base.artifactsDir + "TxtSaveOptions.maxCharactersPerLine.txt", saveOptions);
    //ExEnd
  });
});
