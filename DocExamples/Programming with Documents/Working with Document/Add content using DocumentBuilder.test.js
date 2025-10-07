// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("AddContentUsingDocumentBuilder", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('CreateNewDocument', () => {
    //ExStart:CreateNewDocument
    //GistId:96e42cb4a611465927f8e7b1b3d546d3
    let doc = new aw.Document();

    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello World!");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.CreateNewDocument.docx");
    //ExEnd:CreateNewDocument
  });

  test('InsertBookmark', () => {
    //ExStart:InsertBookmark
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("FineBookmark");
    builder.writeln("This is just a fine bookmark.");
    builder.endBookmark("FineBookmark");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertBookmark.docx");
    //ExEnd:InsertBookmark
  });

  test('BuildTable', () => {
    //ExStart:BuildTable
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();

    builder.insertCell();
    builder.cellFormat.verticalAlignment = "Center";
    builder.write("This is row 1 cell 1");

    builder.insertCell();
    builder.write("This is row 1 cell 2");
    builder.endRow();

    builder.insertCell();
    builder.rowFormat.height = 100;
    builder.rowFormat.heightRule = "Exactly";
    builder.cellFormat.orientation = "Upward";
    builder.writeln("This is row 2 cell 1");

    builder.insertCell();
    builder.cellFormat.orientation = "Downward";
    builder.writeln("This is row 2 cell 2");
    builder.endRow();

    builder.endTable();

    table.autoFit(aw.Tables.AutoFitBehavior.FixedColumnWidths);

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.BuildTable.docx");
    //ExEnd:BuildTable
  });

  test('InsertHorizontalRule', () => {
    //ExStart:InsertHorizontalRule
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Insert a horizontal rule shape into the document.");
    builder.insertHorizontalRule();

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertHorizontalRule.docx");
    //ExEnd:InsertHorizontalRule
  });

  test('HorizontalRuleFormat', () => {
    //ExStart:HorizontalRuleFormat
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let builder = new aw.DocumentBuilder();

    let shape = builder.insertHorizontalRule();

    let horizontalRuleFormat = shape.horizontalRuleFormat;
    horizontalRuleFormat.alignment = "Center"; // Use string for alignment
    horizontalRuleFormat.widthPercent = 70;
    horizontalRuleFormat.height = 3;
    horizontalRuleFormat.color = "#0000FF"; // Use color code as string
    horizontalRuleFormat.noShade = true;

    builder.document.save(base.artifactsDir + "AddContentUsingDocumentBuilder.HorizontalRuleFormat.docx");
    //ExEnd:HorizontalRuleFormat
  });

  test('InsertBreak', () => {
    //ExStart:InsertBreak
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("This is page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);

    builder.writeln("This is page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);

    builder.writeln("This is page 3.");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertBreak.docx");
    //ExEnd:InsertBreak
  });

  test('InsertTextInputFormField', () => {
    //ExStart:InsertTextInputFormField
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertTextInput("TextInput", aw.Fields.TextFormFieldType.Regular, "", "Hello", 0);

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertTextInputFormField.docx");
    //ExEnd:InsertTextInputFormField
  });

  test('InsertCheckBoxFormField', () => {
    //ExStart:InsertCheckBoxFormField
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertCheckBox("CheckBox", true, true, 0);

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertCheckBoxFormField.docx");
    //ExEnd:InsertCheckBoxFormField
  });

  test('InsertComboBoxFormField', () => {
    //ExStart:InsertComboBoxFormField
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let items = ["One", "Two", "Three"];

    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertComboBox("DropDown", items, 0);

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertComboBoxFormField.docx");
    //ExEnd:InsertComboBoxFormField
  });

  test('InsertHtml', () => {
    //ExStart:InsertHtml
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertHtml(
        "<P align='right'>Paragraph right</P>" +
        "<b>Implicit paragraph left</b>" +
        "<div align='center'>Div center</div>" +
        "<h1 align='left'>Heading 1 left.</h1>"
    );

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertHtml.docx");
    //ExEnd:InsertHtml
  });

  test('InsertHyperlink', () => {
    //ExStart:InsertHyperlink
    //GistId:9b6efb87f331ae61c0100e106c9c1738
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Please make sure to visit ");

    builder.font.style = doc.styles.at(aw.StyleIdentifier.Hyperlink);
    builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);
    builder.font.clearFormatting();

    builder.write(" for more information.");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertHyperlink.docx");
    //ExEnd:InsertHyperlink
  });

  test('InsertTableOfContents', () => {
    //ExStart:InsertTableOfContents
    //GistId:e0ccef8441be6a8e2de5810acdefd25a
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

    // Start the actual document content on the second page.
    builder.insertBreak(aw.BreakType.PageBreak);

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    builder.writeln("Heading 1");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;

    builder.writeln("Heading 1.1");
    builder.writeln("Heading 1.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    builder.writeln("Heading 2");
    builder.writeln("Heading 3");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;

    builder.writeln("Heading 3.1");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading3;

    builder.writeln("Heading 3.1.1");
    builder.writeln("Heading 3.1.2");
    builder.writeln("Heading 3.1.3");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;

    builder.writeln("Heading 3.2");
    builder.writeln("Heading 3.3");

    //ExStart:UpdateFields
    //GistId:e0ccef8441be6a8e2de5810acdefd25a
    // The newly inserted table of contents will be initially empty.
    // It needs to be populated by updating the fields in the document.
    doc.updateFields();
    //ExEnd:UpdateFields

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertTableOfContents.docx");
    //ExEnd:InsertTableOfContents
  });

  test('InsertInlineImage', () => {
    //ExStart:InsertInlineImage
    //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imagesDir + "Transparent background logo.png");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertInlineImage.docx");
    //ExEnd:InsertInlineImage
  });

  test('InsertFloatingImage', () => {
    //ExStart:InsertFloatingImage
    //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imagesDir + "Transparent background logo.png",
        "Margin",
        100,
        "Margin",
        100,
        200,
        100,
        "Square");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertFloatingImage.docx");
    //ExEnd:InsertFloatingImage
  });

  test('InsertParagraph', () => {
    //ExStart:InsertParagraph
    //GistId:410919c9c1056a587ed5f2a86f328e7a
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let font = builder.font;
    font.size = 16;
    font.bold = true;
    font.color = "#0000FF";
    font.name = "Arial";
    font.underline = "Dash";

    let paragraphFormat = builder.paragraphFormat;
    paragraphFormat.firstLineIndent = 8;
    paragraphFormat.alignment = "Justify";
    paragraphFormat.keepTogether = true;

    builder.writeln("A whole paragraph.");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertParagraph.docx");
    //ExEnd:InsertParagraph
  });

  test('InsertTcField', () => {
    //ExStart:InsertTcField
    //GistId:e0ccef8441be6a8e2de5810acdefd25a
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField("TC \"Entry Text\" \\f t");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.InsertTcField.docx");
    //ExEnd:InsertTcField
  });

  test('CursorPosition', () => {
    //ExStart:CursorPosition
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let curNode = builder.currentNode;
    let curParagraph = builder.currentParagraph;
    //ExEnd:CursorPosition

    console.log("\nCursor move to paragraph: " + curParagraph.getText());
  });

  test('MoveToNode', () => {
    //ExStart:MoveToNode
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    //ExStart:MoveToBookmark
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Start a bookmark and add content to it using a DocumentBuilder.
    builder.startBookmark("MyBookmark");
    builder.writeln("Bookmark contents.");
    builder.endBookmark("MyBookmark");

    // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
    expect(builder.currentParagraph.firstChild).toEqual(doc.range.bookmarks.at(0).bookmarkEnd);

    // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
    builder.moveToBookmark("MyBookmark");

    // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
    expect(builder.currentParagraph.firstChild).toEqual(doc.range.bookmarks.at(0).bookmarkStart);

    // We can move the builder to an individual node,
    // which in this case will be the first node of the first paragraph, like this.
    builder.moveTo(doc.firstSection.body.firstParagraph.getChildNodes(aw.NodeType.Any, false).at(0));
    //ExEnd:MoveToBookmark

    expect(builder.currentNode.nodeType).toEqual(aw.NodeType.BookmarkStart);
    expect(builder.isAtStartOfParagraph).toBe(true);

    // A shorter way of moving the very start/end of a document is with these methods.
    builder.moveToDocumentEnd();
    expect(builder.isAtEndOfParagraph).toBe(true);
    builder.moveToDocumentStart();
    expect(builder.isAtStartOfParagraph).toBe(true);
    //ExEnd:MoveToNode
  });

  test('MoveToDocumentStartEnd', () => {
    //ExStart:MoveToDocumentStartEnd
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Move the cursor position to the beginning of your document.
    builder.moveToDocumentStart();
    console.log("\nThis is the beginning of the document.");

    // Move the cursor position to the end of your document.
    builder.moveToDocumentEnd();
    console.log("\nThis is the end of the document.");
    //ExEnd:MoveToDocumentStartEnd
  });

  test('MoveToSection', () => {
    //ExStart:MoveToSection
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document();
    doc.appendChild(new aw.Section(doc));

    // Move a DocumentBuilder to the second section and add text.
    let builder = new aw.DocumentBuilder(doc);
    builder.moveToSection(1);
    builder.writeln("Text added to the 2nd section.");

    // Create document with paragraphs.
    doc = new aw.Document(base.myDir + "Paragraphs.docx");
    let paragraphs = doc.firstSection.body.paragraphs;
    expect(paragraphs.count).toBe(22);

    // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
    // and any content added by the DocumentBuilder will just be prepended to the document.
    builder = new aw.DocumentBuilder(doc);
    expect(paragraphs.indexOf(builder.currentParagraph)).toBe(0);

    // You can move the cursor to any position in a paragraph.
    builder.moveToParagraph(2, 10);
    expect(paragraphs.indexOf(builder.currentParagraph)).toBe(2);
    builder.writeln("This is a new third paragraph. ");
    expect(paragraphs.indexOf(builder.currentParagraph)).toBe(3);
    //ExEnd:MoveToSection
  });

  test('MoveToHeadersFooters', () => {
    //ExStart:MoveToHeadersFooters
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify that we want headers and footers different for first, even and odd pages.
    builder.pageSetup.differentFirstPageHeaderFooter = true;
    builder.pageSetup.oddAndEvenPagesHeaderFooter = true;

    // Create the headers.
    builder.moveToHeaderFooter("HeaderFirst");
    builder.write("Header for the first page");
    builder.moveToHeaderFooter("HeaderEven");
    builder.write("Header for even pages");
    builder.moveToHeaderFooter("HeaderPrimary");
    builder.write("Header for all other pages");

    // Create two pages in the document.
    builder.moveToSection(0);
    builder.writeln("Page1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page2");

    doc.save(base.artifactsDir + "AddContentUsingDocumentBuilder.MoveToHeadersFooters.docx");
    //ExEnd:MoveToHeadersFooters
  });

  test('MoveToParagraph', () => {
    //ExStart:MoveToParagraph
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToParagraph(2, 0);
    builder.writeln("This is the 3rd paragraph.");
    //ExEnd:MoveToParagraph
  });

  test('MoveToTableCell', () => {
    //ExStart:MoveToTableCell
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let builder = new aw.DocumentBuilder(doc);

    // Move the builder to row 3, cell 4 of the first table.
    builder.moveToCell(0, 2, 3, 0);
    builder.write("\nCell contents added by DocumentBuilder");
    let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

    expect(table.rows.at(2).cells.at(3).getText().trim()).toMatch("Cell contents added by DocumentBuilderCell 3 contents");
    //ExEnd:MoveToTableCell
  });

  test('MoveToBookmarkEnd', () => {
    //ExStart:MoveToBookmarkEnd
    //GistId:410919c9c1056a587ed5f2a86f328e7a
    let doc = new aw.Document(base.myDir + "Bookmarks.docx");
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToBookmark("MyBookmark1", false, true);
    builder.writeln("This is a bookmark.");
    //ExEnd:MoveToBookmarkEnd
  });

  test('MoveToMergeField', () => {
    //ExStart:MoveToMergeField
    //GistId:811254d9cf25578ceaa32a1d990f70fa
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a field using the DocumentBuilder and add a run of text after it.
    let field = builder.insertField("MERGEFIELD field");
    builder.write(" Text after the field.");

    // The builder's cursor is currently at end of the document.
    expect(builder.currentNode).toBeNull();
    // We can move the builder to a field like this, placing the cursor at immediately after the field.
    builder.moveToField(field, true);

    // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
    // If we wish to move the DocumentBuilder to inside a field,
    // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method.
    expect(builder.currentNode.previousSibling).toEqual(field.end);
    builder.write(" Text immediately after the field.");
    //ExEnd:MoveToMergeField
  });
});
