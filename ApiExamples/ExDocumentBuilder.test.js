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
const MemoryStream = require('memorystream');
const jimp = require("jimp");

describe("ExDocumentBuilder", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  beforeEach(() => {
    base.setUnlimitedLicense();
  });

  test('WriteAndFont', () => {
    //ExStart
    //ExFor:aw.Font.size
    //ExFor:aw.Font.bold
    //ExFor:aw.Font.name
    //ExFor:aw.Font.color
    //ExFor:aw.Font.underline
    //ExFor:DocumentBuilder.#ctor
    //ExSummary:Shows how to insert formatted text using DocumentBuilder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify font formatting, then add text.
    let font = builder.font;
    font.size = 16;
    font.bold = true;
    font.color = "#0000FF";
    font.name = "Courier New";
    font.underline = aw.Underline.Dash;

    builder.write("Hello world!");
    //ExEnd

    doc = DocumentHelper.saveOpen(builder.document);
    let firstRun = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(firstRun.getText().trim()).toEqual("Hello world!");
    expect(firstRun.font.size).toEqual(16.0);
    expect(firstRun.font.bold).toEqual(true);
    expect(firstRun.font.name).toEqual("Courier New");
    expect(firstRun.font.color).toEqual("#0000FF");
    expect(firstRun.font.underline).toEqual(aw.Underline.Dash);
  });

  test('HeadersAndFooters', () => {
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:DocumentBuilder.#ctor(Document)
    //ExFor:aw.DocumentBuilder.moveToHeaderFooter
    //ExFor:aw.DocumentBuilder.moveToSection
    //ExFor:aw.DocumentBuilder.insertBreak
    //ExFor:aw.DocumentBuilder.writeln
    //ExFor:HeaderFooterType
    //ExFor:aw.PageSetup.differentFirstPageHeaderFooter
    //ExFor:aw.PageSetup.oddAndEvenPagesHeaderFooter
    //ExFor:BreakType
    //ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify that we want different headers and footers for first, even and odd pages.
    builder.pageSetup.differentFirstPageHeaderFooter = true;
    builder.pageSetup.oddAndEvenPagesHeaderFooter = true;

    // Create the headers, then add three pages to the document to display each header type.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderFirst);
    builder.write("Header for the first page");
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderEven);
    builder.write("Header for even pages");
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("Header for all other pages");

    builder.moveToSection(0);
    builder.writeln("Page1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page2");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page3");

    doc.save(base.artifactsDir + "DocumentBuilder.HeadersAndFooters.docx");
    //ExEnd

    let headersFooters = 
      new aw.Document(base.artifactsDir + "DocumentBuilder.HeadersAndFooters.docx").firstSection.headersFooters;

    expect(headersFooters.count).toEqual(3);
    expect(headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderFirst).getText().trim()).toEqual("Header for the first page");
    expect(headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderEven).getText().trim()).toEqual("Header for even pages");
    expect(headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getText().trim()).toEqual("Header for all other pages");
  });

  test('MergeFields', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertField(String)
    //ExFor:aw.DocumentBuilder.moveToMergeField(String, Boolean, Boolean)
    //ExSummary:Shows how to insert fields, and move the document builder's cursor to them.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.insertField("MERGEFIELD MyMergeField1 \* MERGEFORMAT");
    builder.insertField("MERGEFIELD MyMergeField2 \* MERGEFORMAT");

    // Move the cursor to the first MERGEFIELD.
    builder.moveToMergeField("MyMergeField1", true, false);

    // Note that the cursor is placed immediately after the first MERGEFIELD, and before the second.
    expect(builder.currentNode).toEqual(doc.range.fields.at(1).start);
    expect(builder.currentNode.previousSibling).toEqual(doc.range.fields.at(0).end);

    // If we wish to edit the field's field code or contents using the builder,
    // its cursor would need to be inside a field.
    // To place it inside a field, we would need to call the document builder's MoveTo method
    // and pass the field's start or separator node as an argument.
    builder.write(" Text between our merge fields. ");

    doc.save(base.artifactsDir + "DocumentBuilder.MergeFields.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MergeFields.docx");

    expect(doc.getText().trim()).toEqual("\u0013MERGEFIELD MyMergeField1 \* MERGEFORMAT\u0014«MyMergeField1»\u0015" +
            " Text between our merge fields. " +
            "\u0013MERGEFIELD MyMergeField2 \* MERGEFORMAT\u0014«MyMergeField2»\u0015");
    expect(doc.range.fields.count).toEqual(2);
    TestUtil.verifyField(aw.Fields.FieldType.FieldMergeField, "MERGEFIELD MyMergeField1 \* MERGEFORMAT", 
      "«MyMergeField1»", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldMergeField, "MERGEFIELD MyMergeField2 \* MERGEFORMAT", 
      "«MyMergeField2»", doc.range.fields.at(1));
  });

  test('InsertHorizontalRule', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertHorizontalRule
    //ExFor:aw.Drawing.ShapeBase.isHorizontalRule
    //ExFor:aw.Drawing.Shape.horizontalRuleFormat
    //ExFor:HorizontalRuleFormat
    //ExFor:aw.Drawing.HorizontalRuleFormat.alignment
    //ExFor:aw.Drawing.HorizontalRuleFormat.widthPercent
    //ExFor:aw.Drawing.HorizontalRuleFormat.height
    //ExFor:aw.Drawing.HorizontalRuleFormat.color
    //ExFor:aw.Drawing.HorizontalRuleFormat.noShade
    //ExSummary:Shows how to insert a horizontal rule shape, and customize its formatting.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let shape = builder.insertHorizontalRule();

    let horizontalRuleFormat = shape.horizontalRuleFormat;
    horizontalRuleFormat.alignment = aw.Drawing.HorizontalRuleAlignment.Center;
    horizontalRuleFormat.widthPercent = 70;
    horizontalRuleFormat.height = 3;
    horizontalRuleFormat.color = "#0000FF";
    horizontalRuleFormat.noShade = true;

    expect(shape.isHorizontalRule).toEqual(true);
    expect(shape.horizontalRuleFormat.noShade).toEqual(true);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    shape = doc.getShape(0, true);

    expect(shape.horizontalRuleFormat.alignment).toEqual(aw.Drawing.HorizontalRuleAlignment.Center);
    expect(shape.horizontalRuleFormat.widthPercent).toEqual(70);
    expect(shape.horizontalRuleFormat.height).toEqual(3);
    expect(shape.horizontalRuleFormat.color).toEqual("#0000FF");
  });

  test('HorizontalRuleFormatExceptions', () => {
    let builder = new aw.DocumentBuilder();
    let shape = builder.insertHorizontalRule();

    let horizontalRuleFormat = shape.horizontalRuleFormat;
    horizontalRuleFormat.widthPercent = 1;
    horizontalRuleFormat.widthPercent = 100;
    expect(() => horizontalRuleFormat.widthPercent = 0).toThrow("Specified argument was out of the range of valid values.");
    expect(() => horizontalRuleFormat.widthPercent = 101).toThrow("Specified argument was out of the range of valid values.");
            
    horizontalRuleFormat.height = 0;
    horizontalRuleFormat.height = 1584;
    expect(() => horizontalRuleFormat.height = -1).toThrow("Specified argument was out of the range of valid values.");
    expect(() => horizontalRuleFormat.height = 1585).toThrow("Specified argument was out of the range of valid values.");
  });

  test('InsertHyperlink', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertHyperlink
    //ExFor:aw.Font.clearFormatting
    //ExFor:aw.Font.color
    //ExFor:aw.Font.underline
    //ExFor:Underline
    //ExSummary:Shows how to insert a hyperlink field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("For more information, please visit the ");

    // Insert a hyperlink and emphasize it with custom formatting.
    // The hyperlink will be a clickable piece of text which will take us to the location specified in the URL.
    builder.font.color = "#0000FF";
    builder.font.underline = aw.Underline.Single;
    builder.insertHyperlink("Google website", "https://www.google.com", false);
    builder.font.clearFormatting();
    builder.writeln(".");

    // Ctrl + left clicking the link in the text in Microsoft Word will take us to the URL via a new web browser window.
    doc.save(base.artifactsDir + "DocumentBuilder.insertHyperlink.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertHyperlink.docx");

    let hyperlink = doc.range.fields.at(0).asFieldHyperlink();
    expect(hyperlink.address).toEqual("https://www.google.com");

    let fieldContents = hyperlink.start.nextSibling.asRun();

    expect(fieldContents.font.color).toEqual("#0000FF");
    expect(fieldContents.font.underline).toEqual(aw.Underline.Single);
    expect(fieldContents.getText().trim()).toEqual("HYPERLINK \"https://www.google.com\"");
  });

  test('PushPopFont', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.pushFont
    //ExFor:aw.DocumentBuilder.popFont
    //ExFor:aw.DocumentBuilder.insertHyperlink
    //ExSummary:Shows how to use a document builder's formatting stack.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set up font formatting, then write the text that goes before the hyperlink.
    builder.font.name = "Arial";
    builder.font.size = 24;
    builder.write("To visit Google, hold Ctrl and click ");

    // Preserve our current formatting configuration on the stack.
    builder.pushFont();

    // Alter the builder's current formatting by applying a new style.
    builder.font.styleIdentifier = aw.StyleIdentifier.Hyperlink;
    builder.insertHyperlink("here", "http://www.google.com", false);

    expect(builder.font.color).toEqual("#0000FF");
    expect(builder.font.underline).toEqual(aw.Underline.Single);

    // Restore the font formatting that we saved earlier and remove the element from the stack.
    builder.popFont();

    expect(builder.font.color).toEqual(base.emptyColor);
    expect(builder.font.underline).toEqual(aw.Underline.None);

    builder.write(". We hope you enjoyed the example.");

    doc.save(base.artifactsDir + "DocumentBuilder.PushPopFont.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.PushPopFont.docx");
    let runs = doc.firstSection.body.firstParagraph.runs;

    expect(runs.count).toEqual(4);

    expect(runs.at(0).getText().trim()).toEqual("To visit Google, hold Ctrl and click");
    expect(runs.at(3).getText().trim()).toEqual(". We hope you enjoyed the example.");
    expect(runs.at(3).font.color).toEqual(runs.at(0).font.color);
    expect(runs.at(3).font.underline).toEqual(runs.at(0).font.underline);

    expect(runs.at(2).getText().trim()).toEqual("here");
    expect(runs.at(2).font.color).toEqual("#0000FF");
    expect(runs.at(2).font.underline).toEqual(aw.Underline.Single);
    expect(runs.at(2).font.color).not.toEqual(runs.at(0).font.color);
    expect(runs.at(2).font.underline).not.toEqual(runs.at(0).font.underline);

    expect(doc.range.fields.at(0).asFieldHyperlink().address).toEqual("http://www.google.com");
  });


  test('InsertWatermark', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToHeaderFooter
    //ExFor:aw.PageSetup.pageWidth
    //ExFor:aw.PageSetup.pageHeight
    //ExFor:WrapType
    //ExFor:RelativeHorizontalPosition
    //ExFor:RelativeVerticalPosition
    //ExSummary:Shows how to insert an image, and use it as a watermark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert the image into the header so that it will be visible on every page.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    let shape = builder.insertImage(base.imageDir + "Transparent background logo.png");
    shape.wrapType = aw.Drawing.WrapType.None;
    shape.behindText = true;

    // Place the image at the center of the page.
    shape.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
    shape.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    shape.left = (builder.pageSetup.pageWidth - shape.width) / 2;
    shape.top = (builder.pageSetup.pageHeight - shape.height) / 2;

    doc.save(base.artifactsDir + "DocumentBuilder.InsertWatermark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertWatermark.docx");
    shape = doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Png, shape);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(shape.behindText).toEqual(true);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Page);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Page);
    expect(shape.left).toEqual((doc.firstSection.pageSetup.pageWidth - shape.width) / 2);
    expect(shape.top).toEqual((doc.firstSection.pageSetup.pageHeight - shape.height) / 2);
  });


  test('InsertOleObject', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOleObject(String, Boolean, Boolean, Stream)
    //ExFor:aw.DocumentBuilder.insertOleObject(String, String, Boolean, Boolean, Stream)
    //ExFor:aw.DocumentBuilder.insertOleObjectAsIcon(String, Boolean, String, String)
    //ExSummary:Shows how to insert an OLE object into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // OLE objects are links to files in our local file system that can be opened by other installed applications.
    // Double clicking these shapes will launch the application, and then use it to open the linked object.
    // There are three ways of using the InsertOleObject method to insert these shapes and configure their appearance.
    // 1 -  Image taken from the local file system:
    let imageStream = base.loadFileToBuffer(base.imageDir + "Logo.jpg");
    // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    // the icon according to the file extension and uses the filename for the icon caption.
    builder.insertOleObject(base.myDir + "Spreadsheet.xlsx", false, false, imageStream); 

    // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    // the icon according to 'progId' and uses the filename for the icon caption.
    // 2 -  Icon based on the application that will open the object:
    builder.insertOleObject(base.myDir + "Spreadsheet.xlsx", "Excel.Sheet", false, true, null);

    // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    // the icon according to 'progId' and uses the predefined icon caption.
    // 3 -  Image icon that's 32 x 32 pixels or smaller from the local file system, with a custom caption:
    builder.insertOleObjectAsIcon(base.myDir + "Presentation.pptx", false, base.imageDir + "Logo icon.ico",
      "Double click to view presentation!");

    doc.save(base.artifactsDir + "DocumentBuilder.insertOleObject.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertOleObject.docx");
    let shape = doc.getShape(0, true);

    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.OleObject);
    expect(shape.oleFormat.progId).toEqual("Excel.Sheet.12");
    expect(shape.oleFormat.suggestedExtension).toEqual(".xlsx");

    shape = doc.getShape(1, true);

    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.OleObject);
    expect(shape.oleFormat.progId).toEqual("Package");
    expect(shape.oleFormat.suggestedExtension).toEqual(".xlsx");

    shape = doc.getShape(2, true);

    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.OleObject);
    expect(shape.oleFormat.progId).toEqual("PowerPoint.Show.12");
    expect(shape.oleFormat.suggestedExtension).toEqual(".pptx");
  });


  test('InsertHtml', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertHtml(String)
    //ExSummary:Shows how to use a document builder to insert html content into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    const html = "<p align='right'>Paragraph right</p>" + 
              "<b>Implicit paragraph left</b>" +
              "<div align='center'>Div center</div>" + 
              "<h1 align='left'>Heading 1 left.</h1>";

    builder.insertHtml(html);

    // Inserting HTML code parses the formatting of each element into equivalent document text formatting.
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).getText().trim()).toEqual("Paragraph right");
    expect(paragraphs.at(0).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Right);

    expect(paragraphs.at(1).getText().trim()).toEqual("Implicit paragraph left");
    expect(paragraphs.at(1).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Left);
    expect(paragraphs.at(1).runs.at(0).font.bold).toEqual(true);

    expect(paragraphs.at(2).getText().trim()).toEqual("Div center");
    expect(paragraphs.at(2).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);

    expect(paragraphs.at(3).getText().trim()).toEqual("Heading 1 left.");
    expect(paragraphs.at(3).paragraphFormat.style.name).toEqual("Heading 1");

    doc.save(base.artifactsDir + "DocumentBuilder.insertHtml.docx");
    //ExEnd
  });

  test.each([false,
    true])('InsertHtmlWithFormatting', (useBuilderFormatting) => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertHtml(String, Boolean)
    //ExSummary:Shows how to apply a document builder's formatting while inserting HTML content.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set a text alignment for the builder, insert an HTML paragraph with a specified alignment, and one without.
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Distributed;
    builder.insertHtml(
      "<p align='right'>Paragraph 1.</p>" +
      "<p>Paragraph 2.</p>", useBuilderFormatting);

    let paragraphs = doc.firstSection.body.paragraphs;

    // The first paragraph has an alignment specified. When InsertHtml parses the HTML code,
    // the paragraph alignment value found in the HTML code always supersedes the document builder's value.
    expect(paragraphs.at(0).getText().trim()).toEqual("Paragraph 1.");
    expect(paragraphs.at(0).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Right);

    // The second paragraph has no alignment specified. It can have its alignment value filled in
    // by the builder's value depending on the flag we passed to the InsertHtml method.
    expect(paragraphs.at(1).getText().trim()).toEqual("Paragraph 2.");
    expect(paragraphs.at(1).paragraphFormat.alignment).toEqual(
      useBuilderFormatting ? aw.ParagraphAlignment.Distributed : aw.ParagraphAlignment.Left);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertHtmlWithFormatting.docx");
    //ExEnd
  });

  test('MathMl', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    const mathMl =
      "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

    builder.insertHtml(mathMl);

    doc.save(base.artifactsDir + "DocumentBuilder.mathML.docx");
    doc.save(base.artifactsDir + "DocumentBuilder.mathML.pdf");

    expect(DocumentHelper.compareDocs(base.goldsDir + "DocumentBuilder.mathML Gold.docx", base.artifactsDir + "DocumentBuilder.mathML.docx")).toEqual(true);
  });

  test('InsertTextAndBookmark', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.startBookmark
    //ExFor:aw.DocumentBuilder.endBookmark
    //ExSummary:Shows how create a bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A valid bookmark needs to have document body text enclosed by
    // BookmarkStart and BookmarkEnd nodes created with a matching bookmark name.
    builder.startBookmark("MyBookmark");
    builder.writeln("Hello world!");
    builder.endBookmark("MyBookmark");
            
    expect(doc.range.bookmarks.count).toEqual(1);
    expect(doc.range.bookmarks.at(0).name).toEqual("MyBookmark");
    expect(doc.range.bookmarks.at(0).text.trim()).toEqual("Hello world!");
    //ExEnd
  });

  test('CreateColumnBookmark', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.startColumnBookmark
    //ExFor:aw.DocumentBuilder.endColumnBookmark
    //ExSummary:Shows how to create a column bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startTable();

    builder.insertCell();
    // Cells 1,2,4,5 will be bookmarked.
    builder.startColumnBookmark("MyBookmark_1");
    // Badly formed bookmarks or bookmarks with duplicate names will be ignored when the document is saved.
    builder.startColumnBookmark("MyBookmark_1");
    builder.startColumnBookmark("BadStartBookmark");
    builder.write("Cell 1");

    builder.insertCell();
    builder.write("Cell 2");

    builder.insertCell();
    builder.write("Cell 3");

    builder.endRow();

    builder.insertCell();
    builder.write("Cell 4");

    builder.insertCell();
    builder.write("Cell 5");
    builder.endColumnBookmark("MyBookmark_1");
    builder.endColumnBookmark("MyBookmark_1");

    expect(() => builder.endColumnBookmark("BadEndBookmark")).toThrow("The corresponding bookmark start must be in the same table.");

    builder.insertCell();
    builder.write("Cell 6");

    builder.endRow();
    builder.endTable();

    doc.save(base.artifactsDir + "Bookmarks.CreateColumnBookmark.docx");
    //ExEnd
  });

  test('CreateForm', () => {
    //ExStart
    //ExFor:TextFormFieldType
    //ExFor:aw.DocumentBuilder.insertTextInput
    //ExFor:aw.DocumentBuilder.insertComboBox
    //ExSummary:Shows how to create form fields.
    let builder = new aw.DocumentBuilder();

    // Form fields are objects in the document that the user can interact with by being prompted to enter values.
    // We can create them using a document builder, and below are two ways of doing so.
    // 1 -  Basic text input:
    builder.insertTextInput("My text input", aw.Fields.TextFormFieldType.Regular, 
      "", "Enter your name here", 30);

    // 2 -  Combo box with prompt text, and a range of possible values:
    let items =
    [
      "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"
    ];

    builder.insertParagraph();
    builder.insertComboBox("My combo box", items, 0);

    builder.document.save(base.artifactsDir + "DocumentBuilder.CreateForm.docx");
    //ExEnd

    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.CreateForm.docx");
    let formField = doc.range.formFields.at(0);

    expect(formField.name).toEqual("My text input");
    expect(formField.textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
    expect(formField.result).toEqual("Enter your name here");

    formField = doc.range.formFields.at(1);

    expect(formField.name).toEqual("My combo box");
    expect(formField.textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
    expect(formField.result).toEqual("-- Select your favorite footwear --");
    expect(formField.dropDownSelectedIndex).toEqual(0);
    expect([...formField.dropDownItems]).toEqual(
    [
      "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other"
    ]);
  });

  test('InsertCheckBox', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertCheckBox(string, bool, bool, int)
    //ExFor:aw.DocumentBuilder.insertCheckBox(String, bool, int)
    //ExSummary:Shows how to insert checkboxes into the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert checkboxes of varying sizes and default checked statuses.
    builder.write("Unchecked check box of a default size: ");
    builder.insertCheckBox('', false, false, 0);
    builder.insertParagraph();

    builder.write("Large checked check box: ");
    builder.insertCheckBox("CheckBox_Default", true, true, 50);
    builder.insertParagraph();

    // Form fields have a name length limit of 20 characters.
    builder.write("Very large checked check box: ");
    builder.insertCheckBox("CheckBox_OnlyCheckedValue", true, 100);

    expect(doc.range.formFields.at(2).name).toEqual("CheckBox_OnlyChecked");

    // We can interact with these check boxes in Microsoft Word by double clicking them.
    doc.save(base.artifactsDir + "DocumentBuilder.insertCheckBox.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertCheckBox.docx");

    let formFields = doc.range.formFields;

    expect(formFields.at(0).name).toEqual('');
    expect(formFields.at(0).checked).toEqual(false);
    expect(formFields.at(0).default).toEqual(false);
    expect(formFields.at(0).checkBoxSize).toEqual(10);

    expect(formFields.at(1).name).toEqual("CheckBox_Default");
    expect(formFields.at(1).checked).toEqual(true);
    expect(formFields.at(1).default).toEqual(true);
    expect(formFields.at(1).checkBoxSize).toEqual(50);

    expect(formFields.at(2).name).toEqual("CheckBox_OnlyChecked");
    expect(formFields.at(2).checked).toEqual(true);
    expect(formFields.at(2).default).toEqual(true);
    expect(formFields.at(2).checkBoxSize).toEqual(100);
  });

  test('InsertCheckBoxEmptyName', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Checking that the checkbox insertion with an empty name working correctly
    builder.insertCheckBox("", true, false, 1);
    builder.insertCheckBox('', false, 1);
  });

  
  test('WorkingWithNodes', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveTo(Node)
    //ExFor:aw.DocumentBuilder.moveToBookmark(String)
    //ExFor:aw.DocumentBuilder.currentParagraph
    //ExFor:aw.DocumentBuilder.currentNode
    //ExFor:aw.DocumentBuilder.moveToDocumentStart
    //ExFor:aw.DocumentBuilder.moveToDocumentEnd
    //ExFor:aw.DocumentBuilder.isAtEndOfParagraph
    //ExFor:aw.DocumentBuilder.isAtStartOfParagraph
    //ExSummary:Shows how to move a document builder's cursor to different nodes in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a valid bookmark, an entity that consists of nodes enclosed by a bookmark start node,
    // and a bookmark end node. 
    builder.startBookmark("MyBookmark");
    builder.write("Bookmark contents.");
    builder.endBookmark("MyBookmark");

    let firstParagraphNodes = doc.firstSection.body.firstParagraph.getChildNodes(aw.NodeType.Any, false);

    expect(firstParagraphNodes.at(0).nodeType).toEqual(aw.NodeType.BookmarkStart);
    expect(firstParagraphNodes.at(1).nodeType).toEqual(aw.NodeType.Run);
    expect(firstParagraphNodes.at(1).getText().trim()).toEqual("Bookmark contents.");
    expect(firstParagraphNodes.at(2).nodeType).toEqual(aw.NodeType.BookmarkEnd);

    // The document builder's cursor is always ahead of the node that we last added with it.
    // If the builder's cursor is at the end of the document, its current node will be null.
    // The previous node is the bookmark end node that we last added.
    // Adding new nodes with the builder will append them to the last node.
    expect(builder.currentNode).toBe(null);

    // If we wish to edit a different part of the document with the builder,
    // we will need to bring its cursor to the node we wish to edit.
    builder.moveToBookmark("MyBookmark");

    // Moving it to a bookmark will move it to the first node within the bookmark start and end nodes, the enclosed run.
    expect(builder.currentNode.referenceEquals(firstParagraphNodes.at(1))).toBe(true);

    // We can also move the cursor to an individual node like this.
    builder.moveTo(doc.firstSection.body.firstParagraph.getChildNodes(aw.NodeType.Any, false).at(0));

    expect(builder.currentNode.nodeType).toEqual(aw.NodeType.BookmarkStart);
    expect(builder.currentParagraph).toEqual(doc.firstSection.body.firstParagraph);
    expect(builder.isAtStartOfParagraph).toEqual(true);

    // We can use specific methods to move to the start/end of a document.
    builder.moveToDocumentEnd();

    expect(builder.isAtEndOfParagraph).toEqual(true);

    builder.moveToDocumentStart();

    expect(builder.isAtStartOfParagraph).toEqual(true);
    //ExEnd
  });


  test('FillMergeFields', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToMergeField(String)
    //ExFor:aw.DocumentBuilder.bold
    //ExFor:aw.DocumentBuilder.italic
    //ExSummary:Shows how to fill MERGEFIELDs with data with a document builder instead of a mail merge.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge,
    // and then fill them manually.
    builder.insertField(" MERGEFIELD Chairman ");
    builder.insertField(" MERGEFIELD ChiefFinancialOfficer ");
    builder.insertField(" MERGEFIELD ChiefTechnologyOfficer ");

    builder.moveToMergeField("Chairman");
    builder.bold = true;
    builder.writeln("John Doe");

    builder.moveToMergeField("ChiefFinancialOfficer");
    builder.italic = true;
    builder.writeln("Jane Doe");

    builder.moveToMergeField("ChiefTechnologyOfficer");
    builder.italic = true;
    builder.writeln("John Bloggs");

    doc.save(base.artifactsDir + "DocumentBuilder.FillMergeFields.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.FillMergeFields.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).runs.at(0).font.bold).toEqual(true);
    expect(paragraphs.at(0).runs.at(0).getText().trim()).toEqual("John Doe");

    expect(paragraphs.at(1).runs.at(0).font.italic).toEqual(true);
    expect(paragraphs.at(1).runs.at(0).getText().trim()).toEqual("Jane Doe");

    expect(paragraphs.at(2).runs.at(0).font.italic).toEqual(true);
    expect(paragraphs.at(2).runs.at(0).getText().trim()).toEqual("John Bloggs");
  });

  test('InsertToc', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertTableOfContents
    //ExFor:aw.Document.updateFields
    //ExFor:DocumentBuilder.#ctor(Document)
    //ExFor:aw.ParagraphFormat.styleIdentifier
    //ExFor:aw.DocumentBuilder.insertBreak
    //ExFor:BreakType
    //ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a table of contents for the first page of the document.
    // Configure the table to pick up paragraphs with headings of levels 1 to 3.
    // Also, set its entries to be hyperlinks that will take us
    // to the location of the heading when left-clicked in Microsoft Word.
    builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
    builder.insertBreak(aw.BreakType.PageBreak);

    // Populate the table of contents by adding paragraphs with heading styles.
    // Each such heading with a level between 1 and 3 will create an entry in the table.
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

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading4;
    builder.writeln("Heading 3.1.3.1");
    builder.writeln("Heading 3.1.3.2");

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;
    builder.writeln("Heading 3.2");
    builder.writeln("Heading 3.3");

    // A table of contents is a field of a type that needs to be updated to show an up-to-date result.
    doc.updateFields();
    doc.save(base.artifactsDir + "DocumentBuilder.InsertToc.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertToc.docx");
    let tableOfContents = doc.range.fields.at(0).asFieldToc();

    expect(tableOfContents.headingLevelRange).toEqual("1-3");
    expect(tableOfContents.insertHyperlinks).toEqual(true);
    expect(tableOfContents.hideInWebLayout).toEqual(true);
    expect(tableOfContents.useParagraphOutlineLevel).toEqual(true);
  });

  test('InsertTable', () => {
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:aw.DocumentBuilder.write
    //ExFor:aw.DocumentBuilder.startTable
    //ExFor:aw.DocumentBuilder.insertCell
    //ExFor:aw.DocumentBuilder.endRow
    //ExFor:aw.DocumentBuilder.endTable
    //ExFor:aw.DocumentBuilder.cellFormat
    //ExFor:aw.DocumentBuilder.rowFormat
    //ExFor:CellFormat
    //ExFor:aw.Tables.CellFormat.fitText
    //ExFor:aw.Tables.CellFormat.width
    //ExFor:aw.Tables.CellFormat.verticalAlignment
    //ExFor:aw.Tables.CellFormat.shading
    //ExFor:aw.Tables.CellFormat.orientation
    //ExFor:aw.Tables.CellFormat.wrapText
    //ExFor:RowFormat
    //ExFor:aw.Tables.RowFormat.borders
    //ExFor:aw.Tables.RowFormat.clearFormatting
    //ExFor:aw.Shading.clearFormatting
    //ExSummary:Shows how to build a table with custom borders.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startTable();

    // Setting table formatting options for a document builder
    // will apply them to every row and cell that we add with it.
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;

    builder.cellFormat.clearFormatting();
    builder.cellFormat.width = 150;
    builder.cellFormat.verticalAlignment = aw.Tables.CellVerticalAlignment.Center;
    builder.cellFormat.shading.backgroundPatternColor = "#ADFF2F";
    builder.cellFormat.wrapText = false;
    builder.cellFormat.fitText = true;

    builder.rowFormat.clearFormatting();
    builder.rowFormat.heightRule = aw.HeightRule.Exactly;
    builder.rowFormat.height = 50;
    builder.rowFormat.borders.lineStyle = aw.LineStyle.Engrave3D;
    builder.rowFormat.borders.color = "#FFA500";

    builder.insertCell();
    builder.write("Row 1, Col 1");

    builder.insertCell();
    builder.write("Row 1, Col 2");
    builder.endRow();

    // Changing the formatting will apply it to the current cell,
    // and any new cells that we create with the builder afterward.
    // This will not affect the cells that we have added previously.
    builder.cellFormat.shading.clearFormatting();

    builder.insertCell();
    builder.write("Row 2, Col 1");

    builder.insertCell();
    builder.write("Row 2, Col 2");

    builder.endRow();

    // Increase row height to fit the vertical text.
    builder.insertCell();
    builder.rowFormat.height = 150;
    builder.cellFormat.orientation = aw.TextOrientation.Upward;
    builder.write("Row 3, Col 1");

    builder.insertCell();
    builder.cellFormat.orientation = aw.TextOrientation.Downward;
    builder.write("Row 3, Col 2");

    builder.endRow();
    builder.endTable();

    doc.save(base.artifactsDir + "DocumentBuilder.InsertTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertTable.docx");
    let table = doc.firstSection.body.tables.at(0);

    expect(table.rows.at(0).cells.at(0).getText().trim()).toEqual("Row 1, Col 1\u0007");
    expect(table.rows.at(0).cells.at(1).getText().trim()).toEqual("Row 1, Col 2\u0007");
    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.Exactly);
    expect(table.rows.at(0).rowFormat.height).toEqual(50.0);
    expect(table.rows.at(0).rowFormat.borders.lineStyle).toEqual(aw.LineStyle.Engrave3D);
    expect(table.rows.at(0).rowFormat.borders.color).toEqual("#FFA500");

    for (var c of table.rows.at(0).cells.toArray())
    {
      expect(c.cellFormat.width).toEqual(150);
      expect(c.cellFormat.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);
      expect(c.cellFormat.shading.backgroundPatternColor).toEqual("#ADFF2F");
      expect(c.cellFormat.wrapText).toEqual(false);
      expect(c.cellFormat.fitText).toEqual(true);

      expect(c.firstParagraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);
    }

    expect(table.rows.at(1).cells.at(0).getText().trim()).toEqual("Row 2, Col 1\u0007");
    expect(table.rows.at(1).cells.at(1).getText().trim()).toEqual("Row 2, Col 2\u0007");


    for (var c of table.rows.at(1).cells.toArray())
    {
      expect(c.cellFormat.width).toEqual(150);
      expect(c.cellFormat.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);
      expect(c.cellFormat.shading.backgroundPatternColor).toEqual(base.emptyColor);
      expect(c.cellFormat.wrapText).toEqual(false);
      expect(c.cellFormat.fitText).toEqual(true);

      expect(c.firstParagraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);
    }

    expect(table.rows.at(2).rowFormat.height).toEqual(150);

    expect(table.rows.at(2).cells.at(0).getText().trim()).toEqual("Row 3, Col 1\u0007");
    expect(table.rows.at(2).cells.at(0).cellFormat.orientation).toEqual(aw.TextOrientation.Upward);
    expect(table.rows.at(2).cells.at(0).firstParagraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);

    expect(table.rows.at(2).cells.at(1).getText().trim()).toEqual("Row 3, Col 2\u0007");
    expect(table.rows.at(2).cells.at(1).cellFormat.orientation).toEqual(aw.TextOrientation.Downward);
    expect(table.rows.at(2).cells.at(1).firstParagraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);
  });


  test('InsertTableWithStyle', () => {
    //ExStart
    //ExFor:aw.Tables.Table.styleIdentifier
    //ExFor:aw.Tables.Table.styleOptions
    //ExFor:TableStyleOptions
    //ExFor:aw.Tables.Table.autoFit
    //ExFor:AutoFitBehavior
    //ExSummary:Shows how to build a new table while applying a style.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let table = builder.startTable();

    // We must insert at least one row before setting any table formatting.
    builder.insertCell();

    // Set the table style used based on the style identifier.
    // Note that not all table styles are available when saving to .doc format.
    table.styleIdentifier = aw.StyleIdentifier.MediumShading1Accent1;

    // Partially apply the style to features of the table based on predicates, then build the table.
    table.styleOptions =
      aw.Tables.TableStyleOptions.FirstColumn | aw.Tables.TableStyleOptions.RowBands | aw.Tables.TableStyleOptions.FirstRow;
    table.autoFit(aw.Tables.AutoFitBehavior.AutoFitToContents);

    builder.writeln("Item");
    builder.cellFormat.rightPadding = 40;
    builder.insertCell();
    builder.writeln("Quantity (kg)");
    builder.endRow();

    builder.insertCell();
    builder.writeln("Apples");
    builder.insertCell();
    builder.writeln("20");
    builder.endRow();

    builder.insertCell();
    builder.writeln("Bananas");
    builder.insertCell();
    builder.writeln("40");
    builder.endRow();

    builder.insertCell();
    builder.writeln("Carrots");
    builder.insertCell();
    builder.writeln("50");
    builder.endRow();

    doc.save(base.artifactsDir + "DocumentBuilder.InsertTableWithStyle.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertTableWithStyle.docx");

    doc.expandTableStylesToDirectFormatting();

    expect(table.style.name).toEqual("Medium Shading 1 Accent 1");
    expect(table.styleOptions).toEqual(aw.Tables.TableStyleOptions.FirstColumn | aw.Tables.TableStyleOptions.RowBands | aw.Tables.TableStyleOptions.FirstRow);
    expect(table.firstRow.firstCell.cellFormat.shading.backgroundPatternColor).toEqual("#4F81BD");
    expect(table.firstRow.firstCell.firstParagraph.runs.at(0).font.color).toEqual("#FFFFFF");
    expect(table.lastRow.firstCell.cellFormat.shading.backgroundPatternColor).not.toEqual("#ADD8E6");
    expect(table.lastRow.firstCell.firstParagraph.runs.at(0).font.color).toEqual(base.emptyColor);
  });

  test('InsertTableSetHeadingRow', () => {
    //ExStart
    //ExFor:aw.Tables.RowFormat.headingFormat
    //ExSummary:Shows how to build a table with rows that repeat on every page. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();

    // Any rows inserted while the "HeadingFormat" flag is set to "true"
    // will show up at the top of the table on every page that it spans.
    builder.rowFormat.headingFormat = true;
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.cellFormat.width = 100;
    builder.insertCell();
    builder.write("Heading row 1");
    builder.endRow();
    builder.insertCell();
    builder.write("Heading row 2");
    builder.endRow();

    builder.cellFormat.width = 50;
    builder.paragraphFormat.clearFormatting();
    builder.rowFormat.headingFormat = false;

    // Add enough rows for the table to span two pages.
    for (let i = 0; i < 50; i++)
    {
      builder.insertCell();
      builder.write(`Row ${table.rows.count}, column 1.`);
      builder.insertCell();
      builder.write(`Row ${table.rows.count}, column 2.`);
      builder.endRow();
    }

    doc.save(base.artifactsDir + "DocumentBuilder.InsertTableSetHeadingRow.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertTableSetHeadingRow.docx");
    table = doc.firstSection.body.tables.at(0);

    for (let i = 0; i < table.rows.count; i++)
      expect(table.rows.at(i).rowFormat.headingFormat).toEqual(i < 2);
  });

  test('InsertTableWithPreferredWidth', () => {
    //ExStart
    //ExFor:aw.Tables.Table.preferredWidth
    //ExFor:aw.Tables.PreferredWidth.fromPercent
    //ExFor:PreferredWidth
    //ExSummary:Shows how to set a table to auto fit to 50% of the width of the page.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Cell #1");
    builder.insertCell();
    builder.write("Cell #2");
    builder.insertCell();
    builder.write("Cell #3");

    table.preferredWidth = aw.Tables.PreferredWidth.fromPercent(50);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.preferredWidth.type).toEqual(aw.Tables.PreferredWidthType.Percent);
    expect(table.preferredWidth.value).toEqual(50);
  });

  test('InsertCellsWithPreferredWidths', () => {
    //ExStart
    //ExFor:aw.Tables.CellFormat.preferredWidth
    //ExFor:PreferredWidth
    //ExFor:aw.Tables.PreferredWidth.auto
    //ExFor:aw.Tables.PreferredWidth.equals(PreferredWidth)
    //ExFor:aw.Tables.PreferredWidth.equals(Object)
    //ExFor:aw.Tables.PreferredWidth.fromPoints
    //ExFor:aw.Tables.PreferredWidth.fromPercent
    //ExFor:aw.Tables.PreferredWidth.getHashCode
    //ExFor:aw.Tables.PreferredWidth.toString
    //ExSummary:Shows how to set a preferred width for table cells.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let table = builder.startTable();

    // There are two ways of applying the "PreferredWidth" class to table cells.
    // 1 -  Set an absolute preferred width based on points:
    builder.insertCell();
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPoints(40);
    builder.cellFormat.shading.backgroundPatternColor = "#FFFFE0";
    builder.writeln(`Cell with a width of ${builder.cellFormat.preferredWidth}.`);

    // 2 -  Set a relative preferred width based on percent of the table's width:
    builder.insertCell();
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPercent(20);
    builder.cellFormat.shading.backgroundPatternColor = "#ADD8E6";
    builder.writeln(`Cell with a width of ${builder.cellFormat.preferredWidth}.`);

    builder.insertCell();

    // A cell with no preferred width specified will take up the rest of the available space.
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.auto;

    builder.cellFormat.shading.backgroundPatternColor = "#90EE90";
    builder.writeln("Automatically sized cell.");

    doc.save(base.artifactsDir + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    //ExEnd

    expect(aw.Tables.PreferredWidth.fromPercent(100).value).toEqual(100.0);
    expect(aw.Tables.PreferredWidth.fromPoints(100).value).toEqual(100.0);

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.firstRow.cells.at(0).cellFormat.preferredWidth.type).toEqual(aw.Tables.PreferredWidthType.Points);
    expect(table.firstRow.cells.at(0).cellFormat.preferredWidth.value).toEqual(40.0);
    expect(table.firstRow.cells.at(0).getText().trim()).toEqual("Cell with a width of 800.\r\u0007");

    expect(table.firstRow.cells.at(1).cellFormat.preferredWidth.type).toEqual(aw.Tables.PreferredWidthType.Percent);
    expect(table.firstRow.cells.at(1).cellFormat.preferredWidth.value).toEqual(20.0);
    expect(table.firstRow.cells.at(1).getText().trim()).toEqual("Cell with a width of 20%.\r\u0007");

    expect(table.firstRow.cells.at(2).cellFormat.preferredWidth.type).toEqual(aw.Tables.PreferredWidthType.Auto);
    expect(table.firstRow.cells.at(2).cellFormat.preferredWidth.value).toEqual(0.0);
    expect(table.firstRow.cells.at(2).getText().trim()).toEqual("Automatically sized cell.\r\u0007");
  });

  test('InsertTableFromHtml', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
    // inserted from HTML.
    builder.insertHtml("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" +
            "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>");

    doc.save(base.artifactsDir + "DocumentBuilder.InsertTableFromHtml.docx");

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertTableFromHtml.docx");

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(1);
    expect(doc.getChildNodes(aw.NodeType.Row, true).count).toEqual(2);
    expect(doc.getChildNodes(aw.NodeType.Cell, true).count).toEqual(4);
  });

  test('InsertNestedTable', () => {
    //ExStart
    //ExFor:aw.Tables.Cell.firstParagraph
    //ExSummary:Shows how to create a nested table using a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Build the outer table.
    let cell = builder.insertCell();
    builder.writeln("Outer Table Cell 1");
    builder.insertCell();
    builder.writeln("Outer Table Cell 2");
    builder.endTable();

    // Move to the first cell of the outer table, the build another table inside the cell.
    builder.moveTo(cell.firstParagraph);
    builder.insertCell();
    builder.writeln("Inner Table Cell 1");
    builder.insertCell();
    builder.writeln("Inner Table Cell 2");
    builder.endTable();

    doc.save(base.artifactsDir + "DocumentBuilder.InsertNestedTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertNestedTable.docx");

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(2);
    expect(doc.getChildNodes(aw.NodeType.Cell, true).count).toEqual(4);
    expect(cell.tables.at(0).count).toEqual(1);
    expect(cell.tables.at(0).firstRow.cells.count).toEqual(2);
  });

  test('CreateTable', () => {
    //ExStart
    //ExFor:DocumentBuilder
    //ExFor:aw.DocumentBuilder.write
    //ExFor:aw.DocumentBuilder.insertCell
    //ExSummary:Shows how to use a document builder to create a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Start the table, then populate the first row with two cells.
    builder.startTable();
    builder.insertCell();
    builder.write("Row 1, Cell 1.");
    builder.insertCell();
    builder.write("Row 1, Cell 2.");

    // Call the builder's "EndRow" method to start a new row.
    builder.endRow();
    builder.insertCell();
    builder.write("Row 2, Cell 1.");
    builder.insertCell();
    builder.write("Row 2, Cell 2.");
    builder.endTable();

    doc.save(base.artifactsDir + "DocumentBuilder.CreateTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.CreateTable.docx");
    let table = doc.firstSection.body.tables.at(0);

    expect(table.getChildNodes(aw.NodeType.Cell, true).count).toEqual(4);

    expect(table.rows.at(0).cells.at(0).getText().trim()).toEqual("Row 1, Cell 1.\u0007");
    expect(table.rows.at(0).cells.at(1).getText().trim()).toEqual("Row 1, Cell 2.\u0007");
    expect(table.rows.at(1).cells.at(0).getText().trim()).toEqual("Row 2, Cell 1.\u0007");
    expect(table.rows.at(1).cells.at(1).getText().trim()).toEqual("Row 2, Cell 2.\u0007");
  });

  test('BuildFormattedTable', () => {
    //ExStart
    //ExFor:aw.Tables.RowFormat.height
    //ExFor:aw.Tables.RowFormat.heightRule
    //ExFor:aw.Tables.Table.leftIndent
    //ExFor:aw.DocumentBuilder.paragraphFormat
    //ExFor:aw.DocumentBuilder.font
    //ExSummary:Shows how to create a formatted table using DocumentBuilder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    table.leftIndent = 20;

    // Set some formatting options for text and table appearance.
    builder.rowFormat.height = 40;
    builder.rowFormat.heightRule = aw.HeightRule.AtLeast;
    builder.cellFormat.shading.backgroundPatternColor = "#C6D9F1";

    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.font.size = 16;
    builder.font.name = "Arial";
    builder.font.bold = true;

    // Configuring the formatting options in a document builder will apply them
    // to the current cell/row its cursor is in,
    // as well as any new cells and rows created using that builder.
    builder.write("Header Row,\n Cell 1");
    builder.insertCell();
    builder.write("Header Row,\n Cell 2");
    builder.insertCell();
    builder.write("Header Row,\n Cell 3");
    builder.endRow();

    // Reconfigure the builder's formatting objects for new rows and cells that we are about to make.
    // The builder will not apply these to the first row already created so that it will stand out as a header row.
    builder.cellFormat.shading.backgroundPatternColor = "#FFFFFF";
    builder.cellFormat.verticalAlignment = aw.Tables.CellVerticalAlignment.Center;
    builder.rowFormat.height = 30;
    builder.rowFormat.heightRule = aw.HeightRule.Auto;
    builder.insertCell();
    builder.font.size = 12;
    builder.font.bold = false;

    builder.write("Row 1, Cell 1.");
    builder.insertCell();
    builder.write("Row 1, Cell 2.");
    builder.insertCell();
    builder.write("Row 1, Cell 3.");
    builder.endRow();
    builder.insertCell();
    builder.write("Row 2, Cell 1.");
    builder.insertCell();
    builder.write("Row 2, Cell 2.");
    builder.insertCell();
    builder.write("Row 2, Cell 3.");
    builder.endRow();
    builder.endTable();

    doc.save(base.artifactsDir + "DocumentBuilder.CreateFormattedTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.CreateFormattedTable.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.leftIndent).toEqual(20.0);

    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.AtLeast);
    expect(table.rows.at(0).rowFormat.height).toEqual(40.0);

    for (let c of doc.getChildNodes(aw.NodeType.Cell, true))
    {
      let cell = c.asCell();
      expect(cell.firstParagraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);

      for (let r of cell.firstParagraph.runs.toArray())
      {
        expect(r.font.name).toEqual("Arial");

        if (cell.parentRow.referenceEquals(table.firstRow))
        {
          expect(r.font.size).toEqual(16);
          expect(r.font.bold).toEqual(true);
        }
        else
        {
          expect(r.font.size).toEqual(12);
          expect(r.font.bold).toEqual(false);
        }
      }
    }
  });

  test('TableBordersAndShading', () => {
    //ExStart
    //ExFor:Shading
    //ExFor:aw.Tables.Table.setBorders
    //ExFor:aw.BorderCollection.left
    //ExFor:aw.BorderCollection.right
    //ExFor:aw.BorderCollection.top
    //ExFor:aw.BorderCollection.bottom
    //ExSummary:Shows how to apply border and shading color while building a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Start a table and set a default color/thickness for its borders.
    let table = builder.startTable();
    table.setBorders(aw.LineStyle.Single, 2.0, "#000000");

    // Create a row with two cells with different background colors.
    builder.insertCell();
    builder.cellFormat.shading.backgroundPatternColor = "#87CEFA";
    builder.writeln("Row 1, Cell 1.");
    builder.insertCell();
    builder.cellFormat.shading.backgroundPatternColor = "#FFA500";
    builder.writeln("Row 1, Cell 2.");
    builder.endRow();

    // Reset cell formatting to disable the background colors
    // set a custom border thickness for all new cells created by the builder,
    // then build a second row.
    builder.cellFormat.clearFormatting();
    builder.cellFormat.borders.left.lineWidth = 4.0;
    builder.cellFormat.borders.right.lineWidth = 4.0;
    builder.cellFormat.borders.top.lineWidth = 4.0;
    builder.cellFormat.borders.bottom.lineWidth = 4.0;

    builder.insertCell();
    builder.writeln("Row 2, Cell 1.");
    builder.insertCell();
    builder.writeln("Row 2, Cell 2.");

    doc.save(base.artifactsDir + "DocumentBuilder.TableBordersAndShading.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.TableBordersAndShading.docx");
    table = doc.firstSection.body.tables.at(0);

    for (let c of table.firstRow.cells.toArray())
    {
      expect(c.cellFormat.borders.top.lineWidth).toEqual(0.5);
      expect(c.cellFormat.borders.bottom.lineWidth).toEqual(0.5);
      expect(c.cellFormat.borders.left.lineWidth).toEqual(0.5);
      expect(c.cellFormat.borders.right.lineWidth).toEqual(0.5);

      expect(c.cellFormat.borders.left.color).toEqual(base.emptyColor);
      expect(c.cellFormat.borders.left.lineStyle).toEqual(aw.LineStyle.Single);
    }

    expect(table.firstRow.firstCell.cellFormat.shading.backgroundPatternColor).toEqual("#87CEFA");
    expect(table.firstRow.cells.at(1).cellFormat.shading.backgroundPatternColor).toEqual("#FFA500");

    for (let c of table.lastRow.cells.toArray())
    {
      expect(c.cellFormat.borders.top.lineWidth).toEqual(4.0);
      expect(c.cellFormat.borders.bottom.lineWidth).toEqual(4.0);
      expect(c.cellFormat.borders.left.lineWidth).toEqual(4.0);
      expect(c.cellFormat.borders.right.lineWidth).toEqual(4.0);

      expect(c.cellFormat.borders.left.color).toEqual(base.emptyColor);
      expect(c.cellFormat.borders.left.lineStyle).toEqual(aw.LineStyle.Single);
      expect(c.cellFormat.shading.backgroundPatternColor).toEqual(base.emptyColor);
    }
  });

  test('SetPreferredTypeConvertUtil', () => {
    //ExStart
    //ExFor:aw.Tables.PreferredWidth.fromPoints
    //ExSummary:Shows how to use unit conversion tools while specifying a preferred width for a cell.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPoints(aw.ConvertUtil.inchToPoint(3));
    builder.insertCell();

    expect(table.firstRow.firstCell.cellFormat.preferredWidth.value).toEqual(216.0);
    //ExEnd
  });

  test('InsertHyperlinkToLocalBookmark', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.startBookmark
    //ExFor:aw.DocumentBuilder.endBookmark
    //ExFor:aw.DocumentBuilder.insertHyperlink
    //ExSummary:Shows how to insert a hyperlink which references a local bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("Bookmark1");
    builder.write("Bookmarked text. ");
    builder.endBookmark("Bookmark1");
    builder.writeln("Text outside of the bookmark.");

    // Insert a HYPERLINK field that links to the bookmark. We can pass field switches
    // to the "InsertHyperlink" method as part of the argument containing the referenced bookmark's name.
    builder.font.color = "#0000FF";
    builder.font.underline = aw.Underline.Single;
    let hyperlink = builder.insertHyperlink("Link to Bookmark1", "Bookmark1", true).asFieldHyperlink();
    hyperlink.screenTip = "Hyperlink Tip";

    doc.save(base.artifactsDir + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
    hyperlink = doc.range.fields.at(0).asFieldHyperlink();

    TestUtil.verifyField(aw.Fields.FieldType.FieldHyperlink, " HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", "Link to Bookmark1", hyperlink);
    expect(hyperlink.subAddress).toEqual("Bookmark1");
    expect(hyperlink.screenTip).toEqual("Hyperlink Tip");
    expect([...doc.range.bookmarks].some(b => b.name == "Bookmark1")).toEqual(true);
  });

  test('CursorPosition', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Hello world!");

    // If the builder's cursor is at the end of the document,
    // there will be no nodes in front of it so that the current node will be null.
    expect(builder.currentNode).toBe(null);

    expect(builder.currentParagraph.getText().trim()).toEqual("Hello world!");

    // Move to the beginning of the document and place the cursor at an existing node.
    builder.moveToDocumentStart();
    expect(builder.currentNode.nodeType).toEqual(aw.NodeType.Run);
  });

  test('MoveTo', () => {
    //ExStart
    //ExFor:aw.Story.lastParagraph
    //ExFor:aw.DocumentBuilder.moveTo(Node)
    //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Run 1. ");

    // The document builder has a cursor, which acts as the part of the document
    // where the builder appends new nodes when we use its document construction methods.
    // This cursor functions in the same way as Microsoft Word's blinking cursor,
    // and it also always ends up immediately after any node that the builder just inserted.
    // To append content to a different part of the document,
    // we can move the cursor to a different node with the "MoveTo" method.
    expect(builder.currentParagraph).toEqual(doc.firstSection.body.lastParagraph);
    builder.moveTo(doc.firstSection.body.firstParagraph.runs.at(0));
    expect(builder.currentParagraph).toEqual(doc.firstSection.body.firstParagraph);

    // The cursor is now in front of the node that we moved it to.
    // Adding a second run will insert it in front of the first run.
    builder.writeln("Run 2. ");

    expect(doc.getText().trim()).toEqual("Run 2. \rRun 1.");

    // Move the cursor to the end of the document to continue appending text to the end as before.
    builder.moveTo(doc.lastSection.body.lastParagraph);
    builder.writeln("Run 3. ");

    expect(doc.getText().trim()).toEqual("Run 2. \rRun 1. \rRun 3.");
    expect(builder.currentParagraph).toEqual(doc.firstSection.body.lastParagraph);

    //ExEnd
  });

  test('MoveToParagraph', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToParagraph
    //ExSummary:Shows how to move a builder's cursor position to a specified paragraph.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.count).toEqual(22);

    // Create document builder to edit the document. The builder's cursor,
    // which is the point where it will insert new nodes when we call its document construction methods,
    // is currently at the beginning of the document.
    let builder = new aw.DocumentBuilder(doc);

    expect(paragraphs.indexOf(builder.currentParagraph)).toEqual(0);

    // Move that cursor to a different paragraph will place that cursor in front of that paragraph.
    builder.moveToParagraph(2, 0);
    expect(paragraphs.indexOf(builder.currentParagraph)).toEqual(2);

    // Any new content that we add will be inserted at that point.
    builder.writeln("This is a new third paragraph. ");
    //ExEnd

    expect(paragraphs.indexOf(builder.currentParagraph)).toEqual(3);

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.firstSection.body.paragraphs.at(2).getText().trim()).toEqual("This is a new third paragraph.");
  });

  test('MoveToCell', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToCell
    //ExSummary:Shows how to move a document builder's cursor to a cell in a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an empty 2x2 table.
    builder.startTable();
    builder.insertCell();
    builder.insertCell();
    builder.endRow();
    builder.insertCell();
    builder.insertCell();
    builder.endTable();

    // Because we have ended the table with the EndTable method,
    // the document builder's cursor is currently outside the table.
    // This cursor has the same function as Microsoft Word's blinking text cursor.
    // It can also be moved to a different location in the document using the builder's MoveTo methods.
    // We can move the cursor back inside the table to a specific cell.
    builder.moveToCell(0, 1, 1, 0);
    builder.write("Column 2, cell 2.");

    doc.save(base.artifactsDir + "DocumentBuilder.moveToCell.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.moveToCell.docx");

    let table = doc.firstSection.body.tables.at(0);

    expect(table.rows.at(1).cells.at(1).getText().trim()).toEqual("Column 2, cell 2.\u0007");
  });

  test('MoveToBookmark', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToBookmark(String, Boolean, Boolean)
    //ExSummary:Shows how to move a document builder's node insertion point cursor to a bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A valid bookmark consists of a BookmarkStart node, a BookmarkEnd node with a
    // matching bookmark name somewhere afterward, and contents enclosed by those nodes.
    builder.startBookmark("MyBookmark");
    builder.write("Hello world! ");
    builder.endBookmark("MyBookmark");

    // There are 4 ways of moving a document builder's cursor to a bookmark.
    // If we are between the BookmarkStart and BookmarkEnd nodes, the cursor will be inside the bookmark.
    // This means that any text added by the builder will become a part of the bookmark.
    // 1 -  Outside of the bookmark, in front of the BookmarkStart node:
    expect(builder.moveToBookmark("MyBookmark", true, false)).toEqual(true);
    builder.write("1. ");

    expect(doc.range.bookmarks.at("MyBookmark").text).toEqual("Hello world! ");
    expect(doc.getText().trim()).toEqual("1. Hello world!");

    // 2 -  Inside the bookmark, right after the BookmarkStart node:
    expect(builder.moveToBookmark("MyBookmark", true, true)).toEqual(true);
    builder.write("2. ");

    expect(doc.range.bookmarks.at("MyBookmark").text).toEqual("2. Hello world! ");
    expect(doc.getText().trim()).toEqual("1. 2. Hello world!");

    // 2 -  Inside the bookmark, right in front of the BookmarkEnd node:
    expect(builder.moveToBookmark("MyBookmark", false, false)).toEqual(true);
    builder.write("3. ");

    expect(doc.range.bookmarks.at("MyBookmark").text).toEqual("2. Hello world! 3. ");
    expect(doc.getText().trim()).toEqual("1. 2. Hello world! 3.");

    // 4 -  Outside of the bookmark, after the BookmarkEnd node:
    expect(builder.moveToBookmark("MyBookmark", false, true)).toEqual(true);
    builder.write("4.");

    expect(doc.range.bookmarks.at("MyBookmark").text).toEqual("2. Hello world! 3. ");
    expect(doc.getText().trim()).toEqual("1. 2. Hello world! 3. 4.");
    //ExEnd
  });

  test('BuildTable', () => {
    //ExStart
    //ExFor:Table
    //ExFor:aw.DocumentBuilder.startTable
    //ExFor:aw.DocumentBuilder.endRow
    //ExFor:aw.DocumentBuilder.endTable
    //ExFor:aw.DocumentBuilder.cellFormat
    //ExFor:aw.DocumentBuilder.rowFormat
    //ExFor:aw.DocumentBuilder.write(String)
    //ExFor:aw.DocumentBuilder.writeln(String)
    //ExFor:CellVerticalAlignment
    //ExFor:aw.Tables.CellFormat.orientation
    //ExFor:TextOrientation
    //ExFor:AutoFitBehavior
    //ExSummary:Shows how to build a formatted 2x2 table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.cellFormat.verticalAlignment = aw.Tables.CellVerticalAlignment.Center;
    builder.write("Row 1, cell 1.");
    builder.insertCell();
    builder.write("Row 1, cell 2.");
    builder.endRow();

    // While building the table, the document builder will apply its current RowFormat/CellFormat property values
    // to the current row/cell that its cursor is in and any new rows/cells as it creates them.
    expect(table.rows.at(0).cells.at(0).cellFormat.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);
    expect(table.rows.at(0).cells.at(1).cellFormat.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);

    builder.insertCell();
    builder.rowFormat.height = 100;
    builder.rowFormat.heightRule = aw.HeightRule.Exactly;
    builder.cellFormat.orientation = aw.TextOrientation.Upward;
    builder.write("Row 2, cell 1.");
    builder.insertCell();
    builder.cellFormat.orientation = aw.TextOrientation.Downward;
    builder.write("Row 2, cell 2.");
    builder.endRow();
    builder.endTable();

    // Previously added rows and cells are not retroactively affected by changes to the builder's formatting.
    expect(table.rows.at(0).rowFormat.height).toEqual(0);
    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.Auto);
    expect(table.rows.at(1).rowFormat.height).toEqual(100);
    expect(table.rows.at(1).rowFormat.heightRule).toEqual(aw.HeightRule.Exactly);
    expect(table.rows.at(1).cells.at(0).cellFormat.orientation).toEqual(aw.TextOrientation.Upward);
    expect(table.rows.at(1).cells.at(1).cellFormat.orientation).toEqual(aw.TextOrientation.Downward);

    doc.save(base.artifactsDir + "DocumentBuilder.BuildTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.BuildTable.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.rows.count).toEqual(2);
    expect(table.rows.at(0).cells.count).toEqual(2);
    expect(table.rows.at(1).cells.count).toEqual(2);

    expect(table.rows.at(0).rowFormat.height).toEqual(0);
    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.Auto);
    expect(table.rows.at(1).rowFormat.height).toEqual(100);
    expect(table.rows.at(1).rowFormat.heightRule).toEqual(aw.HeightRule.Exactly);

    expect(table.rows.at(0).cells.at(0).getText().trim()).toEqual("Row 1, cell 1.\u0007");
    expect(table.rows.at(0).cells.at(0).cellFormat.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);

    expect(table.rows.at(0).cells.at(1).getText().trim()).toEqual("Row 1, cell 2.\u0007");

    expect(table.rows.at(1).cells.at(0).getText().trim()).toEqual("Row 2, cell 1.\u0007");
    expect(table.rows.at(1).cells.at(0).cellFormat.orientation).toEqual(aw.TextOrientation.Upward);

    expect(table.rows.at(1).cells.at(1).getText().trim()).toEqual("Row 2, cell 2.\u0007");
    expect(table.rows.at(1).cells.at(1).cellFormat.orientation).toEqual(aw.TextOrientation.Downward);
  });

  test('TableCellVerticalRotatedFarEastTextOrientation', () => {
    let doc = new aw.Document(base.myDir + "Rotated cell text.docx");

    let table = doc.firstSection.body.tables.at(0);
    let cell = table.firstRow.firstCell;

    expect(cell.cellFormat.orientation).toEqual(aw.TextOrientation.VerticalRotatedFarEast);

    doc = DocumentHelper.saveOpen(doc);

    table = doc.firstSection.body.tables.at(0);
    cell = table.firstRow.firstCell;

    expect(cell.cellFormat.orientation).toEqual(aw.TextOrientation.VerticalRotatedFarEast);
  });

  test('InsertFloatingImage', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // There are two ways of using a document builder to source an image and then insert it as a floating shape.
    // 1 -  From a file in the local file system:
    builder.insertImage(base.imageDir + "Transparent background logo.png", aw.Drawing.RelativeHorizontalPosition.Margin, 100,
      aw.Drawing.RelativeVerticalPosition.Margin, 0, 200, 200, aw.Drawing.WrapType.Square);

    // 2 -  From a URL:
    builder.insertImage(base.imageUrl.toString(), aw.Drawing.RelativeHorizontalPosition.Margin, 100,
      aw.Drawing.RelativeVerticalPosition.Margin, 250, 200, 200, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertFloatingImage.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertFloatingImage.docx");
    let image = doc.getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Png, image);
    expect(image.left).toEqual(100.0);
    expect(image.top).toEqual(0.0);
    expect(image.width).toEqual(200.0);
    expect(image.height).toEqual(200.0);
    expect(image.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(image.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(image.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);

    image = doc.getShape(1, true);

    TestUtil.verifyImageInShape(272, 92, aw.Drawing.ImageType.Png, image);
    expect(image.left).toEqual(100.0);
    expect(image.top).toEqual(250.0);
    expect(image.width).toEqual(200.0);
    expect(image.height).toEqual(200.0);
    expect(image.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(image.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(image.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);
  });

  test('InsertImageOriginalSize', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from the local file system into a document while preserving its dimensions.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // The InsertImage method creates a floating shape with the passed image in its image data.
    // We can specify the dimensions of the shape can be passing them to this method.
    let imageShape = builder.insertImage(base.imageDir + "Logo.jpg", aw.Drawing.RelativeHorizontalPosition.Margin, 0,
      aw.Drawing.RelativeVerticalPosition.Margin, 0, -1, -1, aw.Drawing.WrapType.Square);

    // Passing negative values as the intended dimensions will automatically define
    // the shape's dimensions based on the dimensions of its image.
    expect(imageShape.width).toEqual(300.0);
    expect(imageShape.height).toEqual(300.0);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertImageOriginalSize.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertImageOriginalSize.docx");
    imageShape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);
    expect(imageShape.width).toEqual(300.0);
    expect(imageShape.height).toEqual(300.0);
    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);
  });

  test('InsertTextInput', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertTextInput
    //ExSummary:Shows how to insert a text input form field into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a form that prompts the user to enter text.
    builder.insertTextInput("TextInput", aw.Fields.TextFormFieldType.Regular, "", "Enter your text here", 0);

    doc.save(base.artifactsDir + "DocumentBuilder.insertTextInput.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertTextInput.docx");
    let formField = doc.range.formFields.at(0);

    expect(formField.enabled).toEqual(true);
    expect(formField.name).toEqual("TextInput");
    expect(formField.maxLength).toEqual(0);
    expect(formField.result).toEqual("Enter your text here");
    expect(formField.type).toEqual(aw.Fields.FieldType.FieldFormTextInput);
    expect(formField.textInputFormat).toEqual("");
    expect(formField.textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
  });

  test('InsertComboBox', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertComboBox
    //ExSummary:Shows how to insert a combo box form field into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a form that prompts the user to pick one of the items from the menu.
    builder.write("Pick a fruit: ");
    let items = [ "Apple", "Banana", "Cherry" ];
    builder.insertComboBox("DropDown", items, 0);

    doc.save(base.artifactsDir + "DocumentBuilder.insertComboBox.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertComboBox.docx");
    let formField = doc.range.formFields.at(0);

    expect(formField.enabled).toEqual(true);
    expect(formField.name).toEqual("DropDown");
    expect(formField.dropDownSelectedIndex).toEqual(0);
    expect([...formField.dropDownItems]).toEqual(items);
    expect(formField.type).toEqual(aw.Fields.FieldType.FieldFormDropDown);
  });

  /*
    [Description("WORDSNET-16868")]
    [AotTests.IgnoreAot("CertificateHolder.Create and DigitalSignatureUtil.Sign are not used in AW.NET directly.")]
  test('SignatureLineProviderId', () => {
    //ExStart
    //ExFor:aw.Drawing.SignatureLine.isSigned
    //ExFor:aw.Drawing.SignatureLine.isValid
    //ExFor:aw.Drawing.SignatureLine.providerId
    //ExFor:aw.SignatureLineOptions.showDate
    //ExFor:aw.SignatureLineOptions.email
    //ExFor:aw.SignatureLineOptions.defaultInstructions
    //ExFor:aw.SignatureLineOptions.instructions
    //ExFor:aw.SignatureLineOptions.allowComments
    //ExFor:aw.DocumentBuilder.insertSignatureLine(SignatureLineOptions)
    //ExFor:aw.DigitalSignatures.SignOptions.providerId
    //ExSummary:Shows how to sign a document with a personal certificate and a signature line.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let signatureLineOptions = new aw.SignatureLineOptions
    {
      Signer = "vderyushev",
      SignerTitle = "QA",
      Email = "vderyushev@aspose.com",
      ShowDate = true,
      DefaultInstructions = false,
      Instructions = "Please sign here.",
      AllowComments = true
    };

    let signatureLine = builder.insertSignatureLine(signatureLineOptions).SignatureLine;
    signatureLine.providerId = Guid.parse("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2");

    expect(signatureLine.isSigned).toEqual(false);
    expect(signatureLine.isValid).toEqual(false);

    doc.save(base.artifactsDir + "DocumentBuilder.SignatureLineProviderId.docx");

    let signOptions = new aw.DigitalSignatures.SignOptions
    {
      SignatureLineId = signatureLine.id,
      ProviderId = signatureLine.providerId,
      Comments = "Document was signed by vderyushev",
      SignTime = Date.now()
    };

    let certHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    aw.DigitalSignatures.DigitalSignatureUtil.sign(base.artifactsDir + "DocumentBuilder.SignatureLineProviderId.docx", 
      base.artifactsDir + "DocumentBuilder.SignatureLineProviderId.signed.docx", certHolder, signOptions);

    // Re-open our saved document, and verify that the "IsSigned" and "IsValid" properties both equal "true",
    // indicating that the signature line contains a signature.
    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SignatureLineProviderId.signed.docx");
    let shape = (Shape)doc.getShape(0, true);
    signatureLine = shape.signatureLine;

    expect(signatureLine.isSigned).toEqual(true);
    expect(signatureLine.isValid).toEqual(true);
    //ExEnd

    expect(signatureLine.signer).toEqual("vderyushev");
    expect(signatureLine.signerTitle).toEqual("QA");
    expect(signatureLine.email).toEqual("vderyushev@aspose.com");
    expect(signatureLine.showDate).toEqual(true);
    expect(signatureLine.defaultInstructions).toEqual(false);
    expect(signatureLine.instructions).toEqual("Please sign here.");
    expect(signatureLine.allowComments).toEqual(true);
    expect(signatureLine.isSigned).toEqual(true);
    expect(signatureLine.isValid).toEqual(true);

    let signatures = aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(
      base.artifactsDir + "DocumentBuilder.SignatureLineProviderId.signed.docx");

    expect(signatures.count).toEqual(1);
    expect(signatures.at(0).isValid).toEqual(true);
    expect(signatures.at(0).comments).toEqual("Document was signed by vderyushev");
    expect(signatures.at(0).signTime.date).toEqual(DateTime.Today);
    expect(signatures.at(0).issuerName).toEqual("CN=Morzal.Me");
    expect(signatures.at(0).signatureType).toEqual(aw.DigitalSignatures.DigitalSignatureType.XmlDsig);
  });
*/

  test('SignatureLineInline', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
    //ExSummary:Shows how to insert an inline signature line into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let options = new aw.SignatureLineOptions();
    options.signer = "John Doe",
    options.signerTitle = "Manager",
    options.email = "johndoe@aspose.com",
    options.showDate = true,
    options.defaultInstructions = false,
    options.instructions = "Please sign here.",
    options.allowComments = true

    builder.insertSignatureLine(options, aw.Drawing.RelativeHorizontalPosition.RightMargin, 2.0,
      aw.Drawing.RelativeVerticalPosition.Page, 3.0, aw.Drawing.WrapType.Inline);

    // The signature line can be signed in Microsoft Word by double clicking it.
    doc.save(base.artifactsDir + "DocumentBuilder.SignatureLineInline.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SignatureLineInline.docx");

    let shape = doc.getShape(0, true);
    let signatureLine = shape.signatureLine;

    expect(signatureLine.signer).toEqual("John Doe");
    expect(signatureLine.signerTitle).toEqual("Manager");
    expect(signatureLine.email).toEqual("johndoe@aspose.com");
    expect(signatureLine.showDate).toEqual(true);
    expect(signatureLine.defaultInstructions).toEqual(false);
    expect(signatureLine.instructions).toEqual("Please sign here.");
    expect(signatureLine.allowComments).toEqual(true);
    expect(signatureLine.isSigned).toEqual(false);
    expect(signatureLine.isValid).toEqual(false);
  });

  test('SetParagraphFormatting', () => {
    //ExStart
    //ExFor:aw.ParagraphFormat.rightIndent
    //ExFor:aw.ParagraphFormat.leftIndent
    //ExSummary:Shows how to configure paragraph formatting to create off-center text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Center all text that the document builder writes, and set up indents.
    // The indent configuration below will create a body of text that will sit asymmetrically on the page.
    // The "center" that we align the text to will be the middle of the body of text, not the middle of the page.
    let paragraphFormat = builder.paragraphFormat;
    paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    paragraphFormat.leftIndent = 100;
    paragraphFormat.rightIndent = 50;
    paragraphFormat.spaceAfter = 25;

    builder.writeln(
      "This paragraph demonstrates how left and right indentation affects word wrapping.");
    builder.writeln(
      "The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

    doc.save(base.artifactsDir + "DocumentBuilder.SetParagraphFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SetParagraphFormatting.docx");

    for (let paragraph of doc.firstSection.body.paragraphs.toArray())
    {
      expect(paragraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);
      expect(paragraph.paragraphFormat.leftIndent).toEqual(100.0);
      expect(paragraph.paragraphFormat.rightIndent).toEqual(50.0);
      expect(paragraph.paragraphFormat.spaceAfter).toEqual(25.0);
    }
  });

  test('SetCellFormatting', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.cellFormat
    //ExFor:aw.Tables.CellFormat.width
    //ExFor:aw.Tables.CellFormat.leftPadding
    //ExFor:aw.Tables.CellFormat.rightPadding
    //ExFor:aw.Tables.CellFormat.topPadding
    //ExFor:aw.Tables.CellFormat.bottomPadding
    //ExFor:aw.DocumentBuilder.startTable
    //ExFor:aw.DocumentBuilder.endTable
    //ExSummary:Shows how to format cells with a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1.");

    // Insert a second cell, and then configure cell text padding options.
    // The builder will apply these settings at its current cell, and any new cells creates afterwards.
    builder.insertCell();

    let cellFormat = builder.cellFormat;
    cellFormat.width = 250;
    cellFormat.leftPadding = 30;
    cellFormat.rightPadding = 30;
    cellFormat.topPadding = 30;
    cellFormat.bottomPadding = 30;

    builder.write("Row 1, cell 2.");
    builder.endRow();
    builder.endTable();

    let cells = table.firstRow.cells.toArray();
    // The first cell was unaffected by the padding reconfiguration, and still holds the default values.
    expect(cells.at(0).cellFormat.width).toEqual(0.0);
    expect(cells.at(0).cellFormat.leftPadding).toEqual(5.4);
    expect(cells.at(0).cellFormat.rightPadding).toEqual(5.4);
    expect(cells.at(0).cellFormat.topPadding).toEqual(0.0);
    expect(cells.at(0).cellFormat.bottomPadding).toEqual(0.0);

    expect(cells.at(1).cellFormat.width).toEqual(250.0);
    expect(cells.at(1).cellFormat.leftPadding).toEqual(30.0);
    expect(cells.at(1).cellFormat.rightPadding).toEqual(30.0);
    expect(cells.at(1).cellFormat.topPadding).toEqual(30.0);
    expect(cells.at(1).cellFormat.bottomPadding).toEqual(30.0);

    // The first cell will still grow in the output document to match the size of its neighboring cell.
    doc.save(base.artifactsDir + "DocumentBuilder.SetCellFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SetCellFormatting.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.firstRow.cells.at(0).cellFormat.width).toEqual(157);
    expect(table.firstRow.cells.at(0).cellFormat.leftPadding).toEqual(5.4);
    expect(table.firstRow.cells.at(0).cellFormat.rightPadding).toEqual(5.4);
    expect(table.firstRow.cells.at(0).cellFormat.topPadding).toEqual(0.0);
    expect(table.firstRow.cells.at(0).cellFormat.bottomPadding).toEqual(0.0);

    expect(table.firstRow.cells.at(1).cellFormat.width).toEqual(310.0);
    expect(table.firstRow.cells.at(1).cellFormat.leftPadding).toEqual(30.0);
    expect(table.firstRow.cells.at(1).cellFormat.rightPadding).toEqual(30.0);
    expect(table.firstRow.cells.at(1).cellFormat.topPadding).toEqual(30.0);
    expect(table.firstRow.cells.at(1).cellFormat.bottomPadding).toEqual(30.0);
  });


  test('SetRowFormatting', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.rowFormat
    //ExFor:HeightRule
    //ExFor:aw.Tables.RowFormat.height
    //ExFor:aw.Tables.RowFormat.heightRule
    //ExSummary:Shows how to format rows with a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1.");

    // Start a second row, and then configure its height. The builder will apply these settings to
    // its current row, as well as any new rows it creates afterwards.
    builder.endRow();

    let rowFormat = builder.rowFormat;
    rowFormat.height = 100;
    rowFormat.heightRule = aw.HeightRule.Exactly;

    builder.insertCell();
    builder.write("Row 2, cell 1.");
    builder.endTable();

    // The first row was unaffected by the padding reconfiguration and still holds the default values.
    expect(table.rows.at(0).rowFormat.height).toEqual(0.0);
    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.Auto);

    expect(table.rows.at(1).rowFormat.height).toEqual(100.0);
    expect(table.rows.at(1).rowFormat.heightRule).toEqual(aw.HeightRule.Exactly);

    doc.save(base.artifactsDir + "DocumentBuilder.SetRowFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SetRowFormatting.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.rows.at(0).rowFormat.height).toEqual(0.0);
    expect(table.rows.at(0).rowFormat.heightRule).toEqual(aw.HeightRule.Auto);

    expect(table.rows.at(1).rowFormat.height).toEqual(100.0);
    expect(table.rows.at(1).rowFormat.heightRule).toEqual(aw.HeightRule.Exactly);
  });


  test('InsertFootnote', () => {
    //ExStart
    //ExFor:FootnoteType
    //ExFor:aw.DocumentBuilder.insertFootnote(FootnoteType,String)
    //ExFor:aw.DocumentBuilder.insertFootnote(FootnoteType,String,String)
    //ExSummary:Shows how to reference text with a footnote and an endnote.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert some text and mark it with a footnote with the IsAuto property set to "true" by default,
    // so the marker seen in the body text will be auto-numbered at "1",
    // and the footnote will appear at the bottom of the page.
    builder.write("This text will be referenced by a footnote.");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote comment regarding referenced text.");

    // Insert more text and mark it with an endnote with a custom reference mark,
    // which will be used in place of the number "2" and set "IsAuto" to false.
    builder.write("This text will be referenced by an endnote.");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote comment regarding referenced text.", "CustomMark");

    // Footnotes always appear at the bottom of their referenced text,
    // so this page break will not affect the footnote.
    // On the other hand, endnotes are always at the end of the document
    // so that this page break will push the endnote down to the next page.
    builder.insertBreak(aw.BreakType.PageBreak);

    doc.save(base.artifactsDir + "DocumentBuilder.insertFootnote.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertFootnote.docx");

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote comment regarding referenced text.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, false, "CustomMark",
      "CustomMark Endnote comment regarding referenced text.", doc.getFootnote(1, true));
  });


  test('ApplyBordersAndShading', () => {
    //ExStart
    //ExFor:aw.BorderCollection.item(BorderType)
    //ExFor:Shading
    //ExFor:TextureIndex
    //ExFor:aw.ParagraphFormat.shading
    //ExFor:aw.Shading.texture
    //ExFor:aw.Shading.backgroundPatternColor
    //ExFor:aw.Shading.foregroundPatternColor
    //ExSummary:Shows how to decorate text with borders and shading.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let borders = builder.paragraphFormat.borders;
    borders.distanceFromText = 20;
    borders.at(aw.BorderType.Left).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Right).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Top).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Bottom).lineStyle = aw.LineStyle.Double;

    let shading = builder.paragraphFormat.shading;
    shading.texture = aw.TextureIndex.TextureDiagonalCross;
    shading.backgroundPatternColor = "#F08080";
    shading.foregroundPatternColor = "#FFA07A";

    builder.write("This paragraph is formatted with a double border and shading.");
    doc.save(base.artifactsDir + "DocumentBuilder.ApplyBordersAndShading.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.ApplyBordersAndShading.docx");
    borders = doc.firstSection.body.firstParagraph.paragraphFormat.borders;

    expect(borders.distanceFromText).toEqual(20.0);
    expect(borders.at(aw.BorderType.Left).lineStyle).toEqual(aw.LineStyle.Double);
    expect(borders.at(aw.BorderType.Right).lineStyle).toEqual(aw.LineStyle.Double);
    expect(borders.at(aw.BorderType.Top).lineStyle).toEqual(aw.LineStyle.Double);
    expect(borders.at(aw.BorderType.Bottom).lineStyle).toEqual(aw.LineStyle.Double);

    expect(shading.texture).toEqual(aw.TextureIndex.TextureDiagonalCross);
    expect(shading.backgroundPatternColor).toEqual("#F08080");
    expect(shading.foregroundPatternColor).toEqual("#FFA07A");
  });


  test('DeleteRow', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.deleteRow
    //ExSummary:Shows how to delete a row from a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1.");
    builder.insertCell();
    builder.write("Row 1, cell 2.");
    builder.endRow();
    builder.insertCell();
    builder.write("Row 2, cell 1.");
    builder.insertCell();
    builder.write("Row 2, cell 2.");
    builder.endTable();

    expect(table.rows.count).toEqual(2);

    // Delete the first row of the first table in the document.
    builder.deleteRow(0, 0);

    expect(table.rows.count).toEqual(1);
    expect(table.getText().trim()).toEqual("Row 2, cell 1.\u0007Row 2, cell 2.\u0007\u0007");
    //ExEnd
  });


  test.each([false,
    true])('AppendDocumentAndResolveStyles', (keepSourceNumbering) => {
    //ExStart
    //ExFor:aw.Document.appendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while appending a document.
    // Load a document with text in a custom style and clone it.
    let srcDoc = new aw.Document(base.myDir + "Custom list numbering.docx");
    let dstDoc = srcDoc.clone();

    // We now have two documents, each with an identical style named "CustomStyle".
    // Change the text color for one of the styles to set it apart from the other.
    dstDoc.styles.at("CustomStyle").font.color = "#8B0000";

    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    let options = new aw.ImportFormatOptions();
    options.keepSourceNumbering = keepSourceNumbering;

    // Joining two documents that have different styles that share the same name causes a style clash.
    // We can specify an import format mode while appending documents to resolve this clash.
    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepDifferentStyles, options);
    dstDoc.updateListLabels();

    dstDoc.save(base.artifactsDir + "DocumentBuilder.AppendDocumentAndResolveStyles.docx");
    //ExEnd
  });


  test.each([false,
    true])('InsertDocumentAndResolveStyles', (keepSourceNumbering) => {
    //ExStart
    //ExFor:aw.Document.appendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while inserting a document.
    let dstDoc = new aw.Document();
    let builder = new aw.DocumentBuilder(dstDoc);
    builder.insertBreak(aw.BreakType.ParagraphBreak);

    dstDoc.lists.add(aw.Lists.ListTemplate.NumberDefault);
    let list = dstDoc.lists.at(0);

    builder.listFormat.list = list;

    for (let i = 1; i <= 15; i++)
      builder.write(`List Item ${i}\n`);

    let attachDoc = dstDoc.clone();

    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    let importOptions = new aw.ImportFormatOptions();
    importOptions.keepSourceNumbering = keepSourceNumbering;

    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.insertDocument(attachDoc, aw.ImportFormatMode.KeepSourceFormatting, importOptions);

    dstDoc.save(base.artifactsDir + "DocumentBuilder.InsertDocumentAndResolveStyles.docx");
    //ExEnd
  });


  test.each([false,
    true])('LoadDocumentWithListNumbering', (keepSourceNumbering) => {
    //ExStart
    //ExFor:aw.Document.appendDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to manage list style clashes while appending a clone of a document to itself.
    let srcDoc = new aw.Document(base.myDir + "List item.docx");
    let dstDoc = new aw.Document(base.myDir + "List item.docx");

    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    let builder = new aw.DocumentBuilder(dstDoc);
    builder.moveToDocumentEnd();
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    let options = new aw.ImportFormatOptions();
    options.keepSourceNumbering = keepSourceNumbering;
    builder.insertDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);

    dstDoc.updateListLabels();
    //ExEnd
  });


  test.each([true,
    false])('IgnoreTextBoxes', (ignoreTextBoxes) => {
    //ExStart
    //ExFor:aw.ImportFormatOptions.ignoreTextBoxes
    //ExSummary:Shows how to manage text box formatting while appending a document.
    // Create a document that will have nodes from another document inserted into it.
    let dstDoc = new aw.Document();
    let builder = new aw.DocumentBuilder(dstDoc);

    builder.writeln("Hello world!");

    // Create another document with a text box, which we will import into the first document.
    let srcDoc = new aw.Document();
    builder = new aw.DocumentBuilder(srcDoc);

    let textBox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 300, 100);
    builder.moveTo(textBox.firstParagraph);
    builder.paragraphFormat.style.font.name = "Courier New";
    builder.paragraphFormat.style.font.size = 24;
    builder.write("Textbox contents");

    // Set a flag to specify whether to clear or preserve text box formatting
    // while importing them to other documents.
    let importFormatOptions = new aw.ImportFormatOptions();
    importFormatOptions.ignoreTextBoxes = ignoreTextBoxes;

    // Import the text box from the source document into the destination document,
    // and then verify whether we have preserved the styling of its text contents.
    let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KeepSourceFormatting, importFormatOptions);
    let importedTextBox = importer.importNode(textBox, true).asShape();
    dstDoc.firstSection.body.paragraphs.at(1).appendChild(importedTextBox);

    if (ignoreTextBoxes)
    {
      expect(importedTextBox.firstParagraph.runs.at(0).font.size).toEqual(12.0);
      expect(importedTextBox.firstParagraph.runs.at(0).font.name).toEqual("Times New Roman");
    }
    else
    {
      expect(importedTextBox.firstParagraph.runs.at(0).font.size).toEqual(24.0);
      expect(importedTextBox.firstParagraph.runs.at(0).font.name).toEqual("Courier New");
    }

    dstDoc.save(base.artifactsDir + "DocumentBuilder.ignoreTextBoxes.docx");
    //ExEnd
  });


  test.each([false,
    true])('MoveToField', (moveCursorToAfterTheField) => {
    //ExStart
    //ExFor:aw.DocumentBuilder.moveToField
    //ExSummary:Shows how to move a document builder's node insertion point cursor to a specific field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a field using the DocumentBuilder and add a run of text after it.
    let field = builder.insertField(" AUTHOR \"John Doe\" ");

    // The builder's cursor is currently at end of the document.
    expect(builder.currentNode).toBe(null);

    // Move the cursor to the field while specifying whether to place that cursor before or after the field.
    builder.moveToField(field, moveCursorToAfterTheField);

    // Note that the cursor is outside of the field in both cases.
    // This means that we cannot edit the field using the builder like this.
    // To edit a field, we can use the builder's MoveTo method on a field's FieldStart
    // or FieldSeparator node to place the cursor inside.
    if (moveCursorToAfterTheField)
    {
      expect(builder.currentNode).toBe(null);
      builder.write(" Text immediately after the field.");

      expect(doc.getText().trim()).toEqual("\u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015 Text immediately after the field.");
    }
    else
    {
      expect(builder.currentNode).toEqual(field.start);
      builder.write("Text immediately before the field. ");

      expect(doc.getText().trim()).toEqual("Text immediately before the field. \u0013 AUTHOR \"John Doe\" \u0014John Doe\u0015");
    }
    //ExEnd
  });


  test('InsertOleObjectException', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    expect(() => builder.insertOleObject("", "checkbox", false, true, null)).toThrow("The value cannot be an empty string. (Parameter 'path')");
  });


  test('InsertPieChart', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertChart(ChartType, Double, Double)
    //ExSummary:Shows how to insert a pie chart into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let chart = builder.insertChart(aw.Drawing.Charts.ChartType.Pie, aw.ConvertUtil.pixelToPoint(300), 
      aw.ConvertUtil.pixelToPoint(300)).chart;
    expect(aw.ConvertUtil.pixelToPoint(300)).toEqual(225.0);
    chart.series.clear();
    chart.series.add("My fruit",
      [ "Apples", "Bananas", "Cherries" ],
      [ 1.3, 2.2, 1.5 ]);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertPieChart.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertPieChart.docx");
    let chartShape = doc.getShape(0, true);

    expect(chartShape.chart.title.text).toEqual("Chart Title");
    expect(chartShape.width).toEqual(225.0);
    expect(chartShape.height).toEqual(225.0);
  });


  test('InsertChartRelativePosition', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to specify position and wrapping while inserting a chart.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertChart(aw.Drawing.Charts.ChartType.Pie, aw.Drawing.RelativeHorizontalPosition.Margin, 100, aw.Drawing.RelativeVerticalPosition.Margin,
      100, 200, 100, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertedChartRelativePosition.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertedChartRelativePosition.docx");
    let chartShape = doc.getShape(0, true);

    expect(chartShape.top).toEqual(100.0);
    expect(chartShape.left).toEqual(100.0);
    expect(chartShape.width).toEqual(200.0);
    expect(chartShape.height).toEqual(100.0);
    expect(chartShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(chartShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(chartShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);
  });


  test('InsertField', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertField(String)
    //ExFor:Field
    //ExFor:aw.Fields.Field.result
    //ExFor:aw.Fields.Field.getFieldCode
    //ExFor:aw.Fields.Field.type
    //ExFor:FieldType
    //ExSummary:Shows how to insert a field into a document using a field code.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField("DATE \\@ \"dddd, MMMM dd, yyyy\"");

    expect(field.type).toEqual(aw.Fields.FieldType.FieldDate);
    expect(field.getFieldCode()).toEqual("DATE \\@ \"dddd, MMMM dd, yyyy\"");

    // This overload of the InsertField method automatically updates inserted fields.
    expect((Date.now() - Date.parse(field.result)) / 86400000).toBeLessThanOrEqual(1);
    //ExEnd
  });

  
  test.each([false,
    true])('InsertFieldAndUpdate', (updateInsertedFieldsImmediately) => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertField(FieldType, Boolean)
    //ExFor:aw.Fields.Field.update
    //ExSummary:Shows how to insert a field into a document using FieldType.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert two fields while passing a flag which determines whether to update them as the builder inserts them.
    // In some cases, updating fields could be computationally expensive, and it may be a good idea to defer the update.
    doc.builtInDocumentProperties.author = "John Doe";
    builder.write("This document was written by ");
    builder.insertField(aw.Fields.FieldType.FieldAuthor, updateInsertedFieldsImmediately);

    builder.insertParagraph();
    builder.write("\nThis is page ");
    builder.insertField(aw.Fields.FieldType.FieldPage, updateInsertedFieldsImmediately);

    expect(doc.range.fields.at(0).getFieldCode()).toEqual(" AUTHOR ");
    expect(doc.range.fields.at(1).getFieldCode()).toEqual(" PAGE ");

    if (updateInsertedFieldsImmediately)
    {
      expect(doc.range.fields.at(0).result).toEqual("John Doe");
      expect(doc.range.fields.at(1).result).toEqual("1");
    }
    else
    {
      expect(doc.range.fields.at(0).result).toEqual('');
      expect(doc.range.fields.at(1).result).toEqual('');

      // We will need to update these fields using the update methods manually.
      doc.range.fields.at(0).update();

      expect(doc.range.fields.at(0).result).toEqual("John Doe");

      doc.updateFields();

      expect(doc.range.fields.at(1).result).toEqual("1");
    }
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.getText().trim()).toEqual("This document was written by \u0013 AUTHOR \u0014John Doe\u0015" +
                            "\r\rThis is page \u0013 PAGE \u00141\u0015");

    TestUtil.verifyField(aw.Fields.FieldType.FieldAuthor, " AUTHOR ", "John Doe", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldPage, " PAGE ", "1", doc.range.fields.at(1));
  });

/* TODO IFieldResultFormatter not supported
    //ExStart
    //ExFor:IFieldResultFormatter
    //ExFor:IFieldResultFormatter.Format(Double, GeneralFormat)
    //ExFor:IFieldResultFormatter.Format(String, GeneralFormat)
    //ExFor:IFieldResultFormatter.FormatDateTime(DateTime, String, CalendarType)
    //ExFor:IFieldResultFormatter.FormatNumeric(Double, String)
    //ExFor:FieldOptions.ResultFormatter
    //ExFor:CalendarType
    //ExSummary:Shows how to automatically apply a custom format to field results as the fields are updated.
  test('FieldResultFormatting', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let formatter = new FieldResultFormatter("${0}", "Date: {0}", "Item # {0}:");
    doc.fieldOptions.resultFormatter = formatter;

    // Our field result formatter applies a custom format to newly created fields of three types of formats.
    // Field result formatters apply new formatting to fields as they are updated,
    // which happens as soon as we create them using this InsertField method overload.
    // 1 -  Numeric:
    builder.insertField(" = 2 + 3 \\# $###");

    expect(doc.range.fields.at(0).result).toEqual("$5");
    expect(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.numeric)).toEqual(1);

    // 2 -  Date/time:
    builder.insertField("DATE \\@ \"d MMMM yyyy\"");

    expect(doc.range.fields.at(1).result.StartsWith("Date: ")).toEqual(true);
    expect(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.dateTime)).toEqual(1);

    // 3 -  General:
    builder.insertField("QUOTE \"2\" \\* Ordinal");

    expect(doc.range.fields.at(2).result).toEqual("Item # 2:");
    expect(formatter.CountFormatInvocations(FieldResultFormatter.FormatInvocationType.general)).toEqual(1);

    formatter.PrintFormatInvocations();
  });


    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
  private class FieldResultFormatter : IFieldResultFormatter
  {
    public FieldResultFormatter(string numberFormat, string dateFormat, string generalFormat)
    {
      mNumberFormat = numberFormat;
      mDateFormat = dateFormat;
      mGeneralFormat = generalFormat;
    }

    public string FormatNumeric(double value, string format)
    {
      if (string.IsNullOrEmpty(mNumberFormat)) 
        return null;

      string newValue = String.format(mNumberFormat, value);
      FormatInvocations.add(new FormatInvocation(FormatInvocationType.numeric, value, format, newValue));
      return newValue;
    }

    public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
    {
      if (string.IsNullOrEmpty(mDateFormat))
        return null;

      string newValue = String.format(mDateFormat, value);
      FormatInvocations.add(new FormatInvocation(FormatInvocationType.dateTime, `${value} (${calendarType})`, format, newValue));
      return newValue;
    }

    public string Format(string value, GeneralFormat format)
    {
      return Format((object)value, format);
    }

    public string Format(double value, GeneralFormat format)
    {
      return Format((object)value, format);
    }

    private string Format(object value, GeneralFormat format)
    {
      if (string.IsNullOrEmpty(mGeneralFormat))
        return null;

      string newValue = String.format(mGeneralFormat, value);
      FormatInvocations.add(new FormatInvocation(FormatInvocationType.general, value, format.toString(), newValue));
      return newValue;
    }

    public int CountFormatInvocations(FormatInvocationType formatInvocationType)
    {
      if (formatInvocationType == FormatInvocationType.all)
        return FormatInvocations.count;
      return FormatInvocations.count(f => f.FormatInvocationType == formatInvocationType);
    }

    public void PrintFormatInvocations()
    {
      for (let f of FormatInvocations)
        console.log(`Invocation type:\t${f.FormatInvocationType}\n` +
                `\tOriginal value:\t\t${f.value}\n` +
                `\tOriginal format:\t${f.OriginalFormat}\n` +
                `\tNew value:\t\t\t${f.newValue}\n`);
    }

    private readonly string mNumberFormat;
    private readonly string mDateFormat;
    private readonly string mGeneralFormat; 
    private List<FormatInvocation> FormatInvocations { get; } = new aw.Lists.List<FormatInvocation>();

    private class FormatInvocation
    {
      public FormatInvocationType FormatInvocationType { get; }
      public object Value { get; }
      public string OriginalFormat { get; }
      public string NewValue { get; }

      public FormatInvocation(FormatInvocationType formatInvocationType, object value, string originalFormat, string newValue)
      {
        Value = value;
        FormatInvocationType = formatInvocationType;
        OriginalFormat = originalFormat;
        NewValue = newValue;
      }
    }

    public enum FormatInvocationType
    {
      Numeric, DateTime, General, All
    }
  }
    //ExEnd
*/    


  //[Test, Ignore("Failed")]
  test.skip('InsertVideoWithUrl - Original test failed.', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOnlineVideo(String, Double, Double)
    //ExSummary:Shows how to insert an online video into a document using a URL.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertOnlineVideo("https://youtu.be/g1N9ke8Prmk", 360, 270);

    // We can watch the video from Microsoft Word by clicking on the shape.
    doc.save(base.artifactsDir + "DocumentBuilder.InsertVideoWithUrl.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertVideoWithUrl.docx");
    let shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(480, 360, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.href).toEqual("https://youtu.be/t_1LYZ102RA");

    expect(shape.width).toEqual(360.0);
    expect(shape.height).toEqual(270.0);
  });


  test('InsertUnderline', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.underline
    //ExSummary:Shows how to format text inserted by a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.underline = aw.Underline.Dash;
    builder.font.color = "#0000FF";
    builder.font.size = 32;

    // The builder applies formatting to its current paragraph and any new text added by it afterward.
    builder.writeln("Large, blue, and underlined text.");

    doc.save(base.artifactsDir + "DocumentBuilder.InsertUnderline.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertUnderline.docx");
    let firstRun = doc.firstSection.body.firstParagraph.runs.at(0);

    expect(firstRun.getText().trim()).toEqual("Large, blue, and underlined text.");
    expect(firstRun.font.underline).toEqual(aw.Underline.Dash);
    expect(firstRun.font.color).toEqual("#0000FF");
    expect(firstRun.font.size).toEqual(32.0);
  });


  test('CurrentStory', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.currentStory
    //ExSummary:Shows how to work with a document builder's current story.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A Story is a type of node that has child Paragraph nodes, such as a Body.
    expect(doc.firstSection.body.referenceEquals(builder.currentStory)).toBe(true);
    expect(builder.currentParagraph.parentNode.referenceEquals(builder.currentStory)).toBe(true);
    expect(builder.currentStory.storyType == aw.StoryType.MainText).toBe(true);

    builder.currentStory.appendParagraph("Text added to current Story.");

    // A Story can also contain tables.
    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1");
    builder.insertCell();
    builder.write("Row 1, cell 2");
    builder.endTable();

    expect(builder.currentStory.tables.contains(table)).toEqual(true);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    expect(doc.firstSection.body.tables.count).toEqual(1);
    expect(doc.firstSection.body.getText().trim()).toEqual("Row 1, cell 1\u0007Row 1, cell 2\u0007\u0007\rText added to current Story.");
  });


  test('InsertOleObjects', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOleObject(Stream, String, Boolean, Stream)
    //ExSummary:Shows how to use document builder to embed OLE objects in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a Microsoft Excel spreadsheet from the local file system
    // into the document while keeping its default appearance.
    let spreadsheetData = base.loadFileToBuffer(base.myDir + "Spreadsheet.xlsx");
    builder.writeln("Spreadsheet Ole object:");
    // If 'presentation' is omitted and 'asIcon' is set, this overloaded method selects
    // the icon according to 'progId' and uses the predefined icon caption.
    builder.insertOleObject(spreadsheetData, "OleObject.xlsx", false, null);

    // Insert a Microsoft Powerpoint presentation as an OLE object.
    // This time, it will have an image downloaded from the web for an icon.
    let powerpointData = base.loadFileToBuffer(base.myDir + "Presentation.pptx");
    let imgData = base.loadFileToBuffer(base.imageDir + "Logo.jpg");

    builder.insertParagraph();
    builder.writeln("Powerpoint Ole object:");
    builder.insertOleObject(powerpointData, "OleObject.pptx", true, imgData);

    // Double-click these objects in Microsoft Word to open
    // the linked files using their respective applications.
    doc.save(base.artifactsDir + "DocumentBuilder.InsertOleObjects.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertOleObjects.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(2);

    let shape = doc.getShape(0, true);
    expect(shape.oleFormat.iconCaption).toEqual("");
    expect(shape.oleFormat.oleIcon).toEqual(false);

    shape = doc.getShape(1, true);
    expect(shape.oleFormat.iconCaption).toEqual("Unknown");
    expect(shape.oleFormat.oleIcon).toEqual(true);
  });


  test('InsertStyleSeparator', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertStyleSeparator
    //ExSummary:Shows how to work with style separators.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Each paragraph can only have one style.
    // The InsertStyleSeparator method allows us to work around this limitation.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.write("This text is in a Heading style. ");
    builder.insertStyleSeparator();

    let paraStyle = builder.document.styles.add(aw.StyleType.Paragraph, "MyParaStyle");
    paraStyle.font.bold = false;
    paraStyle.font.size = 8;
    paraStyle.font.name = "Arial";

    builder.paragraphFormat.styleName = paraStyle.name;
    builder.write("This text is in a custom style. ");

    // Calling the InsertStyleSeparator method creates another paragraph,
    // which can have a different style to the previous. There will be no break between paragraphs.
    // The text in the output document will look like one paragraph with two styles.
    expect(doc.firstSection.body.paragraphs.count).toEqual(2);
    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.style.name).toEqual("Heading 1");
    expect(doc.firstSection.body.paragraphs.at(1).paragraphFormat.style.name).toEqual("MyParaStyle");

    doc.save(base.artifactsDir + "DocumentBuilder.insertStyleSeparator.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertStyleSeparator.docx");

    expect(doc.firstSection.body.paragraphs.count).toEqual(2);
    expect(doc.getText().trim()).toEqual("This text is in a Heading style. \r This text is in a custom style.");
    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.style.name).toEqual("Heading 1");
    expect(doc.firstSection.body.paragraphs.at(1).paragraphFormat.style.name).toEqual("MyParaStyle");
    expect(doc.firstSection.body.paragraphs.at(1).runs.at(0).getText()).toEqual(" ");
    TestUtil.docPackageFileContainsString("w:rPr><w:vanish /><w:specVanish /></w:rPr>", 
      base.artifactsDir + "DocumentBuilder.insertStyleSeparator.docx", "document.xml");
    TestUtil.docPackageFileContainsString("<w:t xml:space=\"preserve\"> </w:t>", 
      base.artifactsDir + "DocumentBuilder.insertStyleSeparator.docx", "document.xml");
  });


  test.skip('InsertDocument - Original test failed. Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks.', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertDocument(Document, ImportFormatMode)
    //ExFor:ImportFormatMode
    //ExSummary:Shows how to insert a document into another document.
    let doc = new aw.Document(base.myDir + "Document.docx");

    let builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.insertBreak(aw.BreakType.PageBreak);

    let docToInsert = new aw.Document(base.myDir + "Formatted elements.docx");

    builder.insertDocument(docToInsert, aw.ImportFormatMode.KeepSourceFormatting);
    builder.document.save(base.artifactsDir + "DocumentBuilder.insertDocument.docx");
    //ExEnd

    expect(doc.styles.count).toEqual(29);
    expect(DocumentHelper.compareDocs(base.artifactsDir + "DocumentBuilder.insertDocument.docx", base.goldsDir + "DocumentBuilder.insertDocument Gold.docx")).toEqual(true);
  });


  test('SmartStyleBehavior', () => {
    //ExStart
    //ExFor:ImportFormatOptions
    //ExFor:aw.ImportFormatOptions.smartStyleBehavior
    //ExFor:aw.DocumentBuilder.insertDocument(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve duplicate styles while inserting documents.
    let dstDoc = new aw.Document();
    let builder = new aw.DocumentBuilder(dstDoc);

    let myStyle = builder.document.styles.add(aw.StyleType.Paragraph, "MyStyle");
    myStyle.font.size = 14;
    myStyle.font.name = "Courier New";
    myStyle.font.color = "#0000FF";

    builder.paragraphFormat.styleName = myStyle.name;
    builder.writeln("Hello world!");

    // Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
    // If we insert the clone into the original document, the two styles with the same name will cause a clash.
    let srcDoc = dstDoc.clone(); // TODO cast to Document
    srcDoc.styles.at("MyStyle").font.color = "#FF0000";

    // When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
    // Aspose.words will resolve style clashes by converting source document styles.
    // with the same names as destination styles into direct paragraph attributes.
    let options = new aw.ImportFormatOptions();
    options.smartStyleBehavior = true;

    builder.insertDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);

    dstDoc.save(base.artifactsDir + "DocumentBuilder.smartStyleBehavior.docx");
    //ExEnd

    dstDoc = new aw.Document(base.artifactsDir + "DocumentBuilder.smartStyleBehavior.docx");

    expect(dstDoc.styles.at("MyStyle").font.color).toEqual("#0000FF");
    expect(dstDoc.firstSection.body.paragraphs.at(0).paragraphFormat.style.name).toEqual("MyStyle");

    expect(dstDoc.firstSection.body.paragraphs.at(1).paragraphFormat.style.name).toEqual("Normal");
    expect(dstDoc.firstSection.body.paragraphs.at(1).runs.at(0).font.size).toEqual(14);
    expect(dstDoc.firstSection.body.paragraphs.at(1).runs.at(0).font.name).toEqual("Courier New");
    expect(dstDoc.firstSection.body.paragraphs.at(1).runs.at(0).font.color).toEqual("#FF0000");
  });


  test.skip('EmphasesWarningSourceMarkdown - TODO: warningCallback not supported yet.', () => {
    let doc = new aw.Document(base.myDir + "Emphases markdown warning.docx");
            
    let warnings = new aw.WarningInfoCollection();
    doc.warningCallback = warnings;
    doc.save(base.artifactsDir + "DocumentBuilder.EmphasesWarningSourceMarkdown.md");
 
    for (let warningInfo of warnings)
    {
      if (warningInfo.source == aw.WarningSource.Markdown)
        expect(warningInfo.description).toEqual("The (*, 0:11) cannot be properly written into Markdown.");
    }
  });


  test('DoNotIgnoreHeaderFooter', () => {
    //ExStart
    //ExFor:aw.ImportFormatOptions.ignoreHeaderFooter
    //ExSummary:Shows how to specifies ignoring or not source formatting of headers/footers content.
    let dstDoc = new aw.Document(base.myDir + "Document.docx");
    let srcDoc = new aw.Document(base.myDir + "Header and footer types.docx");
 
    let importFormatOptions = new aw.ImportFormatOptions();
    importFormatOptions.ignoreHeaderFooter = false;
 
    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, importFormatOptions);

    dstDoc.save(base.artifactsDir + "DocumentBuilder.DoNotIgnoreHeaderFooter.docx");
    //ExEnd
  });


  function markdownDocumentEmphases()
  {
    let builder = new aw.DocumentBuilder();

      // Bold and Italic are represented as Font.Bold and Font.Italic.
    builder.font.italic = true;
    builder.writeln("This text will be italic");

      // Use clear formatting if we don't want to combine styles between paragraphs.
    builder.font.clearFormatting();

    builder.font.bold = true;
    builder.writeln("This text will be bold");

    builder.font.clearFormatting();

    builder.font.italic = true;
    builder.write("You ");
    builder.font.bold = true;
    builder.write("can");
    builder.font.bold = false;
    builder.writeln(" combine them");

    builder.font.clearFormatting();

    builder.font.strikeThrough = true;
    builder.writeln("This text will be strikethrough");

      // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis.
    builder.document.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentInlineCode()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

      // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`).
      // If number of backticks is missed, then one backtick will be used by default.
    let inlineCode1BackTicks = doc.styles.add(aw.StyleType.Character, "InlineCode");
    builder.font.style = inlineCode1BackTicks;
    builder.writeln("Text with InlineCode style with one backtick");

      // Use optional dot (.) and number of backticks (`).
      // There will be 3 backticks.
    let inlineCode3BackTicks = doc.styles.add(aw.StyleType.Character, "InlineCode.3");
    builder.font.style = inlineCode3BackTicks;
    builder.writeln("Text with InlineCode style with 3 backticks");

    builder.document.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentHeadings()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

      // By default, Heading styles in Word may have bold and italic formatting.
      // If we do not want text to be emphasized, set these properties explicitly to false.
      // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true.
    builder.font.bold = false;
    builder.font.italic = false;

      // Create for one heading for each level.
    builder.paragraphFormat.styleName = "Heading 1";
    builder.font.italic = true;
    builder.writeln("This is an italic H1 tag");

      // Reset our styles from the previous paragraph to not combine styles between paragraphs.
    builder.font.bold = false;
    builder.font.italic = false;

      // Structure-enhanced text heading can be added through style inheritance.
    let setextHeading1 = doc.styles.add(aw.StyleType.Paragraph, "SetextHeading1");
    builder.paragraphFormat.style = setextHeading1;
    doc.styles.at("SetextHeading1").baseStyleName = "Heading 1";
    builder.writeln("SetextHeading 1");

    builder.paragraphFormat.styleName = "Heading 2";
    builder.writeln("This is an H2 tag");

    builder.font.bold = false;
    builder.font.italic = false;

    let setextHeading2 = doc.styles.add(aw.StyleType.Paragraph, "SetextHeading2");
    builder.paragraphFormat.style = setextHeading2;
    doc.styles.at("SetextHeading2").baseStyleName = "Heading 2";
    builder.writeln("SetextHeading 2");

    builder.paragraphFormat.style = doc.styles.at("Heading 3");
    builder.writeln("This is an H3 tag");

    builder.font.bold = false;
    builder.font.italic = false;

    builder.paragraphFormat.style = doc.styles.at("Heading 4");
    builder.font.bold = true;
    builder.writeln("This is an bold H4 tag");

    builder.font.bold = false;
    builder.font.italic = false;

    builder.paragraphFormat.style = doc.styles.at("Heading 5");
    builder.font.italic = true;
    builder.font.bold = true;
    builder.writeln("This is an italic and bold H5 tag");

    builder.font.bold = false;
    builder.font.italic = false;

    builder.paragraphFormat.style = doc.styles.at("Heading 6");
    builder.writeln("This is an H6 tag");

    doc.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentBlockquotes()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

      // By default, the document stores blockquote style for the first level.
    builder.paragraphFormat.styleName = "Quote";
    builder.writeln("Blockquote");

      // Create styles for nested levels through style inheritance.
    let quoteLevel2 = doc.styles.add(aw.StyleType.Paragraph, "Quote1");
    builder.paragraphFormat.style = quoteLevel2;
    doc.styles.at("Quote1").baseStyleName = "Quote";
    builder.writeln("1. Nested blockquote");

    let quoteLevel3 = doc.styles.add(aw.StyleType.Paragraph, "Quote2");
    builder.paragraphFormat.style = quoteLevel3;
    doc.styles.at("Quote2").baseStyleName = "Quote1";
    builder.font.italic = true;
    builder.writeln("2. Nested italic blockquote");

    let quoteLevel4 = doc.styles.add(aw.StyleType.Paragraph, "Quote3");
    builder.paragraphFormat.style = quoteLevel4;
    doc.styles.at("Quote3").baseStyleName = "Quote2";
    builder.font.italic = false;
    builder.font.bold = true;
    builder.writeln("3. Nested bold blockquote");

    let quoteLevel5 = doc.styles.add(aw.StyleType.Paragraph, "Quote4");
    builder.paragraphFormat.style = quoteLevel5;
    doc.styles.at("Quote4").baseStyleName = "Quote3";
    builder.font.bold = false;
    builder.writeln("4. Nested blockquote");

    let quoteLevel6 = doc.styles.add(aw.StyleType.Paragraph, "Quote5");
    builder.paragraphFormat.style = quoteLevel6;
    doc.styles.at("Quote5").baseStyleName = "Quote4";
    builder.writeln("5. Nested blockquote");

    let quoteLevel7 = doc.styles.add(aw.StyleType.Paragraph, "Quote6");
    builder.paragraphFormat.style = quoteLevel7;
    doc.styles.at("Quote6").baseStyleName = "Quote5";
    builder.font.italic = true;
    builder.font.bold = true;
    builder.writeln("6. Nested italic bold blockquote");

    doc.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentIndentedCode()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.writeln("\n");
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

    let indentedCode = doc.styles.add(aw.StyleType.Paragraph, "IndentedCode");
    builder.paragraphFormat.style = indentedCode;
    builder.writeln("This is an indented code");

    doc.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentFencedCode()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.writeln("\n");
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

    let fencedCode = doc.styles.add(aw.StyleType.Paragraph, "FencedCode");
    builder.paragraphFormat.style = fencedCode;
    builder.writeln("This is a fenced code");

    let fencedCodeWithInfo = doc.styles.add(aw.StyleType.Paragraph, "FencedCode.C#");
    builder.paragraphFormat.style = fencedCodeWithInfo;
    builder.writeln("This is a fenced code with info string");

    doc.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentHorizontalRule()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

      // Insert HorizontalRule that will be present in .md file as '-----'.
    builder.insertHorizontalRule();

    builder.document.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  function markdownDocumentBulletedList()
  {
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let builder = new aw.DocumentBuilder(doc);

      // Prepare our created document for further work
      // and clear paragraph formatting not to use the previous styles.
    builder.moveToDocumentEnd();
    builder.paragraphFormat.clearFormatting();
    builder.writeln("\n");

      // Bulleted lists are represented using paragraph numbering.
    builder.listFormat.applyBulletDefault();
      // There can be 3 types of bulleted lists.
      // The only diff in a numbering format of the very first level are ‘-’, ‘+’ or ‘*’ respectively.
    builder.listFormat.list.listLevels.at(0).numberFormat = "-";

    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.listIndent();
    builder.writeln("Item 2a");
    builder.writeln("Item 2b");

    builder.document.save(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
  }

  test.each([["Italic", "Normal", true, false],
    ["Bold", "Normal", false, true],
    ["ItalicBold", "Normal", true, true],
    ["Text with InlineCode style with one backtick", "InlineCode", false, false],
    ["Text with InlineCode style with 3 backticks", "InlineCode.3", false, false],
    ["This is an italic H1 tag", "Heading 1", true, false],
    ["SetextHeading 1", "SetextHeading1", false, false],
    ["This is an H2 tag", "Heading 2", false, false],
    ["SetextHeading 2", "SetextHeading2", false, false],
    ["This is an H3 tag", "Heading 3", false, false],
    ["This is an bold H4 tag", "Heading 4", false, true],
    ["This is an italic and bold H5 tag", "Heading 5", true, true],
    ["This is an H6 tag", "Heading 6", false, false],
    ["Blockquote", "Quote", false, false],
    ["1. Nested blockquote", "Quote1", false, false],
    ["2. Nested italic blockquote", "Quote2", true, false],
    ["3. Nested bold blockquote", "Quote3", false, true],
    ["4. Nested blockquote", "Quote4", false, false],
    ["5. Nested blockquote", "Quote5", false, false],
    ["6. Nested italic bold blockquote", "Quote6", true, true],
    ["This is an indented code", "IndentedCode", false, false],
    ["This is a fenced code", "FencedCode", false, false],
    ["This is a fenced code with info string", "FencedCode.C#", false, false],
    ["Item 1", "Normal", false, false]])('LoadMarkdownDocumentAndAssertContent', (text, styleName, isItalic, isBold) => {
    // Prepeare document to test.
    markdownDocumentEmphases();
    markdownDocumentInlineCode();
    markdownDocumentHeadings();
    markdownDocumentBlockquotes();
    markdownDocumentIndentedCode();
    markdownDocumentFencedCode();
    markdownDocumentHorizontalRule();
    markdownDocumentBulletedList();

    // Load created document from previous tests.
    let doc = new aw.Document(base.artifactsDir + "DocumentBuilder.MarkdownDocument.md");
    let paragraphs = doc.firstSection.body.paragraphs.toArray();

    for (let paragraph of paragraphs)
    {
      if (paragraph.runs.count != 0)
      {
        // Check that all document text has the necessary styles.
        if (paragraph.runs.at(0).text == text && !text.includes("InlineCode"))
        {
          expect(paragraph.paragraphFormat.style.name).toEqual(styleName);
          expect(paragraph.runs.at(0).font.italic).toEqual(isItalic);
          expect(paragraph.runs.at(0).font.bold).toEqual(isBold);
        }
        else if (paragraph.runs.at(0).text == text && text.includes("InlineCode"))
        {
          expect(paragraph.runs.at(0).font.styleName).toEqual(styleName);
        }
      }

      // Check that document also has a HorizontalRule present as a shape.
      let shapesCollection = doc.firstSection.body.getChildNodes(aw.NodeType.Shape, true);
      let horizontalRuleShape = shapesCollection.at(0).asShape();

      expect(shapesCollection.count == 1).toEqual(true);
      expect(horizontalRuleShape.isHorizontalRule).toEqual(true);
    }
  });


  test('InsertOnlineVideo', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an online video into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let videoUrl = "https://vimeo.com/52477838";

    // Insert a shape that plays a video from the web when clicked in Microsoft Word.
    // This rectangular shape will contain an image based on the first frame of the linked video
    // and a "play button" visual prompt. The video has an aspect ratio of 16:9.
    // We will set the shape's size to that ratio, so the image does not appear stretched.
    builder.insertOnlineVideo(videoUrl, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 0,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 0, 320, 180, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilder.insertOnlineVideo.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.insertOnlineVideo.docx");
    let shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(640, 360, aw.Drawing.ImageType.Jpeg, shape);

    expect(shape.width).toEqual(320.0);
    expect(shape.height).toEqual(180.0);
    expect(shape.left).toEqual(0.0);
    expect(shape.top).toEqual(0.0);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.TopMargin);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.LeftMargin);

    expect(shape.href).toEqual("https://vimeo.com/52477838");
  });


  test('InsertOnlineVideoCustomThumbnail', async () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOnlineVideo(String, String, Byte[], Double, Double)
    //ExFor:aw.DocumentBuilder.insertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an online video into a document with a custom thumbnail.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let videoUrl = "https://vimeo.com/52477838";
    let videoEmbedCode =
      "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " +
      "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

    const imageFileName = base.imageDir + "Logo.jpg";  
    const thumbnailImageBytes = base.loadFileToArray(imageFileName);
    const image = await jimp.Jimp.read(imageFileName); 
    const imageWidth = image.bitmap.width;
    const imageHeight = image.bitmap.height;

    // Below are two ways of creating a shape with a custom thumbnail, which links to an online video
    // that will play when we click on the shape in Microsoft Word.
    // 1 -  Insert an inline shape at the builder's node insertion cursor:
    builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, imageWidth, imageHeight);

    builder.insertBreak(aw.BreakType.PageBreak);

    // 2 -  Insert a floating shape:
    let left = builder.pageSetup.rightMargin - imageWidth;
    let top = builder.pageSetup.bottomMargin - imageHeight;

    builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes,
      aw.Drawing.RelativeHorizontalPosition.RightMargin, left, aw.Drawing.RelativeVerticalPosition.BottomMargin, top,
      imageWidth, imageHeight, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.InsertOnlineVideoCustomThumbnail.docx");
    let shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.width).toEqual(400.0);
    expect(shape.height).toEqual(400.0);
    expect(shape.left).toEqual(0.0);
    expect(shape.top).toEqual(0.0);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);

    expect(shape.href).toEqual("https://vimeo.com/52477838");

    shape = doc.getShape(1, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.width).toEqual(400.0);
    expect(shape.height).toEqual(400.0);
    expect(shape.left).toEqual(-328);
    expect(shape.top).toEqual(-328);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.BottomMargin);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.RightMargin);

    expect(shape.href).toEqual("https://vimeo.com/52477838");
  });


  test('InsertOleObjectAsIcon', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertOleObjectAsIcon(String, String, Boolean, String, String)
    //ExFor:aw.DocumentBuilder.insertOleObjectAsIcon(Stream, String, String, String)
    //ExSummary:Shows how to insert an embedded or linked OLE object as icon into the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    // the icon according to 'progId' and uses the filename for the icon caption.
    builder.insertOleObjectAsIcon(base.myDir + "Presentation.pptx", "Package", false, base.imageDir + "Logo icon.ico", "My embedded file");

    builder.insertBreak(aw.BreakType.LineBreak);

    let stream = base.loadFileToBuffer(base.myDir + "Presentation.pptx");
    // If 'iconFile' and 'iconCaption' are omitted, this overloaded method selects
    // the icon according to the file extension and uses the filename for the icon caption.
    let shape = builder.insertOleObjectAsIcon(stream, "PowerPoint.Application", base.imageDir + "Logo icon.ico",
      "My embedded file stream");

    let setOlePackage = shape.oleFormat.olePackage;
    setOlePackage.fileName = "Presentation.pptx";
    setOlePackage.displayName = "Presentation.pptx";

    doc.save(base.artifactsDir + "DocumentBuilder.insertOleObjectAsIcon.docx");
    //ExEnd
  });


  test('PreserveBlocks', () => {
    //ExStart
    //ExFor:HtmlInsertOptions
    //ExSummary:Shows how to allows better preserve borders and margins seen.
    const html = `
      <html>
        <div style='border:dotted'>
        <div style='border:solid'>
          <p>paragraph 1</p>
          <p>paragraph 2</p>
        </div>
        </div>
      </html>`;

    // Set the new mode of import HTML block-level elements.
    let insertOptions = aw.HtmlInsertOptions.PreserveBlocks;

    let builder = new aw.DocumentBuilder();
    builder.insertHtml(html, insertOptions);
    builder.document.save(base.artifactsDir + "DocumentBuilder.preserveBlocks.docx");
    //ExEnd
  });


  test('PhoneticGuide', () => {
    //ExStart
    //ExFor:aw.Run.isPhoneticGuide
    //ExFor:aw.Run.phoneticGuide
    //ExFor:aw.PhoneticGuide.baseText
    //ExFor:aw.PhoneticGuide.rubyText
    //ExSummary:Shows how to get properties of the phonetic guide.
    let doc = new aw.Document(base.myDir + "Phonetic guide.docx");

    let runs = doc.firstSection.body.firstParagraph.runs;
    // Use phonetic guide in the Asian text.
    expect(runs.at(0).isPhoneticGuide).toEqual(true);
    expect(runs.at(0).phoneticGuide.baseText).toEqual("base");
    expect(runs.at(0).phoneticGuide.rubyText).toEqual("ruby");
    //ExEnd
  });

});
