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

/// <summary>
/// Insert field into the first paragraph of the current document using field type.
/// </summary>
function insertFieldUsingFieldType(doc, fieldType, updateField, refNode, isAfter, paraIndex) {
  let para = DocumentHelper.getParagraph(doc, paraIndex);
  para.insertField(fieldType, updateField, refNode, isAfter);
}

/// <summary>
/// Insert field into the first paragraph of the current document using field code.
/// </summary>
function insertFieldUsingFieldCode(doc, fieldCode, refNode, isAfter, paraIndex) {
  let para = DocumentHelper.getParagraph(doc, paraIndex);
  para.insertField(fieldCode, refNode, isAfter);
}

/// <summary>
/// Insert field into the first paragraph of the current document using field code and field String.
/// </summary>
function insertFieldUsingFieldCodeFieldString(doc, fieldCode, fieldValue, refNode, isAfter, paraIndex) {
  let para = DocumentHelper.getParagraph(doc, paraIndex);
  para.insertField(fieldCode, fieldValue, refNode, isAfter);
}

describe("ExParagraph", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('DocumentBuilderInsertParagraph', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertParagraph
    //ExFor:aw.ParagraphFormat.firstLineIndent
    //ExFor:aw.ParagraphFormat.alignment
    //ExFor:aw.ParagraphFormat.keepTogether
    //ExFor:aw.ParagraphFormat.addSpaceBetweenFarEastAndAlpha
    //ExFor:aw.ParagraphFormat.addSpaceBetweenFarEastAndDigit
    //ExFor:aw.Paragraph.isEndOfDocument
    //ExSummary:Shows how to insert a paragraph into the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let font = builder.font;
    font.size = 16;
    font.bold = true;
    font.color = "#0000FF";
    font.name = "Arial";
    font.underline = aw.Underline.Dash;

    let paragraphFormat = builder.paragraphFormat;
    paragraphFormat.firstLineIndent = 8;
    paragraphFormat.alignment = aw.ParagraphAlignment.Justify;
    paragraphFormat.addSpaceBetweenFarEastAndAlpha = true;
    paragraphFormat.addSpaceBetweenFarEastAndDigit = true;
    paragraphFormat.keepTogether = true;

    // The "Writeln" method ends the paragraph after appending text
    // and then starts a new line, adding a new paragraph.
    builder.writeln("Hello world!");

    expect(builder.currentParagraph.isEndOfDocument).toEqual(true);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    let paragraph = doc.firstSection.body.firstParagraph;

    expect(paragraph.paragraphFormat.firstLineIndent).toEqual(8);
    expect(paragraph.paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Justify);
    expect(paragraph.paragraphFormat.addSpaceBetweenFarEastAndAlpha).toEqual(true);
    expect(paragraph.paragraphFormat.addSpaceBetweenFarEastAndDigit).toEqual(true);
    expect(paragraph.paragraphFormat.keepTogether).toEqual(true);
    expect(paragraph.getText().trim()).toEqual("Hello world!");

    let runFont = paragraph.runs.at(0).font;

    expect(runFont.size).toEqual(16.0);
    expect(runFont.bold).toEqual(true);
    expect(runFont.color).toEqual("#0000FF");
    expect(runFont.name).toEqual("Arial");
    expect(runFont.underline).toEqual(aw.Underline.Dash);
  });


  test('AppendField', () => {
    //ExStart
    //ExFor:aw.Paragraph.appendField(FieldType, Boolean)
    //ExFor:aw.Paragraph.appendField(String)
    //ExFor:aw.Paragraph.appendField(String, String)
    //ExSummary:Shows various ways of appending fields to a paragraph.
    let doc = new aw.Document();
    let paragraph = doc.firstSection.body.firstParagraph;

    // Below are three ways of appending a field to the end of a paragraph.
    // 1 -  Append a DATE field using a field type, and then update it:
    paragraph.appendField(aw.Fields.FieldType.FieldDate, true);

    // 2 -  Append a TIME field using a field code: 
    paragraph.appendField(" TIME  \\@ \"HH:mm:ss\" ");

    // 3 -  Append a QUOTE field using a field code, and get it to display a placeholder value:
    paragraph.appendField(" QUOTE \"Real value\"", "Placeholder value");

    expect(doc.range.fields.at(2).result).toEqual("Placeholder value");

    // This field will display its placeholder value until we update it.
    doc.updateFields();

    expect(doc.range.fields.at(2).result).toEqual("Real value");

    doc.save(base.artifactsDir + "Paragraph.appendField.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Paragraph.appendField.docx");

    let today = new Date();
    today.setHours(0, 0, 0, 0);
    let now = new Date();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, " DATE ", today, doc.range.fields.at(0), 0);
    TestUtil.verifyField(aw.Fields.FieldType.FieldTime, " TIME  \\@ \"HH:mm:ss\" ", now, doc.range.fields.at(1), 5000);
    TestUtil.verifyField(aw.Fields.FieldType.FieldQuote, " QUOTE \"Real value\"", "Real value", doc.range.fields.at(2));
  });


  test('InsertField', () => {
    //ExStart
    //ExFor:aw.Paragraph.insertField(string, Node, bool)
    //ExFor:aw.Paragraph.insertField(FieldType, bool, Node, bool)
    //ExFor:aw.Paragraph.insertField(string, string, Node, bool)
    //ExSummary:Shows various ways of adding fields to a paragraph.
    let doc = new aw.Document();
    let para = doc.firstSection.body.firstParagraph;

    // Below are three ways of inserting a field into a paragraph.
    // 1 -  Insert an AUTHOR field into a paragraph after one of the paragraph's child nodes:
    let run = new aw.Run(doc);
    run.text = "This run was written by ";
    para.appendChild(run);

    doc.builtInDocumentProperties.author = "John Doe";
    para.insertField(aw.Fields.FieldType.FieldAuthor, true, run, true);

    // 2 -  Insert a QUOTE field after one of the paragraph's child nodes:
    run = new aw.Run(doc);
    run.text = ".";
    para.appendChild(run);

    let field = para.insertField(" QUOTE \" Real value\" ", run, true);

    // 3 -  Insert a QUOTE field before one of the paragraph's child nodes,
    // and get it to display a placeholder value:
    para.insertField(" QUOTE \" Real value.\"", " Placeholder value.", field.start, false);

    expect(doc.range.fields.at(1).result).toEqual(" Placeholder value.");

    // This field will display its placeholder value until we update it.
    doc.updateFields();

    expect(doc.range.fields.at(1).result).toEqual(" Real value.");

    doc.save(base.artifactsDir + "Paragraph.insertField.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Paragraph.insertField.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldAuthor, " AUTHOR ", "John Doe", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldQuote, " QUOTE \" Real value.\"", " Real value.", doc.range.fields.at(1));
    TestUtil.verifyField(aw.Fields.FieldType.FieldQuote, " QUOTE \" Real value\" ", " Real value", doc.range.fields.at(2));
  });


  test('InsertFieldBeforeTextInParagraph', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldCode(doc, " AUTHOR ", null, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r");
  });


  test('InsertFieldAfterTextInParagraph', () => {
    let date = new Date().toLocaleDateString();

    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldCode(doc, " DATE ", null, true, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual(`Hello World!\u0013 DATE \u0014${date}\u0015\r`);
  });


  test('InsertFieldBeforeTextInParagraphWithoutUpdateField', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldType(doc, aw.Fields.FieldType.FieldAuthor, false, null, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r");
  });


  test('InsertFieldAfterTextInParagraphWithoutUpdateField', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldType(doc, aw.Fields.FieldType.FieldAuthor, false, null, true, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r");
  });


  test('InsertFieldWithoutSeparator', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldType(doc, aw.Fields.FieldType.FieldListNum, true, null, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("\u0013 LISTNUM \u0015Hello World!\r");
  });


  test('InsertFieldBeforeParagraphWithoutDocumentAuthor', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();
    doc.builtInDocumentProperties.author = "";

    insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r");
  });


  test('InsertFieldAfterParagraphWithoutChangingDocumentAuthor', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r");
  });


  test('InsertFieldBeforeRunText', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    //Add some text into the paragraph
    let run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

    insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r");
  });


  test('InsertFieldAfterRunText', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    // Add some text into the paragraph
    let run = DocumentHelper.insertNewRun(doc, " Hello World!", 1);

    insertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r");
  });


  // WORDSNET-12396
  test('InsertFieldEmptyParagraphWithoutUpdateField', () => {
    let doc = DocumentHelper.createDocumentWithoutDummyText();

    insertFieldUsingFieldType(doc, aw.Fields.FieldType.FieldAuthor, false, null, false, 1);

    expect(DocumentHelper.getParagraphText(doc, 1)).toEqual("\u0013 AUTHOR \u0014\u0015\f");
  });


  // WORDSNET-12397
  test('InsertFieldEmptyParagraphWithUpdateField', () => {
    let doc = DocumentHelper.createDocumentWithoutDummyText();

    insertFieldUsingFieldType(doc, aw.Fields.FieldType.FieldAuthor, true, null, false, 0);

    expect(DocumentHelper.getParagraphText(doc, 0)).toEqual("\u0013 AUTHOR \u0014Test Author\u0015\r");
  });


  test('CompositeNodeChildren', () => {
    //ExStart
    //ExFor:aw.CompositeNode.count
    //ExFor:aw.CompositeNode.getChildNodes(NodeType, Boolean)
    //ExFor:aw.CompositeNode.insertAfter``1(``0, Node)
    //ExFor:aw.CompositeNode.insertBefore``1(``0, Node)
    //ExFor:aw.CompositeNode.prependChild``1(``0)
    //ExFor:aw.Paragraph.getText
    //ExFor:Run
    //ExSummary:Shows how to add, update and delete child nodes in a CompositeNode's collection of children.
    let doc = new aw.Document();

    // An empty document, by default, has one paragraph.
    expect(doc.firstSection.body.paragraphs.count).toEqual(1);

    // Composite nodes such as our paragraph can contain other composite and inline nodes as children.
    let paragraph = doc.firstSection.body.firstParagraph;
    let paragraphText = new aw.Run(doc, "Initial text. ");
    paragraph.appendChild(paragraphText);

    // Create three more run nodes.
    let run1 = new aw.Run(doc, "Run 1. ");
    let run2 = new aw.Run(doc, "Run 2. ");
    let run3 = new aw.Run(doc, "Run 3. ");

    // The document body will not display these runs until we insert them into a composite node
    // that itself is a part of the document's node tree, as we did with the first run.
    // We can determine where the text contents of nodes that we insert
    // appears in the document by specifying an insertion location relative to another node in the paragraph.
    expect(paragraph.getText().trim()).toEqual("Initial text.");

    // Insert the second run into the paragraph in front of the initial run.
    paragraph.insertBefore(run2, paragraphText);

    expect(paragraph.getText().trim()).toEqual("Run 2. Initial text.");

    // Insert the third run after the initial run.
    paragraph.insertAfter(run3, paragraphText);

    expect(paragraph.getText().trim()).toEqual("Run 2. Initial text. Run 3.");

    // Insert the first run to the start of the paragraph's child nodes collection.
    paragraph.prependChild(run1);

    expect(paragraph.getText().trim()).toEqual("Run 1. Run 2. Initial text. Run 3.");
    expect(paragraph.getChildNodes(aw.NodeType.Any, true).count).toEqual(4);

    // We can modify the contents of the run by editing and deleting existing child nodes.
    //paragraph.getChildNodes(aw.NodeType.Run, true).toArray().at(1).text = "Updated run 2. ";
    //paragraph.getChildNodes(aw.NodeType.Run, true).remove(paragraphText);
    paragraph.getChildNodes(aw.NodeType.Run, true).toArray().at(1).asRun().text = "Updated run 2. ";
    paragraph.getChildNodes(aw.NodeType.Run, true).remove(paragraphText);

    expect(paragraph.getText().trim()).toEqual("Run 1. Updated run 2. Run 3.");
    expect(paragraph.getChildNodes(aw.NodeType.Any, true).count).toEqual(3);
    //ExEnd
  });


  test('MoveRevisions', () => {
    //ExStart
    //ExFor:aw.Paragraph.isMoveFromRevision
    //ExFor:aw.Paragraph.isMoveToRevision
    //ExFor:ParagraphCollection
    //ExFor:aw.ParagraphCollection.item(Int32)
    //ExFor:aw.Story.paragraphs
    //ExSummary:Shows how to check whether a paragraph is a move revision.
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    // This document contains "Move" revisions, which appear when we highlight text with the cursor,
    // and then drag it to move it to another location
    // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
    expect(Array.from(doc.revisions).filter(r => r.revisionType == aw.RevisionType.Moving).length).toEqual(6);

    let paragraphs = doc.firstSection.body.paragraphs;

    // Move revisions consist of pairs of "Move from", and "Move to" revisions. 
    // These revisions are potential changes to the document that we can either accept or reject.
    // Before we accept/reject a move revision, the document
    // must keep track of both the departure and arrival destinations of the text.
    // The second and the fourth paragraph define one such revision, and thus both have the same contents.
    expect(paragraphs.at(3).getText()).toEqual(paragraphs.at(1).getText());

    // The "Move from" revision is the paragraph where we dragged the text from.
    // If we accept the revision, this paragraph will disappear,
    // and the other will remain and no longer be a revision.
    expect(paragraphs.at(1).isMoveFromRevision).toEqual(true);

    // The "Move to" revision is the paragraph where we dragged the text to.
    // If we reject the revision, this paragraph instead will disappear, and the other will remain.
    expect(paragraphs.at(3).isMoveToRevision).toEqual(true);
    //ExEnd
  });


  test('RangeRevisions', () => {
    //ExStart
    //ExFor:aw.Range.revisions
    //ExSummary:Shows how to work with revisions in range.
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    let paragraph = doc.firstSection.body.firstParagraph;
    for (let revision of paragraph.range.revisions)
    {
      if (revision.revisionType == aw.RevisionType.Deletion)
        revision.accept();
    }

    // Reject the first section revisions.
    doc.firstSection.range.revisions.rejectAll();
    //ExEnd
  });


  test('GetFormatRevision', () => {
    //ExStart
    //ExFor:aw.Paragraph.isFormatRevision
    //ExSummary:Shows how to check whether a paragraph is a format revision.
    let doc = new aw.Document(base.myDir + "Format revision.docx");

    // This paragraph is a "Format" revision, which occurs when we change the formatting of existing text
    // while tracking revisions in Microsoft Word via "Review" -> "Track changes".
    expect(doc.firstSection.body.firstParagraph.isFormatRevision).toEqual(true);
    //ExEnd
  });


  test('GetFrameProperties', () => {
    //ExStart
    //ExFor:aw.Paragraph.frameFormat
    //ExFor:FrameFormat
    //ExFor:aw.FrameFormat.isFrame
    //ExFor:aw.FrameFormat.width
    //ExFor:aw.FrameFormat.height
    //ExFor:aw.FrameFormat.heightRule
    //ExFor:aw.FrameFormat.horizontalAlignment
    //ExFor:aw.FrameFormat.verticalAlignment
    //ExFor:aw.FrameFormat.horizontalPosition
    //ExFor:aw.FrameFormat.relativeHorizontalPosition
    //ExFor:aw.FrameFormat.horizontalDistanceFromText
    //ExFor:aw.FrameFormat.verticalPosition
    //ExFor:aw.FrameFormat.relativeVerticalPosition
    //ExFor:aw.FrameFormat.verticalDistanceFromText
    //ExSummary:Shows how to get information about formatting properties of paragraphs that are frames.
    let doc = new aw.Document(base.myDir + "Paragraph frame.docx");
    let paragraphFrame = doc.firstSection.body.paragraphs.toArray().filter(p => p.frameFormat.isFrame).at(0);

    expect(paragraphFrame.frameFormat.width).toEqual(233.3);
    expect(paragraphFrame.frameFormat.height).toEqual(138.8);
    expect(paragraphFrame.frameFormat.heightRule).toEqual(aw.HeightRule.AtLeast);
    expect(paragraphFrame.frameFormat.horizontalAlignment).toEqual(aw.Drawing.HorizontalAlignment.Default);
    expect(paragraphFrame.frameFormat.verticalAlignment).toEqual(aw.Drawing.VerticalAlignment.Default);
    expect(paragraphFrame.frameFormat.horizontalPosition).toEqual(34.05);
    expect(paragraphFrame.frameFormat.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Page);
    expect(paragraphFrame.frameFormat.horizontalDistanceFromText).toEqual(9.0);
    expect(paragraphFrame.frameFormat.verticalPosition).toEqual(20.5);
    expect(paragraphFrame.frameFormat.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);
    expect(paragraphFrame.frameFormat.verticalDistanceFromText).toEqual(0.0);
    //ExEnd
  });

  test('IsRevision', () => {
    //ExStart
    //ExFor:aw.Paragraph.isDeleteRevision
    //ExFor:aw.Paragraph.isInsertRevision
    //ExSummary:Shows how to work with revision paragraphs.
    let doc = new aw.Document();
    let body = doc.firstSection.body;
    let para = body.firstParagraph;

    para.appendChild(new aw.Run(doc, "Paragraph 1. "));
    body.appendParagraph("Paragraph 2. ");
    body.appendParagraph("Paragraph 3. ");

    // The above paragraphs are not revisions.
    // Paragraphs that we add after starting revision tracking will register as "Insert" revisions.
    doc.startTrackRevisions("John Doe", Date.now());

    para = body.appendParagraph("Paragraph 4. ");

    expect(para.isInsertRevision).toEqual(true);

    // Paragraphs that we remove after starting revision tracking will register as "Delete" revisions.
    let paragraphs = body.paragraphs;

    expect(paragraphs.count).toEqual(4);

    para = paragraphs.at(2);
    para.remove();

    // Such paragraphs will remain until we either accept or reject the delete revision.
    // Accepting the revision will remove the paragraph for good,
    // and rejecting the revision will leave it in the document as if we never deleted it.
    expect(paragraphs.count).toEqual(4);
    expect(para.isDeleteRevision).toEqual(true);

    // Accept the revision, and then verify that the paragraph is gone.
    doc.acceptAllRevisions();

    expect(paragraphs.count).toEqual(3);
    expect(para.count).toEqual(0);
    expect("Paragraph 1. \rParagraph 2. \rParagraph 4.").toEqual(doc.getText().trim());
    //ExEnd
  });


  test('BreakIsStyleSeparator', () => {
    //ExStart
    //ExFor:aw.Paragraph.breakIsStyleSeparator
    //ExSummary:Shows how to write text to the same line as a TOC heading and have it not show up in the TOC.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertTableOfContents("\\o \\h \\z \\u");
    builder.insertBreak(aw.BreakType.PageBreak);

    // Insert a paragraph with a style that the TOC will pick up as an entry.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;

    // Both these strings are in the same paragraph and will therefore show up on the same TOC entry.
    builder.write("Heading 1. ");
    builder.write("Will appear in the TOC. ");

    // If we insert a style separator, we can write more text in the same paragraph
    // and use a different style without showing up in the TOC.
    // If we use a heading type style after the separator, we can draw multiple TOC entries from one document text line.
    builder.insertStyleSeparator();
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Quote;
    builder.write("Won't appear in the TOC. ");

    expect(doc.firstSection.body.firstParagraph.breakIsStyleSeparator).toEqual(true);

    doc.updateFields();
    doc.save(base.artifactsDir + "Paragraph.breakIsStyleSeparator.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Paragraph.breakIsStyleSeparator.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOC, "TOC \\o \\h \\z \\u", 
      "\u0013 HYPERLINK \\l \"_Toc256000000\" \u0014Heading 1. Will appear in the TOC.\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\u0015\r", doc.range.fields.at(0));
    expect(doc.firstSection.body.firstParagraph.breakIsStyleSeparator).toEqual(false);
  });


  test('TabStops', () => {
    //ExStart
    //ExFor:aw.Paragraph.getEffectiveTabStops
    //ExSummary:Shows how to set custom tab stops for a paragraph.
    let doc = new aw.Document();
    let para = doc.firstSection.body.firstParagraph;

    // If we are in a paragraph with no tab stops in this collection,
    // the cursor will jump 36 points each time we press the Tab key in Microsoft Word.
    expect(doc.firstSection.body.firstParagraph.getEffectiveTabStops().length).toEqual(0);

    // We can add custom tab stops in Microsoft Word if we enable the ruler via the "View" tab.
    // Each unit on this ruler is two default tab stops, which is 72 points.
    // We can add custom tab stops programmatically like this.
    let tabStops = doc.firstSection.body.firstParagraph.paragraphFormat.tabStops;
    tabStops.add(72, aw.TabAlignment.Left, aw.TabLeader.Dots);
    tabStops.add(216, aw.TabAlignment.Center, aw.TabLeader.Dashes);
    tabStops.add(360, aw.TabAlignment.Right, aw.TabLeader.Line);

    // We can see these tab stops in Microsoft Word by enabling the ruler via "View" -> "Show" -> "Ruler".
    expect(para.getEffectiveTabStops().length).toEqual(3);

    // Any tab characters we add will make use of the tab stops on the ruler and may,
    // depending on the tab leader's value, leave a line between the tab departure and arrival destinations.
    para.appendChild(new aw.Run(doc, "\tTab 1\tTab 2\tTab 3"));

    doc.save(base.artifactsDir + "Paragraph.tabStops.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Paragraph.tabStops.docx");
    tabStops = doc.firstSection.body.firstParagraph.paragraphFormat.tabStops;

    TestUtil.verifyTabStop(72.0, aw.TabAlignment.Left, aw.TabLeader.Dots, false, tabStops.at(0));
    TestUtil.verifyTabStop(216.0, aw.TabAlignment.Center, aw.TabLeader.Dashes, false, tabStops.at(1));
    TestUtil.verifyTabStop(360.0, aw.TabAlignment.Right, aw.TabLeader.Line, false, tabStops.at(2));
  });


  test('JoinRuns', () => {
    //ExStart
    //ExFor:aw.Paragraph.joinRunsWithSameFormatting
    //ExSummary:Shows how to simplify paragraphs by merging superfluous runs.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert four runs of text into the paragraph.
    builder.write("Run 1. ");
    builder.write("Run 2. ");
    builder.write("Run 3. ");
    builder.write("Run 4. ");

    // If we open this document in Microsoft Word, the paragraph will look like one seamless text body.
    // However, it will consist of four separate runs with the same formatting. Fragmented paragraphs like this
    // may occur when we manually edit parts of one paragraph many times in Microsoft Word.
    let para = builder.currentParagraph;

    expect(para.runs.count).toEqual(4);

    // Change the style of the last run to set it apart from the first three.
    para.runs.at(3).font.styleIdentifier = aw.StyleIdentifier.Emphasis;

    // We can run the "JoinRunsWithSameFormatting" method to optimize the document's contents
    // by merging similar runs into one, reducing their overall count.
    // This method also returns the number of runs that this method merged.
    // These two merges occurred to combine Runs #1, #2, and #3,
    // while leaving out Run #4 because it has an incompatible style.
    expect(para.joinRunsWithSameFormatting()).toEqual(2);

    // The number of runs left will equal the original count
    // minus the number of run merges that the "JoinRunsWithSameFormatting" method carried out.
    expect(para.runs.count).toEqual(2);
    expect(para.runs.at(0).text).toEqual("Run 1. Run 2. Run 3. ");
    expect(para.runs.at(1).text).toEqual("Run 4. ");
    //ExEnd
  });


});
