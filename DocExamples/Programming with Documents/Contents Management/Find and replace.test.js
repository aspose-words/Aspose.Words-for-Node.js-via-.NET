// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("FindAndReplace", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('SimpleFindReplace', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello _CustomerName_,");
    console.log("Original document text: " + doc.range.text);

    let options = new aw.Replacing.FindReplaceOptions(aw.Replacing.FindReplaceDirection.Forward);
    doc.range.replace("_CustomerName_", "James Bond", options);

    console.log("Document text after replace: " + doc.range.text);

    doc.save(base.artifactsDir + "FindAndReplace.SimpleFindReplace.docx");
  });

  test('FindAndHighlight', () => {
    //ExStart:FindAndHighlight
    let doc = new aw.Document(base.myDir + "Find and highlight.docx");

    let options = new aw.Replacing.FindReplaceOptions();
    options.direction = aw.Replacing.FindReplaceDirection.Backward;
    options.applyFont.highlightColor = "#FFFF00"; // Yellow.

    doc.range.replace("your document", "i", options);

    doc.save(base.artifactsDir + "FindAndReplace.FindAndHighlight.docx");
    //ExEnd:FindAndHighlight
  });

  test('MetaCharactersInSearchPattern', () => {
    /* meta-characters
    &p - paragraph break
    &b - section break
    &m - page break
    &l - manual line break
    */

    //ExStart:MetaCharactersInSearchPattern
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("This is Line 1");
    builder.writeln("This is Line 2");

    doc.range.replace("This is Line 1&pThis is Line 2", "This is replaced line");

    builder.moveToDocumentEnd();
    builder.write("This is Line 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("This is Line 2");

    doc.range.replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.");

    doc.save(base.artifactsDir + "FindAndReplace.MetaCharactersInSearchPattern.docx");
    //ExEnd:MetaCharactersInSearchPattern
  });

  test('ReplaceTextContainingMetaCharacters', () => {
    //ExStart:ReplaceTextContainingMetaCharacters
    //GistId:a652819331ab7eff5560cac42bb71ad2
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("First section");
    builder.writeln("  1st paragraph");
    builder.writeln("  2nd paragraph");
    builder.writeln("{insert-section}");
    builder.writeln("Second section");
    builder.writeln("  1st paragraph");

    const findReplaceOptions = new aw.Replacing.FindReplaceOptions();
    findReplaceOptions.applyParagraphFormat.alignment = aw.ParagraphAlignment.Center;

    // Double each paragraph break after word "section", add kind of underline and make it centered.
    let count = doc.range.replace("section&p", "section&p----------------------&p", findReplaceOptions);

    // Insert section break instead of custom text tag.
    count = doc.range.replace("{insert-section}", "&b", findReplaceOptions);

    doc.save(base.artifactsDir + "FindAndReplace.ReplaceTextContainingMetaCharacters.docx");
    //ExEnd:ReplaceTextContainingMetaCharacters
  });

  test('HighlightColor', () => {
    //ExStart:HighlightColor
    //GistId:a652819331ab7eff5560cac42bb71ad2
    let doc = new aw.Document(base.myDir + "Footer.docx");

    let options = new aw.Replacing.FindReplaceOptions();
    options.applyFont.highlightColor = "#FF8C00";

    doc.range.replace("header", "footer", options);
    //ExEnd:HighlightColor
  });

  test('IgnoreTextInsideFields', () => {
    //ExStart:IgnoreTextInsideFields
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert field with text inside.
    builder.insertField("INCLUDETEXT", "Text in field");

    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreFields = true;

    doc.range.replace("e", "*", options);

    console.log(doc.getText());

    options.ignoreFields = false;
    doc.range.replace("e", "*", options);

    console.log(doc.getText());
    //ExEnd:IgnoreTextInsideFields
  });

  test('IgnoreTextInsideDeleteRevisions', () => {
    //ExStart:IgnoreTextInsideDeleteRevisions
    //GistId:a652819331ab7eff5560cac42bb71ad2
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert non-revised text.
    builder.writeln("Deleted");
    builder.write("Text");

    // Remove first paragraph with tracking revisions.
    doc.startTrackRevisions("author", new Date());
    doc.firstSection.body.firstParagraph.remove();
    doc.stopTrackRevisions();

    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreDeleted = true;

    doc.range.replace("e", "*", options);

    console.log(doc.getText());

    options.ignoreDeleted = false;
    doc.range.replace("e", "*", options);

    console.log(doc.getText());
    //ExEnd:IgnoreTextInsideDeleteRevisions
  });

  test('IgnoreTextInsideInsertRevisions', () => {
    //ExStart:IgnoreTextInsideInsertRevisions
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert text with tracking revisions.
    doc.startTrackRevisions("author", new Date());
    builder.writeln("Inserted");
    doc.stopTrackRevisions();

    // Insert non-revised text.
    builder.write("Text");

    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreInserted = true;

    doc.range.replace("e", "*", options);

    console.log(doc.getText());

    options.ignoreInserted = false;
    doc.range.replace("e", "*", options);

    console.log(doc.getText());
    //ExEnd:IgnoreTextInsideInsertRevisions
  });

  test('ReplaceTextInFooter', () => {
    //ExStart:ReplaceTextInFooter
    //GistId:a652819331ab7eff5560cac42bb71ad2
    let doc = new aw.Document(base.myDir + "Footer.docx");

    let headersFooters = doc.firstSection.headersFooters;
    let footer = headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary);

    let options = new aw.Replacing.FindReplaceOptions();
    options.matchCase = false;
    options.findWholeWordsOnly = false;

    footer.range.replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options);

    doc.save(base.artifactsDir + "FindAndReplace.ReplaceTextInFooter.docx");
    //ExEnd:ReplaceTextInFooter
  });

  test('ReplaceWithString', () => {
    //ExStart:ReplaceWithString
    //GistId:a652819331ab7eff5560cac42bb71ad2
    const doc = new aw.Document();
    const builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello _CustomerName_,");

    const options = new aw.Replacing.FindReplaceOptions();
    options.direction = aw.Replacing.FindReplaceDirection.Forward;
    doc.range.replace("_CustomerName_", "James Bond", options);

    doc.save(base.artifactsDir + "FindAndReplace.ReplaceWithString.docx");
    //ExEnd:ReplaceWithString
  });

  test('UsingLegacyOrder', () => {
    //ExStart:UsingLegacyOrder
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("[tag 1]");
    let textBox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 50);
    builder.writeln("[tag 3]");

    builder.moveTo(textBox.firstParagraph);
    builder.write("[tag 2]");

    let options = new aw.Replacing.FindReplaceOptions();
    options.useLegacyOrder = true;

    doc.range.replace("[tag 1]", "", options);

    doc.save(base.artifactsDir + "FindAndReplace.UsingLegacyOrder.docx");
    //ExEnd:UsingLegacyOrder
  });

  test('ReplaceTextInTable', () => {
    //ExStart:ReplaceText
    //GistId:1693b4ac01f19ec81c9618649b62acb8
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

    let options = new aw.Replacing.FindReplaceOptions();
    options.direction = aw.Replacing.FindReplaceDirection.Forward;

    table.range.replace("Carrots", "Eggs", options);
    table.lastRow.lastCell.range.replace("50", "20", options);

    doc.save(base.artifactsDir + "FindAndReplace.ReplaceTextInTable.docx");
    //ExEnd:ReplaceText
  });
});