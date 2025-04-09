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
const path = require('path');


describe("ExInlineStory", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test.each([aw.Notes.FootnotePosition.BeneathText,
    aw.Notes.FootnotePosition.BottomOfPage])('PositionFootnote', (footnotePosition) => {
    //ExStart
    //ExFor:Document.footnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.position
    //ExFor:FootnotePosition
    //ExSummary:Shows how to select a different place where the document collects and displays its footnotes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A footnote is a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow.  
    // Inserting a footnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote.
    // Each footnote also creates an entry at the bottom of the page, consisting of a symbol
    // that matches the reference symbol in the main body text.
    // The reference text that we pass to the document builder's "InsertFootnote" method.
    builder.write("Hello world!");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote contents.");

    // We can use the "Position" property to determine where the document will place all its footnotes.
    // If we set the value of the "Position" property to "FootnotePosition.BottomOfPage",
    // every footnote will show up at the bottom of the page that contains its reference mark. This is the default value.
    // If we set the value of the "Position" property to "FootnotePosition.BeneathText",
    // every footnote will show up at the end of the page's text that contains its reference mark.
    doc.footnoteOptions.position = footnotePosition;

    doc.save(base.artifactsDir + "InlineStory.PositionFootnote.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.PositionFootnote.docx");

    expect(doc.footnoteOptions.position).toEqual(footnotePosition);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote contents.", doc.getFootnote(0, true));
  });


  test.each([aw.Notes.EndnotePosition.EndOfDocument,
    aw.Notes.EndnotePosition.EndOfSection])('PositionEndnote', (endnotePosition) => {
    //ExStart
    //ExFor:Document.endnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.position
    //ExFor:EndnotePosition
    //ExSummary:Shows how to select a different place where the document collects and displays its endnotes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // An endnote is a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting an endnote adds a small superscript reference symbol
    // at the main body text where we insert the endnote.
    // Each endnote also creates an entry at the end of the document, consisting of a symbol
    // that matches the reference symbol in the main body text.
    // The reference text that we pass to the document builder's "InsertEndnote" method.
    builder.write("Hello world!");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote contents.");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("This is the second section.");

    // We can use the "Position" property to determine where the document will place all its endnotes.
    // If we set the value of the "Position" property to "EndnotePosition.EndOfDocument",
    // every footnote will show up in a collection at the end of the document. This is the default value.
    // If we set the value of the "Position" property to "EndnotePosition.EndOfSection",
    // every footnote will show up in a collection at the end of the section whose text contains the endnote's reference mark.
    doc.endnoteOptions.position = endnotePosition;

    doc.save(base.artifactsDir + "InlineStory.PositionEndnote.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.PositionEndnote.docx");

    expect(doc.endnoteOptions.position).toEqual(endnotePosition);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote contents.", doc.getFootnote(0, true));
  });


  test('RefMarkNumberStyle', () => {
    //ExStart
    //ExFor:Document.endnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.numberStyle
    //ExFor:Document.footnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.numberStyle
    //ExSummary:Shows how to change the number style of footnote/endnote reference marks.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Footnotes and endnotes are a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting a footnote/endnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote/endnote.
    // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
    // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
    // Footnote entries, by default, show up at the bottom of each page that contains
    // their reference symbols, and endnotes show up at the end of the document.
    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 2.");
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 3.", "Custom footnote reference mark");

    builder.insertParagraph();

    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 2.");
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 3.", "Custom endnote reference mark");

    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and for endnotes. By default, footnotes display their numbers using Arabic numerals,
    // and endnotes display their numbers in lowercase Roman numerals.
    expect(doc.footnoteOptions.numberStyle).toEqual(aw.NumberStyle.Arabic);
    expect(doc.endnoteOptions.numberStyle).toEqual(aw.NumberStyle.LowercaseRoman);

    // We can use the "NumberStyle" property to apply custom numbering styles to footnotes and endnotes.
    // This will not affect footnotes/endnotes with custom reference marks.
    doc.footnoteOptions.numberStyle = aw.NumberStyle.UppercaseRoman;
    doc.endnoteOptions.numberStyle = aw.NumberStyle.UppercaseLetter;

    doc.save(base.artifactsDir + "InlineStory.RefMarkNumberStyle.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.RefMarkNumberStyle.docx");

    expect(doc.footnoteOptions.numberStyle).toEqual(aw.NumberStyle.UppercaseRoman);
    expect(doc.endnoteOptions.numberStyle).toEqual(aw.NumberStyle.UppercaseLetter);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 1.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 2.", doc.getFootnote(1, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, false, "Custom footnote reference mark",
      "Custom footnote reference mark Footnote 3.", doc.getFootnote(2, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 1.", doc.getFootnote(3, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 2.", doc.getFootnote(4, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, false, "Custom endnote reference mark",
      "Custom endnote reference mark Endnote 3.", doc.getFootnote(5, true));
  });


  test('NumberingRule', () => {
    //ExStart
    //ExFor:Document.endnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.restartRule
    //ExFor:FootnoteNumberingRule
    //ExFor:Document.footnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.restartRule
    //ExSummary:Shows how to restart footnote/endnote numbering at certain places in the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Footnotes and endnotes are a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting a footnote/endnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote/endnote.
    // Each footnote/endnote also creates an entry, which consists of a symbol that matches the reference
    // symbol in the main body text. The reference text that we pass to the document builder's "InsertEndnote" method.
    // Footnote entries, by default, show up at the bottom of each page that contains
    // their reference symbols, and endnotes show up at the end of the document.
    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 3.");
    builder.write("Text 4. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 4.");

    builder.insertBreak(aw.BreakType.PageBreak);

    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 2.");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 3.");
    builder.write("Text 4. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 4.");

    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and endnotes and does not restart these counts at any point.
    expect(aw.Notes.FootnoteNumberingRule.Default).toEqual(doc.footnoteOptions.restartRule);
    expect(aw.Notes.FootnoteNumberingRule.Continuous).toEqual(aw.Notes.FootnoteNumberingRule.Default);

    // We can use the "RestartRule" property to get the document to restart
    // the footnote/endnote counts at a new page or section.
    doc.footnoteOptions.restartRule = aw.Notes.FootnoteNumberingRule.RestartPage;
    doc.endnoteOptions.restartRule = aw.Notes.FootnoteNumberingRule.RestartSection;

    doc.save(base.artifactsDir + "InlineStory.NumberingRule.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.NumberingRule.docx");

    expect(doc.footnoteOptions.restartRule).toEqual(aw.Notes.FootnoteNumberingRule.RestartPage);
    expect(doc.endnoteOptions.restartRule).toEqual(aw.Notes.FootnoteNumberingRule.RestartSection);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 1.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 2.", doc.getFootnote(1, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 3.", doc.getFootnote(2, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 4.", doc.getFootnote(3, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 1.", doc.getFootnote(4, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 2.", doc.getFootnote(5, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 3.", doc.getFootnote(6, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 4.", doc.getFootnote(7, true));
  });


  test('StartNumber', () => {
    //ExStart
    //ExFor:Document.endnoteOptions
    //ExFor:EndnoteOptions
    //ExFor:EndnoteOptions.startNumber
    //ExFor:Document.footnoteOptions
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.startNumber
    //ExSummary:Shows how to set a number at which the document begins the footnote/endnote count.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Footnotes and endnotes are a way to attach a reference or a side comment to text
    // that does not interfere with the main body text's flow. 
    // Inserting a footnote/endnote adds a small superscript reference symbol
    // at the main body text where we insert the footnote/endnote.
    // Each footnote/endnote also creates an entry, which consists of a symbol
    // that matches the reference symbol in the main body text.
    // The reference text that we pass to the document builder's "InsertEndnote" method.
    // Footnote entries, by default, show up at the bottom of each page that contains
    // their reference symbols, and endnotes show up at the end of the document.
    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 2.");
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote 3.");

    builder.insertParagraph();

    builder.write("Text 1. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 1.");
    builder.write("Text 2. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 2.");
    builder.write("Text 3. ");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Endnote 3.");

    // By default, the reference symbol for each footnote and endnote is its index
    // among all the document's footnotes/endnotes. Each document maintains separate counts
    // for footnotes and for endnotes, which both begin at 1.
    expect(doc.footnoteOptions.startNumber).toEqual(1);
    expect(doc.endnoteOptions.startNumber).toEqual(1);

    // We can use the "StartNumber" property to get the document to
    // begin a footnote or endnote count at a different number.
    doc.endnoteOptions.numberStyle = aw.NumberStyle.Arabic;
    doc.endnoteOptions.startNumber = 50;

    doc.save(base.artifactsDir + "InlineStory.startNumber.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.startNumber.docx");

    expect(doc.footnoteOptions.startNumber).toEqual(1);
    expect(doc.endnoteOptions.startNumber).toEqual(50);
    expect(doc.footnoteOptions.numberStyle).toEqual(aw.NumberStyle.Arabic);
    expect(doc.endnoteOptions.numberStyle).toEqual(aw.NumberStyle.Arabic);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 1.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 2.", doc.getFootnote(1, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote 3.", doc.getFootnote(2, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 1.", doc.getFootnote(3, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 2.", doc.getFootnote(4, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote 3.", doc.getFootnote(5, true));
  });


  test('AddFootnote', () => {
    //ExStart
    //ExFor:Footnote
    //ExFor:Footnote.isAuto
    //ExFor:Footnote.referenceMark
    //ExFor:InlineStory
    //ExFor:InlineStory.paragraphs
    //ExFor:InlineStory.firstParagraph
    //ExFor:FootnoteType
    //ExFor:Footnote.#ctor
    //ExSummary:Shows how to insert and customize footnotes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add text, and reference it with a footnote. This footnote will place a small superscript reference
    // mark after the text that it references and create an entry below the main body text at the bottom of the page.
    // This entry will contain the footnote's reference mark and the reference text,
    // which we will pass to the document builder's "InsertFootnote" method.
    builder.write("Main body text.");
    let footnote = builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote text.");

    // If this property is set to "true", then our footnote's reference mark
    // will be its index among all the section's footnotes.
    // This is the first footnote, so the reference mark will be "1".
    expect(footnote.isAuto).toEqual(true);

    // We can move the document builder inside the footnote to edit its reference text. 
    builder.moveTo(footnote.firstParagraph);
    builder.write(" More text added by a DocumentBuilder.");
    builder.moveToDocumentEnd();

    expect(footnote.getText().trim()).toEqual("\u0002 Footnote text. More text added by a DocumentBuilder.");

    builder.write(" More main body text.");
    footnote = builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote text.");

    // We can set a custom reference mark which the footnote will use instead of its index number.
    footnote.referenceMark = "RefMark";

    expect(footnote.isAuto).toEqual(false);

    // A bookmark with the "IsAuto" flag set to true will still show its real index
    // even if previous bookmarks display custom reference marks, so this bookmark's reference mark will be a "3".
    builder.write(" More main body text.");
    footnote = builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote text.");

    expect(footnote.isAuto).toEqual(true);

    doc.save(base.artifactsDir + "InlineStory.AddFootnote.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.AddFootnote.docx");

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '', 
      "Footnote text. More text added by a DocumentBuilder.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, false, "RefMark", 
      "Footnote text.", doc.getFootnote(1, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '', 
      "Footnote text.", doc.getFootnote(2, true));
  });


  test('FootnoteEndnote', () => {
    //ExStart
    //ExFor:Footnote.footnoteType
    //ExSummary:Shows the difference between footnotes and endnotes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways of attaching numbered references to the text. Both these references will add a
    // small superscript reference mark at the location that we insert them.
    // The reference mark, by default, is the index number of the reference among all the references in the document.
    // Each reference will also create an entry, which will have the same reference mark as in the body text
    // and reference text, which we will pass to the document builder's "InsertFootnote" method.
    // 1 -  A footnote, whose entry will appear on the same page as the text that it references:
    builder.write("Footnote referenced main body text.");
    let footnote = builder.insertFootnote(aw.Notes.FootnoteType.Footnote, 
      "Footnote text, will appear at the bottom of the page that contains the referenced text.");

    // 2 -  An endnote, whose entry will appear at the end of the document:
    builder.write("Endnote referenced main body text.");
    let endnote = builder.insertFootnote(aw.Notes.FootnoteType.Endnote, 
      "Endnote text, will appear at the very end of the document.");

    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    expect(footnote.footnoteType).toEqual(aw.Notes.FootnoteType.Footnote);
    expect(endnote.footnoteType).toEqual(aw.Notes.FootnoteType.Endnote);

    doc.save(base.artifactsDir + "InlineStory.FootnoteEndnote.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.FootnoteEndnote.docx");

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '',
      "Footnote text, will appear at the bottom of the page that contains the referenced text.", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Endnote, true, '',
      "Endnote text, will appear at the very end of the document.", doc.getFootnote(1, true));
  });


  test('AddComment', () => {
    //ExStart
    //ExFor:Comment
    //ExFor:InlineStory
    //ExFor:InlineStory.paragraphs
    //ExFor:InlineStory.firstParagraph
    //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
    //ExSummary:Shows how to add a comment to a paragraph.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Hello world!");

    var today = new Date(2024, 11, 26);
    var comment = new aw.Comment(doc, "John Doe", "JD", today);
    builder.currentParagraph.appendChild(comment);
    builder.moveTo(comment.appendChild(new aw.Paragraph(doc)));
    builder.write("Comment text.");

    expect(comment.dateTime).toEqual(today);

    // In Microsoft Word, we can right-click this comment in the document body to edit it, or reply to it. 
    doc.save(base.artifactsDir + "InlineStory.AddComment.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.AddComment.docx");
    comment = doc.getComment(0, true);
            
    expect(comment.getText()).toEqual("Comment text.\r");
    expect(comment.author).toEqual("John Doe");
    expect(comment.initial).toEqual("JD");
    expect(comment.dateTime).toEqual(today);
  });


  test('InlineStoryRevisions', () => {
    //ExStart
    //ExFor:InlineStory.isDeleteRevision
    //ExFor:InlineStory.isInsertRevision
    //ExFor:InlineStory.isMoveFromRevision
    //ExFor:InlineStory.isMoveToRevision
    //ExSummary:Shows how to view revision-related properties of InlineStory nodes.
    let doc = new aw.Document(base.myDir + "Revision footnotes.docx");

    // When we edit the document while the "Track Changes" option, found in via Review -> Tracking,
    // is turned on in Microsoft Word, the changes we apply count as revisions.
    // When editing a document using Aspose.words, we can begin tracking revisions by
    // invoking the document's "StartTrackRevisions" method and stop tracking by using the "StopTrackRevisions" method.
    // We can either accept revisions to assimilate them into the document
    // or reject them to undo and discard the proposed change.
    expect(doc.hasRevisions).toEqual(true);

    var footnotes = doc.getChildNodes(aw.NodeType.Footnote, true);

    expect(footnotes.count).toEqual(5);

    // Below are five types of revisions that can flag an InlineStory node.
    // 1 -  An "insert" revision:
    // This revision occurs when we insert text while tracking changes.
    expect(footnotes.at(2).asFootnote().isInsertRevision).toEqual(true);

    // 2 -  A "move from" revision:
    // When we highlight text in Microsoft Word, and then drag it to a different place in the document
    // while tracking changes, two revisions appear.
    // The "move from" revision is a copy of the text originally before we moved it.
    expect(footnotes.at(4).asFootnote().isMoveFromRevision).toEqual(true);

    // 3 -  A "move to" revision:
    // The "move to" revision is the text that we moved in its new position in the document.
    // "Move from" and "move to" revisions appear in pairs for every move revision we carry out.
    // Accepting a move revision deletes the "move from" revision and its text,
    // and keeps the text from the "move to" revision.
    // Rejecting a move revision conversely keeps the "move from" revision and deletes the "move to" revision.
    expect(footnotes.at(1).asFootnote().isMoveToRevision).toEqual(true);

    // 4 -  A "delete" revision:
    // This revision occurs when we delete text while tracking changes. When we delete text like this,
    // it will stay in the document as a revision until we either accept the revision,
    // which will delete the text for good, or reject the revision, which will keep the text we deleted where it was.
    expect(footnotes.at(3).asFootnote().isDeleteRevision).toEqual(true);
    //ExEnd
  });


  test('InsertInlineStoryNodes', () => {
    //ExStart
    //ExFor:Comment.storyType
    //ExFor:Footnote.storyType
    //ExFor:InlineStory.ensureMinimum
    //ExFor:InlineStory.font
    //ExFor:InlineStory.lastParagraph
    //ExFor:InlineStory.parentParagraph
    //ExFor:InlineStory.storyType
    //ExFor:InlineStory.tables
    //ExSummary:Shows how to insert InlineStory nodes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let footnote = builder.insertFootnote(aw.Notes.FootnoteType.Footnote, null);

    // Table nodes have an "EnsureMinimum()" method that makes sure the table has at least one cell.
    let table = new aw.Tables.Table(doc);
    table.ensureMinimum();

    // We can place a table inside a footnote, which will make it appear at the referencing page's footer.
    expect(footnote.tables.count).toEqual(0);
    footnote.appendChild(table);
    expect(footnote.tables.count).toEqual(1);
    expect(footnote.lastChild.nodeType).toEqual(aw.NodeType.Table);

    // An InlineStory has an "EnsureMinimum()" method as well, but in this case,
    // it makes sure the last child of the node is a paragraph,
    // for us to be able to click and write text easily in Microsoft Word.
    footnote.ensureMinimum();
    expect(footnote.lastChild.nodeType).toEqual(aw.NodeType.Paragraph);

    // Edit the appearance of the anchor, which is the small superscript number
    // in the main text that points to the footnote.
    footnote.font.name = "Arial";
    footnote.font.color = "#008000";

    // All inline story nodes have their respective story types.
    expect(footnote.storyType).toEqual(aw.StoryType.Footnotes);

    // A comment is another type of inline story.
    let comment = builder.currentParagraph.appendChild(new aw.Comment(doc, "John Doe", "J. D.", Date.now())).asComment();

    // The parent paragraph of an inline story node will be the one from the main document body.
    expect(comment.parentParagraph.referenceEquals(doc.firstSection.body.firstParagraph)).toEqual(true);

    // However, the last paragraph is the one from the comment text contents,
    // which will be outside the main document body in a speech bubble.
    // A comment will not have any child nodes by default,
    // so we can apply the EnsureMinimum() method to place a paragraph here as well.
    expect(comment.lastParagraph).toBe(null);
    comment.ensureMinimum();
    expect(comment.lastChild.nodeType).toEqual(aw.NodeType.Paragraph);

    // Once we have a paragraph, we can move the builder to do it and write our comment.
    builder.moveTo(comment.lastParagraph);
    builder.write("My comment.");

    expect(comment.storyType).toEqual(aw.StoryType.Comments);

    doc.save(base.artifactsDir + "InlineStory.InsertInlineStoryNodes.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "InlineStory.InsertInlineStoryNodes.docx");

    footnote = doc.getFootnote(0, true);

    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '', '', 
      doc.getFootnote(0, true));
    expect(footnote.font.name).toEqual("Arial");
    expect(footnote.font.color).toEqual("#008000");

    comment = doc.getComment(0, true);

    expect(comment.toString(aw.SaveFormat.Text).trim()).toEqual("My comment.");
  });


  test('DeleteShapes', () => {
    //ExStart
    //ExFor:Story
    //ExFor:Story.deleteShapes
    //ExFor:Story.storyType
    //ExFor:StoryType
    //ExSummary:Shows how to remove all shapes from a node.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a DocumentBuilder to insert a shape. This is an inline shape,
    // which has a parent Paragraph, which is a child node of the first section's Body.
    builder.insertShape(aw.Drawing.ShapeType.Cube, 100.0, 100.0);

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(1);

    // We can delete all shapes from the child paragraphs of this Body.
    expect(doc.firstSection.body.storyType).toEqual(aw.StoryType.MainText);
    doc.firstSection.body.deleteShapes();

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);
    //ExEnd
  });


  test('UpdateActualReferenceMarks', () => {
    //ExStart:UpdateActualReferenceMarks
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:Document.updateActualReferenceMarks
    //ExFor:Footnote.actualReferenceMark
    //ExSummary:Shows how to get actual footnote reference mark.
    let doc = new aw.Document(base.myDir + "Footnotes and endnotes.docx");

    let footnote = doc.getFootnote(1, true);
    doc.updateFields();
    doc.updateActualReferenceMarks();

    expect(footnote.actualReferenceMark).toEqual("1");
    //ExEnd:UpdateActualReferenceMarks
  });

  
  test('EndnoteSeparator', () => {
    //ExStart:EndnoteSeparator
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBase.footnoteSeparators
    //ExFor:FootnoteSeparatorType
    //ExSummary:Shows how to remove endnote separator.
    let doc = new aw.Document(base.myDir + "Footnotes and endnotes.docx");

    let endnoteSeparator = doc.footnoteSeparators.at(aw.Notes.FootnoteSeparatorType.EndnoteSeparator);
    // Remove endnote separator.
    endnoteSeparator.firstParagraph.firstChild.remove();
    //ExEnd:EndnoteSeparator

    doc.save(base.artifactsDir + "InlineStory.endnoteSeparator.docx");
  });


  test('FootnoteSeparator', () => {
    //ExStart:FootnoteSeparator
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBase.footnoteSeparators
    //ExFor:FootnoteSeparator
    //ExFor:FootnoteSeparatorType
    //ExFor:FootnoteSeparatorCollection
    //ExFor:FootnoteSeparatorCollection.item(FootnoteSeparatorType)
    //ExSummary:Shows how to manage footnote separator format.
    let doc = new aw.Document(base.myDir + "Footnotes and endnotes.docx");

    let footnoteSeparator = doc.footnoteSeparators.at(aw.Notes.FootnoteSeparatorType.FootnoteSeparator);
    // Align footnote separator.
    footnoteSeparator.firstParagraph.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    //ExEnd:FootnoteSeparator

    doc.save(base.artifactsDir + "InlineStory.footnoteSeparator.docx");
  });


});
