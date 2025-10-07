// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithFootnotes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('SetFootnoteColumns', () => {
    //ExStart:SetFootnoteColumns
    //GistId:7717b3a13cf86fabceb5860693073506
    let doc = new aw.Document(base.myDir + "Document.docx");
    // Specify the number of columns with which the footnotes area is formatted.
    doc.footnoteOptions.columns = 3;

    doc.save(base.artifactsDir + "WorkingWithFootnotes.SetFootnoteColumns.docx");
    //ExEnd:SetFootnoteColumns
  });

  test('SetFootnoteAndEndnotePosition', () => {
    //ExStart:SetFootnoteAndEndnotePosition
    //GistId:7717b3a13cf86fabceb5860693073506
    let doc = new aw.Document(base.myDir + "Document.docx");
    doc.footnoteOptions.position = aw.Notes.FootnotePosition.BeneathText;
    doc.endnoteOptions.position = aw.Notes.EndnotePosition.EndOfSection;

    doc.save(base.artifactsDir + "WorkingWithFootnotes.SetFootnoteAndEndnotePosition.docx");
    //ExEnd:SetFootnoteAndEndnotePosition
  });

  test('SetEndnoteOptions', () => {
    //ExStart:SetEndnoteOptions
    //GistId:7717b3a13cf86fabceb5860693073506
    let doc = new aw.Document(base.myDir + "Document.docx");

    let builder = new aw.DocumentBuilder(doc);
    builder.write("Some text");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Footnote text.");

    let option = doc.endnoteOptions;
    option.restartRule = aw.Notes.FootnoteNumberingRule.RestartPage;
    option.position = aw.Notes.EndnotePosition.EndOfSection;

    doc.save(base.artifactsDir + "WorkingWithFootnotes.SetEndnoteOptions.docx");
    //ExEnd:SetEndnoteOptions
  });

});