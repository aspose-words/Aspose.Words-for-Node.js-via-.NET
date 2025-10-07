// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("CompareDocuments", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CompareForEqual', () => {
    //ExStart:CompareForEqual
    //GistId:57808d29628dd1680d4c229e84c5456c
    let docA = new aw.Document(base.myDir + "Document.docx");
    let docB = docA.clone();

    // DocA now contains changes as revisions.
    docA.compare(docB, "user", new Date());
    console.log(docA.revisions.count == 0 ? "Documents are equal" : "Documents are not equal");
    //ExEnd:CompareForEqual
  });


  test('CompareOptions', () => {
    //ExStart:CompareOptions
    //GistId:57808d29628dd1680d4c229e84c5456c
    let docA = new aw.Document(base.myDir + "Document.docx");
    let docB = docA.clone();
    let options = new aw.Comparing.CompareOptions();
    options.ignoreFormatting = true;
    options.ignoreHeadersAndFooters = true;
    options.ignoreCaseChanges = true;
    options.ignoreTables = true;
    options.ignoreFields = true;
    options.ignoreComments = true;
    options.ignoreTextboxes = true;
    options.ignoreFootnotes = true;

    docA.compare(docB, "user", new Date(), options);
    console.log(docA.revisions.count == 0 ? "Documents are equal" : "Documents are not equal");
    //ExEnd:CompareOptions
  });

  test('ComparisonTarget', () => {
    //ExStart:ComparisonTarget
    let docA = new aw.Document(base.myDir + "Document.docx");
    let docB = docA.clone();
    // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
    let options = new aw.Comparing.CompareOptions();
    options.ignoreFormatting = true;
    options.target = aw.Comparing.ComparisonTargetType.New;

    docA.compare(docB, "user", new Date(), options);
    //ExEnd:ComparisonTarget
  });

  test('ComparisonGranularity', () => {
    //ExStart:ComparisonGranularity
    let builderA = new aw.DocumentBuilder(new aw.Document());
    let builderB = new aw.DocumentBuilder(new aw.Document());
    builderA.writeln("This is A simple word");
    builderB.writeln("This is B simple words");
    let compareOptions = new aw.Comparing.CompareOptions();
    compareOptions.granularity = aw.Comparing.Granularity.CharLevel;

    builderA.document.compare(builderB.document, "author", new Date(), compareOptions);
    //ExEnd:ComparisonGranularity
  });
});
