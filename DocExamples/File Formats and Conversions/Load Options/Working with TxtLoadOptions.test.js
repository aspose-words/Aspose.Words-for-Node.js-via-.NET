// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithTxtLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('DetectNumberingWithWhitespaces', () => {
    //ExStart:DetectNumberingWithWhitespaces
    //GistId:ee038b97a80cf17ce52665651e81d832
    // Create a plaintext document in the form of a string with parts that may be interpreted as lists.
    // Upon loading, the first three lists will always be detected by Aspose.Words,
    // and List objects will be created for them after loading.
    let textDoc = "Full stop delimiters:\n" +
                  "1. First list item 1\n" +
                  "2. First list item 2\n" +
                  "3. First list item 3\n\n" +
                  "Right bracket delimiters:\n" +
                  "1) Second list item 1\n" +
                  "2) Second list item 2\n" +
                  "3) Second list item 3\n\n" +
                  "Bullet delimiters:\n" +
                  "• Third list item 1\n" +
                  "• Third list item 2\n" +
                  "• Third list item 3\n\n" +
                  "Whitespace delimiters:\n" +
                  "1 Fourth list item 1\n" +
                  "2 Fourth list item 2\n" +
                  "3 Fourth list item 3";

    // The fourth list, with whitespace inbetween the list number and list item contents,
    // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
    // to avoid paragraphs that start with numbers being mistakenly detected as lists.
    let loadOptions = new aw.Loading.TxtLoadOptions();
    loadOptions.detectNumberingWithWhitespaces = true;

    // Load the document while applying LoadOptions as a parameter and verify the result.
    let doc = new aw.Document(Buffer.from(textDoc, 'utf8'), loadOptions);

    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.DetectNumberingWithWhitespaces.docx");
    //ExEnd:DetectNumberingWithWhitespaces
  });

  test('HandleSpacesOptions', () => {
    //ExStart:HandleSpacesOptions
    //GistId:ee038b97a80cf17ce52665651e81d832
    let textDoc = "      Line 1 \n" +
                  "    Line 2   \n" +
                  " Line 3       ";

    let loadOptions = new aw.Loading.TxtLoadOptions();
    loadOptions.leadingSpacesOptions = aw.Loading.TxtLeadingSpacesOptions.Trim;
    loadOptions.trailingSpacesOptions = aw.Loading.TxtTrailingSpacesOptions.Trim;

    let doc = new aw.Document(Buffer.from(textDoc, 'utf8'), loadOptions);

    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.HandleSpacesOptions.docx");
    //ExEnd:HandleSpacesOptions
  });

  test('DocumentTextDirection', () => {
    //ExStart:DocumentTextDirection
    //GistId:ee038b97a80cf17ce52665651e81d832
    let loadOptions = new aw.Loading.TxtLoadOptions();
    loadOptions.documentDirection = aw.Loading.DocumentDirection.Auto;

    let doc = new aw.Document(base.myDir + "Hebrew text.txt", loadOptions);

    let paragraph = doc.firstSection.body.firstParagraph;
    console.log(paragraph.paragraphFormat.bidi);

    doc.save(base.artifactsDir + "WorkingWithTxtLoadOptions.DocumentTextDirection.docx");
    //ExEnd:DocumentTextDirection

  });

});