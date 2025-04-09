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


describe("ExTxtLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test.each([false,
    true])('DetectNumberingWithWhitespaces', (detectNumberingWithWhitespaces) => {
    //ExStart
    //ExFor:TxtLoadOptions.detectNumberingWithWhitespaces
    //ExSummary:Shows how to detect lists when loading plaintext documents.
    // Create a plaintext document in a string with four separate parts that we may interpret as lists,
    // with different delimiters. Upon loading the plaintext document into a "Document" object,
    // Aspose.words will always detect the first three lists and will add a "List" object
    // for each to the document's "Lists" property.
    const textDoc = "Full stop delimiters:\n" +
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

    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    let loadOptions = new aw.Loading.TxtLoadOptions();

    // Set the "DetectNumberingWithWhitespaces" property to "true" to detect numbered items
    // with whitespace delimiters, such as the fourth list in our document, as lists.
    // This may also falsely detect paragraphs that begin with numbers as lists.
    // Set the "DetectNumberingWithWhitespaces" property to "false"
    // to not create lists from numbered items with whitespace delimiters.
    loadOptions.detectNumberingWithWhitespaces = detectNumberingWithWhitespaces;

    let doc = new aw.Document(Buffer.from(textDoc), loadOptions);

    if (detectNumberingWithWhitespaces)
    {
      expect(doc.lists.count).toEqual(4);
      expect(doc.firstSection.body.paragraphs.toArray().some(p => p.getText().includes("Fourth list") && p.isListItem)).toEqual(true);
    }
    else
    {
      expect(doc.lists.count).toEqual(3);
      expect(doc.firstSection.body.paragraphs.toArray().some(p => p.getText().includes("Fourth list") && p.isListItem)).toEqual(false);
    }
    //ExEnd
  });


  test.each([[aw.Loading.TxtLeadingSpacesOptions.Preserve, aw.Loading.TxtTrailingSpacesOptions.Preserve],
    [aw.Loading.TxtLeadingSpacesOptions.ConvertToIndent, aw.Loading.TxtTrailingSpacesOptions.Preserve],
    [aw.Loading.TxtLeadingSpacesOptions.Trim, aw.Loading.TxtTrailingSpacesOptions.Trim]])
    ('TrailSpaces', (txtLeadingSpacesOptions, txtTrailingSpacesOptions) => {
    //ExStart
    //ExFor:TxtLoadOptions.trailingSpacesOptions
    //ExFor:TxtLoadOptions.leadingSpacesOptions
    //ExFor:TxtTrailingSpacesOptions
    //ExFor:TxtLeadingSpacesOptions
    //ExSummary:Shows how to trim whitespace when loading plaintext documents.
    const textDoc = "      Line 1 \n" +
            "    Line 2   \n" +
            " Line 3       ";

    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    let loadOptions = new aw.Loading.TxtLoadOptions();

    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Preserve"
    // to preserve all whitespace characters at the start of every line.
    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.ConvertToIndent"
    // to remove all whitespace characters from the start of every line,
    // and then apply a left first line indent to the paragraph to simulate the effect of the whitespaces.
    // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Trim"
    // to remove all whitespace characters from every line's start.
    loadOptions.leadingSpacesOptions = txtLeadingSpacesOptions;

    // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Preserve"
    // to preserve all whitespace characters at the end of every line. 
    // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Trim" to 
    // remove all whitespace characters from the end of every line.
    loadOptions.trailingSpacesOptions = txtTrailingSpacesOptions;

    let doc = new aw.Document(Buffer.from(textDoc), loadOptions);
    let paragraphs = doc.firstSection.body.paragraphs.toArray();

    switch (txtLeadingSpacesOptions)
    {
      case aw.Loading.TxtLeadingSpacesOptions.ConvertToIndent:
        expect(paragraphs[0].paragraphFormat.firstLineIndent).toEqual(37.8);
        expect(paragraphs[1].paragraphFormat.firstLineIndent).toEqual(25.2);
        expect(paragraphs[2].paragraphFormat.firstLineIndent).toEqual(6.3);

        expect(paragraphs[0].getText().startsWith("Line 1")).toEqual(true);
        expect(paragraphs[1].getText().startsWith("Line 2")).toEqual(true);
        expect(paragraphs[2].getText().startsWith("Line 3")).toEqual(true);
        break;
      case aw.Loading.TxtLeadingSpacesOptions.Preserve:
        expect(paragraphs.every(p => p.paragraphFormat.firstLineIndent == 0.0)).toEqual(true);

        expect(paragraphs[0].getText().startsWith("      Line 1")).toEqual(true);
        expect(paragraphs[1].getText().startsWith("    Line 2")).toEqual(true);
        expect(paragraphs[2].getText().startsWith(" Line 3")).toEqual(true);
        break;
      case aw.Loading.TxtLeadingSpacesOptions.Trim:
        expect(paragraphs.every(p => p.paragraphFormat.firstLineIndent == 0.0)).toEqual(true);

        expect(paragraphs[0].getText().startsWith("Line 1")).toEqual(true);
        expect(paragraphs[1].getText().startsWith("Line 2")).toEqual(true);
        expect(paragraphs[2].getText().startsWith("Line 3")).toEqual(true);
        break;
    }

    switch (txtTrailingSpacesOptions)
    {
      case aw.Loading.TxtTrailingSpacesOptions.Preserve:
        expect(paragraphs[0].getText().endsWith("Line 1 \r")).toEqual(true);
        expect(paragraphs[1].getText().endsWith("Line 2   \r")).toEqual(true);
        expect(paragraphs[2].getText().endsWith("Line 3       \f")).toEqual(true);
        break;
      case aw.Loading.TxtTrailingSpacesOptions.Trim:
        expect(paragraphs[0].getText().endsWith("Line 1\r")).toEqual(true);
        expect(paragraphs[1].getText().endsWith("Line 2\r")).toEqual(true);
        expect(paragraphs[2].getText().endsWith("Line 3\f")).toEqual(true);
        break;
    }
    //ExEnd
  });


  test('DetectDocumentDirection', () => {
    //ExStart
    //ExFor:TxtLoadOptions.documentDirection
    //ExFor:ParagraphFormat.bidi
    //ExSummary:Shows how to detect plaintext document text direction.
    // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
    // to modify how we load a plaintext document.
    let loadOptions = new aw.Loading.TxtLoadOptions();

    // Set the "DocumentDirection" property to "DocumentDirection.Auto" automatically detects
    // the direction of every paragraph of text that Aspose.words loads from plaintext.
    // Each paragraph's "Bidi" property will store its direction.
    loadOptions.documentDirection = aw.Loading.DocumentDirection.Auto;
 
    // Detect Hebrew text as right-to-left.
    let doc = new aw.Document(base.myDir + "Hebrew text.txt", loadOptions);

    expect(doc.firstSection.body.firstParagraph.paragraphFormat.bidi).toEqual(true);

    // Detect English text as right-to-left.
    doc = new aw.Document(base.myDir + "English text.txt", loadOptions);

    expect(doc.firstSection.body.firstParagraph.paragraphFormat.bidi).toEqual(false);
    //ExEnd
  });


  test('AutoNumberingDetection', () => {
    //ExStart
    //ExFor:TxtLoadOptions.autoNumberingDetection
    //ExSummary:Shows how to disable automatic numbering detection.
    let options = new aw.Loading.TxtLoadOptions();
    options.autoNumberingDetection = false;
    let doc = new aw.Document(base.myDir + "Number detection.txt", options);
    //ExEnd

    let listItemsCount = 0;
    for (let node of doc.getChildNodes(aw.NodeType.Paragraph, true))
    {
      if (node.asParagraph().isListItem)
        listItemsCount++;
    }

    expect(listItemsCount).toEqual(0);
  });


  test('DetectHyperlinks', () => {
    //ExStart:DetectHyperlinks
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:TxtLoadOptions
    //ExFor:TxtLoadOptions.#ctor
    //ExFor:TxtLoadOptions.detectHyperlinks
    //ExSummary:Shows how to read and display hyperlinks.
    const inputText = "Some links in TXT:\n" +
        "https://www.aspose.com/\n" +
        "https://docs.aspose.com/words/net/\n";

    // Load document with hyperlinks.
    let loadOptions = new aw.Loading.TxtLoadOptions();
    loadOptions.detectHyperlinks = true;
    let doc = new aw.Document(Buffer.from(inputText), loadOptions);

    // Print hyperlinks text.
    for (let field of doc.range.fields)
      console.log(field.result);

    expect(doc.range.fields.at(0).result.trim()).toEqual("https://www.aspose.com/");
    expect(doc.range.fields.at(1).result.trim()).toEqual("https://docs.aspose.com/words/net/");
    //ExEnd:DetectHyperlinks
  });
});
