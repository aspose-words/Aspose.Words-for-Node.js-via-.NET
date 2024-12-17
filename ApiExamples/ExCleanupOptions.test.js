// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExCleanupOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('RemoveUnusedResources', () => {
    //ExStart
    //ExFor:aw.Document.cleanup(CleanupOptions)
    //ExFor:CleanupOptions
    //ExFor:aw.CleanupOptions.unusedLists
    //ExFor:aw.CleanupOptions.unusedStyles
    //ExFor:aw.CleanupOptions.unusedBuiltinStyles
    //ExSummary:Shows how to remove all unused custom styles from a document. 
    let doc = new aw.Document();

    doc.styles.add(aw.StyleType.List, "MyListStyle1");
    doc.styles.add(aw.StyleType.List, "MyListStyle2");
    doc.styles.add(aw.StyleType.Character, "MyParagraphStyle1");
    doc.styles.add(aw.StyleType.Character, "MyParagraphStyle2");

    // Combined with the built-in styles, the document now has eight styles.
    // A custom style is marked as "used" while there is any text within the document
    // formatted in that style. This means that the 4 styles we added are currently unused.
    expect(doc.styles.count).toEqual(8);

    // Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
    let builder = new aw.DocumentBuilder(doc);
    builder.font.style = doc.styles.at("MyParagraphStyle1");
    builder.writeln("Hello world!");

    let list = doc.lists.add(doc.styles.at("MyListStyle1"));
    builder.listFormat.list = list;
    builder.writeln("Item 1");
    builder.writeln("Item 2");

    // Now, there is one unused character style and one unused list style.
    // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them.
    let cleanupOptions = new aw.CleanupOptions();
    cleanupOptions.unusedLists = true;
    cleanupOptions.unusedStyles = true;
    cleanupOptions.unusedBuiltinStyles = true;

    doc.cleanup(cleanupOptions);

    expect(doc.styles.count).toEqual(4);

    // Removing every node that a custom style is applied to marks it as "unused" again. 
    // Rerun the Cleanup method to remove them.
    doc.firstSection.body.removeAllChildren();
    doc.cleanup(cleanupOptions);

    expect(doc.styles.count).toEqual(2);
    //ExEnd
  });


  test('RemoveDuplicateStyles', () => {
    //ExStart
    //ExFor:aw.CleanupOptions.duplicateStyle
    //ExSummary:Shows how to remove duplicated styles from the document.
    let doc = new aw.Document();

    // Add two styles to the document with identical properties,
    // but different names. The second style is considered a duplicate of the first.
    let myStyle = doc.styles.add(aw.StyleType.Paragraph, "MyStyle1");
    myStyle.font.size = 14;
    myStyle.font.name = "Courier New";
    myStyle.font.color = "blue";

    let duplicateStyle = doc.styles.add(aw.StyleType.Paragraph, "MyStyle2");
    duplicateStyle.font.size = 14;
    duplicateStyle.font.name = "Courier New";
    duplicateStyle.font.color = "blue";

    expect(doc.styles.count).toEqual(6);

    // Apply both styles to different paragraphs within the document.
    let builder = new aw.DocumentBuilder(doc);
    builder.paragraphFormat.styleName = myStyle.name;
    builder.writeln("Hello world!");

    builder.paragraphFormat.styleName = duplicateStyle.name;
    builder.writeln("Hello again!");

    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.style).toEqual(myStyle);
    expect(paragraphs.at(1).paragraphFormat.style).toEqual(duplicateStyle);

    // Configure a CleanOptions object, then call the Cleanup method to substitute all duplicate styles
    // with the original and remove the duplicates from the document.
    let cleanupOptions = new aw.CleanupOptions();
    cleanupOptions.duplicateStyle = true;

    doc.cleanup(cleanupOptions);

    expect(doc.styles.count).toEqual(5);
    expect(paragraphs.at(0).paragraphFormat.style).toEqual(myStyle);
    expect(paragraphs.at(1).paragraphFormat.style).toEqual(myStyle);
    //ExEnd
  });


});
