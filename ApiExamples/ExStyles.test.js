// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');
const fs = require('fs');
const DocumentHelper = require('./DocumentHelper');


describe("ExStyles", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('Styles', () => {
    //ExStart
    //ExFor:DocumentBase.styles
    //ExFor:Style.document
    //ExFor:Style.name
    //ExFor:Style.isHeading
    //ExFor:Style.isQuickStyle
    //ExFor:Style.nextParagraphStyleName
    //ExFor:Style.styles
    //ExFor:Style.type
    //ExFor:StyleCollection.document
    //ExFor:StyleCollection.getEnumerator
    //ExSummary:Shows how to access a document's style collection.
    let doc = new aw.Document();

    expect(doc.styles.count).toEqual(4);

    // Enumerate and list all the styles that a document created using Aspose.words contains by default.
    for (var style of doc.styles)
    {
      console.log(`Style name:\t\"${style.name}\", of type \"${style.type}\"`);
      console.log(`\tSubsequent style:\t${style.nextParagraphStyleName}`);
      console.log(`\tIs heading:\t\t\t${style.isHeading}`);
      console.log(`\tIs QuickStyle:\t\t${style.isQuickStyle}`);

      expect(style.document.referenceEquals(doc)).toBe(true);
    }
    //ExEnd
  });


  test('CreateStyle', () => {
    //ExStart
    //ExFor:Style.font
    //ExFor:Style
    //ExFor:Style.remove
    //ExFor:Style.automaticallyUpdate
    //ExSummary:Shows how to create and apply a custom style.
    let doc = new aw.Document();

    let style = doc.styles.add(aw.StyleType.Paragraph, "MyStyle");
    style.font.name = "Times New Roman";
    style.font.size = 16;
    style.font.color = "#000080";
    // Automatically redefine style.
    style.automaticallyUpdate = true;

    let builder = new aw.DocumentBuilder(doc);

    // Apply one of the styles from the document to the paragraph that the document builder is creating.
    builder.paragraphFormat.style = doc.styles.at("MyStyle");
    builder.writeln("Hello world!");

    let firstParagraphStyle = doc.firstSection.body.firstParagraph.paragraphFormat.style;

    expect(firstParagraphStyle).toEqual(style);

    // Remove our custom style from the document's styles collection.
    doc.styles.at("MyStyle").remove();

    firstParagraphStyle = doc.firstSection.body.firstParagraph.paragraphFormat.style;

    // Any text that used a removed style reverts to the default formatting.
    expect([...doc.styles].every(s => s.name == "MyStyle")).toEqual(false);
    expect(firstParagraphStyle.font.name).toEqual("Times New Roman");
    expect(firstParagraphStyle.font.size).toEqual(12.0);
    expect(firstParagraphStyle.font.color).toEqual(base.emptyColor);
    //ExEnd
  });


  test('StyleCollection', () => {
    //ExStart
    //ExFor:StyleCollection.add(StyleType,String)
    //ExFor:StyleCollection.count
    //ExFor:StyleCollection.defaultFont
    //ExFor:StyleCollection.defaultParagraphFormat
    //ExFor:StyleCollection.item(StyleIdentifier)
    //ExFor:StyleCollection.item(Int32)
    //ExSummary:Shows how to add a Style to a document's styles collection.
    let doc = new aw.Document();

    let styles = doc.styles;
    // Set default parameters for new styles that we may later add to this collection.
    styles.defaultFont.name = "Courier New";
    // If we add a style of the "StyleType.Paragraph", the collection will apply the values of
    // its "DefaultParagraphFormat" property to the style's "ParagraphFormat" property.
    styles.defaultParagraphFormat.firstLineIndent = 15.0;
    // Add a style, and then verify that it has the default settings.
    styles.add(aw.StyleType.Paragraph, "MyStyle");

    expect(styles.at(4).font.name).toEqual("Courier New");
    expect(styles.at("MyStyle").paragraphFormat.firstLineIndent).toEqual(15.0);
    //ExEnd
  });


  test('RemoveStylesFromStyleGallery', () => {
    //ExStart
    //ExFor:StyleCollection.clearQuickStyleGallery
    //ExSummary:Shows how to remove styles from Style Gallery panel.
    let doc = new aw.Document();
    // Note that remove styles work only with DOCX format for now.
    doc.styles.clearQuickStyleGallery();

    doc.save(base.artifactsDir + "Styles.RemoveStylesFromStyleGallery.docx");
    //ExEnd
  });


  test('ChangeTocsTabStops', () => {
    //ExStart
    //ExFor:TabStop
    //ExFor:ParagraphFormat.tabStops
    //ExFor:Style.styleIdentifier
    //ExFor:TabStopCollection.removeByPosition
    //ExFor:TabStop.alignment
    //ExFor:TabStop.position
    //ExFor:TabStop.leader
    //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
    let doc = new aw.Document(base.myDir + "Table of contents.docx");

    // Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
    for (var paraNode of doc.getChildNodes(aw.NodeType.Paragraph, true)) {
      var para = paraNode.asParagraph();
      if (para.paragraphFormat.style.styleIdentifier >= aw.StyleIdentifier.Toc1 &&
        para.paragraphFormat.style.styleIdentifier <= aw.StyleIdentifier.Toc9)
      {
        // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
        let tab = para.paragraphFormat.tabStops.at(0);

        // Replace the first default tab, stop with a custom tab stop.
        para.paragraphFormat.tabStops.removeByPosition(tab.position);
        para.paragraphFormat.tabStops.add(tab.position - 50, tab.alignment, tab.leader);
      }
    }

    doc.save(base.artifactsDir + "Styles.ChangeTocsTabStops.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Styles.ChangeTocsTabStops.docx");

    for (var paraNode of doc.getChildNodes(aw.NodeType.Paragraph, true)) {
      var para = paraNode.asParagraph();
      if (para.paragraphFormat.style.styleIdentifier >= aw.StyleIdentifier.Toc1 &&
        para.paragraphFormat.style.styleIdentifier <= aw.StyleIdentifier.Toc9)
      {
        let tabStop = para.getEffectiveTabStops()[0];
        expect(tabStop.position).toEqual(400.8);
        expect(tabStop.alignment).toEqual(aw.TabAlignment.Right);
        expect(tabStop.leader).toEqual(aw.TabLeader.Dots);
      }
    }
  });


  test('CopyStyleSameDocument', () => {
    //ExStart
    //ExFor:StyleCollection.addCopy(Style)
    //ExFor:Style.name
    //ExSummary:Shows how to clone a document's style.
    let doc = new aw.Document();

    // The AddCopy method creates a copy of the specified style and
    // automatically generates a new name for the style, such as "Heading 1_0".
    let newStyle = doc.styles.addCopy(doc.styles.at("Heading 1"));

    // Use the style's "Name" property to change the style's identifying name.
    newStyle.name = "My Heading 1";

    // Our document now has two identical looking styles with different names.
    // Changing settings of one of the styles do not affect the other.
    newStyle.font.color = "#FF0000";

    expect(newStyle.name).toEqual("My Heading 1");
    expect(doc.styles.at("Heading 1").name).toEqual("Heading 1");

    expect(newStyle.type).toEqual(doc.styles.at("Heading 1").type);
    expect(newStyle.font.name).toEqual(doc.styles.at("Heading 1").font.name);
    expect(newStyle.font.size).toEqual(doc.styles.at("Heading 1").font.size);
    expect(newStyle.font.color).not.toEqual(doc.styles.at("Heading 1").font.color);
    //ExEnd
  });


  test('CopyStyleDifferentDocument', () => {
    //ExStart
    //ExFor:StyleCollection.addCopy(Style)
    //ExSummary:Shows how to import a style from one document into a different document.
    let srcDoc = new aw.Document();

    // Create a custom style for the source document.
    let srcStyle = srcDoc.styles.add(aw.StyleType.Paragraph, "MyStyle");
    srcStyle.font.color = "#FF0000";

    // Import the source document's custom style into the destination document.
    let dstDoc = new aw.Document();
    let newStyle = dstDoc.styles.addCopy(srcStyle);

    // The imported style has an appearance identical to its source style.
    expect(newStyle.name).toEqual("MyStyle");
    expect(newStyle.font.color).toEqual("#FF0000");
    //ExEnd
  });


  test('DefaultStyles', () => {
    let doc = new aw.Document();

    doc.styles.defaultFont.name = "PMingLiU";
    doc.styles.defaultFont.bold = true;

    doc.styles.defaultParagraphFormat.spaceAfter = 20;
    doc.styles.defaultParagraphFormat.alignment = aw.ParagraphAlignment.Right;

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.styles.defaultFont.bold).toEqual(true);
    expect(doc.styles.defaultFont.name).toEqual("PMingLiU");
    expect(doc.styles.defaultParagraphFormat.spaceAfter).toEqual(20);
    expect(doc.styles.defaultParagraphFormat.alignment).toEqual(aw.ParagraphAlignment.Right);
  });


  test('ParagraphStyleBulletedList', () => {
    //ExStart
    //ExFor:StyleCollection
    //ExFor:DocumentBase.styles
    //ExFor:Style
    //ExFor:Font
    //ExFor:Style.font
    //ExFor:Style.paragraphFormat
    //ExFor:Style.listFormat
    //ExFor:ParagraphFormat.style
    //ExSummary:Shows how to create and use a paragraph style with list formatting.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a custom paragraph style.
    let style = doc.styles.add(aw.StyleType.Paragraph, "MyStyle1");
    style.font.size = 24;
    style.font.name = "Verdana";
    style.paragraphFormat.spaceAfter = 12;

    // Create a list and make sure the paragraphs that use this style will use this list.
    style.listFormat.list = doc.lists.add(aw.Lists.ListTemplate.BulletDefault);
    style.listFormat.listLevelNumber = 0;

    // Apply the paragraph style to the document builder's current paragraph, and then add some text.
    builder.paragraphFormat.style = style;
    builder.writeln("Hello World: MyStyle1, bulleted list.");

    // Change the document builder's style to one that has no list formatting and write another paragraph.
    builder.paragraphFormat.style = doc.styles.at("Normal");
    builder.writeln("Hello World: Normal.");

    builder.document.save(base.artifactsDir + "Styles.ParagraphStyleBulletedList.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Styles.ParagraphStyleBulletedList.docx");

    style = doc.styles.at("MyStyle1");

    expect(style.name).toEqual("MyStyle1");
    expect(style.font.size).toEqual(24);
    expect(style.font.name).toEqual("Verdana");
    expect(style.paragraphFormat.spaceAfter).toEqual(12.0);
  });


  test('StyleAliases', () => {
    //ExStart
    //ExFor:Style.aliases
    //ExFor:Style.baseStyleName
    //ExFor:Style.equals(Style)
    //ExFor:Style.linkedStyleName
    //ExSummary:Shows how to use style aliases.
    let doc = new aw.Document(base.myDir + "Style with alias.docx");

    // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
    // If a style's name has multiple values separated by commas, each clause is a separate alias.
    let style = doc.styles.at("MyStyle");
    expect(style.aliases).toEqual(["MyStyle Alias 1", "MyStyle Alias 2"]);
    expect(style.baseStyleName).toEqual("Title");
    expect(style.linkedStyleName).toEqual("MyStyle Char");

    // We can reference a style using its alias, as well as its name.
    expect(doc.styles.at("MyStyle Alias 2")).toEqual(doc.styles.at("MyStyle Alias 1"));

    let builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.paragraphFormat.style = doc.styles.at("MyStyle Alias 1");
    builder.writeln("Hello world!");
    builder.paragraphFormat.style = doc.styles.at("MyStyle Alias 2");
    builder.write("Hello again!");

    expect(doc.firstSection.body.paragraphs.at(1).paragraphFormat.style).toEqual(doc.firstSection.body.paragraphs.at(0).paragraphFormat.style);
    //ExEnd
  });


  test('LatentStyles', () => {
    // This test is to check that after re-saving a document it doesn't lose LatentStyle information
    // for 4 styles from documents created in Microsoft Word.
    let doc = new aw.Document(base.myDir + "Blank.docx");

    doc.save(base.artifactsDir + "Styles.latentStyles.docx");

    TestUtil.docPackageFileContainsString(
      '<w:lsdException w:name="Mention" w:semiHidden="1" w:unhideWhenUsed="1" />',
      base.artifactsDir + "Styles.latentStyles.docx", "styles.xml");
    TestUtil.docPackageFileContainsString(
      '<w:lsdException w:name="Smart Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1" />',
      base.artifactsDir + "Styles.latentStyles.docx", "styles.xml");
    TestUtil.docPackageFileContainsString(
      '<w:lsdException w:name="Hashtag" w:semiHidden="1" w:unhideWhenUsed="1" />',
      base.artifactsDir + "Styles.latentStyles.docx", "styles.xml");
    TestUtil.docPackageFileContainsString(
      '<w:lsdException w:name="Unresolved Mention" w:semiHidden="1" w:unhideWhenUsed="1" />',
      base.artifactsDir + "Styles.latentStyles.docx", "styles.xml");
  });


  test('LockStyle', () => {
    //ExStart:LockStyle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Style.locked
    //ExSummary:Shows how to lock style.
    let doc = new aw.Document();

    let styleHeading1 = doc.styles.at(aw.StyleIdentifier.Heading1);
    if (!styleHeading1.locked)            
      styleHeading1.locked = true;

    doc.save(base.artifactsDir + "Styles.LockStyle.docx");
    //ExEnd:LockStyle

    doc = new aw.Document(base.artifactsDir + "Styles.LockStyle.docx");
    expect(doc.styles.at(aw.StyleIdentifier.Heading1).locked).toEqual(true);
  });


  test('StylePriority', () => {
    //ExStart:StylePriority
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:Style.priority
    //ExFor:Style.unhideWhenUsed
    //ExFor:Style.semiHidden
    //ExSummary:Shows how to prioritize and hide a style.
    let doc = new aw.Document();
    let styleTitle = doc.styles.at(aw.StyleIdentifier.Subtitle);
            
    if (styleTitle.priority == 9)
      styleTitle.priority = 10;

    if (!styleTitle.unhideWhenUsed)
      styleTitle.unhideWhenUsed = true;

    if (styleTitle.semiHidden)
      styleTitle.semiHidden = true;

    doc.save(base.artifactsDir + "Styles.StylePriority.docx");
    //ExEnd:StylePriority
  });


  test('LinkedStyleName', () => {
    //ExStart:LinkedStyleName
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:Style.linkedStyleName
    //ExSummary:Shows how to link styles among themselves.
    let doc = new aw.Document();

    let styleHeading1 = doc.styles.at(aw.StyleIdentifier.Heading1);

    let styleHeading1Char = doc.styles.add(aw.StyleType.Character, "Heading 1 Char");
    styleHeading1Char.font.name = "Verdana";
    styleHeading1Char.font.bold = true;
    styleHeading1Char.font.border.lineStyle = aw.LineStyle.Dot;
    styleHeading1Char.font.border.lineWidth = 15;

    styleHeading1.linkedStyleName = "Heading 1 Char";

    expect(styleHeading1.linkedStyleName).toEqual("Heading 1 Char");
    expect(styleHeading1Char.linkedStyleName).toEqual("Heading 1");
    //ExEnd:LinkedStyleName
  });


});
