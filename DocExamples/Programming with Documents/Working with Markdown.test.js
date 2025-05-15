// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');
const MemoryStream = require('memorystream');


describe("WorkingWithMarkdown", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('BoldText', () => {
    //ExStart:BoldText
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Make the text Bold.
    builder.font.bold = true;
    builder.writeln("This text will be Bold");
    //ExEnd:BoldText
  });


  test('ItalicText', () => {
    //ExStart:ItalicText
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Make the text Italic.
    builder.font.italic = true;
    builder.writeln("This text will be Italic");
    //ExEnd:ItalicText
  });


  test('Strikethrough', () => {
    //ExStart:Strikethrough
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Make the text Strikethrough.
    builder.font.strikeThrough = true;
    builder.writeln("This text will be StrikeThrough");
    //ExEnd:Strikethrough
  });


  test('InlineCode', () => {
    //ExStart:InlineCode
    //GistId:a2fee7fa3d8e5704ce24f041be9a4821
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Number of backticks is missed, one backtick will be used by default.
    let inlineCode1BackTicks = builder.document.styles.add(aw.StyleType.Character, "InlineCode");
    builder.font.style = inlineCode1BackTicks;
    builder.writeln("Text with InlineCode style with 1 backtick");

    // There will be 3 backticks.
    let inlineCode3BackTicks = builder.document.styles.add(aw.StyleType.Character, "InlineCode.3");
    builder.font.style = inlineCode3BackTicks;
    builder.writeln("Text with InlineCode style with 3 backtick");
    //ExEnd:InlineCode
  });


  test('Autolink', () => {
    //ExStart:Autolink
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Insert hyperlink.
    builder.insertHyperlink("https://www.aspose.com", "https://www.aspose.com", false);
    builder.insertHyperlink("email@aspose.com", "mailto:email@aspose.com", false);
    //ExEnd:Autolink
  });


  test('Link', () => {
    //ExStart:Link
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Insert hyperlink.
    builder.insertHyperlink("Aspose", "https://www.aspose.com", false);
    //ExEnd:Link
  });


  test('Image', () => {
    //ExStart:Image
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Insert image.
    let shape = builder.insertImage(base.imagesDir + "Logo.jpg");
    shape.imageData.title = "title";
    //ExEnd:Image
  });


  test('HorizontalRule', () => {
    //ExStart:HorizontalRule
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Insert horizontal rule.
    builder.insertHorizontalRule();
    //ExEnd:HorizontalRule
  });


  test('Heading', () => {
    //ExStart:Heading
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // By default Heading styles in Word may have Bold and Italic formatting.
    //If we do not want to be emphasized, set these properties explicitly to false.
    builder.font.bold = false;
    builder.font.italic = false;

    builder.writeln("The following produces headings:");
    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Heading1");
    builder.paragraphFormat.style = doc.styles.at("Heading 2");
    builder.writeln("Heading2");
    builder.paragraphFormat.style = doc.styles.at("Heading 3");
    builder.writeln("Heading3");
    builder.paragraphFormat.style = doc.styles.at("Heading 4");
    builder.writeln("Heading4");
    builder.paragraphFormat.style = doc.styles.at("Heading 5");
    builder.writeln("Heading5");
    builder.paragraphFormat.style = doc.styles.at("Heading 6");
    builder.writeln("Heading6");

    // Note, emphases are also allowed inside Headings:
    builder.font.bold = true;
    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Bold Heading1");

    doc.save(base.artifactsDir + "WorkingWithMarkdown.heading.md");
    //ExEnd:Heading
  });


  test('SetextHeading', () => {
    //ExStart:SetextHeading
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    builder.paragraphFormat.styleName = "Heading 1";
    builder.writeln("This is an H1 tag");

    // Reset styles from the previous paragraph to not combine styles between paragraphs.
    builder.font.bold = false;
    builder.font.italic = false;

    let setexHeading1 = builder.document.styles.add(aw.StyleType.Paragraph, "SetextHeading1");
    builder.paragraphFormat.style = setexHeading1;
    builder.document.styles.at("SetextHeading1").baseStyleName = "Heading 1";
    builder.writeln("Setext Heading level 1");

    builder.paragraphFormat.style = builder.document.styles.at("Heading 3");
    builder.writeln("This is an H3 tag");

    // Reset styles from the previous paragraph to not combine styles between paragraphs.
    builder.font.bold = false;
    builder.font.italic = false;

    let setexHeading2 = builder.document.styles.add(aw.StyleType.Paragraph, "SetextHeading2");
    builder.paragraphFormat.style = setexHeading2;
    builder.document.styles.at("SetextHeading2").baseStyleName = "Heading 3";

    // Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
    builder.writeln("Setext Heading level 2");
    //ExEnd:SetextHeading

    builder.document.save(base.artifactsDir + "WorkingWithMarkdown.setextHeading.md");
  });


  test('IndentedCode', () => {
    //ExStart:IndentedCode
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    let indentedCode = builder.document.styles.add(aw.StyleType.Paragraph, "IndentedCode");
    builder.paragraphFormat.style = indentedCode;
    builder.writeln("This is an indented code");
    //ExEnd:IndentedCode
  });


  test('FencedCode', () => {
    //ExStart:FencedCode
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    let fencedCode = builder.document.styles.add(aw.StyleType.Paragraph, "FencedCode");
    builder.paragraphFormat.style = fencedCode;
    builder.writeln("This is an fenced code");

    let fencedCodeWithInfo = builder.document.styles.add(aw.StyleType.Paragraph, "FencedCode.C#");
    builder.paragraphFormat.style = fencedCodeWithInfo;
    builder.writeln("This is a fenced code with info string");
    //ExEnd:FencedCode
  });


  test('Quote', () => {
    //ExStart:Quote
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // By default a document stores blockquote style for the first level.
    builder.paragraphFormat.styleName = "Quote";
    builder.writeln("Blockquote");

    // Create styles for nested levels through style inheritance.
    let quoteLevel2 = builder.document.styles.add(aw.StyleType.Paragraph, "Quote1");
    builder.paragraphFormat.style = quoteLevel2;
    builder.document.styles.at("Quote1").baseStyleName = "Quote";
    builder.writeln("1. Nested blockquote");

    doc.save(base.artifactsDir + "WorkingWithMarkdown.quote.md");
    //ExEnd:Quote
  });


  test('BulletedList', () => {
    //ExStart:BulletedList
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    builder.listFormat.applyBulletDefault();
    builder.listFormat.list.listLevels.at(0).numberFormat = "-";

    builder.writeln("Item 1");
    builder.writeln("Item 2");

    builder.listFormat.listIndent();

    builder.writeln("Item 2a");
    builder.writeln("Item 2b");
    //ExEnd:BulletedList
  });


  test('OrderedList', () => {
    //ExStart:OrderedList
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.applyNumberDefault();

    builder.writeln("Item 1");
    builder.writeln("Item 2");

    builder.listFormat.listIndent();

    builder.writeln("Item 2a");
    builder.writeln("Item 2b");
    //ExEnd:OrderedList
  });


  test('Table', () => {
    //ExStart:Table
    //GistId:6558fa20d4ebd9a86b255fe67ca67eb1
    // Use a document builder to add content to the document.
    let builder = new aw.DocumentBuilder();

    // Add the first row.
    builder.insertCell();
    builder.writeln("a");
    builder.insertCell();
    builder.writeln("b");

    builder.endRow();

    // Add the second row.
    builder.insertCell();
    builder.writeln("c");
    builder.insertCell();
    builder.writeln("d");
    //ExEnd:Table
  });


  test('ReadMarkdownDocument', () => {
    //ExStart:ReadMarkdownDocument
    //GistId:757cf7d3534a39730cf3290d418681ab
    let doc = new aw.Document(base.myDir + "Quotes.md");

    // Let's remove Heading formatting from a Quote in the very last paragraph.
    let paragraph = doc.firstSection.body.lastParagraph;
    paragraph.paragraphFormat.style = doc.styles.at("Quote");

    doc.save(base.artifactsDir + "WorkingWithMarkdown.ReadMarkdownDocument.md");
    //ExEnd:ReadMarkdownDocument
  });


  test('Emphases', () => {
    //ExStart:Emphases
    //GistId:757cf7d3534a39730cf3290d418681ab
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.");
    builder.write("You can write ");

    builder.font.bold = true;
    builder.write("bold");

    builder.font.bold = false;
    builder.write(" or ");

    builder.font.italic = true;
    builder.write("italic");

    builder.font.italic = false;
    builder.writeln(" text. ");

    builder.write("You can also write ");
    builder.font.bold = true;

    builder.font.italic = true;
    builder.write("BoldItalic");

    builder.font.bold = false;
    builder.font.italic = false;
    builder.write("text.");

    builder.document.save(base.artifactsDir + "WorkingWithMarkdown.emphases.md");
    //ExEnd:Emphases
  });


  test.skip('UseWarningSource - TODO: warningCallback not supported yet', () => {
    //ExStart:UseWarningSourceMarkdown
    let doc = new aw.Document(base.myDir + "Emphases markdown warning.docx");

    let warnings = new aw.WarningInfoCollection();
    doc.warningCallback = warnings;

    doc.save(base.artifactsDir + "WorkingWithMarkdown.UseWarningSource.md");

    for (let warningInfo of warnings)
    {
      if (warningInfo.source == aw.WarningSource.Markdown)
        console.log(warningInfo.description);
    }
    //ExEnd:UseWarningSourceMarkdown
  });

});