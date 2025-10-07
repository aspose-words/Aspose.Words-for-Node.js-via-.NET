// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("DocumentFormatting", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('SpaceBetweenAsianAndLatinText', () => {
    //ExStart:SpaceBetweenAsianAndLatinText
    //GistId:a7a098b26ccfd16de13a8e6efba00217
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let paragraphFormat = builder.paragraphFormat;
    paragraphFormat.addSpaceBetweenFarEastAndAlpha = true;
    paragraphFormat.addSpaceBetweenFarEastAndDigit = true;

    builder.writeln("Automatically adjust space between Asian and Latin text");
    builder.writeln("Automatically adjust space between Asian text and numbers");

    doc.save(base.artifactsDir + "DocumentFormatting.SpaceBetweenAsianAndLatinText.docx");
    //ExEnd:SpaceBetweenAsianAndLatinText
  });

  test('AsianTypographyLineBreakGroup', () => {
    //ExStart:AsianTypographyLineBreakGroup
    //GistId:a7a098b26ccfd16de13a8e6efba00217
    let doc = new aw.Document(base.myDir + "Asian typography.docx");

    let format = doc.firstSection.body.paragraphs.at(0).paragraphFormat;
    format.farEastLineBreakControl = false;
    format.wordWrap = true;
    format.hangingPunctuation = false;

    doc.save(base.artifactsDir + "DocumentFormatting.AsianTypographyLineBreakGroup.docx");
    //ExEnd:AsianTypographyLineBreakGroup
  });

  test('ParagraphFormatting', () => {
    //ExStart:ParagraphFormatting
    //GistId:fc7e411a082bdf9bd715a4cf28552213
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let paragraphFormat = builder.paragraphFormat;
    paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    paragraphFormat.leftIndent = 50;
    paragraphFormat.rightIndent = 50;
    paragraphFormat.spaceAfter = 25;

    builder.writeln(
        "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
    builder.writeln(
        "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

    doc.save(base.artifactsDir + "DocumentFormatting.ParagraphFormatting.docx");
    //ExEnd:ParagraphFormatting
  });

  test('MultilevelListFormatting', () => {
    //ExStart:MultilevelListFormatting
    //GistId:d8326242115a099a83c0072f78763ca2
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.applyNumberDefault();
    builder.writeln("Item 1");
    builder.writeln("Item 2");

    builder.listFormat.listIndent();
    builder.writeln("Item 2.1");
    builder.writeln("Item 2.2");

    builder.listFormat.listIndent();
    builder.writeln("Item 2.2.1");
    builder.writeln("Item 2.2.2");

    builder.listFormat.listOutdent();
    builder.writeln("Item 2.3");

    builder.listFormat.listOutdent();
    builder.writeln("Item 3");

    builder.listFormat.removeNumbers();

    doc.save(base.artifactsDir + "DocumentFormatting.MultilevelListFormatting.docx");
    //ExEnd:MultilevelListFormatting
  });

  test('ApplyParagraphStyle', () => {
    //ExStart:ApplyParagraphStyle
    //GistId:fc7e411a082bdf9bd715a4cf28552213
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Title;
    builder.write("Hello");

    doc.save(base.artifactsDir + "DocumentFormatting.ApplyParagraphStyle.docx");
    //ExEnd:ApplyParagraphStyle
  });

  test('ApplyBordersAndShadingToParagraph', () => {
    //ExStart:ApplyBordersAndShadingToParagraph
    //GistId:fc7e411a082bdf9bd715a4cf28552213
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let borders = builder.paragraphFormat.borders;
    borders.distanceFromText = 20;
    borders.at(aw.BorderType.Left).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Right).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Top).lineStyle = aw.LineStyle.Double;
    borders.at(aw.BorderType.Bottom).lineStyle = aw.LineStyle.Double;

    let shading = builder.paragraphFormat.shading;
    shading.texture = aw.TextureIndex.TextureDiagonalCross;
    shading.backgroundPatternColor = "#F08080";
    shading.foregroundPatternColor = "#FFA07A";

    builder.write("I'm a formatted paragraph with double border and nice shading.");

    doc.save(base.artifactsDir + "DocumentFormatting.ApplyBordersAndShadingToParagraph.doc");
    //ExEnd:ApplyBordersAndShadingToParagraph
  });

  test('ChangeAsianParagraphSpacingAndIndents', () => {
    //ExStart:ChangeAsianParagraphSpacingAndIndents
    let doc = new aw.Document(base.myDir + "Asian typography.docx");

    let format = doc.firstSection.body.firstParagraph.paragraphFormat;
    format.characterUnitLeftIndent = 10;       // ParagraphFormat.LeftIndent will be updated
    format.characterUnitRightIndent = 10;      // ParagraphFormat.RightIndent will be updated
    format.characterUnitFirstLineIndent = 20;  // ParagraphFormat.FirstLineIndent will be updated
    format.lineUnitBefore = 5;                 // ParagraphFormat.SpaceBefore will be updated
    format.lineUnitAfter = 10;                 // ParagraphFormat.SpaceAfter will be updated

    doc.save(base.artifactsDir + "DocumentFormatting.ChangeAsianParagraphSpacingAndIndents.doc");
    //ExEnd:ChangeAsianParagraphSpacingAndIndents
  });

  test('SnapToGrid', () => {
    //ExStart:SetSnapToGrid
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Optimize the layout when typing in Asian characters.
    let par = doc.firstSection.body.firstParagraph;
    par.paragraphFormat.snapToGrid = true;

    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
        "tempor incididunt ut labore et dolore magna aliqua.");

    par.runs.at(0).font.snapToGrid = true;

    doc.save(base.artifactsDir + "DocumentFormatting.SnapToGrid.docx");
    //ExEnd:SetSnapToGrid
  });

  test('GetParagraphStyleSeparator', () => {
    //ExStart:GetParagraphStyleSeparator
    //GistId:fc7e411a082bdf9bd715a4cf28552213
    let doc = new aw.Document(base.myDir + "Document.docx");

    let paragraphs = doc.getChildNodes(aw.NodeType.Paragraph, true);
    for (let i = 0; i < paragraphs.count; i++) {
      let paragraph = paragraphs.at(i);
      if (paragraph.breakIsStyleSeparator) {
        console.log("Separator Found!");
      }
    }
    //ExEnd:GetParagraphStyleSeparator
  });
});
