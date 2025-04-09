// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const base = require('./ApiExampleBase').ApiExampleBase;
const aw = require('@aspose/words');


describe("ExBorder", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  beforeEach(() => {
    base.setUnlimitedLicense();
  });

  test('FontBorder', () => {
    //ExStart
    //ExFor:Border
    //ExFor:Border.color
    //ExFor:Border.lineWidth
    //ExFor:Border.lineStyle
    //ExFor:Font.border
    //ExFor:LineStyle
    //ExFor:Font
    //ExFor:DocumentBuilder.font
    //ExFor:DocumentBuilder.write(String)
    //ExSummary:Shows how to insert a string surrounded by a border into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.border.color = "#008000";
    builder.font.border.lineWidth = 2.5;
    builder.font.border.lineStyle = aw.LineStyle.DashDotStroker;

    builder.write("Text surrounded by green border.");

    doc.save(base.artifactsDir + "Border.FontBorder.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.FontBorder.docx");
    let border = doc.firstSection.body.firstParagraph.runs.at(0).font.border;

    expect(border.color).toEqual("#008000");
    expect(border.lineWidth).toEqual(2.5);
    expect(border.lineStyle).toEqual(aw.LineStyle.DashDotStroker);
  });

  test('ParagraphTopBorder', () => {
    //ExStart
    //ExFor:BorderCollection
    //ExFor:Border.themeColor
    //ExFor:Border.tintAndShade
    //ExFor:Border
    //ExFor:BorderType
    //ExFor:ParagraphFormat.borders
    //ExSummary:Shows how to insert a paragraph with a top border.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let topBorder = builder.paragraphFormat.borders.top;
    topBorder.lineWidth = 4.0;
    topBorder.lineStyle = aw.LineStyle.DashSmallGap;
    // Set ThemeColor only when LineWidth or LineStyle setted.
    topBorder.themeColor = aw.Themes.ThemeColor.Accent1;
    topBorder.tintAndShade = 0.25;

    builder.writeln("Text with a top border.");

    doc.save(base.artifactsDir + "Border.ParagraphTopBorder.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.ParagraphTopBorder.docx");
    let border = doc.firstSection.body.firstParagraph.paragraphFormat.borders.top;

    expect(border.lineWidth).toEqual(4.0);
    expect(border.lineStyle).toEqual(aw.LineStyle.DashSmallGap);
    expect(border.themeColor).toEqual(aw.Themes.ThemeColor.Accent1);
    expect(border.tintAndShade).toBeCloseTo(0.25, 2);
  });

  test('ClearFormatting', () => {
    //ExStart
    //ExFor:Border.clearFormatting
    //ExFor:Border.isVisible
    //ExSummary:Shows how to remove borders from a paragraph.
    let doc = new aw.Document(base.myDir + "Borders.docx");

    // Each paragraph has an individual set of borders.
    // We can access the settings for the appearance of these borders via the paragraph format object.
    let borders = doc.firstSection.body.firstParagraph.paragraphFormat.borders;

    expect(borders.at(0).color).toEqual("#FF0000");
    expect(borders.at(0).lineWidth).toEqual(3.0);
    expect(borders.at(0).lineStyle).toEqual(aw.LineStyle.Single);
    expect(borders.at(0).isVisible).toEqual(true);

    // We can remove a border at once by running the ClearFormatting method. 
    // Running this method on every border of a paragraph will remove all its borders.
    //[...borders].forEach((b) => b.clearFormatting());
    for (let b of borders)
      b.clearFormatting();

    expect(borders.at(0).color).toEqual(base.emptyColor);
    expect(borders.at(0).lineWidth).toEqual(0.0);
    expect(borders.at(0).lineStyle).toEqual(aw.LineStyle.None);
    expect(borders.at(0).isVisible).toEqual(false);

    doc.save(base.artifactsDir + "Border.clearFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.clearFormatting.docx");

    for (let testBorder of doc.firstSection.body.firstParagraph.paragraphFormat.borders)
    {
      expect(testBorder.color).toEqual(base.emptyColor);
      expect(testBorder.lineWidth).toEqual(0.0);
      expect(testBorder.lineStyle).toEqual(aw.LineStyle.None);
    }
  });

  test('SharedElements', () => {
    //ExStart
    //ExFor:Border.equals(Object)
    //ExFor:Border.equals(Border)
    //ExFor:Border.getHashCode
    //ExFor:BorderCollection.count
    //ExFor:BorderCollection.equals(BorderCollection)
    //ExFor:BorderCollection.item(Int32)
    //ExSummary:Shows how border collections can share elements.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Paragraph 1.");
    builder.write("Paragraph 2.");

    // Since we used the same border configuration while creating
    // these paragraphs, their border collections share the same elements.
    let firstParagraphBorders = doc.firstSection.body.firstParagraph.paragraphFormat.borders;
    let secondParagraphBorders = builder.currentParagraph.paragraphFormat.borders;
    expect(firstParagraphBorders.count).toEqual(6);

    for (let i = 0; i < firstParagraphBorders.count; i++)
    {
      expect(firstParagraphBorders.at(i).equals(secondParagraphBorders.at(i))).toEqual(true);
      expect(firstParagraphBorders.at(i).isVisible).toEqual(false);
    }

    for (let border of secondParagraphBorders)
      border.lineStyle = aw.LineStyle.DotDash;

    // After changing the line style of the borders in just the second paragraph,
    // the border collections no longer share the same elements.
    for (let i = 0; i < firstParagraphBorders.count; i++)
    {
      expect(firstParagraphBorders.at(i).equals(secondParagraphBorders.at(i))).toEqual(false);

      // Changing the appearance of an empty border makes it visible.
      expect(secondParagraphBorders.at(i).isVisible).toEqual(true);
    }

    doc.save(base.artifactsDir + "Border.SharedElements.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.SharedElements.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    for (let testBorder of paragraphs.at(0).paragraphFormat.borders)
      expect(testBorder.lineStyle).toEqual(aw.LineStyle.None);

    for (let testBorder of paragraphs.at(1).paragraphFormat.borders)
      expect(testBorder.lineStyle).toEqual(aw.LineStyle.DotDash);
  });

  test('HorizontalBorders', () => {
    //ExStart
    //ExFor:BorderCollection.horizontal
    //ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
    let borders = doc.firstSection.body.firstParagraph.paragraphFormat.borders;
    borders.horizontal.color = "#FF0000";
    borders.horizontal.lineStyle = aw.LineStyle.DashSmallGap;
    borders.horizontal.lineWidth = 3;

    // Write text to the document without creating a new paragraph afterward.
    // Since there is no paragraph underneath, the horizontal border will not be visible.
    builder.write("Paragraph above horizontal border.");

    // Once we add a second paragraph, the border of the first paragraph will become visible.
    builder.insertParagraph();
    builder.write("Paragraph below horizontal border.");

    doc.save(base.artifactsDir + "Border.HorizontalBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.HorizontalBorders.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.borders.at(aw.BorderType.Horizontal).lineStyle).toEqual(aw.LineStyle.DashSmallGap);
    expect(paragraphs.at(1).paragraphFormat.borders.at(aw.BorderType.Horizontal).lineStyle).toEqual(aw.LineStyle.DashSmallGap);
  });

  test('VerticalBorders', () => {
    //ExStart
    //ExFor:BorderCollection.horizontal
    //ExFor:BorderCollection.vertical
    //ExFor:Cell.lastParagraph
    //ExSummary:Shows how to apply settings to vertical borders to a table row's format.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a table with red and blue inner borders.
    let table = builder.startTable();

    for (let i = 0; i < 3; i++)
    {
      builder.insertCell();
      builder.write(`Row ${i + 1}, Column 1`);
      builder.insertCell();
      builder.write(`Row ${i + 1}, Column 2`);

      let row = builder.endRow();
      let borders = row.rowFormat.borders;

      // Adjust the appearance of borders that will appear between rows.
      borders.horizontal.color = "#FF0000";
      borders.horizontal.lineStyle = aw.LineStyle.Dot;
      borders.horizontal.lineWidth = 2.0;

      // Adjust the appearance of borders that will appear between cells.
      borders.vertical.color = "#0000FF";
      borders.vertical.lineStyle = aw.LineStyle.Dot;
      borders.vertical.lineWidth = 2.0;
    }

    // A row format, and a cell's inner paragraph use different border settings.
    let border = table.firstRow.firstCell.lastParagraph.paragraphFormat.borders.vertical;

    expect(border.color).toEqual(base.emptyColor);
    expect(border.lineWidth).toEqual(0.0);
    expect(border.lineStyle).toEqual(aw.LineStyle.None);

    doc.save(base.artifactsDir + "Border.VerticalBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Border.VerticalBorders.docx");
    table = doc.firstSection.body.tables.at(0);

    for (var node of table.getChildNodes(aw.NodeType.Row, true))
    {
      var row = node.asRow();
      expect(row.rowFormat.borders.horizontal.color).toEqual("#FF0000");
      expect(row.rowFormat.borders.horizontal.lineStyle).toEqual(aw.LineStyle.Dot);
      expect(row.rowFormat.borders.horizontal.lineWidth).toEqual(2.0);

      expect(row.rowFormat.borders.vertical.color).toEqual("#0000FF");
      expect(row.rowFormat.borders.vertical.lineStyle).toEqual(aw.LineStyle.Dot);
      expect(row.rowFormat.borders.vertical.lineWidth).toEqual(2.0);
    }
  });
});
