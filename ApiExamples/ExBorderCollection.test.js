// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');


describe("ExBorderCollection", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('GetBordersEnumerator', () => {
    //ExStart
    //ExFor:BorderCollection.getEnumerator
    //ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Configure the builder's paragraph format settings to create a green wave border on all sides.
    let borders = builder.paragraphFormat.borders;

    for (let border of borders) {
      border.color = "#008000";
      border.lineStyle = aw.LineStyle.Wave;
      border.lineWidth = 3;
    }

    // Insert a paragraph. Our border settings will determine the appearance of its border.
    builder.writeln("Hello world!");

    doc.save(base.artifactsDir + "BorderCollection.GetBordersEnumerator.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "BorderCollection.GetBordersEnumerator.docx");

    for (let border of doc.firstSection.body.firstParagraph.paragraphFormat.borders) {
      expect(border.color).toEqual("#008000");
      expect(border.lineStyle).toEqual(aw.LineStyle.Wave);
      expect(border.lineWidth).toEqual(3.0);
    }
  });

  test('RemoveAllBorders', () => {
    //ExStart
    //ExFor:BorderCollection.clearFormatting
    //ExSummary:Shows how to remove all borders from all paragraphs in a document.
    let doc = new aw.Document(base.myDir + "Borders.docx");

    // The first paragraph of this document has visible borders with these settings.
    let firstParagraphBorders = doc.firstSection.body.firstParagraph.paragraphFormat.borders;

    expect(firstParagraphBorders.color).toEqual("#FF0000");
    expect(firstParagraphBorders.lineStyle).toEqual(aw.LineStyle.Single);
    expect(firstParagraphBorders.lineWidth).toEqual(3.0);

    // Use the "ClearFormatting" method on each paragraph to remove all borders.
    for (let paragraph of doc.firstSection.body.paragraphs.toArray()) {
      paragraph.paragraphFormat.borders.clearFormatting();

      for (let border of paragraph.paragraphFormat.borders)
      {
        expect(border.color).toEqual(base.emptyColor);
        expect(border.lineStyle).toEqual(aw.LineStyle.None);
        expect(border.lineWidth).toEqual(0.0);
      }
    }
            
    doc.save(base.artifactsDir + "BorderCollection.RemoveAllBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "BorderCollection.RemoveAllBorders.docx");

    for (let border of doc.firstSection.body.firstParagraph.paragraphFormat.borders)
    {
      expect(border.color).toEqual(base.emptyColor);
      expect(border.lineStyle).toEqual(aw.LineStyle.None);
      expect(border.lineWidth).toEqual(0.0);
    }
  });
});
