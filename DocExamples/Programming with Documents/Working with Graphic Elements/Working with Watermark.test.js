// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkWithWatermark", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AddTextWatermark', () => {
    //ExStart:AddTextWatermark
    //GistId:2936fee38a4d9a8256f95e4276400579
    let doc = new aw.Document(base.myDir + "Document.docx");

    let options = new aw.TextWatermarkOptions();
    options.fontFamily = "Arial";
    options.fontSize = 36;
    options.color = "#000000";
    options.layout = aw.WatermarkLayout.Horizontal;
    options.isSemitrasparent = false;

    doc.watermark.setText("Test", options);

    doc.save(base.artifactsDir + "WorkWithWatermark.AddTextWatermark.docx");
    //ExEnd:AddTextWatermark
  });

  test('AddImageWatermark', () => {
    //ExStart:AddImageWatermark
    //GistId:2936fee38a4d9a8256f95e4276400579
    let doc = new aw.Document(base.myDir + "Document.docx");

    let options = new aw.ImageWatermarkOptions();
    options.scale = 5;
    options.isWashout = false;

    doc.watermark.setImage(base.imagesDir + "Transparent background logo.png", options);

    doc.save(base.artifactsDir + "WorkWithWatermark.AddImageWatermark.docx");
    //ExEnd:AddImageWatermark
  });

  test('RemoveDocumentWatermark', () => {
    //ExStart:RemoveDocumentWatermark
    //GistId:2936fee38a4d9a8256f95e4276400579
    let doc = new aw.Document();

    // Add a plain text watermark.
    doc.watermark.setText("Aspose Watermark");

    // If we wish to edit the text formatting using it as a watermark,
    // we can do so by passing a TextWatermarkOptions object when creating the watermark.
    let textWatermarkOptions = new aw.TextWatermarkOptions();
    textWatermarkOptions.fontFamily = "Arial";
    textWatermarkOptions.fontSize = 36;
    textWatermarkOptions.color = "#000000";
    textWatermarkOptions.layout = aw.WatermarkLayout.Diagonal;
    textWatermarkOptions.isSemitrasparent = false;

    doc.watermark.setText("Aspose Watermark", textWatermarkOptions);

    doc.save(base.artifactsDir + "Document.TextWatermark.docx");

    // We can remove a watermark from a document like this.
    if (doc.watermark.type == aw.WatermarkType.Text)
      doc.watermark.remove();

    doc.save(base.artifactsDir + "WorkWithWatermark.RemoveDocumentWatermark.docx");
    //ExEnd:RemoveDocumentWatermark
  });

  //ExStart:AddDocumentWatermark
  //GistId:2936fee38a4d9a8256f95e4276400579
  test('AddAndRemoveWatermark', () => {
    let doc = new aw.Document(base.myDir + "Document.docx");

    insertWatermarkText(doc, "CONFIDENTIAL");
    doc.save(base.artifactsDir + "WorkWithWatermark.AddWatermark.docx");

    removeWatermarkShape(doc);
    doc.save(base.artifactsDir + "WorkWithWatermark.RemoveWatermark.docx");
  });

  // <summary>
  /// Inserts a watermark into a document.
  /// </summary>
  /// <param name="doc">The input document.</param>
  /// <param name="watermarkText">Text of the watermark.</param>
  function insertWatermarkText(doc, watermarkText) {
    //ExStart:SetShapeName
    //GistId:2936fee38a4d9a8256f95e4276400579
    // Create a watermark shape, this will be a WordArt shape.
    let watermark = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextPlainText);
    watermark.name = "Watermark";
    //ExEnd:SetShapeName

    watermark.textPath.text = watermarkText;
    watermark.textPath.fontFamily = "Arial";
    watermark.width = 500;
    watermark.height = 100;

    // Text will be directed from the bottom-left to the top-right corner.
    watermark.rotation = -40;

    // Remove the following two lines if you need a solid black text.
    watermark.fillColor = "#808080";
    watermark.strokeColor = "#808080";

    // Place the watermark in the page center.
    watermark.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
    watermark.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    watermark.wrapType = aw.Drawing.WrapType.None;
    watermark.verticalAlignment = aw.Drawing.VerticalAlignment.Center;
    watermark.horizontalAlignment = aw.Drawing.HorizontalAlignment.Center;

    // Create a new paragraph and append the watermark to this paragraph.
    let watermarkPara = new aw.Paragraph(doc);
    watermarkPara.appendChild(watermark);

    // Insert the watermark into all headers of each document section.
    for (let sect of doc.sections) {
      sect = sect.asSection();
      // There could be up to three different headers in each section.
      // Since we want the watermark to appear on all pages, insert it into all headers.
      insertWatermarkIntoHeader(watermarkPara, sect, aw.HeaderFooterType.HeaderPrimary);
      insertWatermarkIntoHeader(watermarkPara, sect, aw.HeaderFooterType.HeaderFirst);
      insertWatermarkIntoHeader(watermarkPara, sect, aw.HeaderFooterType.HeaderEven);
    }
  }

  function insertWatermarkIntoHeader(watermarkPara, sect, headerType) {
    let header = sect.headersFooters.at(headerType);

    if (header == null) {
      // There is no header of the specified type in the current section, so we need to create it.
      header = new aw.HeaderFooter(sect.document, headerType);
      sect.headersFooters.add(header);
    }

    // Insert a clone of the watermark into the header.
    header.appendChild(watermarkPara.clone(true));
  }
  //ExEnd:AddDocumentWatermark

  //ExStart:RemoveWatermarkShape
  //GistId:2936fee38a4d9a8256f95e4276400579
  function removeWatermarkShape(doc) {
    for (let hf of doc.getChildNodes(aw.NodeType.HeaderFooter, true)) {
      hf = hf.asHeaderFooter();
      for (let shape of hf.getChildNodes(aw.NodeType.Shape, true)) {
        shape = shape.asShape();
        if (shape.name.includes("Watermark")) {
          shape.remove();
        }
      }
    }
  }
  //ExEnd:RemoveWatermarkShape


});
