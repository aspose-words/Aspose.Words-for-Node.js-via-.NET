// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const TestUtil = require('./TestUtil');

describe("ExDocumentBuilderImages", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('InsertImageFromStream', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(Stream)
    //ExFor:aw.DocumentBuilder.insertImage(Stream, Double, Double)
    //ExFor:aw.DocumentBuilder.insertImage(Stream, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from a stream into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let stream = fs.readFileSync(base.imageDir + "Logo.jpg");
    // Below are three ways of inserting an image from a stream.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder.insertImage(stream);

    builder.insertBreak(aw.BreakType.PageBreak);

    // 2 -  Inline shape with custom dimensions:
    builder.insertImage(stream, aw.ConvertUtil.pixelToPoint(250), aw.ConvertUtil.pixelToPoint(144));

    builder.insertBreak(aw.BreakType.PageBreak);

    // 3 -  Floating shape with custom dimensions:
    builder.insertImage(stream, aw.Drawing.RelativeHorizontalPosition.Margin, 100, aw.Drawing.RelativeVerticalPosition.Margin, 100, 200, 100, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertImageFromStream.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilderImages.InsertImageFromStream.docx");

    let imageShape = doc.getShape(0, true).asShape();

    expect(imageShape.height).toEqual(300.0);
    expect(imageShape.width).toEqual(300.0);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(1, true).asShape();

    expect(imageShape.height).toEqual(108.0);
    expect(imageShape.width).toEqual(187.5);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(2, true).asShape();

    expect(imageShape.height).toEqual(100.0);
    expect(imageShape.width).toEqual(200.0);
    expect(imageShape.left).toEqual(100.0);
    expect(imageShape.top).toEqual(100.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);
  });


  test('InsertImageFromFilename', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(String)
    //ExFor:aw.DocumentBuilder.insertImage(String, Double, Double)
    //ExFor:aw.DocumentBuilder.insertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from the local file system into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are three ways of inserting an image from a local system filename.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder.insertImage(base.imageDir + "Logo.jpg");

    builder.insertBreak(aw.BreakType.PageBreak);

    // 2 -  Inline shape with custom dimensions:
    builder.insertImage(base.imageDir + "Transparent background logo.png", aw.ConvertUtil.pixelToPoint(250), aw.ConvertUtil.pixelToPoint(144));

    builder.insertBreak(aw.BreakType.PageBreak);

    // 3 -  Floating shape with custom dimensions:
    builder.insertImage(base.imageDir + "Windows MetaFile.wmf", aw.Drawing.RelativeHorizontalPosition.Margin, 100,  aw.Drawing.RelativeVerticalPosition.Margin, 100, 200, 100, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertImageFromFilename.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilderImages.InsertImageFromFilename.docx");

    let imageShape = doc.getShape(0, true).asShape();

    expect(imageShape.height).toEqual(300.0);
    expect(imageShape.width).toEqual(300.0);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(1, true).asShape();

    expect(imageShape.height).toEqual(108.0);
    expect(imageShape.width).toEqual(187.5);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Png, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(2, true).asShape();

    expect(imageShape.height).toEqual(100.0);
    expect(imageShape.width).toEqual(200.0);
    expect(imageShape.left).toEqual(100.0);
    expect(imageShape.top).toEqual(100.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);

    TestUtil.verifyImageInShape(1600, 1600, aw.Drawing.ImageType.Wmf, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(400.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(400.0);
  });


  test('InsertSvgImage', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(String)
    //ExSummary:Shows how to determine which image will be inserted.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imageDir + "Scalable Vector Graphics.svg");

    // Aspose.words insert SVG image to the document as PNG with svgBlip extension
    // that contains the original vector SVG image representation.
    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertSvgImage.SvgWithSvgBlip.docx");

    // Aspose.words insert SVG image to the document as PNG, just like Microsoft Word does for old format.
    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertSvgImage.svg.doc");

    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2003);

    // Aspose.words insert SVG image to the document as EMF metafile to keep the image in vector representation.
    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertSvgImage.emf.docx");
    //ExEnd
  });


  test('InsertImageFromImageObject', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(Image)
    //ExFor:aw.DocumentBuilder.insertImage(Image, Double, Double)
    //ExFor:aw.DocumentBuilder.insertImage(Image, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from an object into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let imageFile = base.imageDir + "Logo.jpg";

    // Below are three ways of inserting an image from an Image object instance.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder.insertImage(imageFile);

    builder.insertBreak(aw.BreakType.PageBreak);

    // 2 -  Inline shape with custom dimensions:
    builder.insertImage(imageFile, aw.ConvertUtil.pixelToPoint(250), aw.ConvertUtil.pixelToPoint(144));

    builder.insertBreak(aw.BreakType.PageBreak);

    // 3 -  Floating shape with custom dimensions:
    builder.insertImage(imageFile, aw.Drawing.RelativeHorizontalPosition.Margin, 100, aw.Drawing.RelativeVerticalPosition.Margin, 100, 200, 100, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilderImages.InsertImageFromImageObject.docx");

    let imageShape = doc.getShape(0, true).asShape();

    expect(imageShape.height).toEqual(300.0);
    expect(imageShape.width).toEqual(300.0);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(1, true).asShape();

    expect(imageShape.height).toEqual(108.0);
    expect(imageShape.width).toEqual(187.5);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);

    imageShape = doc.getShape(2, true).asShape();

    expect(imageShape.height).toEqual(100.0);
    expect(imageShape.width).toEqual(200.0);
    expect(imageShape.left).toEqual(100.0);
    expect(imageShape.top).toEqual(100.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toEqual(300.0);
    expect(imageShape.imageData.imageSize.widthPoints).toEqual(300.0);
  });


  test('InsertImageFromByteArray', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(Byte[])
    //ExFor:aw.DocumentBuilder.insertImage(Byte[], Double, Double)
    //ExFor:aw.DocumentBuilder.insertImage(Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
    //ExSummary:Shows how to insert an image from a byte array into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let imageByteArray = fs.readFileSync(base.imageDir + "Logo.jpg");

    // Below are three ways of inserting an image from a byte array.
    // 1 -  Inline shape with a default size based on the image's original dimensions:
    builder.insertImage(imageByteArray);

    builder.insertBreak(aw.BreakType.PageBreak);

    // 2 -  Inline shape with custom dimensions:
    builder.insertImage(imageByteArray, aw.ConvertUtil.pixelToPoint(250), aw.ConvertUtil.pixelToPoint(144));

    builder.insertBreak(aw.BreakType.PageBreak);

    // 3 -  Floating shape with custom dimensions:
    builder.insertImage(imageByteArray, aw.Drawing.RelativeHorizontalPosition.Margin, 100, aw.Drawing.RelativeVerticalPosition.Margin, 100, 200, 100, aw.Drawing.WrapType.Square);

    doc.save(base.artifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilderImages.InsertImageFromByteArray.docx");

    let imageShape = doc.getShape(0, true).asShape();

    expect(imageShape.height).toBeCloseTo(300.0, 1);
    expect(imageShape.width).toBeCloseTo(300.0, 1);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toBeCloseTo(300.0, 1);
    expect(imageShape.imageData.imageSize.widthPoints).toBeCloseTo(300.0, 1);

    imageShape = doc.getShape(1, true).asShape();

    expect(imageShape.height).toEqual(108.0);
    expect(imageShape.width).toEqual(187.5);
    expect(imageShape.left).toEqual(0.0);
    expect(imageShape.top).toEqual(0.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Column);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toBeCloseTo(300.0, 1);
    expect(imageShape.imageData.imageSize.widthPoints).toBeCloseTo(300.0, 1);

    imageShape = doc.getShape(2, true).asShape();

    expect(imageShape.height).toEqual(100.0);
    expect(imageShape.width).toEqual(200.0);
    expect(imageShape.left).toEqual(100.0);
    expect(imageShape.top).toEqual(100.0);

    expect(imageShape.wrapType).toEqual(aw.Drawing.WrapType.Square);
    expect(imageShape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
    expect(imageShape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Margin);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    expect(imageShape.imageData.imageSize.heightPoints).toBeCloseTo(300.0, 1);
    expect(imageShape.imageData.imageSize.widthPoints).toBeCloseTo(300.0, 1);
  });


  test('InsertGif', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertImage(String)
    //ExSummary:Shows how to insert gif image to the document.
    let builder = new aw.DocumentBuilder();

    // We can insert gif image using path or bytes array.
    // It works only if DocumentBuilder optimized to Word version 2010 or higher.
    // Note, that access to the image bytes causes conversion Gif to Png.
    let gifImage = builder.insertImage(base.imageDir + "Graphics Interchange Format.gif");

    gifImage = builder.insertImage(fs.readFileSync(base.imageDir + "Graphics Interchange Format.gif"));

    builder.document.save(base.artifactsDir + "InsertGif.docx");
    //ExEnd
  });

});
