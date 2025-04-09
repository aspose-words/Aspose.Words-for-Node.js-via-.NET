// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const TestUtil = require('./TestUtil');

/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
describe("ExImage", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('FromFile', () => {
    //ExStart
    //ExFor:Shape.#ctor(DocumentBase,ShapeType)
    //ExFor:ShapeType
    //ExSummary:Shows how to insert a shape with an image from the local file system into a document.
    let doc = new aw.Document();

    // The "Shape" class's public constructor will create a shape with "ShapeMarkupLanguage.Vml" markup type.
    // If you need to create a shape of a non-primitive type, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
    // please use DocumentBuilder.insertShape.
    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
    shape.imageData.setImage(base.imageDir + "Windows MetaFile.wmf");
    shape.width = 100;
    shape.height = 100;

    doc.firstSection.body.firstParagraph.appendChild(shape);

    doc.save(base.artifactsDir + "Image.FromFile.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.FromFile.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(1600, 1600, aw.Drawing.ImageType.Wmf, shape);
    expect(shape.height).toEqual(100.0);
    expect(shape.width).toEqual(100.0);
  });


  test('FromUrl', () => {
    //ExStart
    //ExFor:DocumentBuilder.insertImage(String)
    //ExSummary:Shows how to insert a shape with an image into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two locations where the document builder's "InsertShape" method
    // can source the image that the shape will display.
    // 1 -  Pass a local file system filename of an image file:
    builder.write("Image from local file: ");
    builder.insertImage(base.imageDir + "Logo.jpg");
    builder.writeln();

    // 2 -  Pass a URL which points to an image.
    builder.write("Image from a URL: ");
    builder.insertImage(base.imageUrl.toString());
    builder.writeln();

    doc.save(base.artifactsDir + "Image.FromUrl.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.FromUrl.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(2);
    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shapes[0]);
    TestUtil.verifyImageInShape(272, 92, aw.Drawing.ImageType.Png, shapes[1]);
  });


  test.skip('FromStream: WORDSNODEJS-99', () => {
    //ExStart
    //ExFor:DocumentBuilder.insertImage(Stream)
    //ExSummary:Shows how to insert a shape with an image from a stream into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let stream = fs.createReadStream(base.imageDir + "Logo.jpg")
    builder.write("Image from stream: ");
    builder.insertImage(stream);

    doc.save(base.artifactsDir + "Image.FromStream.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.FromStream.docx");

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, doc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape());
  });


  test('CreateFloatingPageCenter', () => {
    //ExStart
    //ExFor:DocumentBuilder.insertImage(String)
    //ExFor:Shape
    //ExFor:ShapeBase
    //ExFor:ShapeBase.wrapType
    //ExFor:ShapeBase.behindText
    //ExFor:ShapeBase.relativeHorizontalPosition
    //ExFor:ShapeBase.relativeVerticalPosition
    //ExFor:ShapeBase.horizontalAlignment
    //ExFor:ShapeBase.verticalAlignment
    //ExFor:WrapType
    //ExFor:RelativeHorizontalPosition
    //ExFor:RelativeVerticalPosition
    //ExFor:HorizontalAlignment
    //ExFor:VerticalAlignment
    //ExSummary:Shows how to insert a floating image to the center of a page.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a floating image that will appear behind the overlapping text and align it to the page's center.
    let shape = builder.insertImage(base.imageDir + "Logo.jpg");
    shape.wrapType = aw.Drawing.WrapType.None;
    shape.behindText = true;
    shape.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
    shape.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    shape.horizontalAlignment = aw.Drawing.HorizontalAlignment.Center;
    shape.verticalAlignment = aw.Drawing.VerticalAlignment.Center;

    doc.save(base.artifactsDir + "Image.CreateFloatingPageCenter.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.CreateFloatingPageCenter.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(shape.behindText).toEqual(true);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Page);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Page);
    expect(shape.horizontalAlignment).toEqual(aw.Drawing.HorizontalAlignment.Center);
    expect(shape.verticalAlignment).toEqual(aw.Drawing.VerticalAlignment.Center);
  });


  test('CreateFloatingPositionSize', () => {
    //ExStart
    //ExFor:ShapeBase.left
    //ExFor:ShapeBase.right
    //ExFor:ShapeBase.top
    //ExFor:ShapeBase.bottom
    //ExFor:ShapeBase.width
    //ExFor:ShapeBase.height
    //ExFor:DocumentBuilder.currentSection
    //ExFor:PageSetup.pageWidth
    //ExSummary:Shows how to insert a floating image, and specify its position and size.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imageDir + "Logo.jpg");
    shape.wrapType = aw.Drawing.WrapType.None;

    // Configure the shape's "RelativeHorizontalPosition" property to treat the value of the "Left" property
    // as the shape's horizontal distance, in points, from the left side of the page. 
    shape.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;

    // Set the shape's horizontal distance from the left side of the page to 100.
    shape.left = 100;

    // Use the "RelativeVerticalPosition" property in a similar way to position the shape 80pt below the top of the page.
    shape.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    shape.top = 80;

    // Set the shape's height, which will automatically scale the width to preserve dimensions.
    shape.height = 125;

    expect(shape.width).toEqual(125.0);

    // The "Bottom" and "Right" properties contain the bottom and right edges of the image.
    expect(shape.bottom).toEqual(shape.top + shape.height);
    expect(shape.right).toEqual(shape.left + shape.width);

    doc.save(base.artifactsDir + "Image.CreateFloatingPositionSize.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.CreateFloatingPositionSize.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(shape.relativeHorizontalPosition).toEqual(aw.Drawing.RelativeHorizontalPosition.Page);
    expect(shape.relativeVerticalPosition).toEqual(aw.Drawing.RelativeVerticalPosition.Page);
    expect(shape.left).toEqual(100.0);
    expect(shape.top).toEqual(80.0);
    expect(shape.height).toEqual(125.0);
    expect(shape.width).toEqual(125.0);
    expect(shape.bottom).toEqual(shape.top + shape.height);
    expect(shape.right).toEqual(shape.left + shape.width);
  });


  test('InsertImageWithHyperlink', () => {
    //ExStart
    //ExFor:ShapeBase.hRef
    //ExFor:ShapeBase.screenTip
    //ExFor:ShapeBase.target
    //ExSummary:Shows how to insert a shape which contains an image, and is also a hyperlink.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imageDir + "Logo.jpg");
    shape.href = "https://forum.aspose.com/";
    shape.target = "New Window";
    shape.screenTip = "Aspose.words Support Forums";

    // Ctrl + left-clicking the shape in Microsoft Word will open a new web browser window
    // and take us to the hyperlink in the "HRef" property.
    doc.save(base.artifactsDir + "Image.InsertImageWithHyperlink.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.InsertImageWithHyperlink.docx");
    shape = doc.getShape(0, true);

    expect(shape.href).toEqual("https://forum.aspose.com/");
    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.target).toEqual("New Window");
    expect(shape.screenTip).toEqual("Aspose.words Support Forums");
  });


  test('CreateLinkedImage', () => {
    //ExStart
    //ExFor:Shape.imageData
    //ExFor:ImageData
    //ExFor:ImageData.sourceFullName
    //ExFor:ImageData.setImage(String)
    //ExFor:DocumentBuilder.insertNode
    //ExSummary:Shows how to insert a linked image into a document. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    const imageFileName = base.imageDir + "Windows MetaFile.wmf";

    // Below are two ways of applying an image to a shape so that it can display it.
    // 1 -  Set the shape to contain the image.
    let shape = new aw.Drawing.Shape(builder.document, aw.Drawing.ShapeType.Image);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.imageData.setImage(imageFileName);

    builder.insertNode(shape);

    doc.save(base.artifactsDir + "Image.CreateLinkedImage.embedded.docx");

    // Every image that we store in shape will increase the size of our document.
    expect(70000 < fs.statSync(base.artifactsDir + "Image.CreateLinkedImage.embedded.docx").size).toBeTruthy();

    doc.firstSection.body.firstParagraph.removeAllChildren();

    // 2 -  Set the shape to link to an image file in the local file system.
    shape = new aw.Drawing.Shape(builder.document, aw.Drawing.ShapeType.Image);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.imageData.sourceFullName = imageFileName;

    builder.insertNode(shape);
    doc.save(base.artifactsDir + "Image.CreateLinkedImage.linked.docx");

    // Linking to images will save space and result in a smaller document.
    // However, the document can only display the image correctly while
    // the image file is present at the location that the shape's "SourceFullName" property points to.
    expect(10000 > fs.statSync(base.artifactsDir + "Image.CreateLinkedImage.linked.docx").size).toBeTruthy();
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.CreateLinkedImage.embedded.docx");

    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(1600, 1600, aw.Drawing.ImageType.Wmf, shape);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(shape.imageData.sourceFullName.replace("%20", " ")).toEqual('');

    doc = new aw.Document(base.artifactsDir + "Image.CreateLinkedImage.linked.docx");

    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(0, 0, aw.Drawing.ImageType.Wmf, shape);
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(shape.imageData.sourceFullName.replace("%20", " ")).toEqual(imageFileName);
  });


  test('DeleteAllImages', () => {
    //ExStart
    //ExFor:Shape.hasImage
    //ExFor:Node.remove
    //ExSummary:Shows how to delete all shapes with images from a document.
    let doc = new aw.Document(base.myDir + "Images.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.filter(s => s.hasImage).length).toEqual(9);

    for (let shape of shapes)
      if (shape.hasImage) 
        shape.remove();

    shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());
    expect(shapes.filter(s => s.hasImage).length).toEqual(0);
    //ExEnd
  });


  test('DeleteAllImagesPreOrder', () => {
    //ExStart
    //ExFor:Node.nextPreOrder(Node)
    //ExFor:Node.previousPreOrder(Node)
    //ExSummary:Shows how to traverse the document's node tree using the pre-order traversal algorithm, and delete any encountered shape with an image.
    let doc = new aw.Document(base.myDir + "Images.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()).filter(s => s.hasImage).length).toEqual(9);

    let curNode = doc;
    while (curNode != null)
    {
      let nextNode = curNode.nextPreOrder(doc);

      if (curNode.previousPreOrder(doc) != null && nextNode != null)
        expect(nextNode.previousPreOrder(doc).referenceEquals(curNode)).toBeTruthy();

      if (curNode.nodeType == aw.NodeType.Shape && curNode.asShape().hasImage)
        curNode.remove();
      curNode = nextNode;
    }

    expect(doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()).filter(s => s.hasImage).length).toEqual(0);
    //ExEnd
  });


  test('ScaleImage', () => {
    //ExStart
    //ExFor:ImageData.imageSize
    //ExFor:ImageSize
    //ExFor:ImageSize.widthPoints
    //ExFor:ImageSize.heightPoints
    //ExFor:ShapeBase.width
    //ExFor:ShapeBase.height
    //ExSummary:Shows how to resize a shape with an image.
    // When we insert an image using the "InsertImage" method, the builder scales the shape that displays the image so that,
    // when we view the document using 100% zoom in Microsoft Word, the shape displays the image in its actual size.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let shape = builder.insertImage(base.imageDir + "Logo.jpg");

    // A 400x400 image will create an ImageData object with an image size of 300x300pt.
    let imageSize = shape.imageData.imageSize;

    expect(imageSize.widthPoints).toEqual(300.0);
    expect(imageSize.heightPoints).toEqual(300.0);

    // If a shape's dimensions match the image data's dimensions,
    // then the shape is displaying the image in its original size.
    expect(shape.width).toEqual(300.0);
    expect(shape.height).toEqual(300.0);

    // Reduce the overall size of the shape by 50%. 
    shape.width *= 0.5;

    // Scaling factors apply to both the width and the height at the same time to preserve the shape's proportions. 
    expect(shape.width).toEqual(150.0);
    expect(shape.height).toEqual(150.0);

    // When we resize the shape, the size of the image data remains the same.
    expect(imageSize.widthPoints).toEqual(300.0);
    expect(imageSize.heightPoints).toEqual(300.0);

    // We can reference the image data dimensions to apply a scaling based on the size of the image.
    shape.width = imageSize.widthPoints * 1.1;

    expect(shape.width).toEqual(330.0);
    expect(shape.height).toEqual(330.0);

    doc.save(base.artifactsDir + "Image.ScaleImage.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Image.ScaleImage.docx");
    shape = doc.getShape(0, true);

    expect(shape.width).toEqual(330.0);
    expect(shape.height).toEqual(330.0);

    imageSize = shape.imageData.imageSize;

    expect(imageSize.widthPoints).toEqual(300.0);
    expect(imageSize.heightPoints).toEqual(300.0);
  });


  test('InsertWebpImage', () => {
    //ExStart:InsertWebpImage
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:DocumentBuilder.insertImage(String)
    //ExSummary:Shows how to insert WebP image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    builder.insertImage(base.imageDir + "WebP image.webp");

    doc.save(base.artifactsDir + "Image.InsertWebpImage.docx");
    //ExEnd:InsertWebpImage
  });


  test('ReadWebpImage', () => {
    //ExStart:ReadWebpImage
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:ImageType
    //ExSummary:Shows how to read WebP image.
    let doc = new aw.Document(base.myDir + "Document with WebP image.docx");

    let shape = doc.getShape(0, true);
    expect(shape.imageData.imageType).toEqual(aw.Drawing.ImageType.WebP);
    //ExEnd:ReadWebpImage
  });


});
