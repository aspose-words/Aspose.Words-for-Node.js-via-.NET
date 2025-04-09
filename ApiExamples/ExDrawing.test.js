// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
let path = require('path')
const TestUtil = require('./TestUtil');

describe("ExDrawing", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    //base.oneTimeTearDown();
  });

/*//Commented
#if NET461_OR_GREATER || JAVA
  test('VariousShapes', () => {
    //ExStart
    //ExFor:ArrowLength
    //ExFor:ArrowType
    //ExFor:ArrowWidth
    //ExFor:DashStyle
    //ExFor:EndCap
    //ExFor:Fill.foreColor
    //ExFor:Fill.imageBytes
    //ExFor:Fill.visible
    //ExFor:JoinStyle
    //ExFor:Shape.stroke
    //ExFor:Stroke.color
    //ExFor:Stroke.startArrowLength
    //ExFor:Stroke.startArrowType
    //ExFor:Stroke.startArrowWidth
    //ExFor:Stroke.endArrowLength
    //ExFor:Stroke.endArrowWidth
    //ExFor:Stroke.dashStyle
    //ExFor:Stroke.endArrowType
    //ExFor:Stroke.endCap
    //ExFor:Stroke.opacity
    //ExSummary:Shows to create a variety of shapes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are four examples of shapes that we can insert into our documents.
    // 1 -  Dotted, horizontal, half-transparent red line
    // with an arrow on the left end and a diamond on the right end:
    let arrow = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Line);
    arrow.width = 200;
    arrow.stroke.color = "#FF0000";
    arrow.stroke.startArrowType = aw.Drawing.ArrowType.Arrow;
    arrow.stroke.startArrowLength = aw.Drawing.ArrowLength.Long;
    arrow.stroke.startArrowWidth = aw.Drawing.ArrowWidth.Wide;
    arrow.stroke.endArrowType = aw.Drawing.ArrowType.Diamond;
    arrow.stroke.endArrowLength = aw.Drawing.ArrowLength.Long;
    arrow.stroke.endArrowWidth = aw.Drawing.ArrowWidth.Wide;
    arrow.stroke.dashStyle = aw.Drawing.DashStyle.Dash;
    arrow.stroke.opacity = 0.5;

    expect(arrow.stroke.joinStyle).toEqual(aw.Drawing.JoinStyle.Miter);

    builder.insertNode(arrow);

    // 2 -  Thick black diagonal line with rounded ends:
    let line = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Line);
    line.top = 40;
    line.width = 200;
    line.height = 20;
    line.strokeWeight = 5.0;
    line.stroke.endCap = aw.Drawing.EndCap.Round;

    builder.insertNode(line);

    // 3 -  Arrow with a green fill:
    let filledInArrow = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Arrow);
    filledInArrow.width = 200;
    filledInArrow.height = 40;
    filledInArrow.top = 100;
    filledInArrow.fill.foreColor = "#008000";
    filledInArrow.fill.visible = true;

    builder.insertNode(filledInArrow);

    // 4 -  Arrow with a flipped orientation filled in with the Aspose logo:
    let filledInArrowImg = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Arrow);
    filledInArrowImg.width = 200;
    filledInArrowImg.height = 40;
    filledInArrowImg.top = 160;
    filledInArrowImg.flipOrientation = aw.Drawing.FlipOrientation.Both;

    byte[] imageBytes = File.ReadAllBytes(base.imageDir + "Logo.jpg");

    using (MemoryStream stream = new MemoryStream(imageBytes))
    {
      Image image = Image.FromStream(stream);
      // When we flip the orientation of our arrow, we also flip the image that the arrow contains.
      // Flip the image the other way to cancel this out before getting the shape to display it.
      image.RotateFlip(RotateFlipType.RotateNoneFlipXY);

      filledInArrowImg.imageData.setImage(image);
      filledInArrowImg.stroke.joinStyle = aw.Drawing.JoinStyle.Round;

      builder.insertNode(filledInArrowImg);
    }

    doc.save(base.artifactsDir + "Drawing.VariousShapes.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Drawing.VariousShapes.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(4);

    arrow = (Shape) doc.getShape(0, true);

    expect(arrow.shapeType).toEqual(aw.Drawing.ShapeType.Line);
    expect(arrow.width).toEqual(200.0);
    expect(arrow.stroke.color).toEqual("#FF0000");
    expect(arrow.stroke.startArrowType).toEqual(aw.Drawing.ArrowType.Arrow);
    expect(arrow.stroke.startArrowLength).toEqual(aw.Drawing.ArrowLength.Long);
    expect(arrow.stroke.startArrowWidth).toEqual(aw.Drawing.ArrowWidth.Wide);
    expect(arrow.stroke.endArrowType).toEqual(aw.Drawing.ArrowType.Diamond);
    expect(arrow.stroke.endArrowLength).toEqual(aw.Drawing.ArrowLength.Long);
    expect(arrow.stroke.endArrowWidth).toEqual(aw.Drawing.ArrowWidth.Wide);
    expect(arrow.stroke.dashStyle).toEqual(aw.Drawing.DashStyle.Dash);
    expect(arrow.stroke.opacity).toEqual(0.5);

    line = (Shape) doc.getShape(1, true);

    expect(line.shapeType).toEqual(aw.Drawing.ShapeType.Line);
    expect(line.top).toEqual(40.0);
    expect(line.width).toEqual(200.0);
    expect(line.height).toEqual(20.0);
    expect(line.strokeWeight).toEqual(5.0);
    expect(line.stroke.endCap).toEqual(aw.Drawing.EndCap.Round);

    filledInArrow = (Shape) doc.getShape(2, true);

    expect(filledInArrow.shapeType).toEqual(aw.Drawing.ShapeType.Arrow);
    expect(filledInArrow.width).toEqual(200.0);
    expect(filledInArrow.height).toEqual(40.0);
    expect(filledInArrow.top).toEqual(100.0);
    expect(filledInArrow.fill.foreColor).toEqual("#008000");
    expect(filledInArrow.fill.visible).toEqual(true);

    filledInArrowImg = (Shape) doc.getShape(3, true);

    expect(filledInArrowImg.shapeType).toEqual(aw.Drawing.ShapeType.Arrow);
    expect(filledInArrowImg.width).toEqual(200.0);
    expect(filledInArrowImg.height).toEqual(40.0);
    expect(filledInArrowImg.top).toEqual(160.0);
    expect(filledInArrowImg.flipOrientation).toEqual(aw.Drawing.FlipOrientation.Both);
  });


  test('ImportImage', () => {
    //ExStart
    //ExFor:ImageData.setImage(Image)
    //ExFor:ImageData.setImage(Stream)
    //ExSummary:Shows how to display images from the local file system in a document.
    let doc = new aw.Document();

    // To display an image in a document, we will need to create a shape
    // which will contain an image, and then append it to the document's body.
    Shape imgShape;

    // Below are two ways of getting an image from a file in the local file system.
    // 1 -  Create an image object from an image file:
    using (Image srcImage = Image.FromFile(base.imageDir + "Logo.jpg"))
    {
      imgShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
      doc.firstSection.body.firstParagraph.appendChild(imgShape);
      imgShape.imageData.setImage(srcImage);
    }
            
    // 2 -  Open an image file from the local file system using a stream:
    using (Stream stream = new FileStream(base.imageDir + "Logo.jpg", FileMode.open, FileAccess.read))
    {
      imgShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
      doc.firstSection.body.firstParagraph.appendChild(imgShape);
      imgShape.imageData.setImage(stream);
      imgShape.left = 150.0f;
    }

    doc.save(base.artifactsDir + "Drawing.ImportImage.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Drawing.ImportImage.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(2);

    imgShape = (Shape)doc.getShape(0, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imgShape);
    expect(imgShape.left).toEqual(0.0);
    expect(imgShape.top).toEqual(0.0);
    expect(imgShape.height).toEqual(300.0);
    expect(imgShape.width).toEqual(300.0);
    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imgShape);

    imgShape = (Shape)doc.getShape(1, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imgShape);
    expect(imgShape.left).toEqual(150.0);
    expect(imgShape.top).toEqual(0.0);
    expect(imgShape.height).toEqual(300.0);
    expect(imgShape.width).toEqual(300.0);
  });

#endif
//EndCommented*/

  test('TypeOfImage', () => {
    //ExStart
    //ExFor:ImageType
    //ExSummary:Shows how to add an image to a shape and check its type.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let imgShape = builder.insertImage(base.imageDir + "Logo.jpg");
    expect(imgShape.imageData.imageType).toEqual(aw.Drawing.ImageType.Jpeg);
    //ExEnd
  });


  test('FillSolid', () => {
    //ExStart
    //ExFor:Fill.color()
    //ExFor:FillType
    //ExFor:Fill.fillType
    //ExFor:Fill.solid
    //ExFor:Fill.transparency
    //ExFor:Font.fill
    //ExSummary:Shows how to convert any of the fills back to solid fill.
    let doc = new aw.Document(base.myDir + "Two color gradient.docx");

    // Get Fill object for Font of the first Run.
    let fill = doc.firstSection.body.paragraphs.at(0).runs.at(0).font.fill;

    // Check Fill properties of the Font.
    console.log("The type of the fill is: {0}", fill.fillType);
    console.log("The foreground color of the fill is: {0}", fill.foreColor);
    console.log("The fill is transparent at {0}%", fill.transparency * 100);

    // Change type of the fill to Solid with uniform green color.
    fill.solid();
    console.log("\nThe fill is changed:");
    console.log("The type of the fill is: {0}", fill.fillType);
    console.log("The foreground color of the fill is: {0}", fill.foreColor);
    console.log("The fill transparency is {0}%", fill.transparency * 100);

    doc.save(base.artifactsDir + "Drawing.FillSolid.docx");
    //ExEnd
  });

  const imageTypeMap = new Map([
    [0, "No"],
    [1, "No"],
    [2, "Emf"],
    [3, "Wmf"],
    [4, "Pict"],
    [5, "Jpeg"],
    [6, "Png"],
    [7, "Bmp"],
    [8, "Eps"],
    [9, "Webp"],
    [10, "Gif"],
  ]);

  test('SaveAllImages', async () => {
    //ExStart
    //ExFor:ImageData.hasImage
    //ExFor:ImageData.toImage
    //ExFor:ImageData.save(Stream)
    //ExSummary:Shows how to save all images from a document to the file system.
    let imgSourceDoc = new aw.Document(base.myDir + "Images.docx");

    // Shapes with the "HasImage" flag set store and display all the document's images.
    let shapesWithImages = imgSourceDoc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()).filter(s => s.hasImage);

    // Go through each shape and save its image.
    for (let shapeIndex = 0; shapeIndex < shapesWithImages.length; ++shapeIndex)
    {
      let imageData = shapesWithImages.at(shapeIndex).imageData;
      let filename = base.artifactsDir + `Drawing.SaveAllImages.${shapeIndex + 1}.${imageTypeMap.get(imageData.imageType)}`;
      //imageData.save(filename);
      let fileStream = fs.createWriteStream(filename);
      console.log("NNNNNNNN:" + shapeIndex);
      imageData.save(fileStream);
      await new Promise(resolved => fileStream.on('finish', resolved));
    }
    //ExEnd

    let imageFileNames = fs.readdirSync(base.artifactsDir).filter(name => name.startsWith("Drawing.SaveAllImages."));
    imageFileNames.sort();
    await TestUtil.verifyImage(2467, 1500, base.artifactsDir + imageFileNames.at(0));
    expect(path.extname(imageFileNames.at(0))).toEqual(".Jpeg");
    await TestUtil.verifyImage(400, 400, base.artifactsDir + imageFileNames.at(1));
    expect(path.extname(imageFileNames.at(1))).toEqual(".Png");
    await TestUtil.verifyImage(1260, 660, base.artifactsDir + imageFileNames.at(5));
    expect(path.extname(imageFileNames.at(5))).toEqual(".Jpeg");
    await TestUtil.verifyImage(1125, 1500, base.artifactsDir + imageFileNames.at(6));
    expect(path.extname(imageFileNames.at(6))).toEqual(".Jpeg");
    await TestUtil.verifyImage(1027, 1500, base.artifactsDir + imageFileNames.at(7));
    expect(path.extname(imageFileNames.at(7))).toEqual(".Jpeg");
    await TestUtil.verifyImage(1200, 1500, base.artifactsDir + imageFileNames.at(8));
    expect(path.extname(imageFileNames.at(8))).toEqual(".Jpeg");
  });


  test('StrokePattern', async () => {
    //ExStart
    //ExFor:Stroke.color2
    //ExFor:Stroke.imageBytes
    //ExSummary:Shows how to process shape stroke features.
    let doc = new aw.Document(base.myDir + "Shape stroke pattern border.docx");
    let shape = doc.getShape(0, true).asShape();
    let stroke = shape.stroke;

    // Strokes can have two colors, which are used to create a pattern defined by two-tone image data.
    // Strokes with a single color do not use the Color2 property.
    expect(stroke.color).toEqual("#800000");
    //expect(128, 0, 0), stroke.color).toEqual(Color.FromArgb(255);
    expect(stroke.color2).toEqual("#FFFF00");
    //expect(255, 255, 0), stroke.color2).toEqual(Color.FromArgb(255);

    expect(stroke.imageBytes).not.toBe(null);
    fs.writeFileSync(base.artifactsDir + "Drawing.StrokePattern.png", Buffer.from(stroke.imageBytes));
    //ExEnd

    await TestUtil.verifyImage(8, 8, base.artifactsDir + "Drawing.StrokePattern.png");
  });


  /*//Commented
  //ExStart
    //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:GroupShape
    //ExFor:GroupShape.#ctor(DocumentBase)
    //ExFor:GroupShape.Accept(DocumentVisitor)
    //ExFor:GroupShape.AcceptStart(DocumentVisitor)
    //ExFor:GroupShape.AcceptEnd(DocumentVisitor)
    //ExFor:ShapeBase.IsGroup
    //ExFor:ShapeBase.ShapeType
    //ExSummary:Shows how to create a group of shapes, and print its contents using a document visitor.
  test('GroupOfShapes', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    // If you need to create "NonPrimitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, DiagonalCornersRounded
    // please use DocumentBuilder.insertShape methods.
    let balloon = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Balloon)
    {
      Width = 200,
      Height = 200,
      Stroke = { Color = "#FF0000" }
    };

    let cube = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Cube)
    {
      Width = 100,
      Height = 100,
      Stroke = { Color = "#0000FF" }
    };

    let group = new aw.Drawing.GroupShape(doc);
    group.appendChild(balloon);
    group.appendChild(cube);

    expect(group.isGroup).toEqual(true);

    builder.insertNode(group);

    let printer = new ShapeGroupPrinter();
    group.accept(printer);

    console.log(printer.getText());
    TestGroupShapes(doc); //ExSkip
  });


    /// <summary>
    /// Prints the contents of a visited shape group to the console.
    /// </summary>
  public class ShapeGroupPrinter : DocumentVisitor
  {
    public ShapeGroupPrinter()
    {
      mBuilder = new StringBuilder();
    }

    public string GetText()
    {
      return mBuilder.toString();
    }

    public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
    {
      mBuilder.AppendLine("Shape group started:");
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
    {
      mBuilder.AppendLine("End of shape group");
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitShapeStart(Shape shape)
    {
      mBuilder.AppendLine("\tShape - " + shape.shapeType + ":");
      mBuilder.AppendLine("\t\tWidth: " + shape.width);
      mBuilder.AppendLine("\t\tHeight: " + shape.height);
      mBuilder.AppendLine("\t\tStroke color: " + shape.stroke.color);
      mBuilder.AppendLine("\t\tFill color: " + shape.fill.foreColor);
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitShapeEnd(Shape shape)
    {
      mBuilder.AppendLine("\tEnd of shape");
      return aw.VisitorAction.Continue;
    }

    private readonly StringBuilder mBuilder;
  }
    //ExEnd

  private static void TestGroupShapes(Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);
    let shapes = (GroupShape)doc.getChild(aw.NodeType.GroupShape, 0, true);

    expect(shapes.getChildNodes(aw.NodeType.Any, false).Count).toEqual(2);

    let shape = (Shape)shapes.getChildNodes(aw.NodeType.Any, false)[0];

    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.Balloon);
    expect(shape.width).toEqual(200.0);
    expect(shape.height).toEqual(200.0);
    expect(shape.strokeColor).toEqual("#FF0000");

    shape = (Shape)shapes.getChildNodes(aw.NodeType.Any, false)[1];

    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.Cube);
    expect(shape.width).toEqual(100.0);
    expect(shape.height).toEqual(100.0);
     expect(shape.strokeColor).toEqual("#0000FF");
  }
//EndCommented*/    

  test('TextBox', async () => {
    //ExStart
    //ExFor:LayoutFlow
    //ExSummary:Shows how to add text to a text box, and change its orientation
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textbox = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    textbox.width = 100;
    textbox.height = 100;
    textbox.textBox.layoutFlow = aw.Drawing.LayoutFlow.BottomToTop;

    textbox.appendChild(new aw.Paragraph(doc));
    builder.insertNode(textbox);

    builder.moveTo(textbox.firstParagraph);
    builder.write("This text is flipped 90 degrees to the left.");

    doc.save(base.artifactsDir + "Drawing.textBox.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Drawing.textBox.docx");
    textbox = doc.getShape(0, true).asShape();

    expect(textbox.shapeType).toEqual(aw.Drawing.ShapeType.TextBox);
    expect(textbox.width).toEqual(100.0);
    expect(textbox.height).toEqual(100.0);
    expect(textbox.textBox.layoutFlow).toEqual(aw.Drawing.LayoutFlow.BottomToTop);
    expect(textbox.getText().trim()).toEqual("This text is flipped 90 degrees to the left.");
  });


  test.skip('GetDataFromImage - TODO: Aspose.Words.Drawing.ImageData.ToStream() is skipped', async () => {
    //ExStart
    //ExFor:ImageData.imageBytes
    //ExFor:ImageData.toByteArray
    //ExFor:ImageData.toStream
    //ExSummary:Shows how to create an image file from a shape's raw image data.
    let imgSourceDoc = new aw.Document(base.myDir + "Images.docx");
    expect(imgSourceDoc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(10);

    let imgShape = imgSourceDoc.getShape(0, true).asShape();

    expect(imgShape.hasImage).toEqual(true);

    // ToByteArray() returns the array stored in the ImageBytes property.
    expect(imgShape.imageData.toByteArray()).toEqual(imgShape.imageData.imageBytes);

    // Save the shape's image data to an image file in the local file system.
    /*using (Stream imgStream = imgShape.imageData.toStream())
    {
      using (FileStream outStream = new FileStream(base.artifactsDir + "Drawing.GetDataFromImage.png",
        FileMode.create, FileAccess.ReadWrite))
      {
        imgStream.copyTo(outStream);
      }
    }*/
    //ExEnd

    await TestUtil.verifyImage(2467, 1500, base.artifactsDir + "Drawing.GetDataFromImage.png");
  });


  test('ImageData', () => {
    //ExStart
    //ExFor:ImageData.biLevel
    //ExFor:ImageData.borders
    //ExFor:ImageData.brightness
    //ExFor:ImageData.chromaKey
    //ExFor:ImageData.contrast
    //ExFor:ImageData.cropBottom
    //ExFor:ImageData.cropLeft
    //ExFor:ImageData.cropRight
    //ExFor:ImageData.cropTop
    //ExFor:ImageData.grayScale
    //ExFor:ImageData.isLink
    //ExFor:ImageData.isLinkOnly
    //ExFor:ImageData.title
    //ExSummary:Shows how to edit a shape's image data.
    let imgSourceDoc = new aw.Document(base.myDir + "Images.docx");
    let sourceShape = imgSourceDoc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape();

    let dstDoc = new aw.Document();

    // Import a shape from the source document and append it to the first paragraph.
    let importedShape = dstDoc.importNode(sourceShape, true).asShape();
    dstDoc.firstSection.body.firstParagraph.appendChild(importedShape);

    // The imported shape contains an image. We can access the image's properties and raw data via the ImageData object.
    let imageData = importedShape.imageData;
    imageData.title = "Imported Image";

    expect(imageData.hasImage).toEqual(true);

    // If an image has no borders, its ImageData object will define the border color as empty.
    expect(imageData.borders.count).toEqual(4);
    expect(imageData.borders.at(0).color).toEqual(base.emptyColor);

    // This image does not link to another shape or image file in the local file system.
    expect(imageData.isLink).toEqual(false);
    expect(imageData.isLinkOnly).toEqual(false);

    // The "Brightness" and "Contrast" properties define image brightness and contrast
    // on a 0-1 scale, with the default value at 0.5.
    imageData.brightness = 0.8;
    imageData.contrast = 1.0;

    // The above brightness and contrast values have created an image with a lot of white.
    // We can select a color with the ChromaKey property to replace with transparency, such as white.
    imageData.chromaKey = "#FFFFFF";

    // Import the source shape again and set the image to monochrome.
    importedShape = dstDoc.importNode(sourceShape, true).asShape();
    dstDoc.firstSection.body.firstParagraph.appendChild(importedShape);

    importedShape.imageData.grayScale = true;

    // Import the source shape again to create a third image and set it to BiLevel.
    // BiLevel sets every pixel to either black or white, whichever is closer to the original color.
    importedShape = dstDoc.importNode(sourceShape, true).asShape();
    dstDoc.firstSection.body.firstParagraph.appendChild(importedShape);

    importedShape.imageData.biLevel = true;

    // Cropping is determined on a 0-1 scale. Cropping a side by 0.3
    // will crop 30% of the image out at the cropped side.
    importedShape.imageData.cropBottom = 0.3;
    importedShape.imageData.cropLeft = 0.3;
    importedShape.imageData.cropTop = 0.3;
    importedShape.imageData.cropRight = 0.3;

    dstDoc.save(base.artifactsDir + "Drawing.imageData.docx");
    //ExEnd

    imgSourceDoc = new aw.Document(base.artifactsDir + "Drawing.imageData.docx");
    sourceShape = imgSourceDoc.getShape(0, true).asShape();

    TestUtil.verifyImageInShape(2467, 1500, aw.Drawing.ImageType.Jpeg, sourceShape);
    expect(sourceShape.imageData.title).toEqual("Imported Image");
    expect(sourceShape.imageData.brightness).toBeCloseTo(0.8, 1);
    expect(sourceShape.imageData.contrast).toBeCloseTo(1.0, 1);
    expect(sourceShape.imageData.chromaKey).toEqual("#FFFFFF");

    sourceShape = imgSourceDoc.getShape(1, true).asShape();

    TestUtil.verifyImageInShape(2467, 1500, aw.Drawing.ImageType.Jpeg, sourceShape);
    expect(sourceShape.imageData.grayScale).toEqual(true);

    sourceShape = imgSourceDoc.getShape(2, true).asShape();

    TestUtil.verifyImageInShape(2467, 1500, aw.Drawing.ImageType.Jpeg, sourceShape);
    expect(sourceShape.imageData.biLevel).toEqual(true);
    expect(sourceShape.imageData.cropBottom).toBeCloseTo(0.3, 1);
    expect(sourceShape.imageData.cropLeft).toBeCloseTo(0.3, 1);
    expect(sourceShape.imageData.cropTop).toBeCloseTo(0.3, 1);
    expect(sourceShape.imageData.cropRight).toBeCloseTo(0.3, 1);
  });


  test('ImageSize', () => {
    //ExStart
    //ExFor:ImageSize.heightPixels
    //ExFor:ImageSize.horizontalResolution
    //ExFor:ImageSize.verticalResolution
    //ExFor:ImageSize.widthPixels
    //ExSummary:Shows how to read the properties of an image in a shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a shape into the document which contains an image taken from our local file system.
    let shape = builder.insertImage(base.imageDir + "Logo.jpg");

    // If the shape contains an image, its ImageData property will be valid,
    // and it will contain an ImageSize object.
    let imageSize = shape.imageData.imageSize;

    // The ImageSize object contains read-only information about the image within the shape.
    expect(imageSize.heightPixels).toEqual(400);
    expect(imageSize.widthPixels).toEqual(400);

    const delta = 0.05;
    expect(Math.abs(imageSize.horizontalResolution - 95.98)).toBeLessThanOrEqual(delta);
    expect(Math.abs(imageSize.verticalResolution - 95.98)).toBeLessThanOrEqual(delta);

    // We can base the size of the shape on the size of its image to avoid stretching the image.
    shape.width = imageSize.widthPoints * 2;
    shape.height = imageSize.heightPoints * 2;

    doc.save(base.artifactsDir + "Drawing.imageSize.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Drawing.imageSize.docx");
    shape = doc.getShape(0, true).asShape();

    TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.width).toEqual(600.0);
    expect(shape.height).toEqual(600.0);

    imageSize = shape.imageData.imageSize;

    expect(imageSize.heightPixels).toEqual(400);
    expect(imageSize.widthPixels).toEqual(400);
    expect(Math.abs(imageSize.horizontalResolution - 95.98)).toBeLessThanOrEqual(delta);
    expect(Math.abs(imageSize.verticalResolution - 95.98)).toBeLessThanOrEqual(delta);
  });

});
