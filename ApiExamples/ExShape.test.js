// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const path = require('path');
const TestUtil = require('./TestUtil');
const DocumentHelper = require('./DocumentHelper');
const finished = require('node:stream/promises');

/// <summary>
/// Examples using shapes in documents.
/// </summary>

describe("ExShape", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AltText', () => {
    //ExStart
    //ExFor:ShapeBase.alternativeText
    //ExFor:ShapeBase.name
    //ExSummary:Shows how to use a shape's alternative text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let shape = builder.insertShape(aw.Drawing.ShapeType.Cube, 150, 150);
    shape.name = "MyCube";

    shape.alternativeText = "Alt text for MyCube.";

    // We can access the alternative text of a shape by right-clicking it, and then via "Format AutoShape" -> "Alt Text".
    doc.save(base.artifactsDir + "Shape.AltText.docx");

    // Save the document to HTML, and then delete the linked image that belongs to our shape.
    // The browser that is reading our HTML will display the alt text in place of the missing image.
    doc.save(base.artifactsDir + "Shape.AltText.html");
    expect(fs.existsSync(base.artifactsDir + "Shape.AltText.001.png")).toEqual(true);
    fs.rmSync(base.artifactsDir + "Shape.AltText.001.png");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.AltText.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Cube, "MyCube", 150.0, 150.0, 0, 0, shape);
    expect(shape.alternativeText).toEqual("Alt text for MyCube.");
    expect(shape.font.name).toEqual("Times New Roman");

    doc = new aw.Document(base.artifactsDir + "Shape.AltText.html");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Image, "", 151.5, 151.5, 0, 0, shape);
    expect(shape.alternativeText).toEqual("Alt text for MyCube.");

    TestUtil.fileContainsString(
      "<img src=\"Shape.AltText.001.png\" width=\"202\" height=\"202\" alt=\"Alt text for MyCube.\" " +
      "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />",
      base.artifactsDir + "Shape.AltText.html");
  });


  test.each([false, true])('Font(%o)', (hideShape) => {
    //ExStart
    //ExFor:ShapeBase.font
    //ExFor:ShapeBase.parentParagraph
    //ExSummary:Shows how to insert a text box, and set the font of its contents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    let shape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 300, 50);
    builder.moveTo(shape.lastParagraph);
    builder.write("This text is inside the text box.");

    // Set the "Hidden" property of the shape's "Font" object to "true" to hide the text box from sight
    // and collapse the space that it would normally occupy.
    // Set the "Hidden" property of the shape's "Font" object to "false" to leave the text box visible.
    shape.font.hidden = hideShape;

    // If the shape is visible, we will modify its appearance via the font object.
    if (!hideShape)
    {
      shape.font.highlightColor = "#D3D3D3";
      shape.font.color = "#FF0000";
      shape.font.underline = aw.Underline.Dash;
    }

    // Move the builder out of the text box back into the main document.
    builder.moveTo(shape.parentParagraph);

    builder.writeln("\nThis text is outside the text box.");

    doc.save(base.artifactsDir + "Shape.font.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.font.docx");
    shape = doc.getShape(0, true);

    expect(shape.font.hidden).toEqual(hideShape);

    if (hideShape)
    {
      expect(shape.font.highlightColor).toEqual(base.emptyColor);
      expect(shape.font.color).toEqual(base.emptyColor);
      expect(shape.font.underline).toEqual(aw.Underline.None);
    }
    else
    {
      expect(shape.font.highlightColor).toEqual("#C0C0C0");
      expect(shape.font.color).toEqual("#FF0000");
      expect(shape.font.underline).toEqual(aw.Underline.Dash);
    }

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 300.0, 50.0, 0, 0, shape);
    expect(shape.getText().trim()).toEqual("This text is inside the text box.");
    expect(doc.getText().trim()).toEqual("Hello world!\rThis text is inside the text box.\r\rThis text is outside the text box.");
  });


  test('Rotate', () => {
    //ExStart
    //ExFor:ShapeBase.canHaveImage
    //ExFor:ShapeBase.rotation
    //ExSummary:Shows how to insert and rotate an image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a shape with an image.
    let shape = builder.insertImage(base.imageDir + "Logo.jpg");
    expect(shape.canHaveImage).toEqual(true);
    expect(shape.hasImage).toEqual(true);

    // Rotate the image 45 degrees clockwise.
    shape.rotation = 45;

    doc.save(base.artifactsDir + "Shape.Rotate.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.Rotate.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Image, "", 300.0, 300.0, 0, 0, shape);
    expect(shape.canHaveImage).toEqual(true);
    expect(shape.hasImage).toEqual(true);
    expect(shape.rotation).toEqual(45.0);
  });


  test('Coordinates', () => {
    //ExStart
    //ExFor:ShapeBase.distanceBottom
    //ExFor:ShapeBase.distanceLeft
    //ExFor:ShapeBase.distanceRight
    //ExFor:ShapeBase.distanceTop
    //ExSummary:Shows how to set the wrapping distance for a text that surrounds a shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a rectangle and, get the text to wrap tightly around its bounds.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 150, 150);
    shape.wrapType = aw.Drawing.WrapType.Tight;

    // Set the minimum distance between the shape and surrounding text to 40pt from all sides.
    shape.distanceTop = 40;
    shape.distanceBottom = 40;
    shape.distanceLeft = 40;
    shape.distanceRight = 40;

    // Move the shape closer to the center of the page, and then rotate the shape 60 degrees clockwise.
    shape.top = 75;
    shape.left = 150;
    shape.rotation = 60;

    // Add text that will wrap around the shape.
    builder.font.size = 24;
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    doc.save(base.artifactsDir + "Shape.Coordinates.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.Coordinates.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100002", 150.0, 150.0, 75.0, 150.0, shape);
    expect(shape.distanceBottom).toEqual(40.0);
    expect(shape.distanceLeft).toEqual(40.0);
    expect(shape.distanceRight).toEqual(40.0);
    expect(shape.distanceTop).toEqual(40.0);
    expect(shape.rotation).toEqual(60.0);
  });


  test('GroupShape', () => {
    //ExStart
    //ExFor:ShapeBase.bounds
    //ExFor:ShapeBase.coordOrigin
    //ExFor:ShapeBase.coordSize
    //ExSummary:Shows how to create and populate a group shape.
    let doc = new aw.Document();

    // Create a group shape. A group shape can display a collection of child shape nodes.
    // In Microsoft Word, clicking within the group shape's boundary or on one of the group shape's child shapes will
    // select all the other child shapes within this group and allow us to scale and move all the shapes at once.
    let group = new aw.Drawing.GroupShape(doc);

    expect(group.wrapType).toEqual(aw.Drawing.WrapType.None);

    // Create a 400pt x 400pt group shape and place it at the document's floating shape coordinate origin.
    group.bounds2 = new aw.JSRectangleF(0, 0, 400, 400);

    // Set the group's internal coordinate plane size to 500 x 500pt. 
    // The top left corner of the group will have an x and y coordinate of (0, 0),
    // and the bottom right corner will have an x and y coordinate of (500, 500).
    group.coordSize2 = new aw.JSSize(500, 500);

    // Set the coordinates of the top left corner of the group to (-250, -250). 
    // The group's center will now have an x and y coordinate value of (0, 0),
    // and the bottom right corner will be at (250, 250).
    group.coordOrigin2 = new aw.JSPoint(-250, -250);

    // Create a rectangle that will display the boundary of this group shape and add it to the group.
    let child1 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    child1.width = group.coordSize2.width;
    child1.height = group.coordSize2.height;
    child1.left = group.coordOrigin2.X;
    child1.top = group.coordOrigin2.Y;
    group.appendChild(child1);

    // Once a shape is a part of a group shape, we can access it as a child node and then modify it.
    group.getShape(0, true).stroke.dashStyle = aw.Drawing.DashStyle.Dash;

    // Create a small red star and insert it into the group.
    // Line up the shape with the group's coordinate origin, which we have moved to the center.
    let child2 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Star);
    child2.width = 20;
    child2.height = 20;
    child2.left = -10;
    child2.top = -10;
    child2.fillColor = "#FF0000";
    group.appendChild(child2);

    // Insert a rectangle, and then insert a slightly smaller rectangle in the same place with an image.
    // Newer shapes that we add to the group overlap older shapes. The light blue rectangle will partially overlap the red star,
    // and then the shape with the image will overlap the light blue rectangle, using it as a frame.
    // We cannot use the "ZOrder" properties of shapes to manipulate their arrangement within a group shape.
    let child3 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    child3.width = 250;
    child3.height = 250;
    child3.left = -250;
    child3.top = -250;
    child3.fillColor = "#ADD8E6";
    group.appendChild(child3);

    let child4 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
    child4.width = 200;
    child4.height = 200;
    child4.left = -225;
    child4.top = -225;
    group.appendChild(child4);

    group.getShape(3, true).imageData.setImage(base.imageDir + "Logo.jpg");

    // Insert a text box into the group shape. Set the "Left" property so that the text box's right edge
    // touches the right boundary of the group shape. Set the "Top" property so that the text box sits outside
    // the boundary of the group shape, with its top size lined up along the group shape's bottom margin.
    let child5 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    child5.width = 200;
    child5.height = 50;
    child5.left = group.coordSize2.width + group.coordOrigin2.X - 200;
    child5.top = group.coordSize2.height + group.coordOrigin2.Y;
    group.appendChild(child5);

    let builder = new aw.DocumentBuilder(doc);
    builder.insertNode(group);
    builder.moveTo(group.getShape(4, true).appendChild(new aw.Paragraph(doc)));
    builder.write("Hello world!");

    doc.save(base.artifactsDir + "Shape.groupShape.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.groupShape.docx");
    group = doc.getGroupShape(0, true);

    expect(group.bounds2).toEqual(new aw.JSRectangleF(0, 0, 400, 400));
    expect(group.coordSize2).toEqual(new aw.JSSize(500, 500));
    expect(group.coordOrigin2).toEqual(new aw.JSPoint(-250, -250));

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, '', 500.0, 500.0, -250.0, -250.0, group.getShape(0, true));
    TestUtil.verifyShape(aw.Drawing.ShapeType.Star, '', 20.0, 20.0, -10.0, -10.0, group.getShape(1, true));
    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, '', 250.0, 250.0, -250.0, -250.0, group.getShape(2, true));
    TestUtil.verifyShape(aw.Drawing.ShapeType.Image, '', 200.0, 200.0, -225.0, -225.0, group.getShape(3, true));
    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, '', 200.0, 50.0, 250.0, 50.0, group.getShape(4, true));
  });


  test('IsTopLevel', () => {
    //ExStart
    //ExFor:ShapeBase.isTopLevel
    //ExSummary:Shows how to tell whether a shape is a part of a group shape.
    let doc = new aw.Document();

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shape.width = 200;
    shape.height = 200;
    shape.wrapType = aw.Drawing.WrapType.None;

    // A shape by default is not part of any group shape, and therefore has the "IsTopLevel" property set to "true".
    expect(shape.isTopLevel).toEqual(true);

    let group = new aw.Drawing.GroupShape(doc);
    group.appendChild(shape);

    // Once we assimilate a shape into a group shape, the "IsTopLevel" property changes to "false".
    expect(shape.isTopLevel).toEqual(false);
    //ExEnd
  });


  test('LocalToParent', () => {
    //ExStart
    //ExFor:ShapeBase.coordOrigin
    //ExFor:ShapeBase.coordSize
    //ExFor:ShapeBase.localToParent(PointF)
    //ExSummary:Shows how to translate the x and y coordinate location on a shape's coordinate plane to a location on the parent shape's coordinate plane.
    let doc = new aw.Document();

    // Insert a group shape, and place it 100 points below and to the right of
    // the document's x and Y coordinate origin point.
    let group = new aw.Drawing.GroupShape(doc);
    group.bounds2 = new aw.JSRectangleF(100, 100, 500, 500);

    // Use the "LocalToParent" method to determine that (0, 0) on the group's internal x and y coordinates
    // lies on (100, 100) of its parent shape's coordinate system. The group shape's parent is the document itself.
    expect(group.localToParent(new aw.JSPointF(0, 0))).toEqual(new aw.JSPointF(100, 100));

    // By default, a shape's internal coordinate plane has the top left corner at (0, 0),
    // and the bottom right corner at (1000, 1000). Due to its size, our group shape covers an area of 500pt x 500pt
    // in the document's plane. This means that a movement of 1pt on the document's coordinate plane will translate
    // to a movement of 2pts on the group shape's coordinate plane.
    expect(group.localToParent(new aw.JSPointF(100, 100))).toEqual(new aw.JSPointF(200, 200));
    expect(group.localToParent(new aw.JSPointF(200, 200))).toEqual(new aw.JSPointF(200, 200));
    expect(group.localToParent(new aw.JSPointF(300, 300))).toEqual(new aw.JSPointF(250, 250));

    // Move the group shape's x and y axis origin from the top left corner to the center.
    // This will offset the group's internal coordinates relative to the document's coordinates even further.
    group.coordOrigin2 = new aw.JSPoint(-250, -250);

    expect(group.localToParent(new aw.JSPointF(300, 300))).toEqual(new aw.JSPointF(375, 375));

    // Changing the scale of the coordinate plane will also affect relative locations.
    group.coordSize2 = new aw.JSSize(500, 500);

    expect(group.localToParent(new aw.JSPointF(300, 300))).toEqual(new aw.JSPointF(650, 650));

    // If we wish to add a shape to this group while defining its location based on a location in the document,
    // we will need to first confirm a location in the group shape that will match the document's location.
    expect(group.localToParent(new aw.JSPointF(350, 350))).toEqual(new aw.JSPointF(700, 700));

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shape.width = 100;
    shape.height = 100;
    shape.left = 700;
    shape.top = 700;

    group.appendChild(shape);
    doc.firstSection.body.firstParagraph.appendChild(group);

    doc.save(base.artifactsDir + "Shape.localToParent.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.localToParent.docx");
    group = doc.getGroupShape(0, true);

    expect(group.bounds2).toEqual(new aw.JSRectangleF(100, 100, 500, 500));
    expect(group.coordSize2).toEqual(new aw.JSSize(500, 500));
    expect(group.coordOrigin2).toEqual(new aw.JSPoint(-250, -250));
  });


  test.each([false, true])('AnchorLocked(%o)', (anchorLocked) => {
    //ExStart
    //ExFor:ShapeBase.anchorLocked
    //ExSummary:Shows how to lock or unlock a shape's paragraph anchor.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    builder.write("Our shape will have an anchor attached to this paragraph.");
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 200, 160);
    shape.wrapType = aw.Drawing.WrapType.None;
    builder.insertBreak(aw.BreakType.ParagraphBreak);

    builder.writeln("Hello again!");

    // Set the "AnchorLocked" property to "true" to prevent the shape's anchor
    // from moving when moving the shape in Microsoft Word.
    // Set the "AnchorLocked" property to "false" to allow any movement of the shape
    // to also move its anchor to any other paragraph that the shape ends up close to.
    shape.anchorLocked = anchorLocked;

    // If the shape does not have a visible anchor symbol to its left,
    // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
    doc.save(base.artifactsDir + "Shape.anchorLocked.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.anchorLocked.docx");
    shape = doc.getShape(0, true);

    expect(shape.anchorLocked).toEqual(anchorLocked);
  });


  test('DeleteAllShapes', () => {
    //ExStart
    //ExFor:Shape
    //ExSummary:Shows how to delete all shapes from a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert two shapes along with a group shape with another shape inside it.
    builder.insertShape(aw.Drawing.ShapeType.Rectangle, 400, 200);
    builder.insertShape(aw.Drawing.ShapeType.Star, 300, 300);

    let group = new aw.Drawing.GroupShape(doc);
    group.bounds2 = new aw.JSRectangleF(100, 50, 200, 100);
    group.coordOrigin2 = new aw.JSPoint(-1000, -500);

    let subShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Cube);
    subShape.width = 500;
    subShape.height = 700;
    subShape.left = 0;
    subShape.top = 0;

    group.appendChild(subShape);
    builder.insertNode(group);

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(3);
    expect(doc.getChildNodes(aw.NodeType.GroupShape, true).count).toEqual(1);

    // Remove all Shape nodes from the document.
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
    shapes.clear();

    // All shapes are gone, but the group shape is still in the document.
    expect(doc.getChildNodes(aw.NodeType.GroupShape, true).count).toEqual(1);
    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);

    // Remove all group shapes separately.
    let groupShapes = doc.getChildNodes(aw.NodeType.GroupShape, true);
    groupShapes.clear();

    expect(doc.getChildNodes(aw.NodeType.GroupShape, true).count).toEqual(0);
    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);
    //ExEnd
  });


  test('IsInline', () => {
    //ExStart
    //ExFor:ShapeBase.isInline
    //ExSummary:Shows how to determine whether a shape is inline or floating.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two wrapping types that shapes may have.
    // 1 -  Inline:
    builder.write("Hello world! ");
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 100, 100);
    shape.fillColor = "#ADD8E6";
    builder.write(" Hello again.");

    // An inline shape sits inside a paragraph among other paragraph elements, such as runs of text.
    // In Microsoft Word, we may click and drag the shape to any paragraph as if it is a character.
    // If the shape is large, it will affect vertical paragraph spacing.
    // We cannot move this shape to a place with no paragraph.
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(shape.isInline).toEqual(true);

    // 2 -  Floating:
    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 200,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 200, 100, 100, aw.Drawing.WrapType.None);
    shape.fillColor = "#FFA500";

    // A floating shape belongs to the paragraph that we insert it into,
    // which we can determine by an anchor symbol that appears when we click the shape.
    // If the shape does not have a visible anchor symbol to its left,
    // we will need to enable visible anchors via "Options" -> "Display" -> "Object Anchors".
    // In Microsoft Word, we may left click and drag this shape freely to any location.
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(shape.isInline).toEqual(false);

    doc.save(base.artifactsDir + "Shape.isInline.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.isInline.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100002", 100, 100, 0, 0, shape);
    expect(shape.fillColor).toEqual("#ADD8E6");
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.Inline);
    expect(shape.isInline).toEqual(true);

    shape = doc.getShape(1, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100004", 100, 100, 200, 200, shape);
    expect(shape.fillColor).toEqual("#FFA500");
    expect(shape.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(shape.isInline).toEqual(false);
  });


  test('Bounds', () => {
    //ExStart
    //ExFor:ShapeBase.bounds
    //ExFor:ShapeBase.boundsInPoints
    //ExSummary:Shows how to verify shape containing block boundaries.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Line, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 50,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 50, 100, 100, aw.Drawing.WrapType.None);
    shape.strokeColor = "#FFA500";

    // Even though the line itself takes up little space on the document page,
    // it occupies a rectangular containing block, the size of which we can determine using the "Bounds" properties.
    expect(shape.bounds2).toEqual(new aw.JSRectangleF(50, 50, 100, 100));
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(50, 50, 100, 100));

    // Create a group shape, and then set the size of its containing block using the "Bounds" property.
    let group = new aw.Drawing.GroupShape(doc);
    group.bounds = new aw.JSRectangleF(0, 100, 250, 250);

    expect(group.boundsInPoints2).toEqual(new aw.JSRectangleF(0, 100, 250, 250));

    // Create a rectangle, verify the size of its bounding block, and then add it to the group shape.
    shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shape.width = 100;
    shape.height = 100;
    shape.left = 700;
    shape.top = 700;

    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(700, 700, 100, 100));

    group.appendChild(shape);

    // The group shape's coordinate plane has its origin on the top left-hand side corner of its containing block,
    // and the x and y coordinates of (1000, 1000) on the bottom right-hand side corner.
    // Our group shape is 250x250pt in size, so every 4pt on the group shape's coordinate plane
    // translates to 1pt in the document body's coordinate plane.
    // Every shape that we insert will also shrink in size by a factor of 4.
    // The change in the shape's "BoundsInPoints" property will reflect this.
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(175, 275, 25, 25));

    doc.firstSection.body.firstParagraph.appendChild(group);

    // Insert a shape and place it outside of the bounds of the group shape's containing block.
    shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shape.width = 100;
    shape.height = 100;
    shape.left = 1000;
    shape.top = 1000;

    group.appendChild(shape);

    // The group shape's footprint in the document body has increased, but the containing block remains the same.
    expect(group.boundsInPoints2).toEqual(new aw.JSRectangleF(0, 100, 250, 250));
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(250, 350, 25, 25));

    doc.save(base.artifactsDir + "Shape.bounds.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.bounds.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Line, "Line 100002", 100, 100, 50, 50, shape);
    expect(shape.strokeColor).toEqual("#FFA500");
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(50, 50, 100, 100));

    group = doc.getGroupShape(0, true);

    expect(group.bounds2).toEqual(new aw.JSRectangleF(0, 100, 250, 250));
    expect(group.boundsInPoints2).toEqual(new aw.JSRectangleF(0, 100, 250, 250));
    expect(group.coordSize2).toEqual(new aw.JSSize(1000, 1000));
    expect(group.coordOrigin2).toEqual(new aw.JSPoint(0, 0));

    shape = doc.getShape(1, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, '', 100, 100, 700, 700, shape);
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(175, 275, 25, 25));

    shape = doc.getShape(2, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, '', 100, 100, 1000, 1000, shape);
    expect(shape.boundsInPoints2).toEqual(new aw.JSRectangleF(250, 350, 25, 25));
  });


  test('FlipShapeOrientation', () => {
    //ExStart
    //ExFor:ShapeBase.flipOrientation
    //ExFor:FlipOrientation
    //ExSummary:Shows how to flip a shape on an axis.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert an image shape and leave its orientation in its default state.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 100,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 100, 100, 100, aw.Drawing.WrapType.None);
    shape.imageData.setImage(base.imageDir + "Logo.jpg");

    expect(shape.flipOrientation).toEqual(aw.Drawing.FlipOrientation.None);

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 250,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 100, 100, 100, aw.Drawing.WrapType.None);
    shape.imageData.setImage(base.imageDir + "Logo.jpg");

    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the second shape on the y-axis,
    // making it into a horizontal mirror image of the first shape.
    shape.flipOrientation = aw.Drawing.FlipOrientation.Horizontal;

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 100,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 250, 100, 100, aw.Drawing.WrapType.None);
    shape.imageData.setImage(base.imageDir + "Logo.jpg");

    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the third shape on the x-axis,
    // making it into a vertical mirror image of the first shape.
    shape.flipOrientation = aw.Drawing.FlipOrientation.Vertical;

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 250,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 250, 100, 100, aw.Drawing.WrapType.None);
    shape.imageData.setImage(base.imageDir + "Logo.jpg");

    // Set the "FlipOrientation" property to "FlipOrientation.Horizontal" to flip the fourth shape on both the x and y axes,
    // making it into a horizontal and vertical mirror image of the first shape.
    shape.flipOrientation = aw.Drawing.FlipOrientation.Both;

    doc.save(base.artifactsDir + "Shape.FlipShapeOrientation.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.FlipShapeOrientation.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100002", 100, 100, 100, 100, shape);
    expect(shape.flipOrientation).toEqual(aw.Drawing.FlipOrientation.None);

    shape = doc.getShape(1, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100004", 100, 100, 100, 250, shape);
    expect(shape.flipOrientation).toEqual(aw.Drawing.FlipOrientation.Horizontal);

    shape = doc.getShape(2, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100006", 100, 100, 250, 100, shape);
    expect(shape.flipOrientation).toEqual(aw.Drawing.FlipOrientation.Vertical);

    shape = doc.getShape(3, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "Rectangle 100008", 100, 100, 250, 250, shape);
    expect(shape.flipOrientation).toEqual(aw.Drawing.FlipOrientation.Both);
  });


  test.skip('Fill: WORDSNODEJS-86', () => {
    //ExStart
    //ExFor:ShapeBase.fill
    //ExFor:Shape.fillColor
    //ExFor:Shape.strokeColor
    //ExFor:Fill
    //ExFor:Fill.opacity
    //ExSummary:Shows how to fill a shape with a solid color.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Write some text, and then cover it with a floating shape.
    builder.font.size = 32;
    builder.writeln("Hello world!");

    let shape = builder.insertShape(aw.Drawing.ShapeType.CloudCallout, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 25,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 25, 250, 150, aw.Drawing.WrapType.None);

    // Use the "StrokeColor" property to set the color of the outline of the shape.
    shape.strokeColor = "#5F9EA0";

    // Use the "FillColor" property to set the color of the inside area of the shape.
    shape.fillColor = "#ADD8E6";

    // The "Opacity" property determines how transparent the color is on a 0-1 scale,
    // with 1 being fully opaque, and 0 being invisible.
    // The shape fill by default is fully opaque, so we cannot see the text that this shape is on top of.
    expect(shape.fill.opacity).toEqual(1.0);

    // Set the shape fill color's opacity to a lower value so that we can see the text underneath it.
    shape.fill.opacity = 0.3;

    doc.save(base.artifactsDir + "Shape.fill.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.fill.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.CloudCallout, "CloudCallout 100002", 250.0, 150.0, 25.0, 25.0, shape);
    //Color colorWithOpacity = Color.FromArgb(Convert.ToInt32(255 * shape.fill.opacity), "#ADD8E6".R, "#ADD8E6".G, "#ADD8E6".B);
    let colorWithOpacity = "#4CADD8E6";
    expect(shape.fillColor).toEqual(colorWithOpacity);
    expect(shape.strokeColor).toEqual("#5F9EA0");
    expect(shape.fill.opacity).toBeCloseTo(0.3, 2);
  });


  test('TextureFill', () => {
    //ExStart
    //ExFor:Fill.presetTexture
    //ExFor:Fill.textureAlignment
    //ExFor:TextureAlignment
    //ExSummary:Shows how to fill and tiling the texture inside the shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 80, 80);

    // Apply texture alignment to the shape fill.
    shape.fill.presetTextured(aw.Drawing.PresetTexture.Canvas);
    shape.fill.textureAlignment = aw.Drawing.TextureAlignment.TopRight;

    // Use the compliance option to define the shape using DML if you want to get "TextureAlignment"
    // property after the document saves.
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;

    doc.save(base.artifactsDir + "Shape.TextureFill.docx", saveOptions);

    doc = new aw.Document(base.artifactsDir + "Shape.TextureFill.docx");

    shape = doc.getShape(0, true);

    expect(shape.fill.textureAlignment).toEqual(aw.Drawing.TextureAlignment.TopRight);
    expect(shape.fill.presetTexture).toEqual(aw.Drawing.PresetTexture.Canvas);
    //ExEnd
  });


  test('GradientFill', () => {
    //ExStart
    //ExFor:Fill.oneColorGradient(Color, GradientStyle, GradientVariant, Double)
    //ExFor:Fill.oneColorGradient(GradientStyle, GradientVariant, Double)
    //ExFor:Fill.twoColorGradient(Color, Color, GradientStyle, GradientVariant)
    //ExFor:Fill.twoColorGradient(GradientStyle, GradientVariant)
    //ExFor:Fill.backColor
    //ExFor:Fill.gradientStyle
    //ExFor:Fill.gradientVariant
    //ExFor:Fill.gradientAngle
    //ExFor:GradientStyle
    //ExFor:GradientVariant
    //ExSummary:Shows how to fill a shape with a gradients.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 80, 80);
    // Apply One-color gradient fill to the shape with ForeColor of gradient fill.
    shape.fill.oneColorGradient("#FF0000", aw.Drawing.GradientStyle.Horizontal, aw.Drawing.GradientVariant.Variant2, 0.1);

    expect(shape.fill.foreColor).toEqual("#FF0000");
    expect(shape.fill.gradientStyle).toEqual(aw.Drawing.GradientStyle.Horizontal);
    expect(shape.fill.gradientVariant).toEqual(aw.Drawing.GradientVariant.Variant2);
    expect(shape.fill.gradientAngle).toEqual(270);

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 80, 80);
    // Apply Two-color gradient fill to the shape.
    shape.fill.twoColorGradient(aw.Drawing.GradientStyle.FromCorner, aw.Drawing.GradientVariant.Variant4);
    // Change BackColor of gradient fill.
    shape.fill.backColor = "#FFFF00";
    // Note that changes "GradientAngle" for "GradientStyle.FromCorner/GradientStyle.FromCenter"
    // gradient fill don't get any effect, it will work only for linear gradient.
    shape.fill.gradientAngle = 15;

    expect(shape.fill.backColor).toEqual("#FFFF00");
    expect(shape.fill.gradientStyle).toEqual(aw.Drawing.GradientStyle.FromCorner);
    expect(shape.fill.gradientVariant).toEqual(aw.Drawing.GradientVariant.Variant4);
    expect(shape.fill.gradientAngle).toEqual(0);

    // Use the compliance option to define the shape using DML if you want to get "GradientStyle",
    // "GradientVariant" and "GradientAngle" properties after the document saves.
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;

    doc.save(base.artifactsDir + "Shape.gradientFill.docx", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.gradientFill.docx");
    let firstShape = doc.getShape(0, true);

    expect(firstShape.fill.foreColor).toEqual("#FF0000");
    expect(firstShape.fill.gradientStyle).toEqual(aw.Drawing.GradientStyle.Horizontal);
    expect(firstShape.fill.gradientVariant).toEqual(aw.Drawing.GradientVariant.Variant2);
    expect(firstShape.fill.gradientAngle).toEqual(270);

    let secondShape = doc.getShape(1, true);

    expect(secondShape.fill.backColor).toEqual("#FFFF00");
    expect(secondShape.fill.gradientStyle).toEqual(aw.Drawing.GradientStyle.FromCorner);
    expect(secondShape.fill.gradientVariant).toEqual(aw.Drawing.GradientVariant.Variant4);
    expect(secondShape.fill.gradientAngle).toEqual(0);
  });


  test('GradientStops', () => {
    //ExStart
    //ExFor:Fill.gradientStops
    //ExFor:GradientStopCollection
    //ExFor:GradientStopCollection.insert(Int32, GradientStop)
    //ExFor:GradientStopCollection.add(GradientStop)
    //ExFor:GradientStopCollection.removeAt(Int32)
    //ExFor:GradientStopCollection.remove(GradientStop)
    //ExFor:GradientStopCollection.item(Int32)
    //ExFor:GradientStopCollection.count
    //ExFor:GradientStop
    //ExFor:GradientStop.#ctor(Color, Double)
    //ExFor:GradientStop.#ctor(Color, Double, Double)
    //ExFor:GradientStop.baseColor
    //ExFor:GradientStop.color
    //ExFor:GradientStop.position
    //ExFor:GradientStop.transparency
    //ExFor:GradientStop.remove
    //ExSummary:Shows how to add gradient stops to the gradient fill.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 80, 80);
    shape.fill.twoColorGradient("#008000", "#FF0000", aw.Drawing.GradientStyle.Horizontal, aw.Drawing.GradientVariant.Variant2);

    // Get gradient stops collection.
    let gradientStops = shape.fill.gradientStops;

    // Change first gradient stop.
    gradientStops.at(0).color = "#00FFFF";
    gradientStops.at(0).position = 0.1;
    gradientStops.at(0).transparency = 0.25;

    // Add new gradient stop to the end of collection.
    let gradientStop = new aw.Drawing.GradientStop("#A52A2A", 0.5);
    gradientStops.add(gradientStop);

    // Remove gradient stop at index 1.
    gradientStops.removeAt(1);
    // And insert new gradient stop at the same index 1.
    gradientStops.insert(1, new aw.Drawing.GradientStop("#D2691E", 0.75, 0.3));

    // Remove last gradient stop in the collection.
    gradientStop = gradientStops.at(2);
    gradientStops.remove(gradientStop);

    expect(gradientStops.count).toEqual(2);

    //expect(gradientStops.at(0).baseColor).toEqual("#FF00FFFF");
    expect(gradientStops.at(0).color).toEqual("#00FFFF");
    expect(gradientStops.at(0).position).toBeCloseTo(0.1, 2);
    expect(gradientStops.at(0).transparency).toBeCloseTo(0.25, 2);

    expect(gradientStops.at(1).color).toEqual("#D2691E");
    expect(gradientStops.at(1).position).toBeCloseTo(0.75, 2);
    expect(gradientStops.at(1).transparency).toBeCloseTo(0.3, 2);

    // Use the compliance option to define the shape using DML
    // if you want to get "GradientStops" property after the document saves.
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;

    doc.save(base.artifactsDir + "Shape.gradientStops.docx", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.gradientStops.docx");

    shape = doc.getShape(0, true);
    gradientStops = shape.fill.gradientStops;

    expect(gradientStops.count).toEqual(2);

    expect(gradientStops.at(0).color).toEqual("#00FFFF");
    expect(gradientStops.at(0).position).toBeCloseTo(0.1, 2);
    expect(gradientStops.at(0).transparency).toBeCloseTo(0.25, 2);

    expect(gradientStops.at(1).color).toEqual("#D2691E");
    expect(gradientStops.at(1).position).toBeCloseTo(0.75, 2);
    expect(gradientStops.at(1).transparency).toBeCloseTo(0.3, 2);
  });


  test('FillPattern', () => {
    //ExStart
    //ExFor:PatternType
    //ExFor:Fill.pattern
    //ExFor:Fill.patterned(PatternType)
    //ExFor:Fill.patterned(PatternType, Color, Color)
    //ExSummary:Shows how to set pattern for a shape.
    let doc = new aw.Document(base.myDir + "Shape stroke pattern border.docx");

    let shape = doc.getShape(0, true);
    let fill = shape.fill;

    console.log(`Pattern value is: ${fill.pattern}`);

    // There are several ways specified fill to a pattern.
    // 1 -  Apply pattern to the shape fill:
    fill.patterned(aw.Drawing.PatternType.DiagonalBrick);

    // 2 -  Apply pattern with foreground and background colors to the shape fill:
    fill.patterned(aw.Drawing.PatternType.DiagonalBrick, "#00FFFF", "#FFE4C4");

    doc.save(base.artifactsDir + "Shape.FillPattern.docx");
    //ExEnd
  });


  test('FillThemeColor', () => {
    //ExStart
    //ExFor:Fill.foreThemeColor
    //ExFor:Fill.backThemeColor
    //ExFor:Fill.backTintAndShade
    //ExSummary:Shows how to set theme color for foreground/background shape color.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.RoundRectangle, 80, 80);

    let fill = shape.fill;
    fill.foreThemeColor = aw.Themes.ThemeColor.Dark1;
    fill.backThemeColor = aw.Themes.ThemeColor.Background2;

    // Note: do not use "BackThemeColor" and "BackTintAndShade" for font fill.
    if (fill.backTintAndShade == 0)
      fill.backTintAndShade = 0.2;

    doc.save(base.artifactsDir + "Shape.FillThemeColor.docx");
    //ExEnd
  });


  test('FillTintAndShade', () => {
    //ExStart
    //ExFor:Fill.foreTintAndShade
    //ExSummary:Shows how to manage lightening and darkening foreground font color.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    let textFill = doc.firstSection.body.firstParagraph.runs.at(0).font.fill;
    textFill.foreThemeColor = aw.Themes.ThemeColor.Accent1;
    if (textFill.foreTintAndShade == 0)
      textFill.foreTintAndShade = 0.5;

    doc.save(base.artifactsDir + "Shape.FillTintAndShade.docx");
    //ExEnd
  });


  test('Title', () => {
    //ExStart
    //ExFor:ShapeBase.title
    //ExSummary:Shows how to set the title of a shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a shape, give it a title, and then add it to the document.
    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Cube);
    shape.width = 200;
    shape.height = 200;
    shape.title = "My cube";

    builder.insertNode(shape);

    // When we save a document with a shape that has a title,
    // Aspose.words will store that title in the shape's Alt Text.
    doc.save(base.artifactsDir + "Shape.title.docx");

    doc = new aw.Document(base.artifactsDir + "Shape.title.docx");
    shape = doc.getShape(0, true);

    expect(shape.title).toEqual('');
    expect(shape.alternativeText).toEqual("Title: My cube");
    //ExEnd

    TestUtil.verifyShape(aw.Drawing.ShapeType.Cube, '', 200.0, 200.0, 0.0, 0.0, shape);
  });


  test('ReplaceTextboxesWithImages', () => {
    //ExStart
    //ExFor:WrapSide
    //ExFor:ShapeBase.wrapSide
    //ExFor:NodeCollection
    //ExFor:CompositeNode.insertAfter``1(``0,Node)
    //ExFor:NodeCollection.toArray
    //ExSummary:Shows how to replace all textbox shapes with image shapes.
    let doc = new aw.Document(base.myDir + "Textboxes in drawing canvas.docx");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.filter(s => s.shapeType == aw.Drawing.ShapeType.TextBox).length).toEqual(3);
    expect(shapes.filter(s => s.shapeType == aw.Drawing.ShapeType.Image).length).toEqual(1);

    for (let shape of shapes)
    {
      if (shape.shapeType == aw.Drawing.ShapeType.TextBox)
      {
        let replacementShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
        replacementShape.imageData.setImage(base.imageDir + "Logo.jpg");
        replacementShape.left = shape.left;
        replacementShape.top = shape.top;
        replacementShape.width = shape.width;
        replacementShape.height = shape.height;
        replacementShape.relativeHorizontalPosition = shape.relativeHorizontalPosition;
        replacementShape.relativeVerticalPosition = shape.relativeVerticalPosition;
        replacementShape.horizontalAlignment = shape.horizontalAlignment;
        replacementShape.verticalAlignment = shape.verticalAlignment;
        replacementShape.wrapType = shape.wrapType;
        replacementShape.wrapSide = shape.wrapSide;

        shape.parentNode.insertAfter(replacementShape, shape);
        shape.remove();
      }
    }

    shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.filter(s => s.shapeType == aw.Drawing.ShapeType.TextBox).length).toEqual(0);
    expect(shapes.filter(s => s.shapeType == aw.Drawing.ShapeType.Image).length).toEqual(4);

    doc.save(base.artifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.ReplaceTextboxesWithImages.docx");
    let outShape = doc.getShape(0, true);

    expect(outShape.wrapSide).toEqual(aw.Drawing.WrapSide.Both);
  });


  test('CreateTextBox', () => {
    //ExStart
    //ExFor:Shape.#ctor(DocumentBase, ShapeType)
    //ExFor:Story.firstParagraph
    //ExFor:Shape.firstParagraph
    //ExFor:ShapeBase.wrapType
    //ExSummary:Shows how to create and format a text box.
    let doc = new aw.Document();

    // Create a floating text box.
    let textBox = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    textBox.wrapType = aw.Drawing.WrapType.None;
    textBox.height = 50;
    textBox.width = 200;

    // Set the horizontal, and vertical alignment of the text inside the shape.
    textBox.horizontalAlignment = aw.Drawing.HorizontalAlignment.Center;
    textBox.verticalAlignment = aw.Drawing.VerticalAlignment.Top;

    // Add a paragraph to the text box and add a run of text that the text box will display.
    textBox.appendChild(new aw.Paragraph(doc));
    let para = textBox.firstParagraph;
    para.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    let run = new aw.Run(doc);
    run.text = "Hello world!";
    para.appendChild(run);

    doc.firstSection.body.firstParagraph.appendChild(textBox);

    doc.save(base.artifactsDir + "Shape.CreateTextBox.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.CreateTextBox.docx");
    textBox = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, '', 200.0, 50.0, 0.0, 0.0, textBox);
    expect(textBox.wrapType).toEqual(aw.Drawing.WrapType.None);
    expect(textBox.horizontalAlignment).toEqual(aw.Drawing.HorizontalAlignment.Center);
    expect(textBox.verticalAlignment).toEqual(aw.Drawing.VerticalAlignment.Top);
    expect(textBox.getText().trim()).toEqual("Hello world!");
  });


  test('ZOrder', () => {
    //ExStart
    //ExFor:ShapeBase.zOrder
    //ExSummary:Shows how to manipulate the order of shapes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert three different colored rectangles that partially overlap each other.
    // When we insert a shape that overlaps another shape, Aspose.words places the newer shape on top of the old one.
    // The light green rectangle will overlap the light blue rectangle and partially obscure it,
    // and the light blue rectangle will obscure the orange rectangle.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 100,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 100, 200, 200, aw.Drawing.WrapType.None);
    shape.fillColor = "#FFA500";

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 150,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 150, 200, 200, aw.Drawing.WrapType.None);
    shape.fillColor = "#ADD8E6";

    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 200,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 200, 200, 200, aw.Drawing.WrapType.None);
    shape.fillColor = "#90EE90";

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    // The "ZOrder" property of a shape determines its stacking priority among other overlapping shapes.
    // If two overlapping shapes have different "ZOrder" values,
    // Microsoft Word will place the shape with a higher value over the shape with the lower value. 
    // Set the "ZOrder" values of our shapes to place the first orange rectangle over the second light blue one
    // and the second light blue rectangle over the third light green rectangle.
    // This will reverse their original stacking order.
    shapes.at(0).zOrder = 3;
    shapes.at(1).zOrder = 2;
    shapes.at(2).zOrder = 1;

    doc.save(base.artifactsDir + "Shape.zOrder.docx");
    //ExEnd
  });


  test('GetActiveXControlProperties', () => {
    //ExStart
    //ExFor:OleControl
    //ExFor:OleControl.isForms2OleControl
    //ExFor:OleControl.name
    //ExFor:OleFormat.oleControl
    //ExFor:Forms2OleControl
    //ExFor:Forms2OleControl.caption
    //ExFor:Forms2OleControl.value
    //ExFor:Forms2OleControl.enabled
    //ExFor:Forms2OleControl.type
    //ExFor:Forms2OleControl.childNodes
    //ExFor:Forms2OleControl.groupName
    //ExSummary:Shows how to verify the properties of an ActiveX control.
    let doc = new aw.Document(base.myDir + "ActiveX controls.docx");

    let shape = doc.getShape(0, true);
    let oleControl = shape.oleFormat.oleControl;

    expect(oleControl.name).toEqual("CheckBox1");

    if (oleControl.isForms2OleControl)
    {
      console.log(oleControl);
      let checkBox = oleControl.asForms2OleControl();
      expect(checkBox.caption).toEqual("First");
      expect(checkBox.value).toEqual("0");
      expect(checkBox.enabled).toEqual(true);
      expect(checkBox.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.CheckBox);
      expect(checkBox.childNodes).toEqual(null);
      expect(checkBox.groupName).toEqual('');

      // Note, that you can't set GroupName for a Frame.
      checkBox.groupName = "Aspose group name";
    }
    //ExEnd

    doc.save(base.artifactsDir + "Shape.GetActiveXControlProperties.docx");
    doc = new aw.Document(base.artifactsDir + "Shape.GetActiveXControlProperties.docx");

    shape = doc.getShape(0, true);
    let forms2OleControl = shape.oleFormat.oleControl.asForms2OleControl();

    expect(forms2OleControl.groupName).toEqual("Aspose group name");
  });


  test('GetOleObjectRawData', () => {
    //ExStart
    //ExFor:OleFormat.getRawData
    //ExSummary:Shows how to access the raw data of an embedded OLE object.
    let doc = new aw.Document(base.myDir + "OLE objects.docx");

    for (let shape of doc.getChildNodes(aw.NodeType.Shape, true))
    {
      let oleFormat = shape.oleFormat;
      if (oleFormat != null)
      {
        console.log(`This is ${(oleFormat.isLink ? "a linked" : "an embedded")} object`);
        let oleRawData = oleFormat.getRawData();

        expect(oleRawData.length).toEqual(24576);
      }
    }
    //ExEnd
  });


  test('LinkedChartSourceFullName', () => {
    //ExStart
    //ExFor:Chart.sourceFullName
    //ExSummary:Shows how to get/set the full name of the external xls/xlsx document if the chart is linked.
    let doc = new aw.Document(base.myDir + "Shape with linked chart.docx");

    let shape = doc.getShape(0, true);

    let sourceFullName = shape.chart.sourceFullName;
    expect(sourceFullName.includes("Examples\\Data\\Spreadsheet.xlsx")).toEqual(true);
    //ExEnd
  });


  test.skip('OleControl: WORDSNODEJS-87', async () => {
    //ExStart
    //ExFor:OleFormat
    //ExFor:OleFormat.autoUpdate
    //ExFor:OleFormat.isLocked
    //ExFor:OleFormat.progId
    //ExFor:OleFormat.save(Stream)
    //ExFor:OleFormat.save(String)
    //ExFor:OleFormat.suggestedExtension
    //ExSummary:Shows how to extract embedded OLE objects into files.
    let doc = new aw.Document(base.myDir + "OLE spreadsheet.docm");
    let shape = doc.getShape(0, true);

    // The OLE object in the first shape is a Microsoft Excel spreadsheet.
    let oleFormat = shape.oleFormat;

    expect(oleFormat.progId).toEqual("Excel.Sheet.12");

    // Our object is neither auto updating nor locked from updates.
    expect(oleFormat.autoUpdate).toEqual(false);
    expect(oleFormat.isLocked).toEqual(false);

    // If we plan on saving the OLE object to a file in the local file system,
    // we can use the "SuggestedExtension" property to determine which file extension to apply to the file.
    expect(oleFormat.suggestedExtension).toEqual(".xlsx");

    // Below are two ways of saving an OLE object to a file in the local file system.
    // 1 -  Save it via a stream:
    let writeStream = fs.createWriteStream(base.artifactsDir + "OLE spreadsheet extracted via stream" + oleFormat.suggestedExtension);
    oleFormat.save(writeStream);
    await finished(writeStream);

    // 2 -  Save it directly to a filename:
    oleFormat.save(base.artifactsDir + "OLE spreadsheet saved directly" + oleFormat.suggestedExtension);
    //ExEnd

    expect(fs.statSync(base.artifactsDir + "OLE spreadsheet extracted via stream.xlsx").size).toBeLessThan(8400);
    expect(fs.statSync(base.artifactsDir + "OLE spreadsheet saved directly.xlsx").size).toBeLessThan(8400);
  });


  test.skip('OleLinks: Aspose.Words.Drawing.OleFormat.GetOleEntry(System.String) is skipped', () => {
    //ExStart
    //ExFor:OleFormat.iconCaption
    //ExFor:OleFormat.getOleEntry(String)
    //ExFor:OleFormat.isLink
    //ExFor:OleFormat.oleIcon
    //ExFor:OleFormat.sourceFullName
    //ExFor:OleFormat.sourceItem
    //ExSummary:Shows how to insert linked and unlinked OLE objects.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Embed a Microsoft Visio drawing into the document as an OLE object.
    builder.insertOleObject(base.imageDir + "Microsoft Visio drawing.vsd", "Package", false, false, null);

    // Insert a link to the file in the local file system and display it as an icon.
    builder.insertOleObject(base.imageDir + "Microsoft Visio drawing.vsd", "Package", true, true, null);

    // Inserting OLE objects creates shapes that store these objects.
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(2);
    expect(shapes.filter(s => s.shapeType == aw.Drawing.ShapeType.OleObject).length).toEqual(2);

    // If a shape contains an OLE object, it will have a valid "OleFormat" property,
    // which we can use to verify some aspects of the shape.
    let oleFormat = shapes.at(0).oleFormat;

    expect(oleFormat.isLink).toEqual(false);
    expect(oleFormat.oleIcon).toEqual(false);

    oleFormat = shapes.at(1).oleFormat;

    expect(oleFormat.isLink).toEqual(true);
    expect(oleFormat.oleIcon).toEqual(true);

    expect(oleFormat.sourceFullName.endsWith(`Images${path.sep}Microsoft Visio drawing.vsd`)).toEqual(true);
    expect(oleFormat.sourceItem).toEqual("");

    expect(oleFormat.iconCaption).toEqual("Microsoft Visio drawing.vsd");

    doc.save(base.artifactsDir + "Shape.OleLinks.docx");

    // If the object contains OLE data, we can access it using a stream.
    /*using (Stream stream = oleFormat.getOleEntry("\x0001CompObj"))
    {
      expect(stream.length).toEqual(76);
    }*/
    //ExEnd
  });


  test('OleControlCollection', () => {
    //ExStart
    //ExFor:OleFormat.clsid
    //ExFor:Forms2OleControlCollection
    //ExFor:Forms2OleControlCollection.count
    //ExFor:Forms2OleControlCollection.item(Int32)
    //ExSummary:Shows how to access an OLE control embedded in a document and its child controls.
    let doc = new aw.Document(base.myDir + "OLE ActiveX controls.docm");

    // Shapes store and display OLE objects in the document's body.
    let shape = doc.getShape(0, true);

    expect(shape.oleFormat.clsid.toString().toLowerCase()).toEqual("6e182020-f460-11ce-9bcd-00aa00608e01");

    let oleControl = shape.oleFormat.oleControl.asForms2OleControl();

    // Some OLE controls may contain child controls, such as the one in this document with three options buttons.
    let oleControlCollection = oleControl.childNodes;

    expect(oleControlCollection.count).toEqual(3);

    expect(oleControlCollection.at(0).caption).toEqual("C#");
    expect(oleControlCollection.at(0).value).toEqual("1");

    expect(oleControlCollection.at(1).caption).toEqual("Visual Basic");
    expect(oleControlCollection.at(1).value).toEqual("0");

    expect(oleControlCollection.at(2).caption).toEqual("Delphi");
    expect(oleControlCollection.at(2).value).toEqual("0");
    //ExEnd
  });


  test.skip('SuggestedFileName: WORDSNODEJS-87', async () => {
    //ExStart
    //ExFor:OleFormat.suggestedFileName
    //ExSummary:Shows how to get an OLE object's suggested file name.
    let doc = new aw.Document(base.myDir + "OLE shape.rtf");

    let oleShape = doc.firstSection.body.getShape(0, true);

    // OLE objects can provide a suggested filename and extension,
    // which we can use when saving the object's contents into a file in the local file system.
    let suggestedFileName = oleShape.oleFormat.suggestedFileName;

    expect(suggestedFileName).toEqual("CSV.csv");

    let writeStream = fs.createWriteStream(base.artifactsDir + suggestedFileName);
    oleShape.oleFormat.save(writeStream);
    await finished(writeStream);
    //ExEnd
  });


  test('ObjectDidNotHaveSuggestedFileName', () => {
    let doc = new aw.Document(base.myDir + "ActiveX controls.docx");

    let shape = doc.getShape(0, true);
    expect(shape.oleFormat.suggestedFileName).toEqual('');
  });


  test('RenderOfficeMath', async () => {
    //ExStart
    //ExFor:ImageSaveOptions.scale
    //ExFor:OfficeMath.getMathRenderer
    //ExFor:NodeRendererBase.save(String, ImageSaveOptions)
    //ExSummary:Shows how to render an Office Math object into an image file in the local file system.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);

    // Create an "ImageSaveOptions" object to pass to the node renderer's "Save" method to modify
    // how it renders the OfficeMath node into an image.
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);

    // Set the "Scale" property to 5 to render the object to five times its original size.
    saveOptions.scale = 5;

    officeMath.getMathRenderer().save(base.artifactsDir + "Shape.RenderOfficeMath.png", saveOptions);
    //ExEnd

    await TestUtil.verifyImage(813, 87, base.artifactsDir + "Shape.RenderOfficeMath.png");
  });


  test('OfficeMathDisplayException', () => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);
    officeMath.displayType = aw.Math.OfficeMathDisplayType.Display;

    const expectedError = "Inline justification cannot be set to the Office Math displayed on its own line. Please, use OfficeMath.DisplayType property to change OfficeMathDisplayType.";
    expect(() => officeMath.justification = aw.Math.OfficeMathJustification.Inline).toThrow(expectedError);
  });


  test('OfficeMathDefaultValue', () => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(6, true);

    expect(officeMath.displayType).toEqual(aw.Math.OfficeMathDisplayType.Inline);
    expect(officeMath.justification).toEqual(aw.Math.OfficeMathJustification.Inline);
  });


  test('OfficeMath', () => {
    //ExStart
    //ExFor:OfficeMath
    //ExFor:OfficeMath.displayType
    //ExFor:OfficeMath.justification
    //ExFor:OfficeMath.nodeType
    //ExFor:OfficeMath.parentParagraph
    //ExFor:OfficeMathDisplayType
    //ExFor:OfficeMathJustification
    //ExSummary:Shows how to set office math display formatting.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);

    // OfficeMath nodes that are children of other OfficeMath nodes are always inline.
    // The node we are working with is the base node to change its location and display type.
    expect(officeMath.mathObjectType).toEqual(aw.Math.MathObjectType.OMathPara);
    expect(officeMath.nodeType).toEqual(aw.NodeType.OfficeMath);
    expect(officeMath.parentParagraph.referenceEquals(officeMath.parentNode)).toBeTruthy();

    // Change the location and display type of the OfficeMath node.
    officeMath.displayType = aw.Math.OfficeMathDisplayType.Display;
    officeMath.justification = aw.Math.OfficeMathJustification.Left;

    doc.save(base.artifactsDir + "Shape.officeMath.docx");
    //ExEnd

    expect(DocumentHelper.compareDocs(base.artifactsDir + "Shape.officeMath.docx", base.goldsDir + "Shape.officeMath Gold.docx")).toEqual(true);
  });


  test('CannotBeSetDisplayWithInlineJustification', () => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);
    officeMath.displayType = aw.Math.OfficeMathDisplayType.Display;

    const expectedError = "Inline justification cannot be set to the Office Math displayed on its own line. Please, use OfficeMath.DisplayType property to change OfficeMathDisplayType.";
    expect(() => officeMath.justification = aw.Math.OfficeMathJustification.Inline).toThrow(expectedError);
  });


  test('CannotBeSetInlineDisplayWithJustification', () => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);
    officeMath.displayType = aw.Math.OfficeMathDisplayType.Inline;

    const expectedError = "Justification cannot be set to the Office Math displayed inline with text. Please, use OfficeMath.DisplayType property to change OfficeMathDisplayType.";
    expect(() => officeMath.justification = aw.Math.OfficeMathJustification.Center).toThrow(expectedError);
  });


  test('OfficeMathDisplayNestedObjects', () => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);

    expect(officeMath.displayType).toEqual(aw.Math.OfficeMathDisplayType.Display);
    expect(officeMath.justification).toEqual(aw.Math.OfficeMathJustification.Center);
  });


  test.each([[0, aw.Math.MathObjectType.OMathPara],
    [1, aw.Math.MathObjectType.OMath],
    [2, aw.Math.MathObjectType.Supercript],
    [3, aw.Math.MathObjectType.Argument],
    [4, aw.Math.MathObjectType.SuperscriptPart]])('WorkWithMathObjectType(%o, %o)', (index, objectType) => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(index, true);
    expect(officeMath.mathObjectType).toEqual(objectType);
  });


  test.each([true, false])('AspectRatio(%o)', (lockAspectRatio) => {
    //ExStart
    //ExFor:ShapeBase.aspectRatioLocked
    //ExSummary:Shows how to lock/unlock a shape's aspect ratio.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a shape. If we open this document in Microsoft Word, we can left click the shape to reveal
    // eight sizing handles around its perimeter, which we can click and drag to change its size.
    let shape = builder.insertImage(base.imageDir + "Logo.jpg");

    // Set the "AspectRatioLocked" property to "true" to preserve the shape's aspect ratio
    // when using any of the four diagonal sizing handles, which change both the image's height and width.
    // Using any orthogonal sizing handles that either change the height or width will still change the aspect ratio.
    // Set the "AspectRatioLocked" property to "false" to allow us to
    // freely change the image's aspect ratio with all sizing handles.
    shape.aspectRatioLocked = lockAspectRatio;

    doc.save(base.artifactsDir + "Shape.aspectRatio.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.aspectRatio.docx");
    shape = doc.getShape(0, true);

    expect(shape.aspectRatioLocked).toEqual(lockAspectRatio);
  });


  test('MarkupLanguageByDefault', () => {
    //ExStart
    //ExFor:ShapeBase.markupLanguage
    //ExFor:ShapeBase.sizeInPoints
    //ExSummary:Shows how to verify a shape's size and markup language.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imageDir + "Transparent background logo.png");

    expect(shape.markupLanguage).toEqual(aw.Drawing.ShapeMarkupLanguage.Dml);
    expect(shape.sizeInPoints2).toEqual(new aw.JSSizeF(300, 300));
    //ExEnd
  });


  test.each([[aw.Settings.MsWordVersion.Word2000, aw.Drawing.ShapeMarkupLanguage.Vml],
    [aw.Settings.MsWordVersion.Word2002, aw.Drawing.ShapeMarkupLanguage.Vml],
    [aw.Settings.MsWordVersion.Word2003, aw.Drawing.ShapeMarkupLanguage.Vml],
    [aw.Settings.MsWordVersion.Word2007, aw.Drawing.ShapeMarkupLanguage.Vml],
    [aw.Settings.MsWordVersion.Word2010, aw.Drawing.ShapeMarkupLanguage.Dml],
    [aw.Settings.MsWordVersion.Word2013, aw.Drawing.ShapeMarkupLanguage.Dml],
    [aw.Settings.MsWordVersion.Word2016, aw.Drawing.ShapeMarkupLanguage.Dml]])
    ('MarkupLanguageForDifferentMsWordVersions(%o, %o)', (msWordVersion, shapeMarkupLanguage) => {
    let doc = new aw.Document();
    doc.compatibilityOptions.optimizeFor(msWordVersion);

    let builder = new aw.DocumentBuilder(doc);
    builder.insertImage(base.imageDir + "Transparent background logo.png");

    for (let shape of doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()))
    {
      expect(shape.markupLanguage).toEqual(shapeMarkupLanguage);
    }
  });


  test('Stroke', () => {
    //ExStart
    //ExFor:Stroke
    //ExFor:Stroke.on
    //ExFor:Stroke.weight
    //ExFor:Stroke.joinStyle
    //ExFor:Stroke.lineStyle
    //ExFor:Stroke.fill
    //ExFor:ShapeLineStyle
    //ExSummary:Shows how change stroke properties.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 100,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 100, 200, 200, aw.Drawing.WrapType.None);

    // Basic shapes, such as the rectangle, have two visible parts.
    // 1 -  The fill, which applies to the area within the outline of the shape:
    shape.fill.foreColor = "#FFFFFF";

    // 2 -  The stroke, which marks the outline of the shape:
    // Modify various properties of this shape's stroke.
    let stroke = shape.stroke;
    stroke.on = true;
    stroke.weight = 5;
    stroke.color = "#FF0000";
    stroke.dashStyle = aw.Drawing.DashStyle.ShortDashDotDot;
    stroke.joinStyle = aw.Drawing.JoinStyle.Miter;
    stroke.endCap = aw.Drawing.EndCap.Square;
    stroke.lineStyle = aw.Drawing.ShapeLineStyle.Triple;
    stroke.fill.twoColorGradient("#FF0000", "#0000FF", aw.Drawing.GradientStyle.Vertical, aw.Drawing.GradientVariant.Variant1);

    doc.save(base.artifactsDir + "Shape.stroke.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.stroke.docx");
    shape = doc.getShape(0, true);
    stroke = shape.stroke;

    expect(stroke.on).toEqual(true);
    expect(stroke.weight).toEqual(5);
    expect(stroke.color).toEqual("#FF0000");
    expect(stroke.dashStyle).toEqual(aw.Drawing.DashStyle.ShortDashDotDot);
    expect(stroke.joinStyle).toEqual(aw.Drawing.JoinStyle.Miter);
    expect(stroke.endCap).toEqual(aw.Drawing.EndCap.Square);
    expect(stroke.lineStyle).toEqual(aw.Drawing.ShapeLineStyle.Triple);
  });


  test('InsertOleObjectAsHtmlFile', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

    doc.save(base.artifactsDir + "Shape.InsertOleObjectAsHtmlFile.docx");
  });


  test.skip('InsertOlePackage: WORDSNODEJS-93', () => {
    //ExStart
    //ExFor:OlePackage
    //ExFor:OleFormat.olePackage
    //ExFor:OlePackage.fileName
    //ExFor:OlePackage.displayName
    //ExSummary:Shows how insert an OLE object into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // OLE objects allow us to open other files in the local file system using another installed application
    // in our operating system by double-clicking on the shape that contains the OLE object in the document body.
    // In this case, our external file will be a ZIP archive.
    let readStream = fs.createReadStream(base.databaseDir + "cat001.zip");
    let shape = builder.insertOleObject(readStream, "Package", true, null);
    shape.oleFormat.olePackage.fileName = "Package file name.zip";
    shape.oleFormat.olePackage.displayName = "Package display name.zip";

    doc.save(base.artifactsDir + "Shape.InsertOlePackage.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.InsertOlePackage.docx");
    let getShape = doc.getShape(0, true);

    expect(getShape.oleFormat.olePackage.fileName).toEqual("Package file name.zip");
    expect(getShape.oleFormat.olePackage.displayName).toEqual("Package display name.zip");
  });


  test('GetAccessToOlePackage', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let oleObject = builder.insertOleObject(base.myDir + "Spreadsheet.xlsx", false, false, null);
    let oleObjectAsOlePackage = builder.insertOleObject(base.myDir + "Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

    expect(oleObject.oleFormat.olePackage).toBeNull();
    expect(oleObjectAsOlePackage.oleFormat.olePackage).not.toBeNull();
  });


  test('Resize', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 200, 300);
    shape.height = 300;
    shape.width = 500;
    shape.rotation = 30;

    doc.save(base.artifactsDir + "Shape.Resize.docx");
  });


  test('Calendar', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startTable();
    builder.rowFormat.height = 100;
    builder.rowFormat.heightRule = aw.HeightRule.Exactly;

    for (let i = 0; i < 31; i++)
    {
      if (i != 0 && i % 7 == 0)
        builder.endRow();
      builder.insertCell();
      builder.write("Cell contents");
    }

    builder.endTable();

    let runs = doc.getChildNodes(aw.NodeType.Run, true).toArray().map(node => node.asRun());
    let num = 1;

    for (let run of runs)
    {
      let watermark = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextPlainText);
      watermark.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
      watermark.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
      watermark.width = 30;
      watermark.height = 30;
      watermark.horizontalAlignment = aw.Drawing.HorizontalAlignment.Center;
      watermark.verticalAlignment = aw.Drawing.VerticalAlignment.Center;
      watermark.rotation = -40;

      watermark.fill.foreColor = "#DCDCDC";
      watermark.strokeColor = "#DCDCDC";

      watermark.textPath.text = `${num}`;
      watermark.textPath.fontFamily = "Arial";

      watermark.name = `Watermark_${num++}`;

      watermark.behindText = true;

      builder.moveTo(run);
      builder.insertNode(watermark);
    }

    doc.save(base.artifactsDir + "Shape.Calendar.docx");

    doc = new aw.Document(base.artifactsDir + "Shape.Calendar.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(31);

    for (let shape of shapes)
      TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, `Watermark_${shapes.indexOf(shape) + 1}`, 30.0, 30.0, 0.0, 0.0, shape);
  });


  test.each([false, true])('IsLayoutInCell(%o)', (isLayoutInCell) => {
    //ExStart
    //ExFor:ShapeBase.isLayoutInCell
    //ExSummary:Shows how to determine how to display a shape in a table cell.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.insertCell();
    builder.endTable();

    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
    tableStyle.bottomPadding = 20;
    tableStyle.leftPadding = 10;
    tableStyle.rightPadding = 10;
    tableStyle.topPadding = 20;
    tableStyle.borders.color = "#000000";
    tableStyle.borders.lineStyle = aw.LineStyle.Single;

    table.style = tableStyle;

    builder.moveTo(table.firstRow.firstCell.firstParagraph);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, aw.Drawing.RelativeHorizontalPosition.LeftMargin, 50,
      aw.Drawing.RelativeVerticalPosition.TopMargin, 100, 100, 100, aw.Drawing.WrapType.None);

    // Set the "IsLayoutInCell" property to "true" to display the shape as an inline element inside the cell's paragraph.
    // The coordinate origin that will determine the shape's location will be the top left corner of the shape's cell.
    // If we re-size the cell, the shape will move to maintain the same position starting from the cell's top left.
    // Set the "IsLayoutInCell" property to "false" to display the shape as an independent floating shape.
    // The coordinate origin that will determine the shape's location will be the top left corner of the page,
    // and the shape will not respond to any re-sizing of its cell.
    shape.isLayoutInCell = isLayoutInCell;

    // We can only apply the "IsLayoutInCell" property to floating shapes.
    shape.wrapType = aw.Drawing.WrapType.None;

    doc.save(base.artifactsDir + "Shape.LayoutInTableCell.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.LayoutInTableCell.docx");
    table = doc.firstSection.body.tables.at(0);
    shape = table.firstRow.firstCell.getShape(0, true);

    expect(shape.isLayoutInCell).toEqual(isLayoutInCell);
  });


  test('ShapeInsertion', () => {
    //ExStart
    //ExFor:DocumentBuilder.insertShape(ShapeType, RelativeHorizontalPosition, double, RelativeVerticalPosition, double, double, double, WrapType)
    //ExFor:DocumentBuilder.insertShape(ShapeType, double, double)
    //ExFor:OoxmlCompliance
    //ExFor:OoxmlSaveOptions.compliance
    //ExSummary:Shows how to insert DML shapes into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two wrapping types that shapes may have.
    // 1 -  Floating:
    builder.insertShape(aw.Drawing.ShapeType.TopCornersRounded, aw.Drawing.RelativeHorizontalPosition.Page, 100,
        aw.Drawing.RelativeVerticalPosition.Page, 100, 50, 50, aw.Drawing.WrapType.None);

    // 2 -  Inline:
    builder.insertShape(aw.Drawing.ShapeType.DiagonalCornersRounded, 50, 50);

    // If you need to create "non-primitive" shapes, such as SingleCornerSnipped, TopCornersSnipped, DiagonalCornersSnipped,
    // TopCornersOneRoundedOneSnipped, SingleCornerRounded, TopCornersRounded, or DiagonalCornersRounded,
    // then save the document with "Strict" or "Transitional" compliance, which allows saving shape as DML.
    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;

    doc.save(base.artifactsDir + "Shape.ShapeInsertion.docx", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.ShapeInsertion.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    TestUtil.verifyShape(aw.Drawing.ShapeType.TopCornersRounded, "TopCornersRounded 100002", 50.0, 50.0, 100.0, 100.0, shapes.at(0));
    TestUtil.verifyShape(aw.Drawing.ShapeType.DiagonalCornersRounded, "DiagonalCornersRounded 100004", 50.0, 50.0, 0.0, 0.0, shapes.at(1));
  });


  /*//Commented
  //ExStart
  //ExFor:Shape.Accept(DocumentVisitor)
  //ExFor:Shape.Chart
  //ExFor:Shape.ExtrusionEnabled
  //ExFor:Shape.Filled
  //ExFor:Shape.HasChart
  //ExFor:Shape.OleFormat
  //ExFor:Shape.ShadowEnabled
  //ExFor:Shape.StoryType
  //ExFor:Shape.StrokeColor
  //ExFor:Shape.Stroked
  //ExFor:Shape.StrokeWeight
  //ExSummary:Shows how to iterate over all the shapes in a document.
  test('VisitShapes', () => {
    let doc = new aw.Document(base.myDir + "Revision shape.docx");
    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(2);

    let visitor = new ShapeAppearancePrinter();
    doc.accept(visitor);

    console.log(visitor.getText());
  });


    /// <summary>
    /// Logs appearance-related information about visited shapes.
    /// </summary>
  private class ShapeAppearancePrinter : DocumentVisitor
  {
    public ShapeAppearancePrinter()
    {
      mShapesVisited = 0;
      mTextIndentLevel = 0;
      mStringBuilder = new StringBuilder();
    }

      /// <summary>
      /// Appends a line to the StringBuilder with one prepended tab character for each indent level.
      /// </summary>
    private void AppendLine(string text)
    {
      for (let i = 0; i < mTextIndentLevel; i++) mStringBuilder.append('\t');

      mStringBuilder.AppendLine(text);
    }

      /// <summary>
      /// Return all the text that the StringBuilder has accumulated.
      /// </summary>
    public string GetText()
    {
      return `Shapes visited: ${mShapesVisited}\n${mStringBuilder}`;
    }

      /// <summary>
      /// Called when this visitor visits the start of a Shape node.
      /// </summary>
    public override VisitorAction VisitShapeStart(Shape shape)
    {
      AppendLine(`Shape found: ${shape.shapeType}`);

      mTextIndentLevel++;

      if (shape.hasChart)
        AppendLine(`Has chart: ${shape.chart.title.text}`);

      AppendLine(`Extrusion enabled: ${shape.extrusionEnabled}`);
      AppendLine(`Shadow enabled: ${shape.shadowEnabled}`);
      AppendLine(`StoryType: ${shape.storyType}`);

      if (shape.stroked)
      {
        expect(shape.strokeColor).toEqual(shape.stroke.color);
        AppendLine(`Stroke colors: ${shape.stroke.color}, ${shape.stroke.color2}`);
        AppendLine(`Stroke weight: ${shape.strokeWeight}`);
      }

      if (shape.filled)
        AppendLine(`Filled: ${shape.fillColor}`);

      if (shape.oleFormat != null)
        AppendLine(`Ole found of type: ${shape.oleFormat.progId}`);

      if (shape.signatureLine != null)
        AppendLine(`Found signature line for: ${shape.signatureLine.signer}, ${shape.signatureLine.signerTitle}`);

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when this visitor visits the end of a Shape node.
      /// </summary>
    public override VisitorAction VisitShapeEnd(Shape shape)
    {
      mTextIndentLevel--;
      mShapesVisited++;
      AppendLine(`End of ${shape.shapeType}`);

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when this visitor visits the start of a GroupShape node.
      /// </summary>
    public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
    {
      AppendLine(`Shape group found: ${groupShape.shapeType}`);
      mTextIndentLevel++;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when this visitor visits the end of a GroupShape node.
      /// </summary>
    public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
    {
      mTextIndentLevel--;
      AppendLine(`End of ${groupShape.shapeType}`);

      return aw.VisitorAction.Continue;
    }

    private int mShapesVisited;
    private int mTextIndentLevel;
    private readonly StringBuilder mStringBuilder;
  }
  //ExEnd
  //EndCommented*/

  test('SignatureLine', () => {
    //ExStart
    //ExFor:Shape.signatureLine
    //ExFor:ShapeBase.isSignatureLine
    //ExFor:SignatureLine
    //ExFor:SignatureLine.allowComments
    //ExFor:SignatureLine.defaultInstructions
    //ExFor:SignatureLine.email
    //ExFor:SignatureLine.instructions
    //ExFor:SignatureLine.showDate
    //ExFor:SignatureLine.signer
    //ExFor:SignatureLine.signerTitle
    //ExSummary:Shows how to create a line for a signature and insert it into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let options = new aw.SignatureLineOptions();
    options.allowComments = true,
    options.defaultInstructions = true,
    options.email = "john.doe@management.com",
    options.instructions = "Please sign here",
    options.showDate = true,
    options.signer = "John Doe",
    options.signerTitle = "Senior Manager"

    // Insert a shape that will contain a signature line, whose appearance we will
    // customize using the "SignatureLineOptions" object we have created above.
    // If we insert a shape whose coordinates originate at the bottom right hand corner of the page,
    // we will need to supply negative x and y coordinates to bring the shape into view.
    let shape = builder.insertSignatureLine(options, aw.Drawing.RelativeHorizontalPosition.RightMargin, -170.0,
        aw.Drawing.RelativeVerticalPosition.BottomMargin, -60.0, aw.Drawing.WrapType.None);

    expect(shape.isSignatureLine).toEqual(true);

    // Verify the properties of our signature line via its Shape object.
    let signatureLine = shape.signatureLine;

    expect(signatureLine.email).toEqual("john.doe@management.com");
    expect(signatureLine.signer).toEqual("John Doe");
    expect(signatureLine.signerTitle).toEqual("Senior Manager");
    expect(signatureLine.instructions).toEqual("Please sign here");
    expect(signatureLine.showDate).toEqual(true);
    expect(signatureLine.allowComments).toEqual(true);
    expect(signatureLine.defaultInstructions).toEqual(true);

    doc.save(base.artifactsDir + "Shape.signatureLine.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.signatureLine.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.Image, '', 192.75, 96.75, -60.0, -170.0, shape);
    expect(shape.isSignatureLine).toEqual(true);

    signatureLine = shape.signatureLine;

    expect(signatureLine.email).toEqual("john.doe@management.com");
    expect(signatureLine.signer).toEqual("John Doe");
    expect(signatureLine.signerTitle).toEqual("Senior Manager");
    expect(signatureLine.instructions).toEqual("Please sign here");
    expect(signatureLine.showDate).toEqual(true);
    expect(signatureLine.allowComments).toEqual(true);
    expect(signatureLine.defaultInstructions).toEqual(true);
    expect(signatureLine.isSigned).toEqual(false);
    expect(signatureLine.isValid).toEqual(false);
  });


  test.each([aw.Drawing.LayoutFlow.Vertical,
    aw.Drawing.LayoutFlow.Horizontal,
    aw.Drawing.LayoutFlow.HorizontalIdeographic,
    aw.Drawing.LayoutFlow.BottomToTop,
    aw.Drawing.LayoutFlow.TopToBottom,
    aw.Drawing.LayoutFlow.TopToBottomIdeographic])('TextBoxLayoutFlow(%o)', (layoutFlow) => {
    //ExStart
    //ExFor:Shape.textBox
    //ExFor:Shape.lastParagraph
    //ExFor:TextBox
    //ExFor:TextBox.layoutFlow
    //ExSummary:Shows how to set the orientation of text inside a text box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBoxShape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 150, 100);
    let textBox = textBoxShape.textBox;

    // Move the document builder to inside the TextBox and add text.
    builder.moveTo(textBoxShape.lastParagraph);
    builder.writeln("Hello world!");
    builder.write("Hello again!");

    // Set the "LayoutFlow" property to set an orientation for the text contents of this text box.
    textBox.layoutFlow = layoutFlow;

    doc.save(base.artifactsDir + "Shape.TextBoxLayoutFlow.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.TextBoxLayoutFlow.docx");
    textBoxShape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 150.0, 100.0, 0.0, 0.0, textBoxShape);

    let expectedLayoutFlow;

    switch (layoutFlow)
    {
      case aw.Drawing.LayoutFlow.BottomToTop:
      case aw.Drawing.LayoutFlow.Horizontal:
      case aw.Drawing.LayoutFlow.TopToBottomIdeographic:
      case aw.Drawing.LayoutFlow.Vertical:
        expectedLayoutFlow = layoutFlow;
        break;
      case aw.Drawing.LayoutFlow.TopToBottom:
        expectedLayoutFlow = aw.Drawing.LayoutFlow.Vertical;
        break;
      default:
        expectedLayoutFlow = aw.Drawing.LayoutFlow.Horizontal;
        break;
    }

    TestUtil.verifyTextBox(expectedLayoutFlow, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, textBoxShape.textBox);
    expect(textBoxShape.getText().trim()).toEqual("Hello world!\rHello again!");
  });


  test('TextBoxFitShapeToText', () => {
    //ExStart
    //ExFor:TextBox
    //ExFor:TextBox.fitShapeToText
    //ExSummary:Shows how to get a text box to resize itself to fit its contents tightly.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBoxShape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 150, 100);
    let textBox = textBoxShape.textBox;

    // Apply these values to both these members to get the parent shape to fit
    // tightly around the text contents, ignoring the dimensions we have set.
    textBox.fitShapeToText = true;
    textBox.textBoxWrapMode = aw.Drawing.TextBoxWrapMode.None;

    builder.moveTo(textBoxShape.lastParagraph);
    builder.write("Text fit tightly inside textbox.");

    doc.save(base.artifactsDir + "Shape.TextBoxFitShapeToText.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.TextBoxFitShapeToText.docx");
    textBoxShape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 150.0, 100.0, 0.0, 0.0, textBoxShape);
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, true, aw.Drawing.TextBoxWrapMode.None, 3.6, 3.6, 7.2, 7.2, textBoxShape.textBox);
    expect(textBoxShape.getText().trim()).toEqual("Text fit tightly inside textbox.");
  });


  test('TextBoxMargins', () => {
    //ExStart
    //ExFor:TextBox
    //ExFor:TextBox.internalMarginBottom
    //ExFor:TextBox.internalMarginLeft
    //ExFor:TextBox.internalMarginRight
    //ExFor:TextBox.internalMarginTop
    //ExSummary:Shows how to set internal margins for a text box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert another textbox with specific margins.
    let textBoxShape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    let textBox = textBoxShape.textBox;
    textBox.internalMarginTop = 15;
    textBox.internalMarginBottom = 15;
    textBox.internalMarginLeft = 15;
    textBox.internalMarginRight = 15;

    builder.moveTo(textBoxShape.lastParagraph);
    builder.write("Text placed according to textbox margins.");

    doc.save(base.artifactsDir + "Shape.TextBoxMargins.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.TextBoxMargins.docx");
    textBoxShape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 100.0, 100.0, 0.0, 0.0, textBoxShape);
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 15.0, 15.0, 15.0, 15.0, textBoxShape.textBox);
    expect(textBoxShape.getText().trim()).toEqual("Text placed according to textbox margins.");
  });


  test.each([aw.Drawing.TextBoxWrapMode.None,
    aw.Drawing.TextBoxWrapMode.Square])('TextBoxContentsWrapMode(%o)', (textBoxWrapMode) => {
    //ExStart
    //ExFor:TextBox.textBoxWrapMode
    //ExFor:TextBoxWrapMode
    //ExSummary:Shows how to set a wrapping mode for the contents of a text box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBoxShape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 300, 300);
    let textBox = textBoxShape.textBox;

    // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.None" to increase the text box's width
    // to accommodate text, should it be large enough.
    // Set the "TextBoxWrapMode" property to "TextBoxWrapMode.Square" to
    // wrap all text inside the text box, preserving its dimensions.
    textBox.textBoxWrapMode = textBoxWrapMode;

    builder.moveTo(textBoxShape.lastParagraph);
    builder.font.size = 32;
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    doc.save(base.artifactsDir + "Shape.TextBoxContentsWrapMode.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.TextBoxContentsWrapMode.docx");
    textBoxShape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 300.0, 300.0, 0.0, 0.0, textBoxShape);
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, textBoxWrapMode, 3.6, 3.6, 7.2, 7.2, textBoxShape.textBox);
    expect(textBoxShape.getText().trim()).toEqual("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
  });


  test('TextBoxShapeType', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set compatibility options to correctly using of VerticalAnchor property.
    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2016);

    let textBoxShape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    // Not all formats are compatible with this one.
    // For most of the incompatible formats, AW generated warnings on save, so use doc.warningCallback to check it.
    textBoxShape.textBox.verticalAnchor = aw.Drawing.TextBoxAnchor.Bottom;

    builder.moveTo(textBoxShape.lastParagraph);
    builder.write("Text placed bottom");

    doc.save(base.artifactsDir + "Shape.TextBoxShapeType.docx");
  });


  test('CreateLinkBetweenTextBoxes', () => {
    //ExStart
    //ExFor:TextBox.isValidLinkTarget(TextBox)
    //ExFor:TextBox.next
    //ExFor:TextBox.previous
    //ExFor:TextBox.breakForwardLink
    //ExSummary:Shows how to link text boxes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBoxShape1 = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    let textBox1 = textBoxShape1.textBox;
    builder.writeln();

    let textBoxShape2 = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    let textBox2 = textBoxShape2.textBox;
    builder.writeln();

    let textBoxShape3 = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    let textBox3 = textBoxShape3.textBox;
    builder.writeln();

    let textBoxShape4 = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 100);
    let textBox4 = textBoxShape4.textBox;

    // Create links between some of the text boxes.
    if (textBox1.isValidLinkTarget(textBox2))
      textBox1.next = textBox2;

    if (textBox2.isValidLinkTarget(textBox3))
      textBox2.next = textBox3;

    // Only an empty text box may have a link.
    expect(textBox3.isValidLinkTarget(textBox4)).toEqual(true);

    builder.moveTo(textBoxShape4.lastParagraph);
    builder.write("Hello world!");

    expect(textBox3.isValidLinkTarget(textBox4)).toEqual(false);

    if (textBox1.next != null && textBox1.previous == null)
      console.log("This TextBox is the head of the sequence");

    if (textBox2.next != null && textBox2.previous != null)
      console.log("This TextBox is the middle of the sequence");

    if (textBox3.next == null && textBox3.previous != null)
    {
      console.log("This TextBox is the tail of the sequence");

      // Break the forward link between textBox2 and textBox3, and then verify that they are no longer linked.
      textBox3.previous.breakForwardLink();
      expect(textBox2.next == null).toEqual(true);
      expect(textBox3.previous == null).toEqual(true);
    }

    doc.save(base.artifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.CreateLinkBetweenTextBoxes.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 100.0, 100.0, 0.0, 0.0, shapes.at(0));
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, shapes.at(0).textBox);
    expect(shapes.at(0).getText().trim()).toEqual('');

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100004", 100.0, 100.0, 0.0, 0.0, shapes.at(1));
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, shapes.at(1).textBox);
    expect(shapes.at(1).getText().trim()).toEqual('');

    TestUtil.verifyShape(aw.Drawing.ShapeType.Rectangle, "TextBox 100006", 100.0, 100.0, 0.0, 0.0, shapes.at(2));
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, shapes.at(2).textBox);
    expect(shapes.at(2).getText().trim()).toEqual('');

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100008", 100.0, 100.0, 0.0, 0.0, shapes.at(3));
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, shapes.at(3).textBox);
    expect(shapes.at(3).getText().trim()).toEqual("Hello world!");
  });


  test.each([aw.Drawing.TextBoxAnchor.Top,
    aw.Drawing.TextBoxAnchor.Middle,
    aw.Drawing.TextBoxAnchor.Bottom])('VerticalAnchor(%o)', (verticalAnchor) => {
    //ExStart
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.optimizeFor(MsWordVersion)
    //ExFor:TextBoxAnchor
    //ExFor:TextBox.verticalAnchor
    //ExSummary:Shows how to vertically align the text contents of a text box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 200, 200);

    // Set the "VerticalAnchor" property to "TextBoxAnchor.Top" to
    // align the text in this text box with the top side of the shape.
    // Set the "VerticalAnchor" property to "TextBoxAnchor.Middle" to
    // align the text in this text box to the center of the shape.
    // Set the "VerticalAnchor" property to "TextBoxAnchor.Bottom" to
    // align the text in this text box to the bottom of the shape.
    shape.textBox.verticalAnchor = verticalAnchor;

    builder.moveTo(shape.firstParagraph);
    builder.write("Hello world!");

    // The vertical aligning of text inside text boxes is available from Microsoft Word 2007 onwards.
    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2007);
    doc.save(base.artifactsDir + "Shape.verticalAnchor.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.verticalAnchor.docx");
    shape = doc.getShape(0, true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextBox, "TextBox 100002", 200.0, 200.0, 0.0, 0.0, shape);
    TestUtil.verifyTextBox(aw.Drawing.LayoutFlow.Horizontal, false, aw.Drawing.TextBoxWrapMode.Square, 3.6, 3.6, 7.2, 7.2, shape.textBox);
    expect(shape.textBox.verticalAnchor).toEqual(verticalAnchor);
    expect(shape.getText().trim()).toEqual("Hello world!");
  });


    //ExStart
    //ExFor:Shape.TextPath
    //ExFor:ShapeBase.IsWordArt
    //ExFor:TextPath
    //ExFor:TextPath.Bold
    //ExFor:TextPath.FitPath
    //ExFor:TextPath.FitShape
    //ExFor:TextPath.FontFamily
    //ExFor:TextPath.Italic
    //ExFor:TextPath.Kerning
    //ExFor:TextPath.On
    //ExFor:TextPath.ReverseRows
    //ExFor:TextPath.RotateLetters
    //ExFor:TextPath.SameLetterHeights
    //ExFor:TextPath.Shadow
    //ExFor:TextPath.SmallCaps
    //ExFor:TextPath.Spacing
    //ExFor:TextPath.StrikeThrough
    //ExFor:TextPath.Text
    //ExFor:TextPath.TextPathAlignment
    //ExFor:TextPath.Trim
    //ExFor:TextPath.Underline
    //ExFor:TextPath.XScale
    //ExFor:TextPath.Size
    //ExFor:TextPathAlignment
    //ExSummary:Shows how to work with WordArt.
  test('InsertTextPaths', () => {
    let doc = new aw.Document();

    // Insert a WordArt object to display text in a shape that we can re-size and move by using the mouse in Microsoft Word.
    // Provide a "ShapeType" as an argument to set a shape for the WordArt.
    let shape = appendWordArt(doc, "Hello World! This text is bold, and italic.",
      "Arial", 480, 24, "#FFFFFF", "#000000", aw.Drawing.ShapeType.TextPlainText);

    // Apply the "Bold" and "Italic" formatting settings to the text using the respective properties.
    shape.textPath.bold = true;
    shape.textPath.italic = true;

    // Below are various other text formatting-related properties.
    expect(shape.textPath.underline).toEqual(false);
    expect(shape.textPath.shadow).toEqual(false);
    expect(shape.textPath.strikeThrough).toEqual(false);
    expect(shape.textPath.reverseRows).toEqual(false);
    expect(shape.textPath.xscale).toEqual(false);
    expect(shape.textPath.trim).toEqual(false);
    expect(shape.textPath.smallCaps).toEqual(false);

    expect(shape.textPath.size).toEqual(36.0);
    expect(shape.textPath.text).toEqual("Hello World! This text is bold, and italic.");
    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.TextPlainText);

    // Use the "On" property to show/hide the text.
    shape = appendWordArt(doc, "On set to \"true\"", "Calibri", 150, 24, "#FFFF00", "#FF0000", aw.Drawing.ShapeType.TextPlainText);
    shape.textPath.on = true;

    shape = appendWordArt(doc, "On set to \"false\"", "Calibri", 150, 24, "#FFFF00", "#800080", aw.Drawing.ShapeType.TextPlainText);
    shape.textPath.on = false;

    // Use the "Kerning" property to enable/disable kerning spacing between certain characters.
    shape = appendWordArt(doc, "Kerning: VAV", "Times New Roman", 90, 24, "#FFA500", "#FF0000", aw.Drawing.ShapeType.TextPlainText);
    shape.textPath.kerning = true;

    shape = appendWordArt(doc, "No kerning: VAV", "Times New Roman", 100, 24, "#FFA500", "#FF0000", aw.Drawing.ShapeType.TextPlainText);
    shape.textPath.kerning = false;

    // Use the "Spacing" property to set the custom spacing between characters on a scale from 0.0 (none) to 1.0 (default).
    shape = appendWordArt(doc, "Spacing set to 0.1", "Calibri", 120, 24, "#8A2BE2", "#0000FF", aw.Drawing.ShapeType.TextCascadeDown);
    shape.textPath.spacing = 0.1;

    // Set the "RotateLetters" property to "true" to rotate each character 90 degrees counterclockwise.
    shape = appendWordArt(doc, "RotateLetters", "Calibri", 200, 36, "#ADFF2F", "#008000", aw.Drawing.ShapeType.TextWave);
    shape.textPath.rotateLetters = true;

    // Set the "SameLetterHeights" property to "true" to get the x-height of each character to equal the cap height.
    shape = appendWordArt(doc, "Same character height for lower and UPPER case", "Calibri", 300, 24, "#00BFFF", "#1E90FF", aw.Drawing.ShapeType.TextSlantUp);
    shape.textPath.sameLetterHeights = true;

    // By default, the text's size will always scale to fit the containing shape's size, overriding the text size setting.
    shape = appendWordArt(doc, "FitShape on", "Calibri", 160, 24, "#ADD8E6", "#0000FF", aw.Drawing.ShapeType.TextPlainText);
    expect(shape.textPath.fitShape).toEqual(true);
    shape.textPath.size = 24.0;

    // If we set the "FitShape: property to "false", the text will keep the size
    // which the "Size" property specifies regardless of the size of the shape.
    // Use the "TextPathAlignment" property also to align the text to a side of the shape.
    shape = appendWordArt(doc, "FitShape off", "Calibri", 160, 24, "#ADD8E6", "#0000FF", aw.Drawing.ShapeType.TextPlainText);
    shape.textPath.fitShape = false;
    shape.textPath.size = 24.0;
    shape.textPath.textPathAlignment = aw.Drawing.TextPathAlignment.Right;

    doc.save(base.artifactsDir + "Shape.InsertTextPaths.docx");
    testInsertTextPaths(base.artifactsDir + "Shape.InsertTextPaths.docx"); //ExSkip
  });


  /// <summary>
  /// Insert a new paragraph with a WordArt shape inside it.
  /// </summary>
  function appendWordArt(doc, text, textFontFamily, shapeWidth, shapeHeight, wordArtFill, line, wordArtShapeType) {
    // Create an inline Shape, which will serve as a container for our WordArt.
    // The shape can only be a valid WordArt shape if we assign a WordArt-designated ShapeType to it.
    // These types will have "WordArt object" in the description,
    // and their enumerator constant names will all start with "Text".
    let shape = new aw.Drawing.Shape(doc, wordArtShapeType);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.width = shapeWidth;
    shape.height = shapeHeight;
    shape.fillColor = wordArtFill;
    shape.strokeColor = line;

    shape.textPath.text = text;
    shape.textPath.fontFamily = textFontFamily;

    let para = doc.firstSection.body.appendChild(new aw.Paragraph(doc)).asParagraph();
    para.appendChild(shape);
    return shape;
  }
    //ExEnd

  function testInsertTextPaths(filename) {
    let doc = new aw.Document(filename);
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 480, 24, 0.0, 0.0, shapes.at(0));
    expect(shapes.at(0).textPath.bold).toEqual(true);
    expect(shapes.at(0).textPath.italic).toEqual(true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 150, 24, 0.0, 0.0, shapes.at(1));
    expect(shapes.at(1).textPath.on).toEqual(true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 150, 24, 0.0, 0.0, shapes.at(2));
    expect(shapes.at(2).textPath.on).toEqual(false);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 90, 24, 0.0, 0.0, shapes.at(3));
    expect(shapes.at(3).textPath.kerning).toEqual(true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 100, 24, 0.0, 0.0, shapes.at(4));
    expect(shapes.at(4).textPath.kerning).toEqual(false);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextCascadeDown, '', 120, 24, 0.0, 0.0, shapes.at(5));
    expect(shapes.at(5).textPath.spacing).toBeCloseTo(0.1, 2);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextWave, '', 200, 36, 0.0, 0.0, shapes.at(6));
    expect(shapes.at(6).textPath.rotateLetters).toEqual(true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextSlantUp, '', 300, 24, 0.0, 0.0, shapes.at(7));
    expect(shapes.at(7).textPath.sameLetterHeights).toEqual(true);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 160, 24, 0.0, 0.0, shapes.at(8));
    expect(shapes.at(8).textPath.fitShape).toEqual(true);
    expect(shapes.at(8).textPath.size).toEqual(24.0);

    TestUtil.verifyShape(aw.Drawing.ShapeType.TextPlainText, '', 160, 24, 0.0, 0.0, shapes.at(9));
    expect(shapes.at(9).textPath.fitShape).toEqual(false);
    expect(shapes.at(9).textPath.size).toEqual(24.0);
    expect(shapes.at(9).textPath.textPathAlignment).toEqual(aw.Drawing.TextPathAlignment.Right);
  }

  test('ShapeRevision', () => {
    //ExStart
    //ExFor:ShapeBase.isDeleteRevision
    //ExFor:ShapeBase.isInsertRevision
    //ExSummary:Shows how to work with revision shapes.
    let doc = new aw.Document();

    expect(doc.trackRevisions).toEqual(false);

    // Insert an inline shape without tracking revisions, which will make this shape not a revision of any kind.
    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Cube);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.width = 100.0;
    shape.height = 100.0;
    doc.firstSection.body.firstParagraph.appendChild(shape);

    // Start tracking revisions and then insert another shape, which will be a revision.
    doc.startTrackRevisions("John Doe");

    shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Sun);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.width = 100.0;
    shape.height = 100.0;
    doc.firstSection.body.firstParagraph.appendChild(shape);

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(2);

    shapes.at(0).remove();

    // Since we removed that shape while we were tracking changes,
    // the shape persists in the document and counts as a delete revision.
    // Accepting this revision will remove the shape permanently, and rejecting it will keep it in the document.
    expect(shapes.at(0).shapeType).toEqual(aw.Drawing.ShapeType.Cube);
    expect(shapes.at(0).isDeleteRevision).toEqual(true);

    // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
    // Accepting this revision will assimilate this shape into the document as a non-revision,
    // and rejecting the revision will remove this shape permanently.
    expect(shapes.at(1).shapeType).toEqual(aw.Drawing.ShapeType.Sun);
    expect(shapes.at(1).isInsertRevision).toEqual(true);
    //ExEnd
  });


  test('MoveRevisions', () => {
    //ExStart
    //ExFor:ShapeBase.isMoveFromRevision
    //ExFor:ShapeBase.isMoveToRevision
    //ExSummary:Shows how to identify move revision shapes.
    // A move revision is when we move an element in the document body by cut-and-pasting it in Microsoft Word while
    // tracking changes. If we involve an inline shape in such a text movement, that shape will also be a revision.
    // Copying-and-pasting or moving floating shapes do not create move revisions.
    let doc = new aw.Document(base.myDir + "Revision shape.docx");

    // Move revisions consist of pairs of "Move from", and "Move to" revisions. We moved in this document in one shape,
    // but until we accept or reject the move revision, there will be two instances of that shape.
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(2);

    // This is the "Move to" revision, which is the shape at its arrival destination.
    // If we accept the revision, this "Move to" revision shape will disappear,
    // and the "Move from" revision shape will remain.
    expect(shapes.at(0).isMoveFromRevision).toEqual(false);
    expect(shapes.at(0).isMoveToRevision).toEqual(true);

    // This is the "Move from" revision, which is the shape at its original location.
    // If we accept the revision, this "Move from" revision shape will disappear,
    // and the "Move to" revision shape will remain.
    expect(shapes.at(1).isMoveFromRevision).toEqual(true);
    expect(shapes.at(1).isMoveToRevision).toEqual(false);
    //ExEnd
  });


  test('AdjustWithEffects', () => {
    //ExStart
    //ExFor:ShapeBase.adjustWithEffects(RectangleF)
    //ExFor:ShapeBase.boundsWithEffects
    //ExSummary:Shows how to check how a shape's bounds are affected by shape effects.
    let doc = new aw.Document(base.myDir + "Shape shadow effect.docx");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    expect(shapes.length).toEqual(2);

    // The two shapes are identical in terms of dimensions and shape type.
    expect(shapes.at(1).width).toEqual(shapes.at(0).width);
    expect(shapes.at(1).height).toEqual(shapes.at(0).height);
    expect(shapes.at(1).shapeType).toEqual(shapes.at(0).shapeType);

    // The first shape has no effects, and the second one has a shadow and thick outline.
    // These effects make the size of the second shape's silhouette bigger than that of the first.
    // Even though the rectangle's size shows up when we click on these shapes in Microsoft Word,
    // the visible outer bounds of the second shape are affected by the shadow and outline and thus are bigger.
    // We can use the "AdjustWithEffects" method to see the true size of the shape.
    expect(shapes.at(0).strokeWeight).toEqual(0.0);
    expect(shapes.at(1).strokeWeight).toEqual(20.0);
    expect(shapes.at(0).shadowEnabled).toEqual(false);
    expect(shapes.at(1).shadowEnabled).toEqual(true);

    let shape = shapes.at(0);

    // Create a RectangleF object, representing a rectangle,
    // which we could potentially use as the coordinates and bounds for a shape.
    let rectangleF = new aw.JSRectangleF(200, 200, 1000, 1000);

    // Run this method to get the size of the rectangle adjusted for all our shape effects.
    let rectangleFOut = shape.adjustWithEffects(rectangleF);

    // Since the shape has no border-changing effects, its boundary dimensions are unaffected.
    expect(rectangleFOut.X).toEqual(200);
    expect(rectangleFOut.Y).toEqual(200);
    expect(rectangleFOut.width).toEqual(1000);
    expect(rectangleFOut.height).toEqual(1000);

    // Verify the final extent of the first shape, in points.
    expect(shape.boundsWithEffects2.X).toEqual(0);
    expect(shape.boundsWithEffects2.Y).toEqual(0);
    expect(shape.boundsWithEffects2.width).toEqual(147);
    expect(shape.boundsWithEffects2.height).toEqual(147);

    shape = shapes.at(1);
    rectangleF = new aw.JSRectangleF(200, 200, 1000, 1000);
    rectangleFOut = shape.adjustWithEffects(rectangleF);

    // The shape effects have moved the apparent top left corner of the shape slightly.
    expect(rectangleFOut.X).toEqual(171.5);
    expect(rectangleFOut.Y).toEqual(167);

    // The effects have also affected the visible dimensions of the shape.
    expect(rectangleFOut.width).toEqual(1045);
    expect(rectangleFOut.height).toEqual(1133.5);

    // The effects have also affected the visible bounds of the shape.
    expect(shape.boundsWithEffects2.X).toEqual(-28.5);
    expect(shape.boundsWithEffects2.Y).toEqual(-33);
    expect(shape.boundsWithEffects2.width).toEqual(192);
    expect(shape.boundsWithEffects2.height).toEqual(280.5);
    //ExEnd
  });


  test('RenderAllShapes', () => {
    //ExStart
    //ExFor:ShapeBase.getShapeRenderer
    //ExFor:NodeRendererBase.save(Stream, ImageSaveOptions)
    //ExSummary:Shows how to use a shape renderer to export shapes to files in the local file system.
    let doc = new aw.Document(base.myDir + "Various shapes.docx");
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape())

    expect(shapes.length).toEqual(7);

    // There are 7 shapes in the document, including one group shape with 2 child shapes.
    // We will render every shape to an image file in the local file system
    // while ignoring the group shapes since they have no appearance.
    // This will produce 6 image files.
    for (let shape of doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()))
    {
      let renderer = shape.getShapeRenderer();
      let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
      renderer.save(base.artifactsDir + `Shape.RenderAllShapes.${shape.name}.png`, options);
    }
    //ExEnd
  });


  test('DocumentHasSmartArtObject', () => {
    //ExStart
    //ExFor:Shape.hasSmartArt
    //ExSummary:Shows how to count the number of shapes in a document with SmartArt objects.
    let doc = new aw.Document(base.myDir + "SmartArt.docx");

    let numberOfSmartArtShapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape()).filter(shape => shape.hasSmartArt).length;

    expect(numberOfSmartArtShapes).toEqual(2);
    //ExEnd

  });


  test('OfficeMathRenderer', () => {
    //ExStart
    //ExFor:NodeRendererBase
    //ExFor:NodeRendererBase.boundsInPoints
    //ExFor:NodeRendererBase.getBoundsInPixels(Single, Single)
    //ExFor:NodeRendererBase.getBoundsInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.getOpaqueBoundsInPixels(Single, Single)
    //ExFor:NodeRendererBase.getOpaqueBoundsInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.getSizeInPixels(Single, Single)
    //ExFor:NodeRendererBase.getSizeInPixels(Single, Single, Single)
    //ExFor:NodeRendererBase.opaqueBoundsInPoints
    //ExFor:NodeRendererBase.sizeInPoints
    //ExFor:OfficeMathRenderer
    //ExFor:OfficeMathRenderer.#ctor(OfficeMath)
    //ExSummary:Shows how to measure and scale shapes.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getOfficeMath(0, true);
    let renderer = new aw.Rendering.OfficeMathRenderer(officeMath);

    // Verify the size of the image that the OfficeMath object will create when we render it.
    expect(Math.abs(renderer.sizeInPoints2.width - 122.0)).toBeLessThanOrEqual(0.25);
    expect(Math.abs(renderer.sizeInPoints2.height - 13.0)).toBeLessThanOrEqual(0.15);

    expect(Math.abs(renderer.boundsInPoints2.width - 122.0)).toBeLessThanOrEqual(0.25);
    expect(Math.abs(renderer.boundsInPoints2.height - 13.0)).toBeLessThanOrEqual(0.15);

    // Shapes with transparent parts may contain different values in the "OpaqueBoundsInPoints" properties.
    expect(Math.abs(renderer.opaqueBoundsInPoints2.width - 122.0)).toBeLessThanOrEqual(0.25);
    expect(Math.abs(renderer.opaqueBoundsInPoints2.height - 14.2)).toBeLessThanOrEqual(0.1);

    // Get the shape size in pixels, with linear scaling to a specific DPI.
    let bounds = renderer.getBoundsInPixels2(1.0, 96.0);

    expect(bounds.width).toEqual(163);
    expect(bounds.height).toEqual(18);

    // Get the shape size in pixels, but with a different DPI for the horizontal and vertical dimensions.
    bounds = renderer.getBoundsInPixels2(1.0, 96.0, 150.0);
    expect(bounds.width).toEqual(163);
    expect(bounds.height).toEqual(27);

    // The opaque bounds may vary here also.
    bounds = renderer.getOpaqueBoundsInPixels2(1.0, 96.0);

    expect(bounds.width).toEqual(163);
    expect(bounds.height).toEqual(19);

    bounds = renderer.getOpaqueBoundsInPixels2(1.0, 96.0, 150.0);

    expect(bounds.width).toEqual(163);
    expect(bounds.height).toEqual(29);
    //ExEnd
  });


  test('ShapeTypes', () => {
    //ExStart
    //ExFor:ShapeType
    //ExSummary:Shows how Aspose.words identify shapes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertShape(aw.Drawing.ShapeType.Heptagon, aw.Drawing.RelativeHorizontalPosition.Page, 0,
      aw.Drawing.RelativeVerticalPosition.Page, 0, 0, 0, aw.Drawing.WrapType.None);

    builder.insertShape(aw.Drawing.ShapeType.Cloud, aw.Drawing.RelativeHorizontalPosition.RightMargin, 0,
      aw.Drawing.RelativeVerticalPosition.Page, 0, 0, 0, aw.Drawing.WrapType.None);

    builder.insertShape(aw.Drawing.ShapeType.MathPlus, aw.Drawing.RelativeHorizontalPosition.RightMargin, 0,
      aw.Drawing.RelativeVerticalPosition.Page, 0, 0, 0, aw.Drawing.WrapType.None);

    // To correct identify shape types you need to work with shapes as DML.
    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    // "Strict" or "Transitional" compliance allows to save shape as DML.
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;

    doc.save(base.artifactsDir + "Shape.ShapeTypes.docx", saveOptions);
    doc = new aw.Document(base.artifactsDir + "Shape.ShapeTypes.docx");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true).toArray().map(node => node.asShape());

    for (let shape of shapes)
    {
      console.log(shape.shapeType);
    }
    //ExEnd
  });


  test('IsDecorative', () => {
    //ExStart
    //ExFor:ShapeBase.isDecorative
    //ExSummary:Shows how to set that the shape is decorative.
    let doc = new aw.Document(base.myDir + "Decorative shapes.docx");

    let shape = doc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape();
    expect(shape.isDecorative).toEqual(true);

    // If "AlternativeText" is not empty, the shape cannot be decorative.
    // That's why our value has changed to 'false'.
    shape.alternativeText = "Alternative text.";
    expect(shape.isDecorative).toEqual(false);

    let builder = new aw.DocumentBuilder(doc);

    builder.moveToDocumentEnd();
    // Create a new shape as decorative.
    shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 100, 100);
    shape.isDecorative = true;

    doc.save(base.artifactsDir + "Shape.isDecorative.docx");
    //ExEnd
  });


  test.skip('FillImage - TODO: WORDSNODEJS-92 - Method Aspose.Words.Drawing.Fill.SetImage(Stream) can\'t be call', () => {
    //ExStart
    //ExFor:Fill.setImage(String)
    //ExFor:Fill.setImage(Byte[])
    //ExFor:Fill.setImage(Stream)
    //ExSummary:Shows how to set shape fill type as image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // There are several ways of setting image.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 80, 80);
    // 1 -  Using a local system filename:
    shape.fill.setImage(base.imageDir + "Logo.jpg");
    doc.save(base.artifactsDir + "Shape.FillImage.fileName.docx");

    // 2 -  Load a file into a byte array:
    shape.fill.setImage(fs.readFileSync(base.imageDir + "Logo.jpg"));
    doc.save(base.artifactsDir + "Shape.FillImage.byteArray.docx");

    // 3 -  From a stream:
    let stream = fs.createReadStream(base.imageDir + "Logo.jpg")
    shape.fill.setImage(stream);
    doc.save(base.artifactsDir + "Shape.FillImage.stream.docx");
    //ExEnd
  });


  test('ShadowFormat', () => {
    //ExStart
    //ExFor:ShadowFormat.visible
    //ExFor:ShadowFormat.clear()
    //ExFor:ShadowType
    //ExSummary:Shows how to work with a shadow formatting for the shape.
    let doc = new aw.Document(base.myDir + "Shape stroke pattern border.docx");
    let shape = doc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape();

    if (shape.shadowFormat.visible && shape.shadowFormat.type == aw.Drawing.ShadowType.Shadow2)
      shape.shadowFormat.type = aw.Drawing.ShadowType.Shadow7;

    if (shape.shadowFormat.type == aw.Drawing.ShadowType.ShadowMixed)
      shape.shadowFormat.clear();
    //ExEnd
  });


  test('NoTextRotation', () => {
    //ExStart
    //ExFor:TextBox.noTextRotation
    //ExSummary:Shows how to disable text rotation when the shape is rotate.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.Ellipse, 20, 20);
    shape.textBox.noTextRotation = true;

    doc.save(base.artifactsDir + "Shape.noTextRotation.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Shape.noTextRotation.docx");
    shape = doc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape();

    expect(shape.textBox.noTextRotation).toEqual(true);
  });


  test('RelativeSizeAndPosition', () => {
    //ExStart
    //ExFor:ShapeBase.relativeHorizontalSize
    //ExFor:ShapeBase.relativeVerticalSize
    //ExFor:ShapeBase.widthRelative
    //ExFor:ShapeBase.heightRelative
    //ExFor:ShapeBase.topRelative
    //ExFor:ShapeBase.leftRelative
    //ExFor:RelativeHorizontalSize
    //ExFor:RelativeVerticalSize
    //ExSummary:Shows how to set relative size and position.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Adding a simple shape with absolute size and position.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 100, 40);
    // Set WrapType to WrapType.None since Inline shapes are automatically converted to absolute units.
    shape.wrapType = aw.Drawing.WrapType.None;

    // Checking and setting the relative horizontal size.
    if (shape.relativeHorizontalSize == aw.Drawing.RelativeHorizontalSize.Default)
    {
      // Setting the horizontal size binding to Margin.
      shape.relativeHorizontalSize = aw.Drawing.RelativeHorizontalSize.Margin;
      // Setting the width to 50% of Margin width.
      shape.widthRelative = 50;
    }

    // Checking and setting the relative vertical size.
    if (shape.relativeVerticalSize == aw.Drawing.RelativeVerticalSize.Default)
    {
      // Setting the vertical size binding to Margin.
      shape.relativeVerticalSize = aw.Drawing.RelativeVerticalSize.Margin;
      // Setting the heigh to 30% of Margin height.
      shape.heightRelative = 30;
    }

    // Checking and setting the relative vertical position.
    if (shape.relativeVerticalPosition == aw.Drawing.RelativeVerticalPosition.Paragraph)
    {
      // etting the position binding to TopMargin.
      shape.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.TopMargin;
      // Setting relative Top to 30% of TopMargin position.
      shape.topRelative = 30;
    }

    // Checking and setting the relative horizontal position.
    if (shape.relativeHorizontalPosition == aw.Drawing.RelativeHorizontalPosition.Default)
    {
      // Setting the position binding to RightMargin.
      shape.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.RightMargin;
      // The position relative value can be negative.
      shape.leftRelative = -260;
    }

    doc.save(base.artifactsDir + "Shape.RelativeSizeAndPosition.docx");
    //ExEnd
  });


  test.skip('FillBaseColor: WORDSNODEJS-86', () => {
    //ExStart:FillBaseColor
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Fill.baseForeColor
    //ExFor:Stroke.baseForeColor
    //ExSummary:Shows how to get foreground color without modifiers.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder();

    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 100, 40);
    shape.fill.foreColor = "#FF0000";
    shape.fill.foreTintAndShade = 0.5;
    shape.stroke.fill.foreColor = "#008000";
    shape.stroke.fill.transparency = 0.5;

    expect(shape.fill.foreColor).toEqual("#FFFFBCBC");
    expect(shape.fill.baseForeColor).toEqual("#FF0000");

    expect(shape.stroke.foreColor).toEqual("#80008000");
    expect(shape.stroke.baseForeColor).toEqual("#008000");

    expect(shape.stroke.fill.foreColor).toEqual("#008000");
    expect(shape.stroke.fill.baseForeColor).toEqual("#008000");
    //ExEnd:FillBaseColor
  });


  test('FitImageToShape', () => {
    //ExStart:FitImageToShape
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ImageData.fitImageToShape
    //ExSummary:Shows hot to fit the image data to Shape frame.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert an image shape and leave its orientation in its default state.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 300, 450);
    shape.imageData.setImage(base.imageDir + "Barcode.png");
    shape.imageData.fitImageToShape();

    doc.save(base.artifactsDir + "Shape.fitImageToShape.docx");
    //ExEnd:FitImageToShape
  });


  test('StrokeForeThemeColors', () => {
    //ExStart:StrokeForeThemeColors
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Stroke.foreThemeColor
    //ExFor:Stroke.foreTintAndShade
    //ExSummary:Shows how to set fore theme color and tint and shade.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 40);
    let stroke = shape.stroke;
    stroke.foreThemeColor = aw.Themes.ThemeColor.Dark1;
    stroke.foreTintAndShade = 0.5;

    doc.save(base.artifactsDir + "Shape.StrokeForeThemeColors.docx");
    //ExEnd:StrokeForeThemeColors

    doc = new aw.Document(base.artifactsDir + "Shape.StrokeForeThemeColors.docx");
    shape = doc.getShape(0, true);

    expect(shape.stroke.foreThemeColor).toEqual(aw.Themes.ThemeColor.Dark1);
    expect(shape.stroke.foreTintAndShade).toEqual(0.5);
  });


  test('StrokeBackThemeColors', () => {
    //ExStart:StrokeBackThemeColors
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Stroke.backThemeColor
    //ExFor:Stroke.backTintAndShade
    //ExSummary:Shows how to set back theme color and tint and shade.
    let doc = new aw.Document(base.myDir + "Stroke gradient outline.docx");

    let shape = doc.getShape(0, true);
    let stroke = shape.stroke;
    stroke.backThemeColor = aw.Themes.ThemeColor.Dark2;
    stroke.backTintAndShade = 0.2;

    doc.save(base.artifactsDir + "Shape.StrokeBackThemeColors.docx");
    //ExEnd:StrokeBackThemeColors

    doc = new aw.Document(base.artifactsDir + "Shape.StrokeBackThemeColors.docx");
    shape = doc.getShape(0, true);

    expect(shape.stroke.backThemeColor).toEqual(aw.Themes.ThemeColor.Dark2);
    expect(shape.stroke.backTintAndShade).toBeCloseTo(0.2, 6);
  });


  test('TextBoxOleControl', () => {
    //ExStart:TextBoxOleControl
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:TextBoxControl
    //ExFor:TextBoxControl.text
    //ExFor:TextBoxControl.type
    //ExSummary:Shows how to change text of the TextBox OLE control.
    let doc = new aw.Document(base.myDir + "Textbox control.docm");

    let shape = doc.getShape(0, true);
    let textBoxControl = shape.oleFormat.oleControl.asTextBoxControl();
    expect(textBoxControl.text).toEqual("Aspose.Words test");

    textBoxControl.text = "Updated text";
    expect(textBoxControl.text).toEqual("Updated text");
    expect(textBoxControl.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.Textbox);
    //ExEnd:TextBoxOleControl
  });


  test.skip('Glow: WORDSNODEJS-86', () => {
    //ExStart:Glow
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ShapeBase.glow
    //ExFor:GlowFormat
    //ExFor:GlowFormat.color
    //ExFor:GlowFormat.radius
    //ExFor:GlowFormat.transparency
    //ExFor:GlowFormat.remove()
    //ExSummary:Shows how to interact with glow shape effect.
    let doc = new aw.Document(base.myDir + "Various shapes.docx");
    let shape = doc.getShape(0, true);

    shape.glow.color = "#FA8072";
    shape.glow.radius = 30;
    shape.glow.transparency = 0.15;

    doc.save(base.artifactsDir + "Shape.glow.docx");

    doc = new aw.Document(base.artifactsDir + "Shape.glow.docx");
    shape = doc.getShape(0, true);

    expect(shape.glow.color).toEqual("#D9FA8072");
    expect(shape.glow.radius).toEqual(30);
    expect(shape.glow.transparency, 0.01).toEqual(0.15);

    shape.glow.remove();

    expect(shape.glow.color).toEqual("#000000");
    expect(shape.glow.radius).toEqual(0);
    expect(shape.glow.transparency).toEqual(0);
    //ExEnd:Glow
  });


  test('Reflection', () => {
    //ExStart:Reflection
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ShapeBase.reflection
    //ExFor:ReflectionFormat
    //ExFor:ReflectionFormat.size
    //ExFor:ReflectionFormat.blur
    //ExFor:ReflectionFormat.transparency
    //ExFor:ReflectionFormat.distance
    //ExFor:ReflectionFormat.remove()
    //ExSummary:Shows how to interact with reflection shape effect.
    let doc = new aw.Document(base.myDir + "Various shapes.docx");
    let shape = doc.getShape(0, true);

    shape.reflection.transparency = 0.37;
    shape.reflection.size = 0.48;
    shape.reflection.blur = 17.5;
    shape.reflection.distance = 9.2;

    doc.save(base.artifactsDir + "Shape.reflection.docx");

    doc = new aw.Document(base.artifactsDir + "Shape.reflection.docx");
    shape = doc.getShape(0, true);

    let reflectionFormat = shape.reflection;
    expect(reflectionFormat.transparency).toBeCloseTo(0.37, 2);
    expect(reflectionFormat.size).toBeCloseTo(0.48, 2);
    expect(reflectionFormat.blur).toBeCloseTo(17.5, 2);
    expect(reflectionFormat.distance).toBeCloseTo(9.2, 2);

    reflectionFormat.remove();

    expect(reflectionFormat.transparency).toEqual(0);
    expect(reflectionFormat.size).toEqual(0);
    expect(reflectionFormat.blur).toEqual(0);
    expect(reflectionFormat.distance).toEqual(0);
    //ExEnd:Reflection
  });


  test('SoftEdge', () => {
    //ExStart:SoftEdge
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ShapeBase.softEdge
    //ExFor:SoftEdgeFormat
    //ExFor:SoftEdgeFormat.radius
    //ExFor:SoftEdgeFormat.remove
    //ExSummary:Shows how to work with soft edge formatting.
    let builder = new aw.DocumentBuilder();
    let shape = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 200, 200);

    // Apply soft edge to the shape.
    shape.softEdge.radius = 30;

    builder.document.save(base.artifactsDir + "Shape.softEdge.docx");

    // Load document with rectangle shape with soft edge.
    let doc = new aw.Document(base.artifactsDir + "Shape.softEdge.docx");
    shape = doc.getShape(0, true);
    let softEdgeFormat = shape.softEdge;

    // Check soft edge radius.
    expect(softEdgeFormat.radius).toEqual(30);

    // Remove soft edge from the shape.
    softEdgeFormat.remove();

    // Check radius of the removed soft edge.
    expect(softEdgeFormat.radius).toEqual(0);
    //ExEnd:SoftEdge
  });


  test('Adjustments', () => {
    //ExStart:Adjustments
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:Shape.adjustments
    //ExFor:AdjustmentCollection
    //ExFor:AdjustmentCollection.count
    //ExFor:AdjustmentCollection.item(Int32)
    //ExFor:Adjustment
    //ExFor:Adjustment.name
    //ExFor:Adjustment.value
    //ExSummary:Shows how to work with adjustment raw values.
    let doc = new aw.Document(base.myDir + "Rounded rectangle shape.docx");
    let shape = doc.getShape(0, true);

    let adjustments = shape.adjustments;
    expect(adjustments.count).toEqual(1);

    let adjustment = adjustments.at(0);
    expect(adjustment.name).toEqual("adj");
    expect(adjustment.value).toEqual(16667);

    adjustment.value = 30000;

    doc.save(base.artifactsDir + "Shape.adjustments.docx");
    //ExEnd:Adjustments

    doc = new aw.Document(base.artifactsDir + "Shape.adjustments.docx");
    shape = doc.getShape(0, true);

    adjustments = shape.adjustments;
    expect(adjustments.count).toEqual(1);

    adjustment = adjustments.at(0);
    expect(adjustment.name).toEqual("adj");
    expect(adjustment.value).toEqual(30000);
  });


  test('ShadowFormatColor', () => {
    //ExStart:ShadowFormatColor
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ShapeBase.shadowFormat
    //ExFor:ShadowFormat
    //ExFor:ShadowFormat.color
    //ExFor:ShadowFormat.type
    //ExSummary:Shows how to get shadow color.
    let doc = new aw.Document(base.myDir + "Shadow color.docx");
    let shape = doc.getShape(0, true);

    expect(shape.shadowFormat.color).toEqual("#FF0000");
    //ExEnd:ShadowFormatColor
  });


  test('SetActiveXProperties', () => {
    //ExStart:SetActiveXProperties
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:Forms2OleControl.foreColor
    //ExFor:Forms2OleControl.backColor
    //ExFor:Forms2OleControl.height
    //ExFor:Forms2OleControl.width
    //ExSummary:Shows how to set properties for ActiveX control.
    let doc = new aw.Document(base.myDir + "ActiveX controls.docx");

    let shape = doc.getShape(0, true);
    let oleControl = shape.oleFormat.oleControl.asForms2OleControl();
    oleControl.foreColor = "#17E135";
    oleControl.backColor = "#3397F4";
    oleControl.height = 100.54;
    oleControl.width = 201.06;
    //ExEnd:SetActiveXProperties

    expect(oleControl.foreColor).toEqual("#17E135");
    expect(oleControl.backColor).toEqual("#3397F4");
    expect(oleControl.height).toEqual(100.54);
    expect(oleControl.width).toEqual(201.06);
  });


  test('SelectRadioControl', () => {
    //ExStart:SelectRadioControl
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:OptionButtonControl
    //ExFor:OptionButtonControl.selected
    //ExFor:OptionButtonControl.type
    //ExSummary:Shows how to select radio button.
    let doc = new aw.Document(base.myDir + "Radio buttons.docx");

    let shape1 = doc.getShape(0, true);
    let optionButton1 = shape1.oleFormat.oleControl.asOptionButtonControl();
    // Deselect selected first item.
    optionButton1.selected = false;

    let shape2 = doc.getShape(1, true);
    let optionButton2 = shape2.oleFormat.oleControl.asOptionButtonControl();
    // Select second option button.
    optionButton2.selected = true;

    expect(optionButton1.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.OptionButton);
    expect(optionButton2.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.OptionButton);

    doc.save(base.artifactsDir + "Shape.SelectRadioControl.docx");
    //ExEnd:SelectRadioControl
  });


  test('CheckedCheckBox', () => {
    //ExStart:CheckedCheckBox
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:CheckBoxControl
    //ExFor:CheckBoxControl.checked
    //ExFor:CheckBoxControl.type
    //ExFor:Forms2OleControlType
    //ExSummary:Shows how to change state of the CheckBox control.
    let doc = new aw.Document(base.myDir + "ActiveX controls.docx");

    let shape = doc.getShape(0, true);
    let checkBoxControl = shape.oleFormat.oleControl.asCheckBoxControl();

    checkBoxControl.checked = true;
            
    expect(checkBoxControl.checked).toEqual(true);
    expect(checkBoxControl.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.CheckBox);
    //ExEnd:CheckedCheckBox
  });


  test.skip('InsertGroupShape: WORDSNODEJS-141', () => {
    //ExStart:InsertGroupShape
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBuilder.insertGroupShape(double, double, double, double, ShapeBase[])
    //ExFor:DocumentBuilder.insertGroupShape(ShapeBase[])
    //ExSummary:Shows how to insert DML group shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape1 = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 200, 250);
    shape1.left = 20;
    shape1.top = 20;
    shape1.stroke.color = "#FF0000";

    let shape2 = builder.insertShape(aw.Drawing.ShapeType.Ellipse, 150, 200);
    shape2.left = 40;
    shape2.top = 50;
    shape2.stroke.color = "#008000";

    // Dimensions for the new GroupShape node.
    let left = 10;
    let top = 10;
    let width = 200;
    let height = 300;
    // Insert GroupShape node for the specified size which is inserted into the specified position.
    let groupShape1 = builder.insertGroupShape(left, top, width, height, [ shape1, shape2 ]);

    // Insert GroupShape node which position and dimension will be calculated automatically.
    let shape3 = shape1.clone(true).asShape();
    let groupShape2 = builder.insertGroupShape(shape3);

    doc.save(base.artifactsDir + "Shape.insertGroupShape.docx");
    //ExEnd:InsertGroupShape
  });


  test.skip('CombineGroupShape: WORDSNODEJS-141', () => {
    //ExStart:CombineGroupShape
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:DocumentBuilder.insertGroupShape(ShapeBase[])
    //ExSummary:Shows how to combine group shape with the shape.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape1 = builder.insertShape(aw.Drawing.ShapeType.Rectangle, 200, 250);
    shape1.left = 20;
    shape1.top = 20;
    shape1.stroke.color = "#FF0000";

    let shape2 = builder.insertShape(aw.Drawing.ShapeType.Ellipse, 150, 200);
    shape2.left = 40;
    shape2.top = 50;
    shape2.stroke.color = "#008000";

    // Combine shapes into a GroupShape node which is inserted into the specified position.
    let groupShape1 = builder.insertGroupShape(shape1, shape2);

    // Combine Shape and GroupShape nodes.
    let shape3 = shape1.clone(true).asShape();
    let groupShape2 = builder.insertGroupShape(groupShape1, shape3);

    doc.save(base.artifactsDir + "Shape.CombineGroupShape.docx");
    //ExEnd:CombineGroupShape

    doc = new aw.Document(base.artifactsDir + "Shape.CombineGroupShape.docx");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
    for (let item of shapes)
    {
      let shape = item.asShape();
      expect(shape.width).not.toEqual(0);
      expect(shape.height).not.toEqual(0);
    }
  });


  test('InsertCommandButton', () => {
    //ExStart:InsertCommandButton
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:CommandButtonControl
    //ExFor:CommandButtonControl.#ctor
    //ExFor:CommandButtonControl.type
    //ExFor:DocumentBuilder.insertForms2OleControl(Forms2OleControl)
    //ExSummary:Shows how to insert ActiveX control.
    let builder = new aw.DocumentBuilder();

    let button1 = new aw.Drawing.Ole.CommandButtonControl();
    let shape = builder.insertForms2OleControl(button1);
    expect(button1.type).toEqual(aw.Drawing.Ole.Forms2OleControlType.CommandButton);
    //ExEnd:InsertCommandButton
  });


  test('Hidden', () => {
    //ExStart:Hidden
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:ShapeBase.hidden
    //ExSummary:Shows how to hide the shape.
    let doc = new aw.Document(base.myDir + "Shadow color.docx");

    let shape = doc.getShape(0, true);
    if (!shape.hidden)
      shape.hidden = true;

    doc.save(base.artifactsDir + "Shape.hidden.docx");
    //ExEnd:Hidden
  });


  test('CommandButtonCaption', () => {
    //ExStart:CommandButtonCaption
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:Forms2OleControl.caption
    //ExSummary:Shows how to set caption for ActiveX control.
    let builder = new aw.DocumentBuilder();

    let button1 = new aw.Drawing.Ole.CommandButtonControl();
    button1.caption = "Button caption";
    let shape = builder.insertForms2OleControl(button1);
    expect(button1.caption).toEqual("Button caption");
    //ExEnd:CommandButtonCaption
  });

});
