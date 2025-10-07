// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithShapes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AddGroupShape', () => {
    //ExStart:AddGroupShape
    //GistId:91a118ddf96c535d343b275d397fce3d
    let doc = new aw.Document();
    doc.ensureMinimum();

    let groupShape = new aw.Drawing.GroupShape(doc);
    let accentBorderShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.AccentBorderCallout1);
    accentBorderShape.width = 100;
    accentBorderShape.height = 100;
    groupShape.appendChild(accentBorderShape);

    let actionButtonShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.ActionButtonBeginning);
    actionButtonShape.left = 100;
    actionButtonShape.width = 100;
    actionButtonShape.height = 200;
    groupShape.appendChild(actionButtonShape);

    groupShape.width = 200;
    groupShape.height = 200;
    groupShape.coordSize2 = new aw.JSSize(200, 200);

    let builder = new aw.DocumentBuilder(doc);
    builder.insertNode(groupShape);

    doc.save(base.artifactsDir + "WorkingWithShapes.AddGroupShape.docx");
    //ExEnd:AddGroupShape
  });

  test('InsertShape', () => {
    //ExStart:InsertShape
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertShape(aw.Drawing.ShapeType.TextBox, aw.Drawing.RelativeHorizontalPosition.Page, 100,
        aw.Drawing.RelativeVerticalPosition.Page, 100, 50, 50, aw.Drawing.WrapType.None);
    shape.rotation = 30.0;

    builder.writeln();

    shape = builder.insertShape(aw.Drawing.ShapeType.TextBox, 50, 50);
    shape.rotation = 30.0;

    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;

    doc.save(base.artifactsDir + "WorkingWithShapes.InsertShape.docx", saveOptions);
    //ExEnd:InsertShape
  });

  test('AspectRatioLocked', () => {
    //ExStart:AspectRatioLocked
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imagesDir + "Transparent background logo.png");
    shape.aspectRatioLocked = false;

    doc.save(base.artifactsDir + "WorkingWithShapes.AspectRatioLocked.docx");
    //ExEnd:AspectRatioLocked
  });

  test('LayoutInCell', () => {
    //ExStart:LayoutInCell
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startTable();
    builder.rowFormat.height = 100;
    builder.rowFormat.heightRule = aw.HeightRule.Exactly;

    for (let i = 0; i < 31; i++) {
        if (i != 0 && i % 7 == 0) builder.endRow();
        builder.insertCell();
        builder.write("Cell contents");
    }

    builder.endTable();

    let watermark = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextPlainText);
    watermark.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
    watermark.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    watermark.isLayoutInCell = true; // Display the shape outside of the table cell if it will be placed into a cell.
    watermark.width = 300;
    watermark.height = 70;
    watermark.horizontalAlignment = aw.Drawing.HorizontalAlignment.Center;
    watermark.verticalAlignment = aw.Drawing.VerticalAlignment.Center;
    watermark.rotation = -40;

    watermark.fillColor = "#808080";
    watermark.strokeColor = "#808080";

    watermark.textPath.text = "watermarkText";
    watermark.textPath.fontFamily = "Arial";

    watermark.name = `WaterMark_${crypto.randomUUID()}`;
    watermark.wrapType = aw.Drawing.WrapType.None;

    let runs = doc.getChildNodes(aw.NodeType.Run, true);
    let run = runs.at(runs.count - 1);

    builder.moveTo(run);
    builder.insertNode(watermark);
    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2010);

    doc.save(base.artifactsDir + "WorkingWithShapes.LayoutInCell.docx");
    //ExEnd:LayoutInCell
  });

  test('AddCornersSnipped', () => {
    //ExStart:AddCornersSnipped
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertShape(aw.Drawing.ShapeType.TopCornersSnipped, 50, 50);

    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;

    doc.save(base.artifactsDir + "WorkingWithShapes.AddCornersSnipped.docx", saveOptions);
    //ExEnd:AddCornersSnipped
  });
  
  test('GetActualShapeBoundsPoints', () => {
    //ExStart:GetActualShapeBoundsPoints
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imagesDir + "Transparent background logo.png");
    shape.aspectRatioLocked = false;

    console.log("\nGets the actual bounds of the shape in points: ");
    console.log(shape.getShapeRenderer().boundsInPoints2);
    //ExEnd:GetActualShapeBoundsPoints
  });

  test('VerticalAnchor', () => {
    //ExStart:VerticalAnchor
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 200, 200);
    textBox.textBox.verticalAnchor = aw.Drawing.TextBoxAnchor.Bottom;

    builder.moveTo(textBox.firstParagraph);
    builder.write("Textbox contents");

    doc.save(base.artifactsDir + "WorkingWithShapes.VerticalAnchor.docx");
    //ExEnd:VerticalAnchor
  });

  test('DetectSmartArtShape', () => {
    //ExStart:DetectSmartArtShape
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document(base.myDir + "SmartArt.docx");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
    let count = 0;
    for (let shape of shapes) {
        shape = shape.asShape();
        if (shape.hasSmartArt) {
            count++;
        }
    }

    console.log(`The document has ${count} shapes with SmartArt.`);
    //ExEnd:DetectSmartArtShape
  });

  test('UpdateSmartArtDrawing', () => {
    let doc = new aw.Document(base.myDir + "SmartArt.docx");

    //ExStart:UpdateSmartArtDrawing
    //GistId:ddd1751fbdd164ee4c1861e4eb52a052
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
    for (let shape of shapes) {
        shape = shape.asShape();
        if (shape.hasSmartArt) {
            shape.updateSmartArtDrawing();
        }
    }
    //ExEnd:UpdateSmartArtDrawing
  });
});
