// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExDocumentBase", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.skip('Constructor - Object.GetType() isn\'t supported.', () => {
    //ExStart
    //ExFor:DocumentBase
    //ExSummary:Shows how to initialize the subclasses of DocumentBase.
    let doc = new aw.Document();

    expect(doc.getType().baseType).toEqual(typeof(DocumentBase));

    let glossaryDoc = new aw.BuildingBlocks.GlossaryDocument();
    doc.glossaryDocument = glossaryDoc;

    expect(glossaryDoc.getType().baseType).toEqual(typeof(DocumentBase));
    //ExEnd
  });


  test('SetPageColor', () => {
    //ExStart
    //ExFor:DocumentBase.pageColor
    //ExSummary:Shows how to set the background color for all pages of a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    doc.pageColor = "#D3D3D3";

    doc.save(base.artifactsDir + "DocumentBase.SetPageColor.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBase.SetPageColor.docx");

    expect(doc.pageColor).toEqual("#D3D3D3");
  });


  test('ImportNode', () => {
    //ExStart
    //ExFor:DocumentBase.importNode(Node, Boolean)
    //ExSummary:Shows how to import a node from one document to another.
    let srcDoc = new aw.Document();
    let dstDoc = new aw.Document();

    srcDoc.firstSection.body.firstParagraph.appendChild(
      new aw.Run(srcDoc, "Source document first paragraph text."));
    dstDoc.firstSection.body.firstParagraph.appendChild(
      new aw.Run(dstDoc, "Destination document first paragraph text."));

    // Every node has a parent document, which is the document that contains the node.
    // Inserting a node into a document that the node does not belong to will throw an exception.
    expect(srcDoc.firstSection.document).not.toEqual(dstDoc);
    expect(() => { dstDoc.appendChild(srcDoc.firstSection); }).toThrow("The newChild was created from a different document than the one that created this node.");

    // Use the ImportNode method to create a copy of a node, which will have the document
    // that called the ImportNode method set as its new owner document.
    let importedSection = dstDoc.importNode(srcDoc.firstSection, true).asSection();

    expect(importedSection.document.referenceEquals(dstDoc)).toBeTruthy();

    // We can now insert the node into the document.
    dstDoc.appendChild(importedSection);

    expect(dstDoc.toString(aw.SaveFormat.Text)).toEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n");
    //ExEnd

    expect(srcDoc.firstSection.referenceEquals(importedSection)).toBeFalsy();
    expect(srcDoc.firstSection.document).not.toEqual(importedSection.document);
    expect(srcDoc.firstSection.body.firstParagraph.getText()).toEqual(importedSection.body.firstParagraph.getText());
  });


  test('ImportNodeCustom', () => {
    //ExStart
    //ExFor:DocumentBase.importNode(Node, Boolean, ImportFormatMode)
    //ExSummary:Shows how to import node from source document to destination document with specific options.
    // Create two documents and add a character style to each document.
    // Configure the styles to have the same name, but different text formatting.
    let srcDoc = new aw.Document();
    let srcStyle = srcDoc.styles.add(aw.StyleType.Character, "My style");
    srcStyle.font.name = "Courier New";
    let srcBuilder = new aw.DocumentBuilder(srcDoc);
    srcBuilder.font.style = srcStyle;
    srcBuilder.writeln("Source document text.");

    let dstDoc = new aw.Document();
    let dstStyle = dstDoc.styles.add(aw.StyleType.Character, "My style");
    dstStyle.font.name = "Calibri";
    let dstBuilder = new aw.DocumentBuilder(dstDoc);
    dstBuilder.font.style = dstStyle;
    dstBuilder.writeln("Destination document text.");

    // Import the Section from the destination document into the source document, causing a style name collision.
    // If we use destination styles, then the imported source text with the same style name
    // as destination text will adopt the destination style.
    let importedSection = dstDoc.importNode(srcDoc.firstSection, true, aw.ImportFormatMode.UseDestinationStyles).asSection();
    expect(importedSection.body.paragraphs.at(0).runs.at(0).getText().trim()).toEqual("Source document text.");
    expect(dstDoc.styles.at("My style_0")).toBe(null);
    expect(importedSection.body.firstParagraph.runs.at(0).font.name).toEqual(dstStyle.font.name);
    expect(importedSection.body.firstParagraph.runs.at(0).font.styleName).toEqual(dstStyle.name);

    // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
    // and the naming clash resolves by adding a suffix.
    dstDoc.importNode(srcDoc.firstSection, true, aw.ImportFormatMode.KeepDifferentStyles);
    expect(dstDoc.styles.at("My style").font.name).toEqual(dstStyle.font.name);
    expect(dstDoc.styles.at("My style_0").font.name).toEqual(srcStyle.font.name);
    //ExEnd
  });


  test('BackgroundShape', () => {
    //ExStart
    //ExFor:DocumentBase.backgroundShape
    //ExSummary:Shows how to set a background shape for every page of a document.
    let doc = new aw.Document();

    expect(doc.backgroundShape).toBe(null);

    // The only shape type that we can use as a background is a rectangle.
    let shapeRectangle = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);

    // There are two ways of using this shape as a page background.
    // 1 -  A flat color:
    shapeRectangle.fillColor = "#ADD8E6";
    doc.backgroundShape = shapeRectangle;

    doc.save(base.artifactsDir + "DocumentBase.backgroundShape.FlatColor.docx");

    // 2 -  An image:
    shapeRectangle = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shapeRectangle.imageData.setImage(base.imageDir + "Transparent background logo.png");

    // Adjust the image's appearance to make it more suitable as a watermark.
    shapeRectangle.imageData.contrast = 0.2;
    shapeRectangle.imageData.brightness = 0.7;

    doc.backgroundShape = shapeRectangle;

    expect(doc.backgroundShape.hasImage).toEqual(true);

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.CacheBackgroundGraphics = false;

    // Microsoft Word does not support shapes with images as backgrounds,
    // but we can still see these backgrounds in other save formats such as .pdf.
    doc.save(base.artifactsDir + "DocumentBase.backgroundShape.image.pdf", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBase.backgroundShape.FlatColor.docx");

    expect(doc.backgroundShape.fillColor).toEqual("#ADD8E6");
    expect(() => { doc.backgroundShape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Triangle); }).toThrow("Only a rectangle shape can be set as a document background.");
  });


  /*//Commented
  //ExStart
  //ExFor:DocumentBase.ResourceLoadingCallback
  //ExFor:IResourceLoadingCallback
  //ExFor:IResourceLoadingCallback.ResourceLoading(ResourceLoadingArgs)
  //ExFor:ResourceLoadingAction
  //ExFor:ResourceLoadingArgs
  //ExFor:ResourceLoadingArgs.OriginalUri
  //ExFor:ResourceLoadingArgs.ResourceType
  //ExFor:ResourceLoadingArgs.SetData(Byte[])
  //ExFor:ResourceType
  //ExSummary:Shows how to customize the process of loading external resources into a document.
  test('ResourceLoadingCallback', () => {
    let doc = new aw.Document();
    doc.resourceLoadingCallback = new ImageNameHandler();

    let builder = new aw.DocumentBuilder(doc);

    // Images usually are inserted using a URI, or a byte array.
    // Every instance of a resource load will call our callback's ResourceLoading method.
    builder.insertImage("Google logo");
    builder.insertImage("Aspose logo");
    builder.insertImage("Watermark");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(3);

    doc.save(base.artifactsDir + "DocumentBase.resourceLoadingCallback.docx");
    TestResourceLoadingCallback(new aw.Document(base.artifactsDir + "DocumentBase.resourceLoadingCallback.docx")); //ExSkip
  });


  /// <summary>
  /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
  /// This will separate image loading logic from the rest of the document construction.
  /// </summary>
  private class ImageNameHandler : IResourceLoadingCallback
  {
    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
        // If this callback encounters one of the image shorthands while loading an image,
        // it will apply unique logic for each defined shorthand instead of treating it as a URI.
      if (args.resourceType == aw.Loading.ResourceType.Image)
        switch (args.originalUri)
        {
          case "Google logo":
#pragma warning disable SYSLIB0014
#pragma warning restore SYSLIB0014
            {
              args.setData(webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png"));
            }

            return aw.Loading.ResourceLoadingAction.UserProvided;

          case "Aspose logo":
            args.setData(File.ReadAllBytes(base.imageDir + "Logo.jpg"));

            return aw.Loading.ResourceLoadingAction.UserProvided;

          case "Watermark":
            args.setData(File.ReadAllBytes(base.imageDir + "Transparent background logo.png"));

            return aw.Loading.ResourceLoadingAction.UserProvided;
        }

      return aw.Loading.ResourceLoadingAction.Default;
    }
  }
    //ExEnd

  private void TestResourceLoadingCallback(Document doc)
  {
    foreach (Shape shape in doc.getChildNodes(aw.NodeType.Shape, true))
    {
      expect(shape.hasImage).toEqual(true);
      Assert.IsNotEmpty(shape.imageData.imageBytes);
    }
  }
  //EndCommented*/

});
