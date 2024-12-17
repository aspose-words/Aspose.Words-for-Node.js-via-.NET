// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');
const fs = require('fs');


describe("ExSvgSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('SaveLikeImage', () => {
    //ExStart
    //ExFor:aw.Saving.SvgSaveOptions.fitToViewPort
    //ExFor:aw.Saving.SvgSaveOptions.showPageBorder
    //ExFor:aw.Saving.SvgSaveOptions.textOutputMode
    //ExFor:SvgTextOutputMode
    //ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Configure the SvgSaveOptions object to save with no page borders or selectable text.
    let options = new aw.Saving.SvgSaveOptions();
    options.fitToViewPort = true;
    options.showPageBorder = false;
    options.textOutputMode = aw.Saving.SvgTextOutputMode.UsePlacedGlyphs;

    doc.save(base.artifactsDir + "SvgSaveOptions.SaveLikeImage.svg", options);
    //ExEnd
  });


  //ExStart
  //ExFor:SvgSaveOptions
  //ExFor:SvgSaveOptions.ExportEmbeddedImages
  //ExFor:SvgSaveOptions.ResourceSavingCallback
  //ExFor:SvgSaveOptions.ResourcesFolder
  //ExFor:SvgSaveOptions.ResourcesFolderAlias
  //ExFor:SvgSaveOptions.SaveFormat
  //ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
  test.skip('SvgResourceFolder - TODO: sourceSavingCallback not supported yet', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.SvgSaveOptions();
    options.saveFormat = aw.SaveFormat.Svg;
    options.exportEmbeddedImages = false;
    options.resourcesFolder = base.artifactsDir + "SvgResourceFolder";
    options.resourcesFolderAlias = base.artifactsDir + "SvgResourceFolderAlias";
    options.showPageBorder = false;

    options.sourceSavingCallback = new ResourceUriPrinter()

    if (!fs.existsSync(options.resourcesFolderAlias)) {
      fs.mkdirSync(options.resourcesFolderAlias);
    }

    doc.save(base.artifactsDir + "SvgSaveOptions.SvgResourceFolder.svg", options);
  });

/*
  /// <summary>
  /// Counts and prints URIs of resources contained by as they are converted to .svg.
  /// </summary>
  private class ResourceUriPrinter : IResourceSavingCallback
  {
    void aw.Saving.IResourceSavingCallback.resourceSaving(ResourceSavingArgs args)
    {
      console.log(`Resource #${++mSavedResourceCount} \"${args.resourceFileName}\"`);
      console.log("\t" + args.resourceFileUri);
    }

    private int mSavedResourceCount;
  }
    //ExEnd
   */


  test('SaveOfficeMath', () => {
    //ExStart:SaveOfficeMath
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:aw.Rendering.NodeRendererBase.save(String, SvgSaveOptions)
    //ExSummary:Shows how to pass save options when rendering office math.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let math = doc.getOfficeMath(0, true);

    let options = new aw.Saving.SvgSaveOptions();
    options.textOutputMode = aw.Saving.SvgTextOutputMode.UsePlacedGlyphs;

    math.getMathRenderer().save(base.artifactsDir + "SvgSaveOptions.Output.svg", options);
    //ExEnd:SaveOfficeMath
  });


  test('MaxImageResolution', () => {
    //ExStart:MaxImageResolution
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:aw.Drawing.ShapeBase.softEdge
    //ExFor:aw.Drawing.SoftEdgeFormat.radius
    //ExFor:aw.Drawing.SoftEdgeFormat.remove
    //ExSummary:Shows how to set limit for image resolution.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.SvgSaveOptions();
    saveOptions.maxImageResolution = 72;

    doc.save(base.artifactsDir + "SvgSaveOptions.maxImageResolution.svg", saveOptions);
    //ExEnd:MaxImageResolution
  });

});
