// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("RenderingShapes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('RenderShapeAsEmf', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    let shape = doc.getShape(0, true);

    //ExStart:RenderShapeAsEmf
    //GistId:e9a02a29ae68be63f1fdfa266a642ea1
    let renderer = shape.getShapeRenderer();
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Emf)
    {
        Scale = 1.5
    };
    renderer.save(base.artifactsDir + "RenderShape.RenderShapeAsEmf.emf", saveOptions);
    //ExEnd:RenderShapeAsEmf
  });


  test('RenderShapeAsJpeg', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    let shape = doc.getShape(0, true);

    //ExStart:RenderShapeAsJpeg
    //GistId:e9a02a29ae68be63f1fdfa266a642ea1
    let renderer = shape.getShapeRenderer();
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg)
    {
        ImageColorMode = aw.Saving.ImageColorMode.Grayscale
        ImageBrightness = 0.45

    };
    renderer.save(base.artifactsDir + "RenderShape.RenderShapeAsJpeg.jpg", saveOptions);
    //ExEnd:RenderShapeAsJpeg
  });

  test('RenderShapeImage', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    let shape = doc.getShape(0, true);

    //ExStart:RenderShapeImage
    //GistId:e9a02a29ae68be63f1fdfa266a642ea1
    shape.getShapeRenderer().save(base.artifactsDir + "RenderShape.RenderShapeImage.jpg", new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg));
    //ExEnd:RenderShapeImage
  });

});
