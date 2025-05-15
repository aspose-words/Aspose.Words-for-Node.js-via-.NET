// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');


describe("WorkingWithImageSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('ExposeThresholdControl', () => {
    //ExStart:ExposeThresholdControl
    //GistId:be83b87ff2e9278db3dae459cf6f7987
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff);
    saveOptions.tiffCompression = aw.Saving.TiffCompression.Ccitt3;
    saveOptions.imageColorMode = aw.Saving.ImageColorMode.Grayscale;
    saveOptions.tiffBinarizationMethod = aw.Saving.ImageBinarizationMethod.FloydSteinbergDithering;
    saveOptions.thresholdForFloydSteinbergDithering = 254;

    doc.save(base.artifactsDir + "WorkingWithImageSaveOptions.ExposeThresholdControl.tiff", saveOptions);
    //ExEnd:ExposeThresholdControl
  });


  test('GetTiffPageRange', () => {
    //ExStart:GetTiffPageRange
    //GistId:be83b87ff2e9278db3dae459cf6f7987
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    //ExStart:SaveAsTiff
    //GistId:be83b87ff2e9278db3dae459cf6f7987
    doc.save(base.artifactsDir + "WorkingWithImageSaveOptions.MultipageTiff.tiff");
    //ExEnd:SaveAsTiff

    //ExStart:SaveAsTIFFUsingImageSaveOptions
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff);
    saveOptions.pageSet = new aw.Saving.PageSet([0, 1]);
    saveOptions.tiffCompression = aw.Saving.TiffCompression.Ccitt4;
    saveOptions.resolution = 160;

    doc.save(base.artifactsDir + "WorkingWithImageSaveOptions.GetTiffPageRange.tiff", saveOptions);
    //ExEnd:SaveAsTIFFUsingImageSaveOptions
    //ExEnd:GetTiffPageRange
  });


  test('Format1BppIndexed', () => {
    //ExStart:Format1BppIndexed
    //GistId:03144d2d1bfafb75c89d385616fdf674
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
    saveOptions.pageSet = new aw.Saving.PageSet(1);
    saveOptions.imageColorMode = aw.Saving.ImageColorMode.BlackAndWhite;
    saveOptions.pixelFormat = aw.Saving.ImagePixelFormat.Format1bppIndexed;

    doc.save(base.artifactsDir + "WorkingWithImageSaveOptions.Format1BppIndexed.png", saveOptions);
    //ExEnd:Format1BppIndexed
  });


  test('GetJpegPageRange', () => {
    //ExStart:GetJpegPageRange
    //GistId:05b9bb6f4d96094b4408287596e99a20
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);

    // Set the "PageSet" to "0" to convert only the first page of a document.
    options.pageSet = new aw.Saving.PageSet(0);

    // Change the image's brightness and contrast.
    // Both are on a 0-1 scale and are at 0.5 by default.
    options.imageBrightness = 0.3;
    options.imageContrast = 0.7;

    // Change the horizontal resolution.
    // The default value for these properties is 96.0, for a resolution of 96dpi.
    options.horizontalResolution = 72;

    doc.save(base.artifactsDir + "WorkingWithImageSaveOptions.GetJpegPageRange.jpeg", options);
    //ExEnd:GetJpegPageRange
  });

});