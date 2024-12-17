// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');

describe("ExImageSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('OnePage', async () => {
    //ExStart
    //ExFor:aw.Document.save(String, SaveOptions)
    //ExFor:FixedPageSaveOptions
    //ExFor:aw.Saving.ImageSaveOptions.pageSet
    //ExSummary:Shows how to render one page from a document to a JPEG image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2.");
    builder.insertImage(base.imageDir + "Logo.jpg");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3.");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);
    // Set the "PageSet" to "1" to select the second page via
    // the zero-based index to start rendering the document from.
    options.pageSet = new aw.Saving.PageSet(1);

    // When we save the document to the JPEG format, Aspose.words only renders one page.
    // This image will contain one page starting from page two,
    // which will just be the second page of the original document.
    doc.save(base.artifactsDir + "ImageSaveOptions.OnePage.jpg", options);
    //ExEnd

    await TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.OnePage.jpg");
  });


  test.each([false,
    true])('Renderer', (useGdiEmfRenderer) => {
    //ExStart
    //ExFor:aw.Saving.ImageSaveOptions.useGdiEmfRenderer
    //ExSummary:Shows how to choose a renderer when converting a document to .emf.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");

    // When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
    // If we set the "UseGdiEmfRenderer" flag to "true", Aspose.words will use the GDI+ renderer.
    // If we set the "UseGdiEmfRenderer" flag to "false", Aspose.words will use its own metafile renderer.
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Emf);
    saveOptions.useGdiEmfRenderer = useGdiEmfRenderer;

    doc.save(base.artifactsDir + "ImageSaveOptions.Renderer.emf", saveOptions);
    //ExEnd
  });


  test('PageSet', async () => {
    //ExStart
    //ExFor:aw.Saving.ImageSaveOptions.pageSet
    //ExSummary:Shows how to specify which page in a document to render as an image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world! This is page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("This is page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("This is page 3.");

    expect(doc.pageCount).toEqual(3);

    // When we save the document as an image, Aspose.words only renders the first page by default.
    // We can pass a SaveOptions object to specify a different page to render.
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Gif);
    // Render every page of the document to a separate image file.
    for (let i = 1; i <= doc.pageCount; i++)
    {
      saveOptions.pageSet = new aw.Saving.PageSet(1);

      doc.save(base.artifactsDir + `ImageSaveOptions.pageIndex.page ${i}.gif`, saveOptions);
    }
    //ExEnd

    await TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.pageIndex.page 1.gif");
    await TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.pageIndex.page 2.gif");
    await TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.pageIndex.page 3.gif");
    expect(fs.existsSync(base.artifactsDir + "ImageSaveOptions.pageIndex.page 4.gif")).toEqual(false);
  });

/*
  test('GraphicsQuality', () => {
    //ExStart
    //ExFor:GraphicsQualityOptions
    //ExFor:GraphicsQualityOptions.CompositingMode
    //ExFor:GraphicsQualityOptions.CompositingQuality
    //ExFor:GraphicsQualityOptions.InterpolationMode
    //ExFor:GraphicsQualityOptions.StringFormat
    //ExFor:GraphicsQualityOptions.SmoothingMode
    //ExFor:GraphicsQualityOptions.TextRenderingHint
    //ExFor:ImageSaveOptions.GraphicsQualityOptions
    //ExSummary:Shows how to set render quality options while converting documents to image formats. 
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let qualityOptions = new GraphicsQualityOptions();
    qualityOptions.smoothingMode = SmoothingMode.AntiAlias;
    qualityOptions.textRenderingHint = TextRenderingHint.ClearTypeGridFit;
    qualityOptions.compositingMode = CompositingMode.SourceOver;
    qualityOptions.compositingQuality = CompositingQuality.HighQuality;
    qualityOptions.interpolationMode = InterpolationMode.high;
    qualityOptions.stringFormat = StringFormat.GenericTypographic
    };

    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);
    saveOptions.GraphicsQualityOptions = qualityOptions;

    doc.save(base.artifactsDir + "ImageSaveOptions.GraphicsQuality.jpg", saveOptions);
    //ExEnd

    TestUtil.verifyImage(794, 1122, base.artifactsDir + "ImageSaveOptions.GraphicsQuality.jpg");
  });


  test('UseTileFlipMode', () => {
    //ExStart
    //ExFor:GraphicsQualityOptions.UseTileFlipMode
    //ExSummary:Shows how to prevent the white line appears when rendering with a high resolution.
    let doc = new aw.Document(base.myDir + "Shape high dpi.docx");

    let shape = (Shape)doc.getShape(0, true);
    let renderer = shape.getShapeRenderer();

    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png)
    {
      Resolution = 500, GraphicsQualityOptions = new GraphicsQualityOptions { UseTileFlipMode = true }
    };
    renderer.save(base.artifactsDir + "ImageSaveOptions.UseTileFlipMode.png", saveOptions);
    //ExEnd
  });

*/

  test.each([aw.Saving.MetafileRenderingMode.Vector,
    aw.Saving.MetafileRenderingMode.Bitmap,
    aw.Saving.MetafileRenderingMode.VectorWithFallback])('WindowsMetaFile', async (metafileRenderingMode) => {
    //ExStart
    //ExFor:aw.Saving.ImageSaveOptions.metafileRenderingOptions
    //ExFor:aw.Saving.MetafileRenderingOptions.useGdiRasterOperationsEmulation
    //ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imageDir + "Windows MetaFile.wmf");

    // When we save the document as an image, we can pass a SaveOptions object to
    // determine how the saving operation will process Windows Metafiles in the document.
    // If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
    // or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
    // If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
    options.metafileRenderingOptions.renderingMode = metafileRenderingMode;
    // Aspose.words uses GDI+ for raster operations emulation, when value is set to true.
    options.metafileRenderingOptions.useGdiRasterOperationsEmulation = true;

    doc.save(base.artifactsDir + "ImageSaveOptions.WindowsMetaFile.png", options);
    //ExEnd

    await TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.WindowsMetaFile.png");
  });


  test('PageByPage', () => {
    //ExStart
    //ExFor:aw.Document.save(String, SaveOptions)
    //ExFor:FixedPageSaveOptions
    //ExFor:aw.Saving.ImageSaveOptions.pageSet
    //ExFor:aw.Saving.ImageSaveOptions.imageSize
    //ExSummary:Shows how to render every page of a document to a separate TIFF image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2.");
    builder.insertImage(base.imageDir + "Logo.jpg");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3.");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff);

    for (let i = 0; i < doc.pageCount; i++)
    {
      // Set the "PageSet" property to the number of the first page from
      // which to start rendering the document from.
      options.pageSet = new aw.Saving.PageSet(i);
      // Export page at 2325x5325 pixels and 600 dpi.
      options.resolution = 600;
      options.imageSize2 = new aw.JSSize(2325, 5325);

      doc.save(base.artifactsDir + `ImageSaveOptions.PageByPage.${i + 1}.tiff`, options);
    }
    //ExEnd

    const imageFileNames = fs.readdirSync(base.artifactsDir)
          .filter(item => item.includes("ImageSaveOptions.PageByPage.") && item.endsWith(".tiff"));
    expect(imageFileNames.length).toEqual(3);
  });


  test.each([aw.Saving.ImageColorMode.BlackAndWhite,
    aw.Saving.ImageColorMode.Grayscale,
    aw.Saving.ImageColorMode.None])('ColorMode', (imageColorMode) => {
    //ExStart
    //ExFor:ImageColorMode
    //ExFor:aw.Saving.ImageSaveOptions.imageColorMode
    //ExSummary:Shows how to set a color mode when rendering documents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");

    // When we save the document as an image, we can pass a SaveOptions object to
    // select a color mode for the image that the saving operation will generate.
    // If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
    // the saving operation will apply grayscale color reduction while rendering the document.
    // If we set the "ImageColorMode" property to "ImageColorMode.Grayscale", 
    // the saving operation will render the document into a monochrome image.
    // If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
    // and preserve all the document's colors in the output image.
    let imageSaveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
    imageSaveOptions.imageColorMode = imageColorMode;

    doc.save(base.artifactsDir + "ImageSaveOptions.colorMode.png", imageSaveOptions);
    //ExEnd

    var testedImageLength = fs.statSync(base.artifactsDir + "ImageSaveOptions.colorMode.png").size;

    switch (imageColorMode)
    {
      case aw.Saving.ImageColorMode.None:
        expect(testedImageLength < 132000).toEqual(true);
        break;
      case aw.Saving.ImageColorMode.Grayscale:
        expect(testedImageLength < 90000).toEqual(true);
        break;
      case aw.Saving.ImageColorMode.BlackAndWhite:
        expect(testedImageLength < 11000).toEqual(true);
        break;
    }
  });


  test('PaperColor', async () => {
    //ExStart
    //ExFor:ImageSaveOptions
    //ExFor:aw.Saving.ImageSaveOptions.paperColor
    //ExSummary:Renders a page of a Word document into an image with transparent or colored background.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Times New Roman";
    builder.font.size = 24;
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    builder.insertImage(base.imageDir + "Logo.jpg");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let imgOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
    // Set the "PaperColor" property to a transparent color to apply a transparent
    // background to the document while rendering it to an image.
    imgOptions.paperColor = "rgba(0,0,0,0)";

    doc.save(base.artifactsDir + "ImageSaveOptions.paperColor.transparent.png", imgOptions);

    // Set the "PaperColor" property to an opaque color to apply that color
    // as the background of the document as we render it to an image.
    imgOptions.paperColor = "#F08080";

    doc.save(base.artifactsDir + "ImageSaveOptions.paperColor.LightCoral.png", imgOptions);
    //ExEnd

    expect(await TestUtil.imageContainsTransparency(base.artifactsDir + "ImageSaveOptions.paperColor.transparent.png")).toEqual(true);
    expect(await TestUtil.imageContainsTransparency(base.artifactsDir + "ImageSaveOptions.paperColor.LightCoral.png")).toEqual(false);
  });


  test.each([aw.Saving.ImagePixelFormat.Format1bppIndexed,
    aw.Saving.ImagePixelFormat.Format16BppRgb555,
    aw.Saving.ImagePixelFormat.Format16BppRgb565,
    aw.Saving.ImagePixelFormat.Format24BppRgb,
    aw.Saving.ImagePixelFormat.Format32BppRgb,
    aw.Saving.ImagePixelFormat.Format32BppArgb,
    aw.Saving.ImagePixelFormat.Format32BppPArgb,
    aw.Saving.ImagePixelFormat.Format48BppRgb,
    aw.Saving.ImagePixelFormat.Format64BppArgb,
    aw.Saving.ImagePixelFormat.Format64BppPArgb])('PixelFormat', (imagePixelFormat) => {
    //ExStart
    //ExFor:ImagePixelFormat
    //ExFor:aw.Saving.ImageSaveOptions.clone
    //ExFor:aw.Saving.ImageSaveOptions.pixelFormat
    //ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");

    // When we save the document as an image, we can pass a SaveOptions object to
    // select a pixel format for the image that the saving operation will generate.
    // Various bit per pixel rates will affect the quality and file size of the generated image.
    let imageSaveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);
    imageSaveOptions.pixelFormat = imagePixelFormat;

    // We can clone ImageSaveOptions instances.
    expect(imageSaveOptions.clone()).not.toBe(imageSaveOptions);
    expect(imageSaveOptions.clone()).toEqual(imageSaveOptions);

    doc.save(base.artifactsDir + "ImageSaveOptions.pixelFormat.png", imageSaveOptions);
    //ExEnd

    var testedImageLength = fs.statSync(base.artifactsDir + "ImageSaveOptions.PixelFormat.png").size;

    switch (imagePixelFormat)
    {
      case aw.Saving.ImagePixelFormat.Format1bppIndexed:
        expect(testedImageLength < 7500).toEqual(true);
        break;
      case aw.Saving.ImagePixelFormat.Format24BppRgb:
        expect(testedImageLength < 77000).toEqual(true);
        break;
      case aw.Saving.ImagePixelFormat.Format16BppRgb565:
      case aw.Saving.ImagePixelFormat.Format16BppRgb555:
      case aw.Saving.ImagePixelFormat.Format32BppRgb:
      case aw.Saving.ImagePixelFormat.Format32BppArgb:
      case aw.Saving.ImagePixelFormat.Format48BppRgb:
      case aw.Saving.ImagePixelFormat.Format64BppArgb:
      case aw.Saving.ImagePixelFormat.Format64BppPArgb:
        expect(testedImageLength < 132000).toEqual(true);
        break;
    }
  });

/*
  test('FloydSteinbergDithering', () => {
    //ExStart
    //ExFor:ImageBinarizationMethod
    //ExFor:aw.Saving.ImageSaveOptions.thresholdForFloydSteinbergDithering
    //ExFor:aw.Saving.ImageSaveOptions.tiffBinarizationMethod
    //ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");

    // When we save the document as a TIFF, we can pass a SaveOptions object to
    // adjust the dithering that Aspose.words will apply when rendering this image.
    // The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
    // Higher values tend to produce darker images.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff)
    {
      TiffCompression = aw.Saving.TiffCompression.Ccitt3,
      TiffBinarizationMethod = aw.Saving.ImageBinarizationMethod.FloydSteinbergDithering,
      ThresholdForFloydSteinbergDithering = 240
    };

    doc.save(base.artifactsDir + "ImageSaveOptions.floydSteinbergDithering.tiff", options);
    //ExEnd
#if NET461_OR_GREATER || JAVA
    TestUtil.verifyImage(816, 1056, base.artifactsDir + "ImageSaveOptions.floydSteinbergDithering.tiff");
#endif
  });


    [AotTests.IgnoreAot("Failed on net7")]
  test('EditImage', () => {
    //ExStart
    //ExFor:aw.Saving.ImageSaveOptions.horizontalResolution
    //ExFor:aw.Saving.ImageSaveOptions.imageBrightness
    //ExFor:aw.Saving.ImageSaveOptions.imageContrast
    //ExFor:aw.Saving.ImageSaveOptions.saveFormat
    //ExFor:aw.Saving.ImageSaveOptions.scale
    //ExFor:aw.Saving.ImageSaveOptions.verticalResolution
    //ExSummary:Shows how to edit the image while Aspose.words converts a document to one.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");

    // When we save the document as an image, we can pass a SaveOptions object to
    // edit the image while the saving operation renders it.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png)
    {
      // We can adjust these properties to change the image's brightness and contrast.
      // Both are on a 0-1 scale and are at 0.5 by default.
      ImageBrightness = 0.3f,
      ImageContrast = 0.7f,

      // We can adjust horizontal and vertical resolution with these properties.
      // This will affect the dimensions of the image.
      // The default value for these properties is 96.0, for a resolution of 96dpi.
      HorizontalResolution = 72f,
      VerticalResolution = 72f,

      // We can scale the image using this property. The default value is 1.0, for scaling of 100%.
      // We can use this property to negate any changes in image dimensions that changing the resolution would cause.
      Scale = 96f / 72f
    };

    doc.save(base.artifactsDir + "ImageSaveOptions.EditImage.png", options);
    //ExEnd

#if NET5_0_OR_GREATER
    const int expectedWidth = 816;
    const int expectedHeight = 1056;
#else
    const int expectedWidth = 817;
    const int expectedHeight = 1057;
#endif
    TestUtil.verifyImage(expectedWidth, expectedHeight, base.artifactsDir + "ImageSaveOptions.EditImage.png");
  });


  test('JpegQuality', () => {
    //ExStart
    //ExFor:aw.Document.save(String, SaveOptions)
    //ExFor:aw.Saving.FixedPageSaveOptions.jpegQuality
    //ExFor:ImageSaveOptions
    //ExFor:ImageSaveOptions.#ctor
    //ExFor:aw.Saving.ImageSaveOptions.jpegQuality
    //ExSummary:Shows how to configure compression while saving a document as a JPEG.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.insertImage(base.imageDir + "Logo.jpg");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let imageOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);
    // Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
    // This will reduce the file size of the document, but the image will display more prominent compression artifacts.
    imageOptions.jpegQuality = 10;
    doc.save(base.artifactsDir + "ImageSaveOptions.jpegQuality.HighCompression.jpg", imageOptions);

    // Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
    // This will improve the quality of the image at the cost of an increased file size.
    imageOptions.jpegQuality = 100;
    doc.save(base.artifactsDir + "ImageSaveOptions.jpegQuality.HighQuality.jpg", imageOptions);
    //ExEnd

    expect(new awDynabic.Metering.Metered.BillingServices.Internals.FileInfo(base.artifactsDir + "ImageSaveOptions.jpegQuality.HighCompression.jpg").Length < 18000).toEqual(true);
    expect(new awDynabic.Metering.Metered.BillingServices.Internals.FileInfo(base.artifactsDir + "ImageSaveOptions.jpegQuality.HighQuality.jpg").Length < 75000).toEqual(true);
  });


  test.each([TiffCompression.None), Category("SkipMono",
    TiffCompression.Rle), Category("SkipMono",
    TiffCompression.Lzw), Category("SkipMono",
    TiffCompression.Ccitt3), Category("SkipMono",
    TiffCompression.Ccitt4), Category("SkipMono"])('TiffImageCompression', (TiffCompression tiffCompression) => {
    //ExStart
    //ExFor:TiffCompression
    //ExFor:aw.Saving.ImageSaveOptions.tiffCompression
    //ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertImage(base.imageDir + "Logo.jpg");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff);
    // Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
    // which may result in a very large output file.
    // Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
    // Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
    // Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
    // Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
    options.tiffCompression = tiffCompression;

    doc.save(base.artifactsDir + "ImageSaveOptions.TiffImageCompression.tiff", options);
    //ExEnd

    var testedImageLength = new awDynabic.Metering.Metered.BillingServices.Internals.FileInfo(base.artifactsDir + "ImageSaveOptions.TiffImageCompression.tiff").Length;

    switch (tiffCompression)
    {
      case aw.Saving.TiffCompression.None:
        expect(testedImageLength < 3450000).toEqual(true);
        break;
      case aw.Saving.TiffCompression.Rle:
#if NET5_0_OR_GREATER
        expect(testedImageLength < 7500).toEqual(true);
#else
        expect(testedImageLength < 687000).toEqual(true);
#endif
        break;
      case aw.Saving.TiffCompression.Lzw:
        expect(testedImageLength < 250000).toEqual(true);
        break;
      case aw.Saving.TiffCompression.Ccitt3:
#if NET5_0_OR_GREATER
        expect(testedImageLength < 6100).toEqual(true);
#else
        expect(testedImageLength < 8300).toEqual(true);
#endif
        break;
      case aw.Saving.TiffCompression.Ccitt4:
        expect(testedImageLength < 1700).toEqual(true);
        break;
    }
  });


  test('Resolution', () => {
    //ExStart
    //ExFor:ImageSaveOptions
    //ExFor:aw.Saving.ImageSaveOptions.resolution
    //ExSummary:Shows how to specify a resolution while rendering a document to PNG.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Times New Roman";
    builder.font.size = 24;
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    builder.insertImage(base.imageDir + "Logo.jpg");

    // Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
    // to modify the way in which that method renders the document into an image.
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Png);

    // Set the "Resolution" property to "72" to render the document in 72dpi.
    options.resolution = 72;
    doc.save(base.artifactsDir + "ImageSaveOptions.resolution.72dpi.png", options);

    // Set the "Resolution" property to "300" to render the document in 300dpi.
    options.resolution = 300;
    doc.save(base.artifactsDir + "ImageSaveOptions.resolution.300dpi.png", options);
    //ExEnd

    TestUtil.verifyImage(612, 792, base.artifactsDir + "ImageSaveOptions.resolution.72dpi.png");
    TestUtil.verifyImage(2550, 3300, base.artifactsDir + "ImageSaveOptions.resolution.300dpi.png");
  });


  test('ExportVariousPageRanges', () => {
    //ExStart
    //ExFor:PageSet.#ctor(PageRange[])
    //ExFor:PageRange.#ctor(int, int)
    //ExFor:aw.Saving.ImageSaveOptions.pageSet
    //ExSummary:Shows how to extract pages based on exact page ranges.
    let doc = new aw.Document(base.myDir + "Images.docx");

    let imageOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Tiff);
    let pageSet = new aw.Saving.PageSet(new aw.Saving.PageRange(1, 1), new aw.Saving.PageRange(2, 3), new aw.Saving.PageRange(1, 3),
      new aw.Saving.PageRange(2, 4), new aw.Saving.PageRange(1, 1));

    imageOptions.pageSet = pageSet;
    doc.save(base.artifactsDir + "ImageSaveOptions.ExportVariousPageRanges.tiff", imageOptions);
    //ExEnd
  });


  test('RenderInkObject', () => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.imlRenderingMode
    //ExFor:ImlRenderingMode
    //ExSummary:Shows how to render Ink object.
    let doc = new aw.Document(base.myDir + "Ink object.docx");

    // Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
    // If the rendering result is unsatisfactory,
    // please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg)
    {
      ImlRenderingMode = aw.Saving.ImlRenderingMode.InkML
    };

    doc.save(base.artifactsDir + "ImageSaveOptions.RenderInkObject.jpeg", saveOptions);
    //ExEnd
  });

*/
});
