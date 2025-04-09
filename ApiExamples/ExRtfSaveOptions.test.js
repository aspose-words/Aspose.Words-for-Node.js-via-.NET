// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');

describe("ExRtfSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.skip.each([false, true])('ExportImages(%o) - TODO: WORDSNODEJS-81', (exportImagesForOldReaders) => {
    //ExStart
    //ExFor:RtfSaveOptions
    //ExFor:RtfSaveOptions.exportCompactSize
    //ExFor:RtfSaveOptions.exportImagesForOldReaders
    //ExFor:RtfSaveOptions.saveFormat
    //ExSummary:Shows how to save a document to .rtf with custom options.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
    let options = new aw.Saving.RtfSaveOptions();

    expect(options.saveFormat).toEqual(aw.SaveFormat.Rtf);

    // Set the "ExportCompactSize" property to "true" to
    // reduce the saved document's size at the cost of right-to-left text compatibility.
    options.exportCompactSize = true;

    // Set the "ExportImagesFotOldReaders" property to "true" to use extra keywords to ensure that our document is
    // compatible with pre-Microsoft Word 97 readers and WordPad.
    // Set the "ExportImagesFotOldReaders" property to "false" to reduce the size of the document,
    // but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
    options.exportImagesForOldReaders = exportImagesForOldReaders;

    doc.save(base.artifactsDir + "RtfSaveOptions.ExportImages.rtf", options);
    //ExEnd

    if (exportImagesForOldReaders)
    {
      TestUtil.fileContainsString("nonshppict", base.artifactsDir + "RtfSaveOptions.ExportImages.rtf");
      TestUtil.fileContainsString("shprslt", base.artifactsDir + "RtfSaveOptions.ExportImages.rtf");
    }
    else
    {
      expect(() => { TestUtil.fileContainsString("nonshppict", base.artifactsDir + "RtfSaveOptions.ExportImages.rtf"); }).toThrow();
      expect(() => { TestUtil.fileContainsString("shprslt", base.artifactsDir + "RtfSaveOptions.ExportImages.rtf"); }).toThrow();
    }
  });


  test.each([false, true])('SaveImagesAsWmf(%o)', (saveImagesAsWmf) => {
    //ExStart
    //ExFor:RtfSaveOptions.saveImagesAsWmf
    //ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Jpeg image:");
    let imageShape = builder.insertImage(base.imageDir + "Logo.jpg");

    expect(imageShape.imageData.imageType).toEqual(aw.Drawing.ImageType.Jpeg);

    builder.insertParagraph();
    builder.writeln("Png image:");
    imageShape = builder.insertImage(base.imageDir + "Transparent background logo.png");

    expect(imageShape.imageData.imageType).toEqual(aw.Drawing.ImageType.Png);

    // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
    let rtfSaveOptions = new aw.Saving.RtfSaveOptions();

    // Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
    // Doing so will help readers such as WordPad to read our document.
    // Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
    // as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
    rtfSaveOptions.saveImagesAsWmf = saveImagesAsWmf;

    doc.save(base.artifactsDir + "RtfSaveOptions.saveImagesAsWmf.rtf", rtfSaveOptions);

    doc = new aw.Document(base.artifactsDir + "RtfSaveOptions.saveImagesAsWmf.rtf");

    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);

    if (saveImagesAsWmf)
    {
      expect(shapes.at(0).asShape().imageData.imageType).toEqual(aw.Drawing.ImageType.Wmf);
      expect(shapes.at(1).asShape().imageData.imageType).toEqual(aw.Drawing.ImageType.Wmf);
    }
    else
    {
      expect(shapes.at(0).asShape().imageData.imageType).toEqual(aw.Drawing.ImageType.Jpeg);
      expect(shapes.at(1).asShape().imageData.imageType).toEqual(aw.Drawing.ImageType.Png);
    }
    //ExEnd
  });

});
