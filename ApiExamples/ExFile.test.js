// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const fs = require('fs');


describe("ExFile", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CatchFileCorruptedException', () => {
    //ExStart
    //ExFor:FileCorruptedException
    //ExSummary:Shows how to catch a FileCorruptedException.

    // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
    // chances are that we will get an exception thrown when trying to load that document using Aspose.words.
    expect(() => new aw.Document(base.myDir + "Corrupted document.docx"))
      .toThrow("The document appears to be corrupted and cannot be loaded.");
    //ExEnd
  });


  test('DetectEncoding', () => {
    //ExStart
    //ExFor:aw.FileFormatInfo.encoding
    //ExFor:FileFormatUtil
    //ExSummary:Shows how to detect encoding in an html file.
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Document.html");

    expect(info.loadFormat).toEqual(aw.LoadFormat.Html);

    // The Encoding property is used only when we create a FileFormatInfo object for an html document.
    expect(info.encoding).toEqual("windows-1252");
    //ExEnd

    info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Document.docx");

    expect(info.loadFormat).toEqual(aw.LoadFormat.Docx);
    expect(info.encoding).toBe(null);
  });


  test('FileFormatToString', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.contentTypeToLoadFormat(String)
    //ExFor:aw.FileFormatUtil.contentTypeToSaveFormat(String)
    //ExSummary:Shows how to find the corresponding Aspose load/save format from each media type string.
    // The ContentTypeToSaveFormat/ContentTypeToLoadFormat methods only accept official IANA media type names, also known as MIME types. 
    // All valid media types are listed here: https://www.iana.org/assignments/media-types/media-types.xhtml.

    // Trying to associate a SaveFormat with a partial media type string will not work.
    expect(() => aw.FileFormatUtil.contentTypeToSaveFormat("jpeg")).toThrow("Cannot convert this content type to a save format.");

    // If Aspose.words does not have a corresponding save/load format for a content type, an exception will also be thrown.
    expect(() => aw.FileFormatUtil.contentTypeToSaveFormat("application/zip")).toThrow("Cannot convert this content type to a save format.");

    // Files of the types listed below can be saved, but not loaded using Aspose.words.
    expect(() => aw.FileFormatUtil.contentTypeToLoadFormat("image/jpeg")).toThrow("Cannot convert this content type to a load format.");

    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/jpeg")).toEqual(aw.SaveFormat.Jpeg);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/png")).toEqual(aw.SaveFormat.Png);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/tiff")).toEqual(aw.SaveFormat.Tiff);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/gif")).toEqual(aw.SaveFormat.Gif);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/x-emf")).toEqual(aw.SaveFormat.Emf);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("application/vnd.ms-xpsdocument")).toEqual(aw.SaveFormat.Xps);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("application/pdf")).toEqual(aw.SaveFormat.Pdf);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("image/svg+xml")).toEqual(aw.SaveFormat.Svg);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("application/epub+zip")).toEqual(aw.SaveFormat.Epub);

    // For file types that can be saved and loaded, we can match a media type to both a load format and a save format.
    expect(aw.FileFormatUtil.contentTypeToLoadFormat("application/msword")).toEqual(aw.LoadFormat.Doc);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("application/msword")).toEqual(aw.SaveFormat.Doc);

    expect(aw.FileFormatUtil.contentTypeToLoadFormat( "application/vnd.openxmlformats-officedocument.wordprocessingml.document")).toEqual(aw.LoadFormat.Docx);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat( "application/vnd.openxmlformats-officedocument.wordprocessingml.document")).toEqual(aw.SaveFormat.Docx);

    expect(aw.FileFormatUtil.contentTypeToLoadFormat("text/plain")).toEqual(aw.LoadFormat.Text);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("text/plain")).toEqual(aw.SaveFormat.Text);

    expect(aw.FileFormatUtil.contentTypeToLoadFormat("application/rtf")).toEqual(aw.LoadFormat.Rtf);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("application/rtf")).toEqual(aw.SaveFormat.Rtf);

    expect(aw.FileFormatUtil.contentTypeToLoadFormat("text/html")).toEqual(aw.LoadFormat.Html);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("text/html")).toEqual(aw.SaveFormat.Html);

    expect(aw.FileFormatUtil.contentTypeToLoadFormat("multipart/related")).toEqual(aw.LoadFormat.Mhtml);
    expect(aw.FileFormatUtil.contentTypeToSaveFormat("multipart/related")).toEqual(aw.SaveFormat.Mhtml);
    //ExEnd
  });


  test('DetectDocumentEncryption', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.detectFileFormat(String)
    //ExFor:FileFormatInfo
    //ExFor:aw.FileFormatInfo.loadFormat
    //ExFor:aw.FileFormatInfo.isEncrypted
    //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and encryption.
    let doc = new aw.Document();

    // Configure a SaveOptions object to encrypt the document
    // with a password when we save it, and then save the document.
    let saveOptions = new aw.Saving.OdtSaveOptions(aw.SaveFormat.Odt);
    saveOptions.password = "MyPassword";

    doc.save(base.artifactsDir + "File.DetectDocumentEncryption.odt", saveOptions);

    // Verify the file type of our document, and its encryption status.
    let info = aw.FileFormatUtil.detectFileFormat(base.artifactsDir + "File.DetectDocumentEncryption.odt");

    expect(aw.FileFormatUtil.loadFormatToExtension(info.loadFormat)).toEqual(".odt");
    expect(info.isEncrypted).toEqual(true);
    //ExEnd
  });

/*
    [AotTests.IgnoreAot("CertificateHolder.Create and DigitalSignatureUtil.Sign are not used in AW.NET directly.")]
  test('AddDigitalSignature', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.detectFileFormat(String)
    //ExFor:FileFormatInfo
    //ExFor:aw.FileFormatInfo.loadFormat
    //ExFor:aw.FileFormatInfo.hasDigitalSignature
    //ExSummary:Shows how to add a digital signature to a document.
    string signedFile = base.myDir + "File.DetectDigitalSignatures.docx";
    if (fs.existsSync(signedFile))
      File.delete(signedFile);

    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Document.docx");

    expect(aw.FileFormatUtil.loadFormatToExtension(info.loadFormat)).toEqual(".docx");
    expect(info.hasDigitalSignature).toEqual(false);

    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw", null);
    aw.DigitalSignatures.DigitalSignatureUtil.sign(base.myDir + "Document.docx", signedFile,
      certificateHolder, new aw.DigitalSignatures.SignOptions() { SignTime = Date.now() });
    //ExEnd
  });
*/


  test('DetectDigitalSignatures', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.detectFileFormat(String)
    //ExFor:FileFormatInfo
    //ExFor:aw.FileFormatInfo.loadFormat
    //ExFor:aw.FileFormatInfo.hasDigitalSignature
    //ExSummary:Shows how to use the FileFormatUtil class to detect the document format and presence of digital signatures.
    // Use a FileFormatInfo instance to verify that a document is not digitally signed.
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Document.docx");

    expect(aw.FileFormatUtil.loadFormatToExtension(info.loadFormat)).toEqual(".docx");
    expect(info.hasDigitalSignature).toEqual(false);

    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw", null);
    let signOptions = new aw.DigitalSignatures.SignOptions();
    signOptions.signTime = Date.now();
    aw.DigitalSignatures.DigitalSignatureUtil.sign(base.myDir + "Document.docx", base.artifactsDir + "File.DetectDigitalSignatures.docx",
      certificateHolder, signOptions);

    // Use a new FileFormatInstance to confirm that it is signed.
    info = aw.FileFormatUtil.detectFileFormat(base.artifactsDir + "File.DetectDigitalSignatures.docx");

    expect(info.hasDigitalSignature).toEqual(true);

    // We can load and access the signatures of a signed document in a collection like this.
    expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "File.DetectDigitalSignatures.docx").count).toEqual(1);
    //ExEnd
  });


  test('SaveToDetectedFileFormat', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.detectFileFormat(Stream)
    //ExFor:aw.FileFormatUtil.loadFormatToExtension(LoadFormat)
    //ExFor:aw.FileFormatUtil.extensionToSaveFormat(String)
    //ExFor:aw.FileFormatUtil.saveFormatToExtension(SaveFormat)
    //ExFor:aw.FileFormatUtil.loadFormatToSaveFormat(LoadFormat)
    //ExFor:aw.Document.originalFileName
    //ExFor:aw.FileFormatInfo.loadFormat
    //ExFor:LoadFormat
    //ExSummary:Shows how to use the FileFormatUtil methods to detect the format of a document.
    // Load a document from a file that is missing a file extension, and then detect its file format.
    let docStream = base.loadFileToBuffer(base.myDir + "Word document with missing file extension");
    let info = aw.FileFormatUtil.detectFileFormat(docStream);
    let loadFormat = info.loadFormat;
    expect(loadFormat).toEqual(aw.LoadFormat.Doc);
    // Below are two methods of converting a LoadFormat to its corresponding SaveFormat.
    // 1 -  Get the file extension string for the LoadFormat, then get the corresponding SaveFormat from that string:
    let fileExtension = aw.FileFormatUtil.loadFormatToExtension(loadFormat);
    let saveFormat = aw.FileFormatUtil.extensionToSaveFormat(fileExtension);
    // 2 -  Convert the LoadFormat directly to its SaveFormat:
    saveFormat = aw.FileFormatUtil.loadFormatToSaveFormat(loadFormat);
    // Load a document from the stream, and then save it to the automatically detected file extension.
    let doc = new aw.Document(docStream);
    expect(aw.FileFormatUtil.saveFormatToExtension(saveFormat)).toEqual(".doc");
    doc.save(base.artifactsDir + "File.SaveToDetectedFileFormat" + aw.FileFormatUtil.saveFormatToExtension(saveFormat));
  //ExEnd
  });


  test('DetectFileFormat_SaveFormatToLoadFormat', () => {
    //ExStart
    //ExFor:aw.FileFormatUtil.saveFormatToLoadFormat(SaveFormat)
    //ExSummary:Shows how to convert a save format to its corresponding load format.
    expect(aw.FileFormatUtil.saveFormatToLoadFormat(aw.SaveFormat.Html)).toEqual(aw.LoadFormat.Html);

    // Some file types can have documents saved to, but not loaded from using Aspose.words.
    // If we attempt to convert a save format of such a type to a load format, an exception will be thrown.
    expect(() => aw.FileFormatUtil.saveFormatToLoadFormat(aw.SaveFormat.Jpeg)).toThrow("Cannot convert this save format to a load format.");
    //ExEnd
  });


  test('ExtractImages', () => {
    //ExStart
    //ExFor:Shape
    //ExFor:aw.Drawing.Shape.imageData
    //ExFor:aw.Drawing.Shape.hasImage
    //ExFor:ImageData
    //ExFor:aw.FileFormatUtil.imageTypeToExtension(ImageType)
    //ExFor:aw.Drawing.ImageData.imageType
    //ExFor:aw.Drawing.ImageData.save(String)
    //ExFor:aw.CompositeNode.getChildNodes(NodeType, bool)
    //ExSummary:Shows how to extract images from a document, and save them to the local file system as individual files.
    let doc = new aw.Document(base.myDir + "Images.docx");

    // Get the collection of shapes from the document,
    // and save the image data of every shape with an image as a file to the local file system.
    let nodes = [...doc.getChildNodes(aw.NodeType.Shape, true)];

    expect(nodes.filter(s => s.asShape().hasImage).length).toEqual(9);

    let imageIndex = 0;
    for (let node of nodes)
    {
      let shape = node.asShape();
      if (shape.hasImage)
      {
        // The image data of shapes may contain images of many possible image formats. 
        // We can determine a file extension for each image automatically, based on its format.
        let imageFileName =
          `File.ExtractImages.${imageIndex}${aw.FileFormatUtil.imageTypeToExtension(shape.imageData.imageType)}`;
        shape.imageData.save(base.artifactsDir + imageFileName);
        imageIndex++;
      }
    }
    //ExEnd
    const r = new RegExp("^.+\.(jpeg|png|emf|wmf)$");
    expect(fs.readdirSync(base.artifactsDir)
      .filter(s => r.test(s) && s.startsWith("File.ExtractImages")).length).toEqual(9);
  });
});
