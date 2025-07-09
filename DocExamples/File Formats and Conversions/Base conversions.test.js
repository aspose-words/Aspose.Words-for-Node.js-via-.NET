// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;
const MemoryStream = require('memorystream');
const fs = require('fs');


describe("BaseConversions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('DocToDocx', () => {
    //ExStart:LoadAndSave
    //GistId:757cf7d3534a39730cf3290d418681ab
    //ExStart:OpenDocument
    let doc = new aw.Document(base.myDir + "Document.doc");
    //ExEnd:OpenDocument

    doc.save(base.artifactsDir + "BaseConversions.DocToDocx.docx");
    //ExEnd:LoadAndSave
  });


  test('DocxToRtf', async () => {
    //ExStart:LoadAndSaveToStream
    //GistId:757cf7d3534a39730cf3290d418681ab
    //ExStart:OpenFromStream
    //GistId:96e42cb4a611465927f8e7b1b3d546d3
    // Read only access is enough for Aspose.Words to load a document.
    
    let stream = base.loadFileToBuffer(base.myDir + "Document.docx");
    let doc = new aw.Document(stream);
    //ExEnd:OpenFromStream

    // ... do something with the document.

    // Convert the document to a different format and save to stream.
    const dstFile = base.artifactsDir + "BaseConversions.DocxToRtf.rtf";
    const dstStream = fs.createWriteStream(dstFile);
    doc.save(dstStream, aw.SaveFormat.Rtf);
    await new Promise(resolve => dstStream.on("finish", resolve));

    //ExEnd:LoadAndSaveToStream
  });


  test('DocxToPdf', () => {
    //ExStart:DocxToPdf
    //GistId:38c6608baa855f951a4e117a721bdaae
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.save(base.artifactsDir + "BaseConversions.DocxToPdf.pdf");
    //ExEnd:DocxToPdf
  });


  test('DocxToByte', () => {
    //ExStart:DocxToByte
    //GistId:04839bd5cc7e85e20e0e239b4dcdead3
    let doc = new aw.Document(base.myDir + "Document.docx");

    let outStream = new MemoryStream();

    doc.save(outStream, aw.SaveFormat.Docx);

    //ExEnd:DocxToByte
  });


  test('DocxToEpub', () => {
    //ExStart:DocxToEpub
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.save(base.artifactsDir + "BaseConversions.DocxToEpub.epub");
    //ExEnd:DocxToEpub
  });


  test('DocxToHtml', () => {
    //ExStart:DocxToHtml
    //GistId:e4b272992a7c8fafdd7ff42f8c2de379
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.save(base.artifactsDir + "BaseConversions.DocxToHtml.html");
    //ExEnd:DocxToHtml
  });


  test('DocxToMhtml', () => {
    //ExStart:DocxToMhtml
    //GistId:8e5415d4997270432815aec11112dd8d
    let doc = new aw.Document(base.myDir + "Document.docx");

    let stream = new MemoryStream();
    doc.save(stream, aw.SaveFormat.Mhtml);

    //ExEnd:DocxToMhtml
  });


  test('DocxToMarkdown', () => {
    //ExStart:DocxToMarkdown
    //GistId:a2fee7fa3d8e5704ce24f041be9a4821
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Some text!");

    doc.save(base.artifactsDir + "BaseConversions.DocxToMarkdown.md");
    //ExEnd:DocxToMarkdown
  });


  test('DocxToTxt', () => {
    //ExStart:DocxToTxt
    //GistId:433f5122fe18fdc24a406528b70b0020
    let doc = new aw.Document(base.myDir + "Document.docx");
    doc.save(base.artifactsDir + "BaseConversions.DocxToTxt.txt");
    //ExEnd:DocxToTxt
  });


  test('DocxToXlsx', () => {
    //ExStart:DocxToXlsx
    //GistId:a2974bc90e2ef0579ddce59482175c52
    let doc = new aw.Document(base.myDir + "Document.docx");
    doc.save(base.artifactsDir + "BaseConversions.DocxToXlsx.xlsx");
    //ExEnd:DocxToXlsx
  });

  test('DocxToJpeg', () => {
    //ExStart:DocxToJpeg
    //GistId:05b9bb6f4d96094b4408287596e99a20
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);
    doc.save(base.artifactsDir + "BaseConversions.DocxToJpeg.jpeg", saveOptions);
    //ExEnd:DocxToJpeg
  });


  test('TxtToDocx', () => {
    //ExStart:TxtToDocx
    // The encoding of the text file is automatically detected.
    let doc = new aw.Document(base.myDir + "English text.txt");

    doc.save(base.artifactsDir + "BaseConversions.TxtToDocx.docx");
    //ExEnd:TxtToDocx
  });


  test.skip('PdfToJpeg - TODO: Loading PDF not supported jet', () => {
    //ExStart:PdfToJpeg
    //GistId:ebbb90d74ef57db456685052a18f8e86
    let doc = new aw.Document(base.myDir + "Pdf Document.pdf");

    doc.save(base.artifactsDir + "BaseConversions.PdfToJpeg.jpeg");
    //ExEnd:PdfToJpeg
  });


  test.skip('PdfToDocx - TODO: Loading PDF not supported jet', () => {
    //ExStart:PdfToDocx
    //GistId:a0d52b62c1643faa76a465a41537edfc
    let doc = new aw.Document(base.myDir + "Pdf Document.pdf");

    doc.save(base.artifactsDir + "BaseConversions.PdfToDocx.docx");
    //ExEnd:PdfToDocx
  });


  test.skip('PdfToXlsx - TODO: Loading PDF not supported jet', () => {
    //ExStart:PdfToXlsx
    //GistId:a50652f28531278511605e0fd778bbdf
    let doc = new aw.Document(base.myDir + "Pdf Document.pdf");

    doc.save(base.artifactsDir + "BaseConversions.PdfToXlsx.xlsx");
    //ExEnd:PdfToXlsx
  });


  test('FindReplaceXlsx', () => {
    //ExStart:FindReplaceXlsx
    //GistId:41ba1b5ed95eda5f5dbf8297fc4b5bf0
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Ruby bought a ruby necklace.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
    // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
    options.matchCase = true;

    doc.range.replace("Ruby", "Jade", options);

    doc.save(base.artifactsDir + "BaseConversions.FindReplaceXlsx.xlsx");
    //ExEnd:FindReplaceXlsx
  });


  test('CompressXlsx', () => {
    //ExStart:CompressXlsx
    //GistId:41ba1b5ed95eda5f5dbf8297fc4b5bf0
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.XlsxSaveOptions();
    saveOptions.compressionLevel = aw.Saving.CompressionLevel.Maximum;

    doc.save(base.artifactsDir + "BaseConversions.CompressXlsx.xlsx", saveOptions);
    //ExEnd:CompressXlsx
  });


  test('ImagesToPdf', () => {
    //ExStart:ImageToPdf
    //GistId:38c6608baa855f951a4e117a721bdaae
    convertImageToPdf(base.imagesDir + "Logo.jpg", base.artifactsDir + "BaseConversions.JpgToPdf.pdf");
    convertImageToPdf(base.imagesDir + "Transparent background logo.png", base.artifactsDir + "BaseConversions.PngToPdf.pdf");
    convertImageToPdf(base.imagesDir + "Windows MetaFile.wmf", base.artifactsDir + "BaseConversions.WmfToPdf.pdf");
    convertImageToPdf(base.imagesDir + "Tagged Image File Format.tiff", base.artifactsDir + "BaseConversions.TiffToPdf.pdf");
    convertImageToPdf(base.imagesDir + "Graphics Interchange Format.gif", base.artifactsDir + "BaseConversions.GifToPdf.pdf");
    //ExEnd:ImageToPdf
  });


  //ExStart:ConvertImageToPdf
  //GistId:38c6608baa855f951a4e117a721bdaae
  /// <summary>
  /// Converts an image to PDF using Aspose.Words for .NET.
  /// </summary>
  /// <param name="inputFileName">File name of input image file.</param>
  /// <param name="outputFileName">Output PDF file name.</param>
  function convertImageToPdf(inputFileName, outputFileName) {
    console.log("Converting " + inputFileName + " to PDF ....");

    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert the image into the document and position it at the top left corner of the page.
    builder.insertImage(inputFileName);

    doc.save(outputFileName);            
  };
  //ExEnd:ConvertImageToPdf

});
