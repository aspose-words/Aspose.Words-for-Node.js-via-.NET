// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("WorkingWithLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('UpdateDirtyFields', () => {
    //ExStart:UpdateDirtyFields
    //GistId:757cf7d3534a39730cf3290d418681ab
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.updateDirtyFields = true;

    let doc = new aw.Document(base.myDir + "Dirty field.docx", loadOptions);

    doc.save(base.artifactsDir + "WorkingWithLoadOptions.updateDirtyFields.docx");
    //ExEnd:UpdateDirtyFields
  });


  test('LoadEncryptedDocument', () => {
    //ExStart:LoadSaveEncryptedDocument
    //GistId:50a58d2d88c2177a9a4888b5d0e4de81
    //ExStart:OpenEncryptedDocument
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let doc = new aw.Document(base.myDir + "Encrypted.docx", new aw.Loading.LoadOptions("docPassword"));
    //ExEnd:OpenEncryptedDocument

    doc.save(base.artifactsDir + "WorkingWithLoadOptions.LoadSaveEncryptedDocument.odt", 
      new aw.Saving.OdtSaveOptions("newPassword"));
    //ExEnd:LoadSaveEncryptedDocument
  });


  test('LoadEncryptedDocumentWithoutPassword', () => {
    //ExStart:LoadEncryptedDocumentWithoutPassword
    //GistId:50a58d2d88c2177a9a4888b5d0e4de81
    // We will not be able to open this document with Microsoft Word or
    // Aspose.words without providing the correct password.
    expect(() => new aw.Document(base.myDir + "Encrypted.docx")).toThrow("The document password is incorrect.");
    //ExEnd:LoadEncryptedDocumentWithoutPassword
  });


  test('ConvertShapeToOfficeMath', () => {
    //ExStart:ConvertShapeToOfficeMath
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.convertShapeToOfficeMath = true;

    let doc = new aw.Document(base.myDir + "Office math.docx", loadOptions);

    doc.save(base.artifactsDir + "WorkingWithLoadOptions.convertShapeToOfficeMath.docx", aw.SaveFormat.Docx);
    //ExEnd:ConvertShapeToOfficeMath
  });


  test('SetMsWordVersion', () => {
    //ExStart:SetMsWordVersion
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
    // and change the loading version to Microsoft Word 2010.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.mswVersion = aw.Settings.MsWordVersion.Word2010;
            
    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);

    doc.save(base.artifactsDir + "WorkingWithLoadOptions.SetMsWordVersion.docx");
    //ExEnd:SetMsWordVersion
  });


  test('TempFolder', () => {
    //ExStart:TempFolder
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.tempFolder = base.artifactsDir;

    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);
    //ExEnd:TempFolder
  });


  test('WarningCallback', () => {
    //ExStart:WarningCallback
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let loadOptions = new aw.Loading.LoadOptions();
    // loadOptions.warningCallback = new DocumentLoadingWarningCallback(); TODO: not supported yet
            
    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);
    //ExEnd:WarningCallback
  });


/*  //ExStart:IWarningCallback
  //GistId:40be8275fc43f78f5e5877212e4e1bf3
  public class DocumentLoadingWarningCallback : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
        // Prints warnings and their details as they arise during document loading.
      console.log(`WARNING: ${info.warningType}, source: ${info.source}`);
      console.log(`\tDescription: ${info.description}`);
    }
  }
    //ExEnd:IWarningCallback
*/


  test('ResourceLoadingCallback', () => {
    //ExStart:ResourceLoadingCallback
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let loadOptions = new aw.Loading.LoadOptions();
    // loadOptions.resourceLoadingCallback = new HtmlLinkedResourceLoadingCallback(); TODO: not supported yet.

    // When we open an Html document, external resources such as references to CSS stylesheet files
    // and external images will be handled customarily by the loading callback as the document is loaded.
    let doc = new aw.Document(base.myDir + "Images.html", loadOptions);

    doc.save(base.artifactsDir + "WorkingWithLoadOptions.resourceLoadingCallback.pdf");
    //ExEnd:ResourceLoadingCallback
  });


/*
  //ExStart:IResourceLoadingCallback
  //GistId:40be8275fc43f78f5e5877212e4e1bf3
  private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
  {
    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
      switch (args.resourceType)
      {
        case aw.Loading.ResourceType.CssStyleSheet:
        {
          console.log(`External CSS Stylesheet found upon loading: ${args.originalUri}`);

            // CSS file will don't used in the document.
          return aw.Loading.ResourceLoadingAction.Skip;
        }
        case aw.Loading.ResourceType.Image:
        {
            // Replaces all images with a substitute.
          Image newImage = Image.FromFile(ImagesDir + "Logo.jpg");

          let converter = new ImageConverter();
          byte.at(] imageBytes = (byte[))converter.ConvertTo(newImage, typeof(byte[]));

          args.setData(imageBytes);

            // New images will be used instead of presented in the document.
          return aw.Loading.ResourceLoadingAction.UserProvided;
        }
        case aw.Loading.ResourceType.Document:
        {
          console.log(`External document found upon loading: ${args.originalUri}`);

            // Will be used as usual.
          return aw.Loading.ResourceLoadingAction.Default;
        }
        default:
          throw new InvalidOperationException("Unexpected ResourceType value.");
      }
    }
  }
  //ExEnd:IResourceLoadingCallback
*/


  test('LoadWithEncoding', () => {
    //ExStart:LoadWithEncoding
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let loadOptions = new aw.Loading.LoadOptions;
    loadOptions.encoding = "ascii";

    // Load the document while passing the LoadOptions object, then verify the document's contents.
    let doc = new aw.Document(base.myDir + "English text.txt", loadOptions);
    //ExEnd:LoadWithEncoding
  });


  test.skip('SkipPdfImages - TODO: Loading PDF not supported yet', () => {
    //ExStart:SkipPdfImages
    let loadOptions = new aw.Loading.PdfLoadOptions();
    loadOptions.skipPdfImages = true;

    let doc = new aw.Document(base.myDir + "Pdf Document.pdf", loadOptions);
    //ExEnd:SkipPdfImages
  });


  test('ConvertMetafilesToPng', () => {
    //ExStart:ConvertMetafilesToPng
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.convertMetafilesToPng = true;

    let doc = new aw.Document(base.myDir + "WMF with image.docx", loadOptions);
    //ExEnd:ConvertMetafilesToPng
  });


  test('LoadChm', () => {
    //ExStart:LoadCHM
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.encoding = "windows-1251";

    let doc = new aw.Document(base.myDir + "HTML help.chm", loadOptions);
    //ExEnd:LoadCHM
  });

});