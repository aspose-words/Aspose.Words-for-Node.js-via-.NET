// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');


describe("WorkingWithDocSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('EncryptDocumentWithPassword', () => {
    //ExStart:EncryptDocumentWithPassword
    //GistId:50a58d2d88c2177a9a4888b5d0e4de81
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    builder.write("Hello world!");

    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.password = "password";

    doc.save(base.artifactsDir + "WorkingWithDocSaveOptions.EncryptDocumentWithPassword.docx", saveOptions);
    //ExEnd:EncryptDocumentWithPassword
  });


  test('DoNotCompressSmallMetafiles', () => {
    //ExStart:DoNotCompressSmallMetafiles
    let doc = new aw.Document(base.myDir + "Microsoft equation object.docx");

    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.alwaysCompressMetafiles = false;

    doc.save(base.artifactsDir + "WorkingWithDocSaveOptions.NotCompressSmallMetafiles.docx", saveOptions);
    //ExEnd:DoNotCompressSmallMetafiles
  });


  test('DoNotSavePictureBullet', () => {
    //ExStart:DoNotSavePictureBullet
    let doc = new aw.Document(base.myDir + "Image bullet points.docx");

    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.savePictureBullet = false;

    doc.save(base.artifactsDir + "WorkingWithDocSaveOptions.DoNotSavePictureBullet.docx", saveOptions);
    //ExEnd:DoNotSavePictureBullet
  });

});
