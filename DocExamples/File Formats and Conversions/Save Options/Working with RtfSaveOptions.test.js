// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const MemoryStream = require('memorystream');


describe("WorkingWithRtfSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('SavingImagesAsWmf', () => {
    //ExStart:SavingImagesAsWmf
    //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.RtfSaveOptions();
    saveOptions.saveImagesAsWmf = true;

    doc.save(base.artifactsDir + "WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
    //ExEnd:SavingImagesAsWmf
  });

});