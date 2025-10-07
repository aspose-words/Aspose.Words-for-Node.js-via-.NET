// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithOoxmlSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('EncryptDocxWithPassword', () => {
    //ExStart:EncryptDocxWithPassword
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.password = "password";

    doc.save(base.artifactsDir + "WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
    //ExEnd:EncryptDocxWithPassword
  });


  test('OoxmlComplianceIso29500_2008_Strict', () => {
    //ExStart:OoxmlComplianceIso29500_2008_Strict
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2016);
            
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;

    doc.save(base.artifactsDir + "WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
    //ExEnd:OoxmlComplianceIso29500_2008_Strict
  });


  test('UpdateLastSavedTime', () => {
    //ExStart:UpdateLastSavedTime
    //GistId:03144d2d1bfafb75c89d385616fdf674
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.updateLastSavedTimeProperty = true;

    doc.save(base.artifactsDir + "WorkingWithOoxmlSaveOptions.UpdateLastSavedTime.docx", saveOptions);
    //ExEnd:UpdateLastSavedTime
  });


  test('KeepLegacyControlChars', () => {
    //ExStart:KeepLegacyControlChars
    let doc = new aw.Document(base.myDir + "Legacy control character.doc");

    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.FlatOpc);
    saveOptions.keepLegacyControlChars = true;

    doc.save(base.artifactsDir + "WorkingWithOoxmlSaveOptions.keepLegacyControlChars.docx", saveOptions);
    //ExEnd:KeepLegacyControlChars
  });


  test('SetCompressionLevel', () => {
    //ExStart:SetCompressionLevel
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compressionLevel = aw.Saving.CompressionLevel.SuperFast;

    doc.save(base.artifactsDir + "WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
    //ExEnd:SetCompressionLevel
  });

});
