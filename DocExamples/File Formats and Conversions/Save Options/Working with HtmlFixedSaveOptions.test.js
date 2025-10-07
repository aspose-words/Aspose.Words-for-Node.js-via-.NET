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


describe("WorkingWithHtmlFixedSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('UseFontFromTargetMachine', () => {
    //ExStart:UseFontFromTargetMachine
    let doc = new aw.Document(base.myDir + "Bullet points with alternative font.docx");

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.useTargetMachineFonts = true;

    doc.save(base.artifactsDir + "WorkingWithHtmlFixedSaveOptions.UseFontFromTargetMachine.html", saveOptions);
    //ExEnd:UseFontFromTargetMachine
  });

  test('WriteAllCssRulesInSingleFile', () => {
    //ExStart:WriteAllCssRulesInSingleFile
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
    // All CSS rules are written into single file "styles.css.
    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.saveFontFaceCssSeparately = false;

    doc.save(base.artifactsDir + "WorkingWithHtmlFixedSaveOptions.WriteAllCssRulesInSingleFile.html", saveOptions);
    //ExEnd:WriteAllCssRulesInSingleFile
  });

});
