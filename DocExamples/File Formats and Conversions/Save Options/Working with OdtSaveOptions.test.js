// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const MemoryStream = require('memorystream');


describe("WorkingWithOdtSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('MeasureUnit', () => {
    //ExStart:MeasureUnit
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
    // and content properties in documents whereas MS Office uses inches.
    let saveOptions = new aw.Saving.OdtSaveOptions();
    saveOptions.measureUnit = aw.Saving.OdtSaveMeasureUnit.Inches;

    doc.save(base.artifactsDir + "WorkingWithOdtSaveOptions.MeasureUnit.odt", saveOptions);
    //ExEnd:MeasureUnit
  });

});