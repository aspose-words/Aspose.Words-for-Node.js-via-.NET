// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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


describe("ExXlsxSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CompressXlsx', () => {
    //ExStart
    //ExFor:XlsxSaveOptions
    //ExFor:XlsxSaveOptions.compressionLevel
    //ExFor:XlsxSaveOptions.saveFormat
    //ExSummary:Shows how to compress XLSX document.
    let doc = new aw.Document(base.myDir + "Shape with linked chart.docx");

    let xlsxSaveOptions = new aw.Saving.XlsxSaveOptions();
    xlsxSaveOptions.compressionLevel = aw.Saving.CompressionLevel.Maximum;
    xlsxSaveOptions.saveFormat = aw.SaveFormat.Xlsx;

    doc.save(base.artifactsDir + "XlsxSaveOptions.CompressXlsx.xlsx", xlsxSaveOptions);
    //ExEnd
  });


  test('SelectionMode', () => {
    //ExStart:SelectionMode
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:XlsxSaveOptions.sectionMode
    //ExFor:XlsxSectionMode
    //ExSummary:Shows how to save document as a separate worksheets.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // Each section of a document will be created as a separate worksheet.
    // Use 'SingleWorksheet' to display all document on one worksheet.
    let xlsxSaveOptions = new aw.Saving.XlsxSaveOptions();
    xlsxSaveOptions.sectionMode = aw.Saving.XlsxSectionMode.MultipleWorksheets;

    doc.save(base.artifactsDir + "XlsxSaveOptions.SelectionMode.xlsx", xlsxSaveOptions);
    //ExEnd:SelectionMode
  });


  test('DateTimeParsingMode', () => {
    //ExStart:DateTimeParsingMode
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:XlsxSaveOptions.dateTimeParsingMode
    //ExFor:XlsxDateTimeParsingMode
    //ExSummary:Shows how to specify autodetection of the date time format.
    let doc = new aw.Document(base.myDir + "Xlsx DateTime.docx");

    let saveOptions = new aw.Saving.XlsxSaveOptions();
    // Specify using datetime format autodetection.
    saveOptions.dateTimeParsingMode = aw.Saving.XlsxDateTimeParsingMode.Auto;

    doc.save(base.artifactsDir + "XlsxSaveOptions.dateTimeParsingMode.xlsx", saveOptions);
    //ExEnd:DateTimeParsingMode
  });
});
