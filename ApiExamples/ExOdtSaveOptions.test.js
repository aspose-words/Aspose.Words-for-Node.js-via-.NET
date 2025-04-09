// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');

describe("ExOdtSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.each([false, true])('Odt11Schema(%o)', (exportToOdt11Specs) => {
    //ExStart
    //ExFor:OdtSaveOptions
    //ExFor:OdtSaveOptions.#ctor
    //ExFor:OdtSaveOptions.isStrictSchema11
    //ExFor:RevisionOptions.measurementUnit
    //ExFor:MeasurementUnits
    //ExSummary:Shows how to make a saved document conform to an older ODT schema.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.OdtSaveOptions();
    saveOptions.measureUnit = aw.Saving.OdtSaveMeasureUnit.Centimeters;
    saveOptions.isStrictSchema11 = exportToOdt11Specs;

    doc.save(base.artifactsDir + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "OdtSaveOptions.Odt11Schema.odt");

    expect(doc.layoutOptions.revisionOptions.measurementUnit).toEqual(aw.MeasurementUnits.Centimeters);

    if (exportToOdt11Specs)
    {
      expect(doc.range.formFields.count).toEqual(2);
      expect(doc.range.formFields.at(0).type).toEqual(aw.Fields.FieldType.FieldFormTextInput);
      expect(doc.range.formFields.at(1).type).toEqual(aw.Fields.FieldType.FieldFormCheckBox);
    }
    else
    {
      expect(doc.range.formFields.count).toEqual(3);
      expect(doc.range.formFields.at(0).type).toEqual(aw.Fields.FieldType.FieldFormTextInput);
      expect(doc.range.formFields.at(1).type).toEqual(aw.Fields.FieldType.FieldFormCheckBox);
      expect(doc.range.formFields.at(2).type).toEqual(aw.Fields.FieldType.FieldFormDropDown);
    }
  });


  test.each([aw.Saving.OdtSaveMeasureUnit.Centimeters,
    aw.Saving.OdtSaveMeasureUnit.Inches])('MeasurementUnits(%o)', (odtSaveMeasureUnit) => {
    //ExStart
    //ExFor:OdtSaveOptions
    //ExFor:OdtSaveOptions.measureUnit
    //ExFor:OdtSaveMeasureUnit
    //ExSummary:Shows how to use different measurement units to define style parameters of a saved ODT document.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // When we export the document to .odt, we can use an OdtSaveOptions object to modify how we save the document.
    // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Centimeters"
    // to define content such as style parameters using the metric system, which Open Office uses. 
    // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Inches"
    // to define content such as style parameters using the imperial system, which Microsoft Word uses.
    let saveOptions = new aw.Saving.OdtSaveOptions();
    saveOptions.measureUnit = odtSaveMeasureUnit;

    doc.save(base.artifactsDir + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
    //ExEnd

    switch (odtSaveMeasureUnit)
    {
      case aw.Saving.OdtSaveMeasureUnit.Centimeters:
        TestUtil.docPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"1.27cm\" />",
          base.artifactsDir + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
        break;
      case aw.Saving.OdtSaveMeasureUnit.Inches:
        TestUtil.docPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"0.5in\" />",
          base.artifactsDir + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
        break;
    }
  });


  test.each([aw.SaveFormat.Odt, aw.SaveFormat.Ott])('Encrypt(%o)', (saveFormat) => {
    //ExStart
    //ExFor:OdtSaveOptions.#ctor(SaveFormat)
    //ExFor:OdtSaveOptions.password
    //ExFor:OdtSaveOptions.saveFormat
    //ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.words.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Create a new OdtSaveOptions, and pass either "SaveFormat.Odt",
    // or "SaveFormat.Ott" as the format to save the document in. 
    let saveOptions = new aw.Saving.OdtSaveOptions(saveFormat);
    saveOptions.password = "@sposeEncrypted_1145";

    let extensionString = aw.FileFormatUtil.saveFormatToExtension(saveFormat);

    // If we open this document with an appropriate editor,
    // it will prompt us for the password we specified in the SaveOptions object.
    doc.save(base.artifactsDir + "OdtSaveOptions.encrypt" + extensionString, saveOptions);

    let docInfo = aw.FileFormatUtil.detectFileFormat(base.artifactsDir + "OdtSaveOptions.encrypt" + extensionString);

    expect(docInfo.isEncrypted).toEqual(true);

    // If we wish to open or edit this document again using Aspose.words,
    // we will have to provide a LoadOptions object with the correct password to the loading constructor.
    doc = new aw.Document(base.artifactsDir + "OdtSaveOptions.encrypt" + extensionString,
      new aw.Loading.LoadOptions("@sposeEncrypted_1145"));

    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd
  });

});
