// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const fs = require('fs');

describe("ExDocSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  beforeEach(() => {
    base.setUnlimitedLicense();
  });

  test('SaveAsDoc', () => {
    //ExStart
    //ExFor:DocSaveOptions
    //ExFor:DocSaveOptions.#ctor
    //ExFor:DocSaveOptions.#ctor(SaveFormat)
    //ExFor:aw.Saving.DocSaveOptions.password
    //ExFor:aw.Saving.DocSaveOptions.saveFormat
    //ExFor:aw.Saving.DocSaveOptions.saveRoutingSlip
    //ExSummary:Shows how to set save options for older Microsoft Word formats.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Hello world!");

    let options = new aw.Saving.DocSaveOptions(aw.SaveFormat.Doc);
            
    // Set a password which will protect the loading of the document by Microsoft Word or Aspose.words.
    // Note that this does not encrypt the contents of the document in any way.
    options.password = "MyPassword";

    // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
    options.saveRoutingSlip = true;

    doc.save(base.artifactsDir + "DocSaveOptions.SaveAsDoc.doc", options);

    // To be able to load the document,
    // we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
    expect(() => doc = new aw.Document(base.artifactsDir + "DocSaveOptions.SaveAsDoc.doc")).toThrow("The document password is incorrect.");

    let loadOptions = new aw.Loading.LoadOptions("MyPassword");
    doc = new aw.Document(base.artifactsDir + "DocSaveOptions.SaveAsDoc.doc", loadOptions);

    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd
  });

  test('TempFolder', () => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.tempFolder
    //ExSummary:Shows how to use the hard drive instead of memory when saving a document.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
    // We can use this option to use a temporary folder in the local file system instead,
    // which will reduce our application's memory overhead.
    let options = new aw.Saving.DocSaveOptions();
    options.tempFolder = base.artifactsDir + "TempFiles";

    // The specified temporary folder must exist in the local file system before the save operation.
    if (!fs.existsSync(options.tempFolder))
      fs.mkdirSync(options.tempFolder);

    doc.save(base.artifactsDir + "DocSaveOptions.tempFolder.doc", options);

    // The folder will persist with no residual contents from the load operation.
    const fileList = fs.readdirSync(options.tempFolder);    
    expect(fileList.length).toEqual(0);
    //ExEnd
  });

  test('PictureBullets', () => {
    //ExStart
    //ExFor:aw.Saving.DocSaveOptions.savePictureBullet
    //ExSummary:Shows how to omit PictureBullet data from the document when saving.
    let doc = new aw.Document(base.myDir + "Image bullet points.docx");
    expect(doc.lists.at(0).listLevels.at(0).imageData).not.toBe(null);

    // Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
    // By setting a flag in the SaveOptions object,
    // we can convert all image bullet points to ordinary bullet points while saving.
    let saveOptions = new aw.Saving.DocSaveOptions(aw.SaveFormat.Doc);
    saveOptions.savePictureBullet = false;

    doc.save(base.artifactsDir + "DocSaveOptions.PictureBullets.doc", saveOptions);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocSaveOptions.PictureBullets.doc");

    expect(doc.lists.at(0).listLevels.at(0).imageData).toBe(null);
  });

  test.each([true,
    false])('UpdateLastPrintedProperty', (isUpdateLastPrintedProperty) => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.updateLastPrintedProperty
    //ExSummary:Shows how to update a document's "Last printed" property when saving.
    let doc = new aw.Document();
    doc.builtInDocumentProperties.lastPrinted =new Date(2019, 12, 20);

    // This flag determines whether the last printed date, which is a built-in property, is updated.
    // If so, then the date of the document's most recent save operation
    // with this SaveOptions object passed as a parameter is used as the print date.
    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.updateLastPrintedProperty = isUpdateLastPrintedProperty;

    // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
    // It can also be displayed in the document's body by using a PRINTDATE field.
    doc.save(base.artifactsDir + "DocSaveOptions.updateLastPrintedProperty.doc", saveOptions);

    // Open the saved document, then verify the value of the property.
    doc = new aw.Document(base.artifactsDir + "DocSaveOptions.updateLastPrintedProperty.doc");

    if (isUpdateLastPrintedProperty)
      expect(doc.builtInDocumentProperties.lastPrinted).not.toEqual(new Date(2019, 12, 20));
    else
      expect(doc.builtInDocumentProperties.lastPrinted).toEqual(new Date(2019, 12, 20));
    //ExEnd
  });

  test.each([true,
    false])('UpdateCreatedTimeProperty', (isUpdateCreatedTimeProperty) => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.updateLastPrintedProperty
    //ExSummary:Shows how to update a document's "CreatedTime" property when saving.
    let doc = new aw.Document();
    doc.builtInDocumentProperties.createdTime = new Date(2019, 12, 20);

    // This flag determines whether the created time, which is a built-in property, is updated.
    // If so, then the date of the document's most recent save operation
    // with this SaveOptions object passed as a parameter is used as the created time.
    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.updateCreatedTimeProperty = isUpdateCreatedTimeProperty;

    doc.save(base.artifactsDir + "DocSaveOptions.updateCreatedTimeProperty.docx", saveOptions);

    // Open the saved document, then verify the value of the property.
    doc = new aw.Document(base.artifactsDir + "DocSaveOptions.updateCreatedTimeProperty.docx");

    if (isUpdateCreatedTimeProperty)
      expect(doc.builtInDocumentProperties.createdTime).not.toEqual(new Date(2019, 12, 20));
    else
      expect(doc.builtInDocumentProperties.createdTime).toEqual(new Date(2019, 12, 20));
    //ExEnd
  });

  test.each([false,
    true])('AlwaysCompressMetafiles', (compressAllMetafiles) => {
    //ExStart
    //ExFor:aw.Saving.DocSaveOptions.alwaysCompressMetafiles
    //ExSummary:Shows how to change metafiles compression in a document while saving.
    // Open a document that contains a Microsoft Equation 3.0 formula.
    let doc = new aw.Document(base.myDir + "Microsoft equation object.docx");

    // When we save a document, smaller metafiles are not compressed for performance reasons.
    // We can set a flag in a SaveOptions object to compress every metafile when saving.
    // Some editors such as LibreOffice cannot read uncompressed metafiles.
    let saveOptions = new aw.Saving.DocSaveOptions();
    saveOptions.alwaysCompressMetafiles = compressAllMetafiles;

    doc.save(base.artifactsDir + "DocSaveOptions.alwaysCompressMetafiles.docx", saveOptions);
    //ExEnd

    var testedFileLength = fs.statSync(base.artifactsDir + "DocSaveOptions.alwaysCompressMetafiles.docx").size;

    if (compressAllMetafiles)
      expect(testedFileLength).toBeLessThan(14000);
    else
      expect(testedFileLength).toBeLessThan(22000);            
  });
});
