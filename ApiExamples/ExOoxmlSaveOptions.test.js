// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const TestUtil = require('./TestUtil');
const jimp = require("jimp");

describe("ExOoxmlSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Password', () => {
    //ExStart
    //ExFor:aw.Saving.OoxmlSaveOptions.password
    //ExSummary:Shows how to create a password encrypted Office Open XML document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.password = "MyPassword";

    doc.save(base.artifactsDir + "OoxmlSaveOptions.password.docx", saveOptions);

    // We will not be able to open this document with Microsoft Word or
    // Aspose.words without providing the correct password.
    //Assert.Throws<IncorrectPasswordException>(() => doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.password.docx"));
    expect(() => { doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.password.docx"); }).toThrow();

    // Open the encrypted document by passing the correct password in a LoadOptions object.
    doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.password.docx", new aw.Loading.LoadOptions("MyPassword"));

    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('Iso29500Strict', () => {
    //ExStart
    //ExFor:CompatibilityOptions
    //ExFor:aw.Settings.CompatibilityOptions.optimizeFor(MsWordVersion)
    //ExFor:OoxmlSaveOptions
    //ExFor:OoxmlSaveOptions.#ctor
    //ExFor:aw.Saving.OoxmlSaveOptions.saveFormat
    //ExFor:OoxmlCompliance
    //ExFor:aw.Saving.OoxmlSaveOptions.compliance
    //ExFor:ShapeMarkupLanguage
    //ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If we configure compatibility options to comply with Microsoft Word 2003,
    // inserting an image will define its shape using VML.
    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2003);
    builder.insertImage(base.imageDir + "Transparent background logo.png");

    expect(doc.getShape(0, true).markupLanguage).toEqual(aw.Drawing.ShapeMarkupLanguage.Vml);

    // The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
    // If we set the "Compliance" property of the SaveOptions object to "OoxmlCompliance.Iso29500_2008_Strict",
    // any document we save while passing this object will have to follow that standard. 
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;
    saveOptions.saveFormat = aw.SaveFormat.Docx;

    doc.save(base.artifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

    // Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
    doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx");

    expect(doc.getShape(0, true).markupLanguage).toEqual(aw.Drawing.ShapeMarkupLanguage.Dml);
    //ExEnd
  });


  test.each([false, true])('RestartingDocumentList(%o)', (restartListAtEachSection) => {
    //ExStart
    //ExFor:aw.Lists.List.isRestartAtEachSection
    //ExFor:OoxmlCompliance
    //ExFor:aw.Saving.OoxmlSaveOptions.compliance
    //ExSummary:Shows how to configure a list to restart numbering at each section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    doc.lists.add(aw.Lists.ListTemplate.NumberDefault);

    let list = doc.lists.at(0);
    list.isRestartAtEachSection = restartListAtEachSection;

    // The "IsRestartAtEachSection" property will only be applicable when
    // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
    let options = new aw.Saving.OoxmlSaveOptions();
    options.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;

    builder.listFormat.list = list;

    builder.writeln("List item 1");
    builder.writeln("List item 2");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.writeln("List item 3");
    builder.writeln("List item 4");

    doc.save(base.artifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx", options);

    doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx");

    expect(doc.lists.at(0).isRestartAtEachSection).toEqual(restartListAtEachSection);
    //ExEnd
  });


  test.each([false, true])('LastSavedTime(%o)', (updateLastSavedTimeProperty) => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.updateLastSavedTimeProperty
    //ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
    let doc = new aw.Document(base.myDir + "Document.docx");

    expect(doc.builtInDocumentProperties.lastSavedTime).toEqual(new Date(2021, 5 - 1, 11, 6, 32, 0));

    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "UpdateLastSavedTimeProperty" property to "true" to
    // set the output document's "Last saved time" built-in property to the current date/time.
    // Set the "UpdateLastSavedTimeProperty" property to "false" to
    // preserve the original value of the input document's "Last saved time" built-in property.
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.updateLastSavedTimeProperty = updateLastSavedTimeProperty;

    doc.save(base.artifactsDir + "OoxmlSaveOptions.lastSavedTime.docx", saveOptions);

    doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.lastSavedTime.docx");
    let lastSavedTimeNew = doc.builtInDocumentProperties.lastSavedTime;

    if (updateLastSavedTimeProperty)
      expect(lastSavedTimeNew.getUTCDate()).toEqual(new Date().getUTCDate())
    else
      expect(lastSavedTimeNew).toEqual(new Date(2021, 5 - 1, 11, 6, 32, 0));
    //ExEnd
  });


  test.skip.each([false, true])('KeepLegacyControlChars(%o) - TODO: WORDSNODEJS-80', (keepLegacyControlChars) => {
    //ExStart
    //ExFor:aw.Saving.OoxmlSaveOptions.keepLegacyControlChars
    //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
    //ExSummary:Shows how to support legacy control characters when converting to .docx.
    let doc = new aw.Document(base.myDir + "Legacy control character.doc");

    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "KeepLegacyControlChars" property to "true" to preserve
    // the "ShortDateTime" legacy character while saving.
    // Set the "KeepLegacyControlChars" property to "false" to remove
    // the "ShortDateTime" legacy character from the output document.
    let so = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    so.keepLegacyControlChars = keepLegacyControlChars;
 
    doc.save(base.artifactsDir + "OoxmlSaveOptions.keepLegacyControlChars.docx", so);

    doc = new aw.Document(base.artifactsDir + "OoxmlSaveOptions.keepLegacyControlChars.docx");

    expect(doc.firstSection.body.getText()).toEqual(keepLegacyControlChars ? "\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f" : "\u001e\f");
    //ExEnd
  });


  test.each([aw.Saving.CompressionLevel.Maximum,
    aw.Saving.CompressionLevel.Fast,
    aw.Saving.CompressionLevel.Normal,
    aw.Saving.CompressionLevel.SuperFast])('DocumentCompression(%o)', (compressionLevel) => {
    //ExStart
    //ExFor:aw.Saving.OoxmlSaveOptions.compressionLevel
    //ExFor:CompressionLevel
    //ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
    // and then pass it to the document's saving method to modify how we save the document.
    // Set the "CompressionLevel" property to "CompressionLevel.Maximum" to apply the strongest and slowest compression.
    // Set the "CompressionLevel" property to "CompressionLevel.Normal" to apply
    // the default compression that Aspose.words uses while saving OOXML documents.
    // Set the "CompressionLevel" property to "CompressionLevel.Fast" to apply a faster and weaker compression.
    // Set the "CompressionLevel" property to "CompressionLevel.SuperFast" to apply
    // the default compression that Microsoft Word uses.
    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);
    saveOptions.compressionLevel = compressionLevel;

    let timeBefore = new Date();
    doc.save(base.artifactsDir + "OoxmlSaveOptions.DocumentCompression.docx", saveOptions);
    let timeAfter = new Date();

    let testedFileLength = fs.statSync(base.artifactsDir + "OoxmlSaveOptions.DocumentCompression.docx").size;

    console.log(`Saving operation done using the \"${compressionLevel}\" compression level:`);
    console.log(`\tDuration:\t${timeAfter - timeBefore} ms`);
    console.log(`\tFile Size:\t${testedFileLength} bytes`);
    //ExEnd

    switch (compressionLevel)
    {
      case aw.Saving.CompressionLevel.Maximum:
        expect(testedFileLength < 1269000).toEqual(true);
        break;
      case aw.Saving.CompressionLevel.Normal:
        expect(testedFileLength < 1271000).toEqual(true);
        break;
      case aw.Saving.CompressionLevel.Fast:
        expect(testedFileLength < 1280000).toEqual(true);
        break;
      case aw.Saving.CompressionLevel.SuperFast:
        expect(testedFileLength < 1276000).toEqual(true);
        break;
    }
  });


  test('CheckFileSignatures', () => {
    let compressionLevels = [
      aw.Saving.CompressionLevel.Maximum,
      aw.Saving.CompressionLevel.Normal,
      aw.Saving.CompressionLevel.Fast,
      aw.Saving.CompressionLevel.SuperFast
    ];

    let fileSignatures = [
      "50 4B 03 04 14 00 02 00 08 00 ",
      "50 4B 03 04 14 00 00 00 08 00 ",
      "50 4B 03 04 14 00 04 00 08 00 ",
      "50 4B 03 04 14 00 06 00 08 00 "
    ];

    let doc = new aw.Document();
    let saveOptions = new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx);

    let prevFileSize = 0;
    for (let i = 0; i < fileSignatures.length; ++i)
    {
      saveOptions.compressionLevel = compressionLevels.at(i);
      doc.save(base.artifactsDir + "OoxmlSaveOptions.CheckFileSignatures.docx", saveOptions);

      let data = Array.from(fs.readFileSync(base.artifactsDir + "OoxmlSaveOptions.CheckFileSignatures.docx"));
      expect(prevFileSize < data.length).toEqual(true);
      expect(TestUtil.dumpArray(data, 0, 10)).toEqual(fileSignatures.at(i));
      prevFileSize = data.length;
    }
  });

  test('ExportGeneratorName', () => {
    //ExStart
    //ExFor:aw.Saving.SaveOptions.exportGeneratorName
    //ExSummary:Shows how to disable adding name and version of Aspose.words into produced files.
    let doc = new aw.Document();

    // Use https://docs.aspose.com/words/net/generator-or-producer-name-included-in-output-documents/ to know how to check the result.
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.exportGeneratorName = false;

    doc.save(base.artifactsDir + "OoxmlSaveOptions.exportGeneratorName.docx", saveOptions);
    //ExEnd
  });


  /*//ExStart
    //ExFor:SaveOptions.ProgressCallback
    //ExFor:IDocumentSavingCallback
    //ExFor:IDocumentSavingCallback.Notify(DocumentSavingArgs)
    //ExFor:DocumentSavingArgs.EstimatedProgress
    //ExSummary:Shows how to manage a document while saving to docx.
  test.each([SaveFormat.Docx, "docx",
    SaveFormat.Docm, "docm",
    SaveFormat.Dotm, "dotm",
    SaveFormat.Dotx, "dotx",
    SaveFormat.FlatOpc, "flatopc"])('ProgressCallback', (SaveFormat saveFormat, string ext) => {
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // Following formats are supported: Docx, FlatOpc, Docm, Dotm, Dotx.
    let saveOptions = new aw.Saving.OoxmlSaveOptions(saveFormat)
    {
      ProgressCallback = new SavingProgressCallback()
    };

    var exception = Assert.Throws<OperationCanceledException>(() =>
      doc.save(base.artifactsDir + `OoxmlSaveOptions.progressCallback.${ext}`, saveOptions));
    expect(exception?.Message.contains("EstimatedProgress")).toEqual(true);
  });


    /// <summary>
    /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
    /// </summary>
  public class SavingProgressCallback : IDocumentSavingCallback
  {
      /// <summary>
      /// Ctr.
      /// </summary>
    public SavingProgressCallback()
    {
      mSavingStartedAt = Date.now();
    }

      /// <summary>
      /// Callback method which called during document saving.
      /// </summary>
      /// <param name="args">Saving arguments.</param>
    public void Notify(DocumentSavingArgs args)
    {
      DateTime canceledAt = Date.now();
      double ellapsedSeconds = (canceledAt - mSavingStartedAt).TotalSeconds;
      if (ellapsedSeconds > MaxDuration)
        throw new OperationCanceledException(`EstimatedProgress = ${args.estimatedProgress}; CanceledAt = ${canceledAt}`);
    }

      /// <summary>
      /// Date and time when document saving is started.
      /// </summary>
    private readonly DateTime mSavingStartedAt;

      /// <summary>
      /// Maximum allowed duration in sec.
      /// </summary>
    private const double MaxDuration = 0.01;
  }
  //ExEnd*/

  test('Zip64ModeOption', async () => {
    //ExStart:Zip64ModeOption
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:aw.Saving.OoxmlSaveOptions.zip64Mode
    //ExFor:Zip64Mode
    //ExSummary:Shows how to use ZIP64 format extensions.
    //let random = new Random();
    let builder = new aw.DocumentBuilder();

    for (let i = 0; i < 10000; i++)
    {
      let rgbaColor = Math.floor(Math.random() * (1 + 0xffffffff));
      let image = new jimp.Jimp({width: 5, height: 5, color: rgbaColor});
      let buffer = await image.getBuffer("image/bmp", {});
      let data = Array.from(buffer);
      builder.insertImage(data);
    }

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.zip64Mode = aw.Saving.Zip64Mode.Always;
    builder.document.save(base.artifactsDir + "OoxmlSaveOptions.Zip64ModeOption.docx", saveOptions);
    //ExEnd:Zip64ModeOption
  });

  test('DigitalSignature', () => {
    //ExStart:DigitalSignature
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:aw.Saving.OoxmlSaveOptions.digitalSignatureDetails
    //ExSummary:Shows how to sign OOXML document.
    let doc = new aw.Document(base.myDir + "Document.docx");

    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");
    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    let signOptions = new aw.DigitalSignatures.SignOptions();
    signOptions.comments = "Some comments";
    signOptions.signTime = new Date();
    saveOptions.digitalSignatureDetails = new aw.Saving.DigitalSignatureDetails(certificateHolder, signOptions);

    doc.save(base.artifactsDir + "OoxmlSaveOptions.digitalSignature.docx", saveOptions);
    //ExEnd:DigitalSignature
  });

});
