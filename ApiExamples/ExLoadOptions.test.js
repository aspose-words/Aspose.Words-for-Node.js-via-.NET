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


describe("ExLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  //ExStart
  //ExFor:LoadOptions.ResourceLoadingCallback
  //ExSummary:Shows how to handle external resources when loading Html documents.
  test.skip('LoadOptionsCallback - TODO: WORDSNODEJS-121 - Add support of loadOptions.resourceLoadingCallback', () => {
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.resourceLoadingCallback = new HtmlLinkedResourceLoadingCallback();

    // When we load the document, our callback will handle linked resources such as CSS stylesheets and images.
    let doc = new aw.Document(base.myDir + "Images.html", loadOptions);
    doc.save(base.artifactsDir + "LoadOptions.LoadOptionsCallback.pdf");
  });


/*  
    /// <summary>
    /// Prints the filenames of all external stylesheets and substitutes all images of a loaded html document.
    /// </summary>
  private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
  {
    public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
    {
      switch (args.resourceType)
      {
        case aw.Loading.ResourceType.CssStyleSheet:
          console.log(`External CSS Stylesheet found upon loading: ${args.originalUri}`);
          return aw.Loading.ResourceLoadingAction.Default;
        case aw.Loading.ResourceType.Image:
          console.log(`External Image found upon loading: ${args.originalUri}`);

          const string newImageFilename = "Logo.jpg";
          console.log(`\tImage will be substituted with: ${newImageFilename}`);

          Image newImage = Image.FromFile(base.imageDir + newImageFilename);

          let converter = new ImageConverter();
          byte.at(] imageBytes = (byte[))converter.ConvertTo(newImage, typeof(byte[]));
          args.setData(imageBytes);

          return aw.Loading.ResourceLoadingAction.UserProvided;
      }

      return aw.Loading.ResourceLoadingAction.Default;
    }
  }
    //ExEnd */

  test.each([true,
    false])('ConvertShapeToOfficeMath', (isConvertShapeToOfficeMath) => {
    //ExStart
    //ExFor:LoadOptions.convertShapeToOfficeMath
    //ExSummary:Shows how to convert EquationXML shapes to Office Math objects.
    let loadOptions = new aw.Loading.LoadOptions();

    // Use this flag to specify whether to convert the shapes with EquationXML attributes
    // to Office Math objects and then load the document.
    loadOptions.convertShapeToOfficeMath = isConvertShapeToOfficeMath;

    let doc = new aw.Document(base.myDir + "Math shapes.docx", loadOptions);

    if (isConvertShapeToOfficeMath)
    {
      expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(16);
      expect(doc.getChildNodes(aw.NodeType.OfficeMath, true).count).toEqual(34);
    }
    else
    {
      expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(24);
      expect(doc.getChildNodes(aw.NodeType.OfficeMath, true).count).toEqual(0);
    }
    //ExEnd
  });


  test('SetEncoding', () => {
    //ExStart
    //ExFor:LoadOptions.encoding
    //ExSummary:Shows how to set the encoding with which to open a document.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.encoding = "ascii";

    // Load the document while passing the LoadOptions object, then verify the document's contents.
    let doc = new aw.Document(base.myDir + "English text.txt", loadOptions);

    expect(doc.toString(aw.SaveFormat.Text).includes("This is a sample text in English.")).toEqual(true);
    //ExEnd
  });


  test('FontSettings', () => {
    //ExStart
    //ExFor:LoadOptions.fontSettings
    //ExSummary:Shows how to apply font substitution settings while loading a document. 
    // Create a FontSettings object that will substitute the "Times New Roman" font
    // with the font "Arvo" from our "MyFonts" folder.
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.setFontsFolder(base.fontsDir, false);
    fontSettings.substitutionSettings.tableSubstitution.addSubstitutes("Times New Roman", [ "Arvo" ]);

    // Set that FontSettings object as a property of a newly created LoadOptions object.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.fontSettings = fontSettings;

    // Load the document, then render it as a PDF with the font substitution.
    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);

    doc.save(base.artifactsDir + "LoadOptions.fontSettings.pdf");
    //ExEnd
  });


  test('LoadOptionsMswVersion', () => {
    //ExStart
    //ExFor:LoadOptions.mswVersion
    //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
    // By default, Aspose.words load documents according to Microsoft Word 2019 specification.
    let loadOptions = new aw.Loading.LoadOptions();

    expect(loadOptions.mswVersion).toEqual(aw.Settings.MsWordVersion.Word2019);

    // This document is missing the default paragraph formatting style.
    // This default style will be regenerated when we load the document either with Microsoft Word or Aspose.words.
    loadOptions.mswVersion = aw.Settings.MsWordVersion.Word2007;
    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);

    // The style's line spacing will have this value when loaded by Microsoft Word 2007 specification.
    expect(doc.styles.defaultParagraphFormat.lineSpacing).toBeCloseTo(12.95, 2);
    //ExEnd
  });


  //ExStart
  //ExFor:LoadOptions.WarningCallback
  //ExSummary:Shows how to print and store warnings that occur during document loading.
  test.skip('LoadOptionsWarningCallback - TODO: WORDSNODEJS-108 - Add support of IWarningCallback', () => {
    // Create a new aw.Loading.LoadOptions object and set its WarningCallback attribute
    // as an instance of our IWarningCallback implementation.
    var loadOptions = new aw.Loading.LoadOptions();
    loadOptions.warningCallback = new DocumentLoadingWarningCallback();

    // Our callback will print all warnings that come up during the load operation.
    let doc = new aw.Document(base.myDir + "Document.docx", loadOptions);

    let warnings = loadOptions.warningCallback.getWarnings();
    expect(warnings.count).toEqual(3);
    testLoadOptionsWarningCallback(warnings); //ExSkip
  });


/*  
    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
  private class DocumentLoadingWarningCallback : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
      console.log(`Warning: ${info.warningType}`);
      console.log(`\tSource: ${info.source}`);
      console.log(`\tDescription: ${info.description}`);
      mWarnings.add(info);
    }

    public List<WarningInfo> GetWarnings()
    {
      return mWarnings;
    }

    private readonly List<WarningInfo> mWarnings = new aw.Lists.List<WarningInfo>();
  }
    //ExEnd

  private static void TestLoadOptionsWarningCallback(List<WarningInfo> warnings)
  {
    expect(warnings.at(0).warningType).toEqual(aw.WarningType.UnexpectedContent);
    expect(warnings.at(0).source).toEqual(aw.WarningSource.Docx);
    expect(warnings.at(0).description).toEqual("3F01");

    expect(warnings.at(1).warningType).toEqual(aw.WarningType.MinorFormattingLoss);
    expect(warnings.at(1).source).toEqual(aw.WarningSource.Docx);
    expect(warnings.at(1).description).toEqual("Import of element 'shapedefaults' is not supported in Docx format by Aspose.words.");

    expect(warnings.at(2).warningType).toEqual(aw.WarningType.MinorFormattingLoss);
    expect(warnings.at(2).source).toEqual(aw.WarningSource.Docx);
    expect(warnings.at(2).description).toEqual("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.words.");
  }
*/


  test('TempFolder', () => {
    //ExStart
    //ExFor:LoadOptions.tempFolder
    //ExSummary:Shows how to use the hard drive instead of memory when loading a document.
    // When we load a document, various elements are temporarily stored in memory as the save operation occurs.
    // We can use this option to use a temporary folder in the local file system instead,
    // which will reduce our application's memory overhead.
    let options = new aw.Loading.LoadOptions();
    options.tempFolder = base.artifactsDir + "TempFiles";

    // The specified temporary folder must exist in the local file system before the load operation.
    if (!fs.existsSync(options.tempFolder)) {
      fs.mkdirSync(options.tempFolder);
    }

    let doc = new aw.Document(base.myDir + "Document.docx", options);

    // The folder will persist with no residual contents from the load operation.
    expect(fs.readdirSync(options.tempFolder).length).toEqual(0);
    //ExEnd
  });


  test('AddEditingLanguage', () => {
    //ExStart
    //ExFor:LanguagePreferences
    //ExFor:LanguagePreferences.addEditingLanguage(EditingLanguage)
    //ExFor:LoadOptions.languagePreferences
    //ExFor:EditingLanguage
    //ExSummary:Shows how to apply language preferences when loading a document.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.languagePreferences.addEditingLanguage(aw.Loading.EditingLanguage.Japanese);

    let doc = new aw.Document(base.myDir + "No default editing language.docx", loadOptions);

    var localeIdFarEast = doc.styles.defaultFont.localeIdFarEast;
    console.log(localeIdFarEast == aw.Loading.EditingLanguage.Japanese
      ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
      : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
    //ExEnd

    expect(doc.styles.defaultFont.localeIdFarEast).toEqual(aw.Loading.EditingLanguage.Japanese);

    doc = new aw.Document(base.myDir + "No default editing language.docx");

    expect(doc.styles.defaultFont.localeIdFarEast).toEqual(aw.Loading.EditingLanguage.EnglishUS);
  });


  test('SetEditingLanguageAsDefault', () => {
    //ExStart
    //ExFor:LanguagePreferences.defaultEditingLanguage
    //ExSummary:Shows how set a default language when loading a document.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.languagePreferences.defaultEditingLanguage = aw.Loading.EditingLanguage.Russian;

    let doc = new aw.Document(base.myDir + "No default editing language.docx", loadOptions);

    var localeId = doc.styles.defaultFont.localeId;
    console.log(localeId == aw.Loading.EditingLanguage.Russian
      ? "The document either has no any language set in defaults or it was set to Russian originally."
      : "The document default language was set to another than Russian language originally, so it is not overridden.");
    //ExEnd

    expect(doc.styles.defaultFont.localeId).toEqual(aw.Loading.EditingLanguage.Russian);

    doc = new aw.Document(base.myDir + "No default editing language.docx");

    expect(doc.styles.defaultFont.localeId).toEqual(aw.Loading.EditingLanguage.EnglishUS);
  });


  test('ConvertMetafilesToPng', () => {
    //ExStart
    //ExFor:LoadOptions.convertMetafilesToPng
    //ExSummary:Shows how to convert WMF/EMF to PNG during loading document.
    let doc = new aw.Document();

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Image);
    shape.imageData.setImage(base.imageDir + "Windows MetaFile.wmf");
    shape.width = 100;
    shape.height = 100;

    doc.firstSection.body.firstParagraph.appendChild(shape);

    doc.save(base.artifactsDir + "Image.CreateImageDirectly.docx");

    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(1600, 1600, aw.Drawing.ImageType.Wmf, shape);

    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.convertMetafilesToPng = true;

    doc = new aw.Document(base.artifactsDir + "Image.CreateImageDirectly.docx", loadOptions);
    shape = doc.getShape(0, true);

    TestUtil.verifyImageInShape(1666, 1666, aw.Drawing.ImageType.Png, shape);
    //ExEnd
  });


  test('OpenChmFile', () => {
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "HTML help.chm");
    expect(aw.LoadFormat.Chm).toEqual(info.loadFormat);

    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.encoding = "windows-1251";

    let doc = new aw.Document(base.myDir + "HTML help.chm", loadOptions);
  });


  //ExStart
  //ExFor:LoadOptions.ProgressCallback
  //ExFor:IDocumentLoadingCallback
  //ExFor:IDocumentLoadingCallback.Notify
  //ExSummary:Shows how to notify the user if document loading exceeded expected loading time.
  test.skip('ProgressCallback - TODO: WORDSNODEJS-122 - Add support of LoadOptions.ProgressCallback', () => {
    let progressCallback = new LoadingProgressCallback();

    let loadOptions = new aw.Loading.LoadOption();
    loadOptions.progressCallback = progressCallback;

    try
    {
      let doc = new aw.Document(base.myDir + "Big document.docx", loadOptions);
    }
    catch (err)
    {
      console.log(err);

      // Handle loading duration issue.
    }
  });

/*
    /// <summary>
    /// Cancel a document loading after the "MaxDuration" seconds.
    /// </summary>
  public class LoadingProgressCallback : IDocumentLoadingCallback
  {
      /// <summary>
      /// Ctr.
      /// </summary>
    public LoadingProgressCallback()
    {
      mLoadingStartedAt = Date.now();
    }

      /// <summary>
      /// Callback method which called during document loading.
      /// </summary>
      /// <param name="args">Loading arguments.</param>
    public void Notify(DocumentLoadingArgs args)
    {
      DateTime canceledAt = Date.now();
      double ellapsedSeconds = (canceledAt - mLoadingStartedAt).TotalSeconds;

      if (ellapsedSeconds > MaxDuration)
        throw new OperationCanceledException(`EstimatedProgress = ${args.estimatedProgress}; CanceledAt = ${canceledAt}`);
    }

      /// <summary>
      /// Date and time when document loading is started.
      /// </summary>
    private readonly DateTime mLoadingStartedAt;

      /// <summary>
      /// Maximum allowed duration in sec.
      /// </summary>
    private const double MaxDuration = 0.5;
  }
    //ExEnd
    */

  test('IgnoreOleData', () => {
    //ExStart
    //ExFor:LoadOptions.ignoreOleData
    //ExSummary:Shows how to ingore OLE data while loading.
    // Ignoring OLE data may reduce memory consumption and increase performance
    // without data lost in a case when destination format does not support OLE objects.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.ignoreOleData = true;
    let doc = new aw.Document(base.myDir + "OLE objects.docx", loadOptions);

    doc.save(base.artifactsDir + "LoadOptions.ignoreOleData.docx");
    //ExEnd
  });
});
