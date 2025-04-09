// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const os = require("os");

describe("ExFontSettings", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('DefaultFontInstance', () => {
    //ExStart
    //ExFor:FontSettings.defaultInstance
    //ExSummary:Shows how to configure the default font settings instance.
    // Configure the default font settings instance to use the "Courier New" font
    // as a backup substitute when we attempt to use an unknown font.
    aw.Fonts.FontSettings.defaultInstance.substitutionSettings.defaultFontSubstitution.defaultFontName = "Courier New";

    expect(aw.Fonts.FontSettings.defaultInstance.substitutionSettings.defaultFontSubstitution.enabled).toEqual(true);

    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Non-existent font";
    builder.write("Hello world!");

    // This document does not have a FontSettings configuration. When we render the document,
    // the default FontSettings instance will resolve the missing font.
    // Aspose.words will use "Courier New" to render text that uses the unknown font.
    expect(doc.fontSettings).toBe(null);

    doc.save(base.artifactsDir + "FontSettings.DefaultFontInstance.pdf");
    //ExEnd
  });


  test.skip('DefaultFontName: Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', () => {
    //ExStart
    //ExFor:DefaultFontSubstitutionRule.defaultFontName
    //ExSummary:Shows how to specify a default font.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Arvo";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    let fontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    // The font sources that the document uses contain the font "Arial", but not "Arvo".
    expect(fontSources.length).toEqual(1);
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arvo")).toEqual(false);

    // Set the "DefaultFontName" property to "Courier New" to,
    // while rendering the document, apply that font in all cases when another font is not available. 
    aw.Fonts.FontSettings.defaultInstance.substitutionSettings.defaultFontSubstitution.defaultFontName = "Courier New";

    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Courier New")).toEqual(true);

    // Aspose.words will now use the default font in place of any missing fonts during any rendering calls.
    doc.save(base.artifactsDir + "FontSettings.defaultFontName.pdf");
    //ExEnd
  });


  /*//Commented
  test('UpdatePageLayoutWarnings', () => {
    // Store the font sources currently used so we can restore them later
    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    // Load the document to render
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
    let callback = new HandleDocumentWarnings();
    doc.warningCallback = callback;

    // We can choose the default font to use in the case of any missing fonts
    aw.Fonts.FontSettings.defaultInstance.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial";

    // For testing we will set Aspose.words to look for fonts only in a folder which does not exist. Since Aspose.words won't
    // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
    // font specified under FontSettings.defaultFontName. We can pick up on this substitution using our callback
    aw.Fonts.FontSettings.defaultInstance.setFontsFolder('', false);

    // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
    // are stored until the document save and then sent to the appropriate WarningCallback
    doc.updatePageLayout();

    // Even though the document was rendered previously, any save warnings are notified to the user during document save
    doc.save(base.artifactsDir + "FontSettings.UpdatePageLayoutWarnings.pdf");

    expect(callback.FontWarnings.count > 0).toEqual(true);
    expect(callback.FontWarnings.at(0).warningType == aw.WarningType.FontSubstitution).toEqual(true);
    expect(callback.FontWarnings.at(0).description.contains("has not been found")).toEqual(true);

    // Restore default fonts
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);
  });


  public class HandleDocumentWarnings : IWarningCallback
  {
      /// <summary>
      /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
      /// potential issue during document processing. The callback can be set to listen for warnings generated during document
      /// load and/or document save.
      /// </summary>
    public void Warning(WarningInfo info)
    {
        // We are only interested in fonts being substituted
      if (info.warningType == aw.WarningType.FontSubstitution)
      {
        console.log("Font substitution: " + info.description);
        FontWarnings.warning(info);
      }
    }

    public WarningInfoCollection FontWarnings = new aw.WarningInfoCollection();
  }*/

  /*  //ExStart
    //ExFor:IWarningCallback
    //ExFor:DocumentBase.WarningCallback
    //ExFor:FontSettings.DefaultInstance
    //ExSummary:Shows how to use the IWarningCallback interface to monitor font substitution warnings.
  test('SubstitutionWarning', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Times New Roman";
    builder.writeln("Hello world!");

    let callback = new FontSubstitutionWarningCollector();
    doc.warningCallback = callback;

    // Store the current collection of font sources, which will be the default font source for every document
    // for which we do not specify a different font source.
    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    // For testing purposes, we will set Aspose.words to look for fonts only in a folder that does not exist.
    aw.Fonts.FontSettings.defaultInstance.setFontsFolder('', false);

    // When rendering the document, there will be no place to find the "Times New Roman" font.
    // This will cause a font substitution warning, which our callback will detect.
    doc.save(base.artifactsDir + "FontSettings.SubstitutionWarning.pdf");

    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);

    expect(callback.FontSubstitutionWarnings.count).toEqual(1);
    expect(callback.FontSubstitutionWarnings.at(0).warningType == aw.WarningType.FontSubstitution).toEqual(true);
    Assert.true(callback.FontSubstitutionWarnings.at(0).description
      .Equals(
        "Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));
  });


  private class FontSubstitutionWarningCollector : IWarningCallback
  {
      /// <summary>
      /// Called every time a warning occurs during loading/saving.
      /// </summary>
    public void Warning(WarningInfo info)
    {
      if (info.warningType == aw.WarningType.FontSubstitution)
        FontSubstitutionWarnings.warning(info);
    }

    public WarningInfoCollection FontSubstitutionWarnings = new aw.WarningInfoCollection();
  }
    //ExEnd*/

  /*  //ExStart
    //ExFor:FontSourceBase.WarningCallback
    //ExSummary:Shows how to call warning callback when the font sources working with.
  test('FontSourceWarning', () => {
    let settings = new aw.Fonts.FontSettings();
    settings.setFontsFolder("bad folder?", false);

    let source = settings.getFontsSources()[0];
    let callback = new FontSourceWarningCollector();
    source.warningCallback = callback;

    // Get the list of fonts to call warning callback.
    IList<PhysicalFontInfo> fontInfos = source.getAvailableFonts();

    Assert.true(callback.FontSubstitutionWarnings.at(0).description
      .includes("Error loading font from the folder \"bad folder?\""));
  });


  private class FontSourceWarningCollector : IWarningCallback
  {
      /// <summary>
      /// Called every time a warning occurs during processing of font source.
      /// </summary>
    public void Warning(WarningInfo info)
    {
      FontSubstitutionWarnings.warning(info);
    }

    public readonly WarningInfoCollection FontSubstitutionWarnings = new aw.WarningInfoCollection();
  }
    //ExEnd*/

  /*  //ExStart
    //ExFor:FontInfoSubstitutionRule
    //ExFor:FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:LayoutOptions.KeepOriginalFontMetrics
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.GetEnumerator
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
  test('EnableFontSubstitution', () => {
    // Open a document that contains text formatted with a font that does not exist in any of our font sources.
    let doc = new aw.Document(base.myDir + "Missing font.docx");

    // Assign a callback for handling font substitution warnings.
    let substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
    doc.warningCallback = substitutionWarningHandler;

    // Set a default font name and enable font substitution.
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial";
    ;
    fontSettings.substitutionSettings.fontInfoSubstitution.enabled = true;

    // Original font metrics should be used after font substitution.
    doc.layoutOptions.keepOriginalFontMetrics = true;

    // We will get a font substitution warning if we save a document with a missing font.
    doc.fontSettings = fontSettings;
    doc.save(base.artifactsDir + "FontSettings.EnableFontSubstitution.pdf");

    using (IEnumerator<WarningInfo> warnings = substitutionWarningHandler.FontWarnings.getEnumerator())
      while (warnings.moveNext())
        console.log(warnings.current.description);

    // We can also verify warnings in the collection and clear them.
    expect(substitutionWarningHandler.FontWarnings.at(0).source).toEqual(aw.WarningSource.Layout);
    Assert.AreEqual(
      "Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.",
      substitutionWarningHandler.FontWarnings.at(0).description);

    substitutionWarningHandler.FontWarnings.clear();

    expect(substitutionWarningHandler.FontWarnings.count).toEqual(0);
  });


  public class HandleDocumentSubstitutionWarnings : IWarningCallback
  {
      /// <summary>
      /// Called every time a warning occurs during loading/saving.
      /// </summary>
    public void Warning(WarningInfo info)
    {
      if (info.warningType == aw.WarningType.FontSubstitution)
        FontWarnings.warning(info);
    }

    public WarningInfoCollection FontWarnings = new aw.WarningInfoCollection();
  }
    //ExEnd

  test('SubstitutionWarningsClosestMatch', () => {
    let doc = new aw.Document(base.myDir + "Bullet points with alternative font.docx");

    let callback = new HandleDocumentSubstitutionWarnings();
    doc.warningCallback = callback;

    doc.save(base.artifactsDir + "FontSettings.SubstitutionWarningsClosestMatch.pdf");

    Assert.true(callback.FontWarnings.at(0).description
      .Equals(
        "Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."));
  });


  test('DisableFontSubstitution', () => {
    let doc = new aw.Document(base.myDir + "Missing font.docx");

    let callback = new HandleDocumentSubstitutionWarnings();
    doc.warningCallback = callback;

    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial";
    fontSettings.substitutionSettings.fontInfoSubstitution.enabled = false;

    doc.fontSettings = fontSettings;
    doc.save(base.artifactsDir + "FontSettings.DisableFontSubstitution.pdf");

    let reg = new Regex(
      "Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");

    for (let fontWarning of callback.FontWarnings)
    {
      Match match = reg.match(fontWarning.description);
      if (match.success)
      {
        Assert.Pass();
      }
    }
  });


    [Category("SkipMono")]
  test('SubstitutionWarnings', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let callback = new HandleDocumentSubstitutionWarnings();
    doc.warningCallback = callback;

    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial";
    fontSettings.setFontsFolder(base.fontsDir, false);
    fontSettings.substitutionSettings.tableSubstitution.addSubstitutes("Arial", "Arvo", "Slab");

    doc.fontSettings = fontSettings;
    doc.save(base.artifactsDir + "FontSettings.SubstitutionWarnings.pdf");

    Assert.AreEqual(
      "Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
      callback.FontWarnings.at(0).description);
    Assert.AreEqual(
      "Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
      callback.FontWarnings.at(1).description);
  });


  test('GetSubstitutionWithoutSuffixes', () => {
    let doc = new aw.Document(base.myDir + "Get substitution without suffixes.docx");

    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    let substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
    doc.warningCallback = substitutionWarningHandler;

    List<FontSourceBase> fontSources = new aw.Lists.List<FontSourceBase>(aw.Fonts.FontSettings.defaultInstance.getFontsSources());
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, true);
    fontSources.add(folderFontSource);

    let updatedFontSources = fontSources.toArray();
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(updatedFontSources);

    doc.save(base.artifactsDir + "Font.GetSubstitutionWithoutSuffixes.pdf");

    Assert.AreEqual(
      "Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
      substitutionWarningHandler.FontWarnings.at(0).description);

    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);
  });
  //EndCommented*/


  test('FontSourceFile', () => {
    //ExStart
    //ExFor:FileFontSource
    //ExFor:FileFontSource.#ctor(String)
    //ExFor:FileFontSource.#ctor(String, Int32)
    //ExFor:FileFontSource.filePath
    //ExFor:FileFontSource.type
    //ExFor:FontSourceBase
    //ExFor:FontSourceBase.priority
    //ExFor:FontSourceBase.type
    //ExFor:FontSourceType
    //ExSummary:Shows how to use a font file in the local file system as a font source.
    let fileFontSource = new aw.Fonts.FileFontSource(base.myDir + "Alte DIN 1451 Mittelschrift.ttf", 0);

    let doc = new aw.Document();
    doc.fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings.setFontsSources([fileFontSource]);

    expect(fileFontSource.filePath).toEqual(base.myDir + "Alte DIN 1451 Mittelschrift.ttf");
    expect(fileFontSource.type).toEqual(aw.Fonts.FontSourceType.FontFile);
    expect(fileFontSource.priority).toEqual(0);
    //ExEnd
  });


  test('FontSourceFolder', () => {
    //ExStart
    //ExFor:FolderFontSource
    //ExFor:FolderFontSource.#ctor(String, Boolean)
    //ExFor:FolderFontSource.#ctor(String, Boolean, Int32)
    //ExFor:FolderFontSource.folderPath
    //ExFor:FolderFontSource.scanSubfolders
    //ExFor:FolderFontSource.type
    //ExSummary:Shows how to use a local system folder which contains fonts as a font source.

    // Create a font source from a folder that contains font files.
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, false, 1);

    let doc = new aw.Document();
    doc.fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings.setFontsSources([folderFontSource]);

    expect(folderFontSource.folderPath).toEqual(base.fontsDir);
    expect(folderFontSource.scanSubfolders).toEqual(false);
    expect(folderFontSource.type).toEqual(aw.Fonts.FontSourceType.FontsFolder);
    expect(folderFontSource.priority).toEqual(1);
    //ExEnd
  });


  test.skip.each([false, true])('SetFontsFolder(recursive = %o): Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', (recursive) => {
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.setFontsFolder(String, Boolean)
    //ExSummary:Shows how to set a font source directory.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arvo";
    builder.writeln("Hello world!");
    builder.font.name = "Amethysta";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    // Our font sources do not contain the font that we have used for text in this document.
    // If we use these font settings while rendering this document,
    // Aspose.words will apply a fallback font to text which has a font that Aspose.words cannot locate.
    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(originalFontSources.length).toEqual(1);
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);

    // The default font sources are missing the two fonts that we are using in this document.
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arvo")).toEqual(false);
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(false);

    // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
    // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
    // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
    // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
    // in the first argument, as well as all the fonts in its subdirectories.
    aw.Fonts.FontSettings.defaultInstance.setFontsFolder(base.fontsDir, recursive);

    let newFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(newFontSources.length).toEqual(1);
    expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(false);
    expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arvo")).toEqual(true);

    // The "Amethysta" font is in a subfolder of the font directory.
    if (recursive)
    {
      expect(newFontSources.at(0).getAvailableFonts().Count).toEqual(25);
      expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(true);
    }
    else
    {
      expect(newFontSources.at(0).getAvailableFonts().Count).toEqual(18);
      expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(false);
    }

    doc.save(base.artifactsDir + "FontSettings.setFontsFolder.pdf");

    // Restore the original font sources.
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);
    //ExEnd
  });


  test.skip.each([false, true])('SetFontsFolders(recursive = %o): Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', (recursive) => {
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.setFontsFolders(String[], Boolean)
    //ExSummary:Shows how to set multiple font source directories.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Amethysta";
    builder.writeln("The quick brown fox jumps over the lazy dog.");
    builder.font.name = "Junction Light";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    // Our font sources do not contain the font that we have used for text in this document.
    // If we use these font settings while rendering this document,
    // Aspose.words will apply a fallback font to text which has a font that Aspose.words cannot locate.
    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(originalFontSources.length).toEqual(1);
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);

    // The default font sources are missing the two fonts that we are using in this document.
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(false);
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Junction Light")).toEqual(false);

    // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
    // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
    // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
    // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
    // in the first argument, as well as all the fonts in their subdirectories.
    aw.Fonts.FontSettings.defaultInstance.setFontsFolders([base.fontsDir + "/Amethysta", base.fontsDir + "/Junction"], recursive);

    let newFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(newFontSources.length).toEqual(2);
    expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(false);
    expect(newFontSources.at(0).getAvailableFonts().Count).toEqual(1);
    expect(newFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(true);

    // The "Junction" folder itself contains no font files, but has subfolders that do.
    if (recursive)
    {
      expect(newFontSources.at(1).getAvailableFonts().Count).toEqual(6);
      expect(newFontSources.at(1).getAvailableFonts().Any(f => f.fullFontName == "Junction Light")).toEqual(true);
    }
    else
    {
      expect(newFontSources.at(1).getAvailableFonts().Count).toEqual(0);
    }

    doc.save(base.artifactsDir + "FontSettings.setFontsFolders.pdf");

    // Restore the original font sources.
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);
    //ExEnd
  });


  test.skip('AddFontSource: Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', () => {
    //ExStart
    //ExFor:FontSettings
    //ExFor:FontSettings.getFontsSources()
    //ExFor:FontSettings.setFontsSources(FontSourceBase[])
    //ExSummary:Shows how to add a font source to our existing font sources.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Amethysta";
    builder.writeln("The quick brown fox jumps over the lazy dog.");
    builder.font.name = "Junction Light";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    let originalFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(originalFontSources.length).toEqual(1);

    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);

    // The default font source is missing two of the fonts that we are using in our document.
    // When we save this document, Aspose.words will apply fallback fonts to all text formatted with inaccessible fonts.
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(false);
    expect(originalFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Junction Light")).toEqual(false);

    // Create a font source from a folder that contains fonts.
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, true);

    // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
    let updatedFontSources = [originalFontSources[0], folderFontSource];
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(updatedFontSources);

    // Verify that Aspose.words has access to all required fonts before we render the document to PDF.
    updatedFontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    expect(updatedFontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);
    expect(updatedFontSources.at(1).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(true);
    expect(updatedFontSources.at(1).getAvailableFonts().Any(f => f.fullFontName == "Junction Light")).toEqual(true);

    doc.save(base.artifactsDir + "FontSettings.AddFontSource.pdf");

    // Restore the original font sources.
    aw.Fonts.FontSettings.defaultInstance.setFontsSources(originalFontSources);
    //ExEnd
  });


  test.skip('SetSpecifyFontFolder: WORDSNODEJS-124', () => {
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.setFontsFolder(base.fontsDir, false);

    // Using load options
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.fontSettings = fontSettings;

    let doc = new aw.Document(base.myDir + "Rendering.docx", loadOptions);

    let folderSource = doc.fontSettings.getFontsSources()[0].asFolderFontSource();

    expect(folderSource.folderPath).toEqual(base.fontsDir);
    expect(folderSource.scanSubfolders).toEqual(false);
  });


  test.skip('TableSubstitution: Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', () => {
    //ExStart
    //ExFor:Document.fontSettings
    //ExFor:TableSubstitutionRule.setSubstitutes(String, String[])
    //ExSummary:Shows how set font substitution rules.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Amethysta";
    builder.writeln("The quick brown fox jumps over the lazy dog.");

    let fontSources = aw.Fonts.FontSettings.defaultInstance.getFontsSources();

    // The default font sources contain the first font that the document uses.
    expect(fontSources.length).toEqual(1);
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arial")).toEqual(true);

    // The second font, "Amethysta", is unavailable.
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Amethysta")).toEqual(false);

    // We can configure a font substitution table which determines
    // which fonts Aspose.words will use as substitutes for unavailable fonts.
    // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
    // If the first substitute is unavailable, Aspose.words attempts to use the second substitute, and so on.
    doc.fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings.substitutionSettings.tableSubstitution.setSubstitutes("Amethysta", ["Arvo", "Courier New"]);

    // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo". 
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Arvo")).toEqual(false);

    // "Arvo" is also unavailable, but "Courier New" is. 
    expect(fontSources.at(0).getAvailableFonts().Any(f => f.fullFontName == "Courier New")).toEqual(true);

    // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
    doc.save(base.artifactsDir + "FontSettings.tableSubstitution.pdf");
    //ExEnd
  });


  test.skip('SetSpecifyFontFolders: WORDSNODEJS-124', () => {
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.setFontsFolders([base.fontsDir, "C:\\Windows\\Fonts\\"], true);

    // Using load options
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.fontSettings = fontSettings;
    let doc = new aw.Document(base.myDir + "Rendering.docx", loadOptions);

    let folderSource = doc.fontSettings.getFontsSources()[0].FolderFontSource();
    expect(folderSource.folderPath).toEqual(base.fontsDir);
    expect(folderSource.scanSubfolders).toEqual(true);

    folderSource = doc.fontSettings.getFontsSources()[1].asFolderFontSource();
    expect(folderSource.folderPath).toEqual("C:\\Windows\\Fonts\\");
    expect(folderSource.scanSubfolders).toEqual(true);
  });


  test('AddFontSubstitutes', () => {
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.substitutionSettings.tableSubstitution.setSubstitutes("Slab", ["Times New Roman", "Arial"]);
    fontSettings.substitutionSettings.tableSubstitution.addSubstitutes("Arvo", ["Open Sans", "Arial"]);

    let doc = new aw.Document(base.myDir + "Rendering.docx");
    doc.fontSettings = fontSettings;

    let alternativeFonts = doc.fontSettings.substitutionSettings.tableSubstitution.getSubstitutes("Slab");
    expect(["Times New Roman", "Arial"]).toEqual(alternativeFonts);

    alternativeFonts = doc.fontSettings.substitutionSettings.tableSubstitution.getSubstitutes("Arvo");
    expect(["Open Sans", "Arial"]).toEqual(alternativeFonts);
  });


  test('FontSourceMemory', () => {
    //ExStart
    //ExFor:MemoryFontSource
    //ExFor:MemoryFontSource.#ctor(Byte[])
    //ExFor:MemoryFontSource.#ctor(Byte[], Int32)
    //ExFor:MemoryFontSource.fontData
    //ExFor:MemoryFontSource.type
    //ExSummary:Shows how to use a byte array with data from a font file as a font source.

    let fontBytes = Array.from(fs.readFileSync(base.myDir + "Alte DIN 1451 Mittelschrift.ttf"));
    let memoryFontSource = new aw.Fonts.MemoryFontSource(fontBytes, 0);

    let doc = new aw.Document();
    doc.fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings.setFontsSources([memoryFontSource]);

    expect(memoryFontSource.type).toEqual(aw.Fonts.FontSourceType.MemoryFont);
    expect(memoryFontSource.priority).toEqual(0);
    //ExEnd
  });


  test.skip('FontSourceSystem: WORDSNODEJS-124', () => {
    //ExStart
    //ExFor:TableSubstitutionRule.addSubstitutes(String, String[])
    //ExFor:FontSubstitutionRule.enabled
    //ExFor:TableSubstitutionRule.getSubstitutes(String)
    //ExFor:FontSettings.resetFontSources
    //ExFor:FontSettings.substitutionSettings
    //ExFor:FontSubstitutionSettings
    //ExFor:FontSubstitutionSettings.fontNameSubstitution
    //ExFor:SystemFontSource
    //ExFor:SystemFontSource.#ctor
    //ExFor:SystemFontSource.#ctor(Int32)
    //ExFor:SystemFontSource.getSystemFontFolders
    //ExFor:SystemFontSource.type
    //ExSummary:Shows how to access a document's system font source and set font substitutes.
    console.log(os.platform());
    let doc = new aw.Document();
    doc.fontSettings = new aw.Fonts.FontSettings();

    // By default, a blank document always contains a system font source.
    expect(doc.fontSettings.getFontsSources().length).toEqual(1);

    let systemFontSource = doc.fontSettings.getFontsSources()[0].asSystemFontSource();
    expect(systemFontSource.type).toEqual(aw.Fonts.FontSourceType.SystemFonts);
    expect(systemFontSource.priority).toEqual(0);

    let platform = os.platform();
    if (platform == "win32")
    {
      let fontsPath = "C:\\WINDOWS\\Fonts";
      expect(aw.Fonts.SystemFontSource.getSystemFontFolders().FirstOrDefault()?.toLower()).toEqual(fontsPath.toLower());
    }

    for (let systemFontFolder of aw.Fonts.SystemFontSource.getSystemFontFolders())
    {
      console.log(systemFontFolder);
    }

    // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
    doc.fontSettings.substitutionSettings.fontInfoSubstitution.enabled = true;
    doc.fontSettings.substitutionSettings.tableSubstitution.addSubstitutes("Kreon-Regular", ["Calibri"]);

    expect(doc.fontSettings.substitutionSettings.tableSubstitution.getSubstitutes("Kreon-Regular").length).toEqual(1);
    //Assert.contains("Calibri", doc.fontSettings.substitutionSettings.tableSubstitution.getSubstitutes("Kreon-Regular").toArray());

    // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, false);
    doc.fontSettings.setFontsSources([systemFontSource, folderFontSource]);
    expect(doc.fontSettings.getFontsSources().length).toEqual(2);

    // Resetting the font sources still leaves us with the system font source as well as our substitutes.
    doc.fontSettings.resetFontSources();

    expect(doc.fontSettings.getFontsSources().length).toEqual(1);
    expect(doc.fontSettings.getFontsSources()[0].type).toEqual(aw.Fonts.FontSourceType.SystemFonts);
    expect(doc.fontSettings.substitutionSettings.tableSubstitution.getSubstitutes("Kreon-Regular").length).toEqual(1);
    //ExEnd*/
  });


  test('LoadFontFallbackSettingsFromFile', () => {
    //ExStart
    //ExFor:FontFallbackSettings.load(String)
    //ExFor:FontFallbackSettings.save(String)
    //ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Load an XML document that defines a set of font fallback settings.
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.fallbackSettings.load(base.myDir + "Font fallback rules.xml");

    doc.fontSettings = fontSettings;
    doc.save(base.artifactsDir + "FontSettings.LoadFontFallbackSettingsFromFile.pdf");

    // Save our document's current font fallback settings as an XML document.
    doc.fontSettings.fallbackSettings.save(base.artifactsDir + "FallbackSettings.xml");
    //ExEnd
  });


  test.skip('LoadFontFallbackSettingsFromStream: WORDSNODEJS-125', () => {
    //ExStart
    //ExFor:FontFallbackSettings.load(Stream)
    //ExFor:FontFallbackSettings.save(Stream)
    //ExSummary:Shows how to load and save font fallback settings to/from a stream.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Load an XML document that defines a set of font fallback settings.
    let fontFallbackStream =  fs.createReadStream(base.myDir + "Font fallback rules.xml")
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.fallbackSettings.load(fontFallbackStream);

    doc.fontSettings = fontSettings;
    doc.save(base.artifactsDir + "FontSettings.LoadFontFallbackSettingsFromStream.pdf");

    // Use a stream to save our document's current font fallback settings as an XML document.
    fontFallbackStream = fs.createWriteStream(base.artifactsDir + "FallbackSettings.xml")
    doc.fontSettings.fallbackSettings.save(fontFallbackStream);
    //ExEnd

    let fallbackSettingsDoc = new XmlDocument();
    fallbackSettingsDoc.loadXml(File.ReadAllText(base.artifactsDir + "FallbackSettings.xml"));
    let manager = new XmlNamespaceManager(fallbackSettingsDoc.nameTable);
    manager.AddNamespace("aw", "Aspose.words");

    let rules = fallbackSettingsDoc.selectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

    expect(rules.at(0).attributes.at("Ranges").value).toEqual("0B80-0BFF");
    expect(rules.at(0).attributes.at("FallbackFonts").value).toEqual("Vijaya");

    expect(rules.at(1).attributes.at("Ranges").value).toEqual("1F300-1F64F");
    expect(rules.at(1).attributes.at("FallbackFonts").value).toEqual("Segoe UI Emoji, Segoe UI Symbol");

    expect(rules.at(2).attributes.at("Ranges").value).toEqual("2000-206F, 2070-209F, 20B9");
    expect(rules.at(2).attributes.at("FallbackFonts").value).toEqual("Arial");

    expect(rules.at(3).attributes.at("Ranges").value).toEqual("3040-309F");
    expect(rules.at(3).attributes.at("FallbackFonts").value).toEqual("MS Gothic");
    expect(rules.at(3).attributes.at("BaseFonts").value).toEqual("Times New Roman");

    expect(rules.at(4).attributes.at("Ranges").value).toEqual("3040-309F");
    expect(rules.at(4).attributes.at("FallbackFonts").value).toEqual("MS Mincho");

    expect(rules.at(5).attributes.at("FallbackFonts").value).toEqual("Arial Unicode MS");
  });


  test('LoadNotoFontsFallbackSettings', () => {
    //ExStart
    //ExFor:FontFallbackSettings.loadNotoFallbackSettings
    //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
    let fontSettings = new aw.Fonts.FontSettings();

    // These are free fonts licensed under the SIL Open Font License.
    // We can download the fonts here:
    // https://www.google.com/get/noto/#sans-lgc
    fontSettings.setFontsFolder(base.fontsDir + "Noto", false);

    // Note that the predefined settings only use Sans-style Noto fonts with regular weight. 
    // Some of the Noto fonts use advanced typography features.
    // Fonts featuring advanced typography may not be rendered correctly as Aspose.words currently do not support them.
    fontSettings.fallbackSettings.loadNotoFallbackSettings();
    fontSettings.substitutionSettings.fontInfoSubstitution.enabled = false;
    fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Noto Sans";

    let doc = new aw.Document();
    doc.fontSettings = fontSettings;
    //ExEnd
  });


  test('DefaultFontSubstitutionRule', () => {
    //ExStart
    //ExFor:DefaultFontSubstitutionRule
    //ExFor:DefaultFontSubstitutionRule.defaultFontName
    //ExFor:FontSubstitutionSettings.defaultFontSubstitution
    //ExSummary:Shows how to set the default font substitution rule.
    let doc = new aw.Document();
    let fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings = fontSettings;

    // Get the default substitution rule within FontSettings.
    // This rule will substitute all missing fonts with "Times New Roman".
    let defaultFontSubstitutionRule = fontSettings.substitutionSettings.defaultFontSubstitution;
    expect(defaultFontSubstitutionRule.enabled).toEqual(true);
    expect(defaultFontSubstitutionRule.defaultFontName).toEqual("Times New Roman");

    // Set the default font substitute to "Courier New".
    defaultFontSubstitutionRule.defaultFontName = "Courier New";

    // Using a document builder, add some text in a font that we do not have to see the substitution take place,
    // and then render the result in a PDF.
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Missing Font";
    builder.writeln("Line written in a missing font, which will be substituted with Courier New.");

    doc.save(base.artifactsDir + "FontSettings.DefaultFontSubstitutionRule.pdf");
    //ExEnd

    expect(defaultFontSubstitutionRule.defaultFontName).toEqual("Courier New");
  });


  test('FontConfigSubstitution', () => {
    //ExStart
    //ExFor:FontConfigSubstitutionRule
    //ExFor:FontConfigSubstitutionRule.enabled
    //ExFor:FontConfigSubstitutionRule.isFontConfigAvailable
    //ExFor:FontConfigSubstitutionRule.resetCache
    //ExFor:FontSubstitutionRule
    //ExFor:FontSubstitutionRule.enabled
    //ExFor:FontSubstitutionSettings.fontConfigSubstitution
    //ExSummary:Shows operating system-dependent font config substitution.
    let fontSettings = new aw.Fonts.FontSettings();
    let fontConfigSubstitution = fontSettings.substitutionSettings.fontConfigSubstitution;

    let platform = os.platform();
    let isWindows = platform == "win32"
    // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
    // On Windows, it is unavailable.
    if (isWindows)
    {
      expect(fontConfigSubstitution.enabled).toEqual(false);
      expect(fontConfigSubstitution.isFontConfigAvailable()).toEqual(false);
    }

    let isLinuxOrMac = (platform == "linux") || (platform == "darwin")
    // On Linux/Mac, we will have access to it, and will be able to perform operations.
    if (isLinuxOrMac)
    {
      expect(fontConfigSubstitution.enabled).toEqual(true);
      expect(fontConfigSubstitution.isFontConfigAvailable()).toEqual(true);

      fontConfigSubstitution.resetCache();
    }

    //ExEnd
  });


  test.skip('FallbackSettings: usage of XmlDocument and XmlNamespaceManager', () => {
    //ExStart
    //ExFor:FontFallbackSettings.loadMsOfficeFallbackSettings
    //ExFor:FontFallbackSettings.loadNotoFallbackSettings
    //ExSummary:Shows how to load pre-defined fallback font settings.
    let doc = new aw.Document();

    let fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings = fontSettings;
    let fontFallbackSettings = fontSettings.fallbackSettings;

    // Save the default fallback font scheme to an XML document.
    // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
    // This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
    // the fallback scheme will use symbols from the "Vani" font substitute.
    fontFallbackSettings.save(base.artifactsDir + "FontSettings.fallbackSettings.default.xml");

    // Below are two pre-defined font fallback schemes we can choose from.
    // 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
    fontFallbackSettings.loadMsOfficeFallbackSettings();
    fontFallbackSettings.save(base.artifactsDir + "FontSettings.fallbackSettings.loadMsOfficeFallbackSettings.xml");

    // 2 -  Use the scheme built from Google Noto fonts:
    fontFallbackSettings.loadNotoFallbackSettings();
    fontFallbackSettings.save(base.artifactsDir + "FontSettings.fallbackSettings.loadNotoFallbackSettings.xml");
    //ExEnd

    let fallbackSettingsDoc = new XmlDocument();
    fallbackSettingsDoc.loadXml(File.ReadAllText(base.artifactsDir + "FontSettings.fallbackSettings.default.xml"));
    let manager = new XmlNamespaceManager(fallbackSettingsDoc.nameTable);
    manager.AddNamespace("aw", "Aspose.words");

    let rules = fallbackSettingsDoc.selectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

    expect(rules.at(9).attributes.at("Ranges").value).toEqual("0C00-0C7F");
    expect(rules.at(9).attributes.at("FallbackFonts").value).toEqual("Vani");
  });


  test.skip('FallbackSettingsCustom: usage of XmlDocument and XmlNamespaceManager', () => {
    //ExStart
    //ExFor:FontSettings.fallbackSettings
    //ExFor:FontFallbackSettings
    //ExFor:FontFallbackSettings.buildAutomatic
    //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
    let doc = new aw.Document();

    let fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings = fontSettings;
    let fontFallbackSettings = fontSettings.fallbackSettings;

    // Configure our font settings to source fonts only from the "MyFonts" folder.
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, false);
    fontSettings.setFontsSources([folderFontSource]);

    // Calling the "BuildAutomatic" method will generate a fallback scheme that
    // distributes accessible fonts across as many Unicode character codes as possible.
    // In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
    fontFallbackSettings.buildAutomatic();
    fontFallbackSettings.save(base.artifactsDir + "FontSettings.FallbackSettingsCustom.buildAutomatic.xml");

    // We can also load a custom substitution scheme from a file like this.
    // This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
    // and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
    fontFallbackSettings.load(base.myDir + "Custom font fallback settings.xml");

    // Create a document builder and set its font to one that does not exist in any of our sources.
    // Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
    let builder = new aw.DocumentBuilder(doc);
    builder.font.name = "Missing Font";

    // Use the builder to print every Unicode character from 0x0021 to 0x052F,
    // with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
    for (let i = 0x0021; i < 0x0530; i++)
    {
      switch (i)
      {
        case 0x0021:
          builder.writeln(
            "\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"AllegroOpen\" font:");
          break;
        case 0x0100:
          builder.writeln(
            "\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"AllegroOpen\" font:");
          break;
        case 0x0250:
          builder.writeln("\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
          break;
      }

      builder.write(`${Convert.ToChar(i)}`);
    }

    doc.save(base.artifactsDir + "FontSettings.FallbackSettingsCustom.pdf");
    //ExEnd

    let fallbackSettingsDoc = new XmlDocument();
    fallbackSettingsDoc.loadXml(fs.readFileSync(base.artifactsDir + "FontSettings.FallbackSettingsCustom.buildAutomatic.xml"));
    let manager = new XmlNamespaceManager(fallbackSettingsDoc.nameTable);
    manager.AddNamespace("aw", "Aspose.words");

    let rules = fallbackSettingsDoc.selectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

    expect(rules.at(0).attributes.at("Ranges").value).toEqual("0000-007F");
    expect(rules.at(0).attributes.at("FallbackFonts").value).toEqual("AllegroOpen");

    expect(rules.at(2).attributes.at("Ranges").value).toEqual("0100-017F");
    expect(rules.at(2).attributes.at("FallbackFonts").value).toEqual("AllegroOpen");

    expect(rules.at(4).attributes.at("Ranges").value).toEqual("0250-02AF");
    expect(rules.at(4).attributes.at("FallbackFonts").value).toEqual("M+ 2m");

    expect(rules.at(7).attributes.at("Ranges").value).toEqual("0370-03FF");
    expect(rules.at(7).attributes.at("FallbackFonts").value).toEqual("Arvo");
  });


  test.skip('TableSubstitutionRule: usage of XmlDocument and XmlNamespaceManager', () => {
    //ExStart
    //ExFor:TableSubstitutionRule
    //ExFor:TableSubstitutionRule.loadLinuxSettings
    //ExFor:TableSubstitutionRule.loadWindowsSettings
    //ExFor:TableSubstitutionRule.save(Stream)
    //ExFor:TableSubstitutionRule.save(String)
    //ExSummary:Shows how to access font substitution tables for Windows and Linux.
    let doc = new aw.Document();
    let fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings = fontSettings;

    // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
    let tableSubstitutionRule = fontSettings.substitutionSettings.tableSubstitution;
    tableSubstitutionRule.loadWindowsSettings();

    // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
    expect(tableSubstitutionRule.getSubstitutes("Times New Roman CE").ToArray()).toEqual(["Times New Roman"]);

    // We can save the table in the form of an XML document.
    tableSubstitutionRule.save(base.artifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml");

    // Linux has its own substitution table.
    // There are multiple substitute fonts for "Times New Roman CE".
    // If the first substitute, "FreeSerif" is also unavailable,
    // this rule will cycle through the others in the array until it finds an available one.
    tableSubstitutionRule.loadLinuxSettings();
    expect(["FreeSerif", "Liberation Serif", "DejaVu Serif"]).toEqual(tableSubstitutionRule.getSubstitutes("Times New Roman CE").toArray());

    // Save the Linux substitution table in the form of an XML document using a stream.
    let fileStream = fs.createWriteStream(base.artifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml");
    tableSubstitutionRule.save(fileStream);
    //ExEnd

    let fallbackSettingsDoc = new XmlDocument();
    fallbackSettingsDoc.loadXml(fs.readFileSync(base.artifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml"));
    let manager = new XmlNamespaceManager(fallbackSettingsDoc.nameTable);
    manager.AddNamespace("aw", "Aspose.words");

    let rules = fallbackSettingsDoc.selectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

    expect(rules.at(16).attributes.at("OriginalFont").value).toEqual("Times New Roman CE");
    expect(rules.at(16).attributes.at("SubstituteFonts").value).toEqual("Times New Roman");

    fallbackSettingsDoc.loadXml(fs.readFileSync(base.artifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml"));
    rules = fallbackSettingsDoc.selectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

    expect(rules.at(31).attributes.at("OriginalFont").value).toEqual("Times New Roman CE");
    expect(rules.at(31).attributes.at("SubstituteFonts").value).toEqual("FreeSerif, Liberation Serif, DejaVu Serif");
  });


  test.skip('TableSubstitutionRuleCustom: WORDSNODEJS-126', () => {
    //ExStart
    //ExFor:FontSubstitutionSettings.tableSubstitution
    //ExFor:TableSubstitutionRule.addSubstitutes(String,String[])
    //ExFor:TableSubstitutionRule.getSubstitutes(String)
    //ExFor:TableSubstitutionRule.load(Stream)
    //ExFor:TableSubstitutionRule.load(String)
    //ExFor:TableSubstitutionRule.setSubstitutes(String,String[])
    //ExSummary:Shows how to work with custom font substitution tables.
    let doc = new aw.Document();
    let fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings = fontSettings;

    // Create a new table substitution rule and load the default Windows font substitution table.
    let tableSubstitutionRule = fontSettings.substitutionSettings.tableSubstitution;

    // If we select fonts exclusively from our folder, we will need a custom substitution table.
    // We will no longer have access to the Microsoft Windows fonts,
    // such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
    let folderFontSource = new aw.Fonts.FolderFontSource(base.fontsDir, false);
    fontSettings.setFontsSources([folderFontSource]);

    // Below are two ways of loading a substitution table from a file in the local file system.
    // 1 -  From a stream:
    let fileStream = fs.createReadStream(base.myDir + "Font substitution rules.xml");
    tableSubstitutionRule.load(fileStream);

    // 2 -  Directly from a file:
    tableSubstitutionRule.load(base.myDir + "Font substitution rules.xml");

    // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
    // We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
    expect(["Missing Font", "Kreon"]).toEqual(tableSubstitutionRule.getSubstitutes("Arial").toArray());

    // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
    expect(tableSubstitutionRule.getSubstitutes("Times New Roman")).toBe(null);
    tableSubstitutionRule.addSubstitutes("Times New Roman", "Arvo");
    expect(tableSubstitutionRule.getSubstitutes("Times New Roman").toArray()).toEqual(["Arvo"]);

    // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
    // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
    tableSubstitutionRule.addSubstitutes("Times New Roman", "M+ 2m");
    expect(["Arvo", "M+ 2m"]).toEqual(tableSubstitutionRule.getSubstitutes("Times New Roman").toArray());

    // SetSubstitutes() can set a new list of substitute fonts for a font.
    tableSubstitutionRule.setSubstitutes("Times New Roman", "Squarish Sans CT", "M+ 2m");
    expect(["Squarish Sans CT", "M+ 2m"]).toEqual(tableSubstitutionRule.getSubstitutes("Times New Roman").toArray());

    // Writing text in fonts that we do not have access to will invoke our substitution rules.
    let builder = new aw.DocumentBuilder(doc);
    builder.font.name = "Arial";
    builder.writeln("Text written in Arial, to be substituted by Kreon.");

    builder.font.name = "Times New Roman";
    builder.writeln("Text written in Times New Roman, to be substituted by Squarish Sans CT.");

    doc.save(base.artifactsDir + "FontSettings.TableSubstitutionRule.custom.pdf");
    //ExEnd
  });


  test('ResolveFontsBeforeLoadingDocument', () => {
    //ExStart
    //ExFor:LoadOptions.fontSettings
    //ExSummary:Shows how to designate font substitutes during loading.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.fontSettings = new aw.Fonts.FontSettings();

    // Set a font substitution rule for a LoadOptions object.
    // If the document we are loading uses a font which we do not have,
    // this rule will substitute the unavailable font with one that does exist.
    // In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
    let substitutionRule = loadOptions.fontSettings.substitutionSettings.tableSubstitution;
    substitutionRule.addSubstitutes("MissingFont", ["Comic Sans MS"]);

    let doc = new aw.Document(base.myDir + "Missing font.html", loadOptions);

    // At this point such text will still be in "MissingFont".
    // Font substitution will take place when we render the document.
    expect(doc.firstSection.body.firstParagraph.runs.at(0).font.name).toEqual("MissingFont");

    doc.save(base.artifactsDir + "FontSettings.ResolveFontsBeforeLoadingDocument.pdf");
    //ExEnd
  });


  /*  //ExStart
    //ExFor:StreamFontSource
    //ExFor:StreamFontSource.OpenFontDataStream
    //ExSummary:Shows how to load fonts from stream.
  test('StreamFontSourceFileRendering', () => {
    let fontSettings = new aw.Fonts.FontSettings();
    fontSettings.setFontsSources(new aw.Fonts.FontSourceBase[] {new StreamFontSourceFile()});

    let builder = new aw.DocumentBuilder();
    builder.document.fontSettings = fontSettings;
    builder.font.name = "Kreon-Regular";
    builder.writeln("Test aspose text when saving to PDF.");

    builder.document.save(base.artifactsDir + "FontSettings.StreamFontSourceFileRendering.pdf");
  });


    /// <summary>
    /// Load the font data only when required instead of storing it in the memory
    /// for the entire lifetime of the "FontSettings" object.
    /// </summary>
  private class StreamFontSourceFile : StreamFontSource
  {
    public override Stream OpenFontDataStream()
    {
      return File.OpenRead(base.fontsDir + "Kreon-Regular.ttf");
    }
  }
    //ExEnd*/

  /*  //ExStart
    //ExFor:FileFontSource.#ctor(String, Int32, String)
    //ExFor:MemoryFontSource.#ctor(Byte[], Int32, String)
    //ExFor:FontSettings.SaveSearchCache(Stream)
    //ExFor:FontSettings.SetFontsSources(FontSourceBase[], Stream)
    //ExFor:FileFontSource.CacheKey
    //ExFor:MemoryFontSource.CacheKey
    //ExFor:StreamFontSource.CacheKey
    //ExSummary:Shows how to speed up the font cache initialization process.
  test('LoadFontSearchCache', () => {
    const string cacheKey1 = "Arvo";
    const string cacheKey2 = "Arvo-Bold";
    let parsedFonts = new aw.Fonts.FontSettings();
    let loadedCache = new aw.Fonts.FontSettings();

    parsedFonts.setFontsSources(new aw.Fonts.FontSourceBase[]
    {
      new aw.Fonts.FileFontSource(base.fontsDir + "Arvo-Regular.ttf", 0, cacheKey1),
      new aw.Fonts.FileFontSource(base.fontsDir + "Arvo-Bold.ttf", 0, cacheKey2)
    });
            
    using (MemoryStream cacheStream = new MemoryStream())
    {
      parsedFonts.saveSearchCache(cacheStream);
      loadedCache.setFontsSources(new aw.Fonts.FontSourceBase[]
      {
        new SearchCacheStream(cacheKey1),
        new aw.Fonts.MemoryFontSource(File.ReadAllBytes(base.fontsDir + "Arvo-Bold.ttf"), 0, cacheKey2)
      }, cacheStream);
    }

    expect(loadedCache.getFontsSources().Length).toEqual(parsedFonts.getFontsSources().Length);
  });


    /// <summary>
    /// Load the font data only when required instead of storing it in the memory
    /// for the entire lifetime of the "FontSettings" object.
    /// </summary>
  private class SearchCacheStream : StreamFontSource
  {
    public SearchCacheStream(string cacheKey):base(0, cacheKey)
    {
    }

    public override Stream OpenFontDataStream()
    {
      return File.OpenRead(base.fontsDir + "Arvo-Regular.ttf");
    }
  }
    //ExEnd*/

});
