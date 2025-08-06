// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');
const MemoryStream = require('memorystream');


describe("WorkingWithFonts", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('LoadOptionFontSettings', () => {
    //ExStart:LoadOptionFontSettings
    //GistId:194889f2a6beb4b2aed7b1ad088392ca
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.fontSettings = new aw.Fonts.FontSettings();

    let doc = new aw.Document(base.myDir + "Rendering.docx", loadOptions);
    //ExEnd:LoadOptionFontSettings
  });

  test('FontSettingsDefaultInstance', () => {
      //ExStart:FontsFolders
      //GistId:412100f144878625758c6f877d9ec584
      //ExStart:FontSettingsFontSource
      //GistId:194889f2a6beb4b2aed7b1ad088392ca
      //ExStart:FontSettingsDefaultInstance
      //GistId:194889f2a6beb4b2aed7b1ad088392ca
      let fontSettings = aw.Fonts.FontSettings.defaultInstance;
      //ExEnd:FontSettingsDefaultInstance
      fontSettings.setFontsSources([new aw.Fonts.SystemFontSource(), new aw.Fonts.FolderFontSource("C:\\MyFonts\\", true)]);
      //ExEnd:FontSettingsFontSource

      let doc = new aw.Document(base.myDir + "Rendering.docx");
      //ExEnd:FontsFolders
  });

  test('FontFallbackSettings', () => {
      //ExStart:FontFallbackSettings
      //GistId:194889f2a6beb4b2aed7b1ad088392ca
      let doc = new aw.Document(base.myDir + "Rendering.docx");

      let fontSettings = aw.Fonts.FontSettings.defaultInstance;
      fontSettings.fallbackSettings.load(base.myDir + "Font fallback rules.xml");

      doc.fontSettings = fontSettings;

      doc.save(base.artifactsDir + "WorkingWithFonts.FontFallbackSettings.pdf")
      //ExEnd:FontFallbackSettings
  });

  test('NotoFallbackSettings', () => {
      //ExStart:NotoFallbackSettings
      //GistId:194889f2a6beb4b2aed7b1ad088392ca
      let doc = new aw.Document(base.myDir + "Rendering.docx");

      let fontSettings = aw.Fonts.FontSettings.defaultInstance;
      fontSettings.fallbackSettings.loadNotoFallbackSettings;

      doc.fontSettings = fontSettings;

      doc.save(base.artifactsDir + "WorkingWithFonts.NotoFallbackSettings.pdf")
      //ExEnd:NotoFallbackSettings
  });

  test('TrueTypeFontsFolder', () => {
        //ExStart:TrueTypeFontsFolder
        //GistId:412100f144878625758c6f877d9ec584
        let doc = new aw.Document(base.myDir + "Rendering.docx");

        let fontSettings = new aw.Fonts.FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.getFontSources and
        // FontSettings.setFontSources instead.
        fontSettings.setFontsFolder(base.fontsDir, false);
        doc.fontSettings = fontSettings;

        doc.save(base.artifactsDir + "WorkingWithFonts.TrueTypeFontsFolder.pdf")
        //ExEnd:TrueTypeFontsFolder
    });

    test('MultipleFolders', () => {
        //ExStart:MultipleFolders
        //GistId:412100f144878625758c6f877d9ec584
        let doc = new aw.Document(base.myDir + "Rendering.docx");

        let fontSettings = new aw.Fonts.FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.getFontSources and
        // FontSettings.setFontSources instead.
        fontSettings.setFontsFolders([base.fontsDir, "D:\\Misc\\Fonts\\"], false);
        doc.fontSettings = fontSettings;

        doc.save(base.artifactsDir + "WorkingWithFonts.MultipleFolders.pdf")
        //ExEnd:MultipleFolders
    });

    test('DefaultInstance', () => {
        //ExStart:DefaultInstance
        //GistId:412100f144878625758c6f877d9ec584
        aw.Fonts.FontSettings.defaultInstance.setFontsFolder(base.fontsDir, true);
        //ExEnd:DefaultInstance

        let doc = new aw.Document(base.myDir + "Rendering.docx");
        doc.save(base.artifactsDir + "WorkingWithFonts.DefaultInstance.pdf");
    });

    test('FontsFoldersWithPriority', () => {
        //ExStart:FontsFoldersWithPriority
        //GistId:412100f144878625758c6f877d9ec584
        aw.Fonts.FontSettings.defaultInstance.setFontsSources([new aw.Fonts.SystemFontSource(), new aw.Fonts.FolderFontSource(base.fontsDir, true, 1)])
        //ExEnd:FontsFoldersWithPriority

        let doc = new aw.Document(base.myDir + "Rendering.docx");
        doc.save(base.artifactsDir + "WorkingWithFonts.FontsFoldersWithPriority.pdf");
    });

});