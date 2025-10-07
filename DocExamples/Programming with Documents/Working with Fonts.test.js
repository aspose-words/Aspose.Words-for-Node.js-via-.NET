// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithFonts", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('FontFormatting', () => {
        //ExStart:WriteAndFont
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let font = builder.font;
        font.size = 16;
        font.bold = true;
        font.color = "#0000FF";
        font.name = "Arial";
        font.underline = aw.Underline.Dash;

        builder.write("Sample text.");

        doc.save(base.artifactsDir + "WorkingWithFonts.FontFormatting.docx");
        //ExEnd:WriteAndFont
    });

    test('GetFontLineSpacing', () => {
        //ExStart:GetFontLineSpacing
        //GistId:4977e1370d5e0deaf48bfa1197bcae98
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.font.name = "Calibri";
        builder.writeln("qText");

        let font = builder.document.firstSection.body.firstParagraph.runs.at(0).font;
        console.log(`lineSpacing = ${font.lineSpacing}`);
        //ExEnd:GetFontLineSpacing
    });

    test('CheckDmlTextEffect', () => {
        //ExStart:CheckDmlTextEffect
        let doc = new aw.Document(base.myDir + "DrawingML text effects.docx");

        let runs = doc.firstSection.body.firstParagraph.runs;
        let runFont = runs.at(0).font;

        // One run might have several Dml text effects applied.
        console.log(runFont.hasDmlEffect(aw.TextDmlEffect.Shadow));
        console.log(runFont.hasDmlEffect(aw.TextDmlEffect.Effect3D));
        console.log(runFont.hasDmlEffect(aw.TextDmlEffect.Reflection));
        console.log(runFont.hasDmlEffect(aw.TextDmlEffect.Outline));
        console.log(runFont.hasDmlEffect(aw.TextDmlEffect.Fill));
        //ExEnd:CheckDmlTextEffect
    });

    test('SetFontFormatting', () => {
        //ExStart:SetFontFormatting
        //GistId:4977e1370d5e0deaf48bfa1197bcae98
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let font = builder.font;
        font.bold = true;
        font.color = "#00008B";
        font.italic = true;
        font.name = "Arial";
        font.size = 24;
        font.spacing = 5;
        font.underline = aw.Underline.Double;

        builder.writeln("I'm a very nice formatted string.");

        doc.save(base.artifactsDir + "WorkingWithFonts.SetFontFormatting.docx");
        //ExEnd:SetFontFormatting
    });

    test('SetFontEmphasisMark', () => {
        //ExStart:SetFontEmphasisMark
        //GistId:4977e1370d5e0deaf48bfa1197bcae98
        let document = new aw.Document();
        let builder = new aw.DocumentBuilder(document);

        builder.font.emphasisMark = aw.EmphasisMark.UnderSolidCircle;

        builder.write("Emphasis text");
        builder.writeln();
        builder.font.clearFormatting();
        builder.write("Simple text");

        document.save(base.artifactsDir + "WorkingWithFonts.SetFontEmphasisMark.docx");
        //ExEnd:SetFontEmphasisMark
    });

    test('EnableDisableFontSubstitution', () => {
        //ExStart:EnableDisableFontSubstitution
        let doc = new aw.Document(base.myDir + "Rendering.docx");

        let fontSettings = new aw.Fonts.FontSettings();
        fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial";
        fontSettings.substitutionSettings.fontInfoSubstitution.enabled = false;

        doc.fontSettings = fontSettings;

        doc.save(base.artifactsDir + "WorkingWithFonts.EnableDisableFontSubstitution.pdf");
        //ExEnd:EnableDisableFontSubstitution
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

    test('DefaultInstance', () => {
        //ExStart:DefaultInstance
        //GistId:412100f144878625758c6f877d9ec584
        aw.Fonts.FontSettings.defaultInstance.setFontsFolder(base.fontsDir, true);
        //ExEnd:DefaultInstance

        let doc = new aw.Document(base.myDir + "Rendering.docx");
        doc.save(base.artifactsDir + "WorkingWithFonts.DefaultInstance.pdf");
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

    test('SetFontsFoldersSystemAndCustomFolder', () => {
        //ExStart:SetFontsFoldersSystemAndCustomFolder
        let doc = new aw.Document(base.myDir + "Rendering.docx");

        let fontSettings = new aw.Fonts.FontSettings();
        // Retrieve the array of environment-dependent font sources that are searched by default.
        // For example this will contain a "Windows\Fonts\" source on a Windows machines.
        // We add this array to a new List to make adding or removing font entries much easier.
        let fontSources = [...fontSettings.getFontsSources()];

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        let folderFontSource = new aw.Fonts.FolderFontSource("C:\\MyFonts\\", true);
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.push(folderFontSource);

        fontSettings.setFontsSources(fontSources);

        doc.fontSettings = fontSettings;

        doc.save(base.artifactsDir + "WorkingWithFonts.SetFontsFoldersSystemAndCustomFolder.pdf");
        //ExEnd:SetFontsFoldersSystemAndCustomFolder
    });

    test('FontsFoldersWithPriority', () => {
        //ExStart:FontsFoldersWithPriority
        //GistId:412100f144878625758c6f877d9ec584
        aw.Fonts.FontSettings.defaultInstance.setFontsSources([new aw.Fonts.SystemFontSource(), new aw.Fonts.FolderFontSource(base.fontsDir, true, 1)])
        //ExEnd:FontsFoldersWithPriority

        let doc = new aw.Document(base.myDir + "Rendering.docx");
        doc.save(base.artifactsDir + "WorkingWithFonts.FontsFoldersWithPriority.pdf");
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

    test('SpecifyDefaultFontWhenRendering', () => {
        //ExStart:SpecifyDefaultFontWhenRendering
        let doc = new aw.Document(base.myDir + "Rendering.docx");

        let fontSettings = new aw.Fonts.FontSettings();
        // If the default font defined here cannot be found during rendering then
        // the closest font on the machine is used instead.
        fontSettings.substitutionSettings.defaultFontSubstitution.defaultFontName = "Arial Unicode MS";

        doc.fontSettings = fontSettings;

        doc.save(base.artifactsDir + "WorkingWithFonts.SpecifyDefaultFontWhenRendering.pdf");
        //ExEnd:SpecifyDefaultFontWhenRendering
    });

    test('FontSettingsWithLoadOptions', () => {
        //ExStart:FontSettingsWithLoadOptions
        let fontSettings = new aw.Fonts.FontSettings();

        let substitutionRule = fontSettings.substitutionSettings.tableSubstitution;
        // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
        substitutionRule.addSubstitutes("UnknownFont1", ["Comic Sans MS"]);

        let loadOptions = new aw.Loading.LoadOptions();
        loadOptions.fontSettings = fontSettings;

        let doc = new aw.Document(base.myDir + "Rendering.docx", loadOptions);
        //ExEnd:FontSettingsWithLoadOptions
    });

    test('SetFontsFolder', () => {
        //ExStart:SetFontsFolder
        let fontSettings = new aw.Fonts.FontSettings();
        fontSettings.setFontsFolder(base.myDir + "Fonts", false);

        let loadOptions = new aw.Loading.LoadOptions();
        loadOptions.fontSettings = fontSettings;

        let doc = new aw.Document(base.myDir + "Rendering.docx", loadOptions);
        //ExEnd:SetFontsFolder
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

});