// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');

describe("ExHtmlSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.each([aw.SaveFormat.Html,
    aw.SaveFormat.Mhtml,
    aw.SaveFormat.Epub,
    aw.SaveFormat.Azw3,
    aw.SaveFormat.Mobi])('ExportPageMarginsEpub(%o)', (saveFormat) => {
    let doc = new aw.Document(base.myDir + "TextBoxes.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.saveFormat = saveFormat;
    saveOptions.exportPageMargins = true;

    doc.save(
      base.artifactsDir + "HtmlSaveOptions.ExportPageMarginsEpub" +
      aw.FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
  });


  test.each([[aw.SaveFormat.Html, aw.Saving.HtmlOfficeMathOutputMode.Image],
    [aw.SaveFormat.Mhtml, aw.Saving.HtmlOfficeMathOutputMode.MathML],
    [aw.SaveFormat.Epub, aw.Saving.HtmlOfficeMathOutputMode.Text],
    [aw.SaveFormat.Azw3, aw.Saving.HtmlOfficeMathOutputMode.Text],
    [aw.SaveFormat.Mobi, aw.Saving.HtmlOfficeMathOutputMode.Text]])('ExportOfficeMathEpub', (saveFormat, outputMode) => {
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.officeMathOutputMode = outputMode;

    doc.save(
      base.artifactsDir + "HtmlSaveOptions.ExportOfficeMathEpub" +
      aw.FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
  });


  test.each([[aw.SaveFormat.Html, true, Description = "TextBox as svg (html)"],
    [aw.SaveFormat.Epub, true, "TextBox as svg (epub)"],
    [aw.SaveFormat.Mhtml, false, "TextBox as img (mhtml)"],
    [aw.SaveFormat.Azw3, false, "TextBox as img (azw3)"],
    [aw.SaveFormat.Mobi, false, "TextBox as img (mobi)"]])('ExportTextBoxAsSvgEpub', (saveFormat, isTextBoxAsSvg, description) => {
    
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textbox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 300, 100);
    builder.moveTo(textbox.firstParagraph);
    builder.write("Hello world!");

    let saveOptions = new aw.Saving.HtmlSaveOptions(saveFormat);
    saveOptions.exportShapesAsSvg = isTextBoxAsSvg;
            
    doc.save(base.artifactsDir + "HtmlSaveOptions.ExportTextBoxAsSvgEpub" + aw.FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);

    let dirFiles;
    switch (saveFormat)
    {
      case aw.SaveFormat.Html:

        dirFiles = fs.readdirSync(base.artifactsDir, { recursive: true }).filter((f) => f.endsWith("HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png"));
        expect(dirFiles.length).toEqual(0);
        return;

      case aw.SaveFormat.Epub:

        dirFiles = fs.readdirSync(base.artifactsDir, { recursive: true }).filter((f) => f.endsWith("HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png"));
        expect(dirFiles.length).toEqual(0);
        return;

      case aw.SaveFormat.Mhtml:

        dirFiles = fs.readdirSync(base.artifactsDir, { recursive: true }).filter((f) => f.endsWith("HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png"));
        expect(dirFiles.length).toEqual(0);
        return;

      case aw.SaveFormat.Azw3:

        dirFiles = fs.readdirSync(base.artifactsDir, { recursive: true }).filter((f) => f.endsWith("HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png"));
        expect(dirFiles.length).toEqual(0);
        return;

      case aw.SaveFormat.Mobi:

        dirFiles = fs.readdirSync(base.artifactsDir, { recursive: true }).filter((f) => f.endsWith("HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png"));
        expect(dirFiles.length).toEqual(0);
        return;
    }
  });


  test('CreateAZW3Toc', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.navigationMapLevel
    //ExSummary:Shows how to generate table of contents for Azw3 documents.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Azw3);
    options.navigationMapLevel = 2;

    doc.save(base.artifactsDir + "HtmlSaveOptions.CreateAZW3Toc.azw3", options);
    //ExEnd
  });


  test('CreateMobiToc', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.navigationMapLevel
    //ExSummary:Shows how to generate table of contents for Mobi documents.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Mobi);
    options.navigationMapLevel = 5;

    doc.save(base.artifactsDir + "HtmlSaveOptions.CreateMobiToc.mobi", options);
    //ExEnd
  });


  test.each([aw.Saving.ExportListLabels.Auto,
    aw.Saving.ExportListLabels.AsInlineText,
    aw.Saving.ExportListLabels.ByHtmlTags])('ControlListLabelsExport', (howExportListLabels) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let bulletedList = doc.lists.add(aw.Lists.ListTemplate.BulletDefault);
    builder.listFormat.list = bulletedList;
    builder.paragraphFormat.leftIndent = 72;
    builder.writeln("Bulleted list item 1.");
    builder.writeln("Bulleted list item 2.");
    builder.paragraphFormat.clearFormatting();

    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it does not cause formatting loss, 
    // otherwise HTML <p> tag is used. This is also the default value.
    // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
    // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
    saveOptions.exportListLabels = howExportListLabels;

    doc.save(base.artifactsDir + "HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
  });


  test.each([true,
    false])('ExportUrlForLinkedImage', (exportOrigUrl) => {
    let doc = new aw.Document(base.myDir + "Linked image.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportOriginalUrlForLinkedImages = exportOrigUrl;

    doc.save(base.artifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

    let dirFiles = fs.readdirSync(base.artifactsDir).filter((f) => f.endsWith("HtmlSaveOptions.ExportUrlForLinkedImage.001.png"));

    DocumentHelper.findTextInFile(base.artifactsDir + "HtmlSaveOptions.ExportUrlForLinkedImage.html",
      dirFiles.length == 0
        ? "<img src=\"http://www.aspose.com/images/aspose-logo.gif\""
        : "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
  });


  test('ExportRoundtripInformation', () => {
    let doc = new aw.Document(base.myDir + "TextBoxes.docx");
    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportRoundtripInformation = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
  });


  test('RoundtripInformationDefaulValue', () => {
    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    expect(saveOptions.exportRoundtripInformation).toEqual(true);

    saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Mhtml);
    expect(saveOptions.exportRoundtripInformation).toEqual(false);

    saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Epub);
    expect(saveOptions.exportRoundtripInformation).toEqual(false);
  });


  test('ExternalResourceSavingConfig', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.exportFontResources = true;
    saveOptions.resourceFolder = "Resources";
    saveOptions.resourceFolderAlias = "https://www.aspose.com/";

    doc.save(base.artifactsDir + "HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);

    let imageFiles = fs.readdirSync(base.artifactsDir + "Resources/", { recursive: true })
      .filter((f) => f.match(/HtmlSaveOptions\.ExternalResourceSavingConfig(.*)\.png/));
    expect(imageFiles.length).toEqual(8);

    let fontFiles = fs.readdirSync(base.artifactsDir + "Resources/", { recursive: true })
      .filter((f) => f.match(/HtmlSaveOptions\.ExternalResourceSavingConfig(.*)\.ttf/));
    expect(fontFiles.length).toEqual(10);

    let cssFiles = fs.readdirSync(base.artifactsDir + "Resources/", {recursive: true})
      .filter((f) => f.match(/HtmlSaveOptions\.ExternalResourceSavingConfig(.*)\.css/));
    expect(cssFiles.length).toEqual(1);

    DocumentHelper.findTextInFile(base.artifactsDir + "HtmlSaveOptions.ExternalResourceSavingConfig.html",
      "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
  });


  test('ConvertFontsAsBase64', () => {
    let doc = new aw.Document(base.myDir + "TextBoxes.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.resourceFolder = "Resources";
    saveOptions.exportFontResources = true;
    saveOptions.exportFontsAsBase64 = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
  });


  test.each([aw.Saving.HtmlVersion.Html5,
    aw.Saving.HtmlVersion.Xhtml])('Html5Support', (htmlVersion) => {
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.htmlVersion = htmlVersion;

    doc.save(base.artifactsDir + "HtmlSaveOptions.Html5Support.html", saveOptions);
  });


  test.each([false,
    true])('ExportFonts', (exportAsBase64) => {
    let fontsFolder = base.artifactsDir + "HtmlSaveOptions.ExportFonts.Resources";

    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportFontResources = true;
    saveOptions.fontsFolder = fontsFolder;
    saveOptions.exportFontsAsBase64 = exportAsBase64;

    switch (exportAsBase64)
    {
      case false:

        doc.save(base.artifactsDir + "HtmlSaveOptions.ExportFonts.false.html", saveOptions);

        expect(fs.readdirSync(fontsFolder, {recursive: true})
          .filter((f) => f.endsWith("HtmlSaveOptions.ExportFonts.false.times.ttf").length)).not.toBe(0);

        fs.rmSync(fontsFolder, {recursive: true, force: true});
        break;

      case true:

        doc.save(base.artifactsDir + "HtmlSaveOptions.ExportFonts.true.html", saveOptions);
        expect(fs.existsSync(fontsFolder)).toBe(false);
        break;
    }
  });


  test('ResourceFolderPriority', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.exportFontResources = true;
    saveOptions.resourceFolder = base.artifactsDir + "Resources";
    saveOptions.resourceFolderAlias = "http://example.com/resources";

    doc.save(base.artifactsDir + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

    expect(fs.readdirSync(base.artifactsDir + "Resources", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderPriority.001.png").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Resources", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderPriority.002.png").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Resources", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderPriority.arial.ttf").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Resources", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderPriority.css").length)).not.toBe(0);
  });


  test('ResourceFolderLowPriority', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.exportFontResources = true;
    saveOptions.fontsFolder = base.artifactsDir + "Fonts";
    saveOptions.imagesFolder = base.artifactsDir + "Images";
    saveOptions.resourceFolder = base.artifactsDir + "Resources";
    saveOptions.resourceFolderAlias = "http://example.com/resources";

    doc.save(base.artifactsDir + "HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);

    expect(fs.readdirSync(base.artifactsDir + "Images", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderLowPriority.001.png").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Images", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderLowPriority.002.png").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Fonts", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf").length)).not.toBe(0);
    expect(fs.readdirSync(base.artifactsDir + "Resources", {recursive: true})
      .filter((f) => f.endsWith("HtmlSaveOptions.ResourceFolderLowPriority.css").length)).not.toBe(0);
  });


  test('SvgMetafileFormat', () => {
    let builder = new aw.DocumentBuilder();

    builder.write("Here is an SVG image: ");
    builder.insertHtml(
      "<svg height='210' width='500'>" +
"        <polygon points='100,10 40,198 190,78 10,78 160,198' " +
"         style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />" +
"      </svg> ");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.metafileFormat = aw.Saving.HtmlMetafileFormat.Png;
    builder.document.save(base.artifactsDir + "HtmlSaveOptions.SvgMetafileFormat.html", saveOptions);
  });


  test('PngMetafileFormat', () => {
    let builder = new aw.DocumentBuilder();

    builder.write("Here is an Png image: ");
    builder.insertHtml(
      `<svg height='210' width='500'>
        <polygon points='100,10 40,198 190,78 10,78 160,198' 
          style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
      </svg> `);

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.metafileFormat = aw.Saving.HtmlMetafileFormat.Png;
    builder.document.save(base.artifactsDir + "HtmlSaveOptions.PngMetafileFormat.html", saveOptions);
  });


  test('EmfOrWmfMetafileFormat', () => {
    let builder = new aw.DocumentBuilder();

    builder.write("Here is an image as is: ");
    builder.insertHtml(
      `<img src=""data:image/png;base64,
        iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
        C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
        AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
        REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
        ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
        vr4MkhoXe0rZigAAAABJRU5ErkJggg=="" alt=""Red dot"" />`);

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.metafileFormat = aw.Saving.HtmlMetafileFormat.EmfOrWmf;
    builder.document.save(base.artifactsDir + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html", saveOptions);
  });


  test('CssClassNamesPrefix', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.cssClassNamePrefix
    //ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.cssClassNamePrefix = "myprefix-";

    doc.save(base.artifactsDir + "HtmlSaveOptions.cssClassNamePrefix.html", saveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.cssClassNamePrefix.html").toString();

    expect(outDocContents.includes("<p class=\"myprefix-Header\">")).toEqual(true);
    expect(outDocContents.includes("<p class=\"myprefix-Footer\">")).toEqual(true);

    outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.cssClassNamePrefix.css").toString();

    expect(outDocContents.includes(".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:footer }")).toEqual(true);
    expect(outDocContents.includes(".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt; -aw-style-name:header }")).toEqual(true);
    //ExEnd
  });


  test('CssClassNamesNotValidPrefix', () => {
    let saveOptions = new aw.Saving.HtmlSaveOptions();
    expect(() => saveOptions.cssClassNamePrefix = "@%-").toThrow("The class name prefix must be a valid CSS identifier.");
  });


  test('CssClassNamesNullPrefix', () => {
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cCssStyleSheetType = aw.Saving.CssStyleSheetType.Embedded;
    saveOptions.cssClassNamePrefix = null;

    doc.save(base.artifactsDir + "HtmlSaveOptions.cssClassNamePrefix.html", saveOptions);
  });


  test('ContentIdScheme', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Mhtml);
    saveOptions.prettyFormat = true;
    saveOptions.exportCidUrlsForMhtmlResources = true;
    
    doc.save(base.artifactsDir + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
  });


/*    [Ignore("Bug")]
  test.each([false,
    true])('ResolveFontNames', (bool resolveFontNames) => {
    //ExStart
    //ExFor:HtmlSaveOptions.resolveFontNames
    //ExSummary:Shows how to resolve all font names before writing them to HTML.
    let doc = new aw.Document(base.myDir + "Missing font.docx");

    // This document contains text that names a font that we do not have.
    expect(doc.fontInfos.at("28 Days Later")).not.toBe(null);

    // If we have no way of getting this font, and we want to be able to display all the text
    // in this document in an output HTML, we can substitute it with another font.
    let fontSettings = new aw.Fonts.FontSettings
    {
      SubstitutionSettings =
      {
        DefaultFontSubstitution =
        {
          DefaultFontName = "Arial",
          Enabled = true
        }
      }
    };

    doc.fontSettings = fontSettings;

    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html)
    {
      // By default, this option is set to 'False' and Aspose.words writes font names as specified in the source document
      ResolveFontNames = resolveFontNames
    };

    doc.save(base.artifactsDir + "HtmlSaveOptions.resolveFontNames.html", saveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.resolveFontNames.html").toString();

    Assert.true(resolveFontNames
      ? Regex.match(outDocContents, "<span style=\"font-family:Arial\">").Success
      : Regex.match(outDocContents, "<span style=\"font-family:\'28 Days Later\'\">").Success);
    //ExEnd
  });
*/

  test('HeadingLevels', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.documentSplitHeadingLevel
    //ExSummary:Shows how to split an output HTML document by headings into several parts.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Every paragraph that we format using a "Heading" style can serve as a heading.
    // Each heading may also have a heading level, determined by the number of its heading style.
    // The headings below are of levels 1-3.
    builder.paragraphFormat.style = builder.document.styles.at("Heading 1");
    builder.writeln("Heading #1");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 2");
    builder.writeln("Heading #2");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 3");
    builder.writeln("Heading #3");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 1");
    builder.writeln("Heading #4");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 2");
    builder.writeln("Heading #5");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 3");
    builder.writeln("Heading #6");

    // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph".
    // These criteria will split the document at paragraphs with "Heading" styles into several smaller documents,
    // and save each document in a separate HTML file in the local file system.
    // We will also set the maximum heading level, which splits the document to 2.
    // Saving the document will split it at headings of levels 1 and 2, but not at 3 to 9.
    let options = new aw.Saving.HtmlSaveOptions();
    options.documentSplitCriteria = aw.Saving.DocumentSplitCriteria.HeadingParagraph;
    options.documentSplitHeadingLevel = 2;

    // Our document has four headings of levels 1 - 2. One of those headings will not be
    // a split point since it is at the beginning of the document.
    // The saving operation will split our document at three places, into four smaller documents.
    doc.save(base.artifactsDir + "HtmlSaveOptions.HeadingLevels.html", options);

    doc = new aw.Document(base.artifactsDir + "HtmlSaveOptions.HeadingLevels.html");

    expect(doc.getText().trim()).toEqual("Heading #1");

    doc = new aw.Document(base.artifactsDir + "HtmlSaveOptions.HeadingLevels-01.html");

    expect(doc.getText().trim()).toEqual("Heading #2\r" +
                            "Heading #3");

    doc = new aw.Document(base.artifactsDir + "HtmlSaveOptions.HeadingLevels-02.html");

    expect(doc.getText().trim()).toEqual("Heading #4");

    doc = new aw.Document(base.artifactsDir + "HtmlSaveOptions.HeadingLevels-03.html");

    expect(doc.getText().trim()).toEqual("Heading #5\r" +
                            "Heading #6");
    //ExEnd
  });


  test.each([false,
    true])('NegativeIndent', (allowNegativeIndent) => {
    //ExStart
    //ExFor:HtmlElementSizeOutputMode
    //ExFor:HtmlSaveOptions.allowNegativeIndent
    //ExFor:HtmlSaveOptions.tableWidthOutputMode
    //ExSummary:Shows how to preserve negative indents in the output .html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a table with a negative indent, which will push it to the left past the left page boundary.
    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, Cell 1");
    builder.insertCell();
    builder.write("Row 1, Cell 2");
    builder.endTable();
    table.leftIndent = -36;
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(144);

    builder.insertBreak(aw.BreakType.ParagraphBreak);

    // Insert a table with a positive indent, which will push the table to the right.
    table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, Cell 1");
    builder.insertCell();
    builder.write("Row 1, Cell 2");
    builder.endTable();
    table.leftIndent = 36;
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(144);

    // When we save a document to HTML, Aspose.words will only preserve negative indents
    // such as the one we have applied to the first table if we set the "AllowNegativeIndent" flag
    // in a SaveOptions object that we will pass to "true".
    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    options.allowNegativeIndent = allowNegativeIndent;
    options.tableWidthOutputMode = aw.Saving.HtmlElementSizeOutputMode.RelativeOnly;

    doc.save(base.artifactsDir + "HtmlSaveOptions.NegativeIndent.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.NegativeIndent.html").toString();

    if (allowNegativeIndent)
    {
      expect(outDocContents.includes(
        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">")).toEqual(true);
      expect(outDocContents.includes(
        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">")).toEqual(true); 
    }
    else
    {
      expect(outDocContents.includes(
        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">")).toEqual(true); 
      expect(outDocContents.includes(
        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">")).toEqual(true);
    }
    //ExEnd
  });


  test('FolderAlias', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportOriginalUrlForLinkedImages
    //ExFor:HtmlSaveOptions.fontsFolder
    //ExFor:HtmlSaveOptions.fontsFolderAlias
    //ExFor:HtmlSaveOptions.imageResolution
    //ExFor:HtmlSaveOptions.imagesFolderAlias
    //ExFor:HtmlSaveOptions.resourceFolder
    //ExFor:HtmlSaveOptions.resourceFolderAlias
    //ExSummary:Shows how to set folders and folder aliases for externally saved resources that Aspose.words will create when saving a document to HTML.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.HtmlSaveOptions();
    options.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    options.exportFontResources = true;
    options.imageResolution = 72;
    options.fontResourcesSubsettingSizeThreshold = 0;
    options.fontsFolder = base.artifactsDir + "Fonts";
    options.imagesFolder = base.artifactsDir + "Images";
    options.resourceFolder = base.artifactsDir + "Resources";
    options.fontsFolderAlias = "http://example.com/fonts";
    options.imagesFolderAlias = "http://example.com/images";
    options.resourceFolderAlias = "http://example.com/resources";
    options.exportOriginalUrlForLinkedImages = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.FolderAlias.html", options);
    //ExEnd
  });


  //ExStart
  //ExFor:HtmlSaveOptions.ExportFontResources
  //ExFor:HtmlSaveOptions.FontSavingCallback
  //ExFor:IFontSavingCallback
  //ExFor:IFontSavingCallback.FontSaving
  //ExFor:FontSavingArgs
  //ExFor:FontSavingArgs.Bold
  //ExFor:FontSavingArgs.Document
  //ExFor:FontSavingArgs.FontFamilyName
  //ExFor:FontSavingArgs.FontFileName
  //ExFor:FontSavingArgs.FontStream
  //ExFor:FontSavingArgs.IsExportNeeded
  //ExFor:FontSavingArgs.IsSubsettingNeeded
  //ExFor:FontSavingArgs.Italic
  //ExFor:FontSavingArgs.KeepFontStreamOpen
  //ExFor:FontSavingArgs.OriginalFileName
  //ExFor:FontSavingArgs.OriginalFileSize
  //ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
  test.skip('SaveExportedFonts - TODO: fontSavingCallback not supported yet', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Configure a SaveOptions object to export fonts to separate files.
    // Set a callback that will handle font saving in a custom manner.
    let options = new aw.Saving.HtmlSaveOptions
    options.exportFontResources = true,
    options.fontSavingCallback = new HandleFontSaving();

    // The callback will export .ttf files and save them alongside the output document.
    doc.save(base.artifactsDir + "HtmlSaveOptions.SaveExportedFonts.html", options);

    for (var fontFilename of fs.readdirSync(base.artifactsDir).filter((s) => s.endsWith(".ttf")))
    {
      console.log(fontFilename);
    }

    expect(fs.readdirSync(base.artifactsDir).filter((s) => s.endsWith(".ttf")).count).toEqual(10);
  });

/*
  /// <summary>
  /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
  /// </summary>
  public class HandleFontSaving : IFontSavingCallback
  {
    void aw.Saving.IFontSavingCallback.fontSaving(FontSavingArgs args)
    {
      Console.write(`Font:\t${args.fontFamilyName}`);
      if (args.bold) Console.write(", bold");
      if (args.italic) Console.write(", italic");
      console.log(`\nSource:\t${args.originalFileName}, ${args.originalFileSize} bytes\n`);

        // We can also access the source document from here.
      expect(args.document.originalFileName.EndsWith("Rendering.docx")).toEqual(true);

      expect(args.isExportNeeded).toEqual(true);
      expect(args.isSubsettingNeeded).toEqual(true);

        // There are two ways of saving an exported font.
        // 1 -  Save it to a local file system location:
      args.fontFileName = args.originalFileName.split(Path.DirectorySeparatorChar).Last();

        // 2 -  Save it to a stream:
      args.fontStream =
        new FileStream(base.artifactsDir + args.originalFileName.split(Path.DirectorySeparatorChar).Last(), FileMode.create);
      expect(args.keepFontStreamOpen).toEqual(false);
    }
  }
  //ExEnd
*/  

  test.each([aw.Saving.HtmlVersion.Html5,
    aw.Saving.HtmlVersion.Xhtml])('HtmlVersions', (htmlVersion) => {
    //ExStart
    //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
    //ExFor:HtmlSaveOptions.htmlVersion
    //ExFor:HtmlVersion
    //ExSummary:Shows how to save a document to a specific version of HTML.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    options.htmlVersion = htmlVersion,
    options.prettyFormat = true

    doc.save(base.artifactsDir + "HtmlSaveOptions.HtmlVersions.html", options);

    // Our HTML documents will have minor differences to be compatible with different HTML versions.
    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.HtmlVersions.html").toString();

    switch (htmlVersion)
    {
      case aw.Saving.HtmlVersion.Html5:
        expect(outDocContents.includes("<a id=\"_Toc76372689\"></a>")).toEqual(true);
        expect(outDocContents.includes("<a id=\"_Toc76372689\"></a>")).toEqual(true);
        expect(outDocContents.includes("<table style=\"padding:0pt; -aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">")).toEqual(true);
        break;
      case aw.Saving.HtmlVersion.Xhtml:
        expect(outDocContents.includes("<a name=\"_Toc76372689\"></a>")).toEqual(true);
        expect(outDocContents.includes("<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">")).toEqual(true);
        expect(outDocContents.includes("<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border:0.5pt single #000000; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\"")).toEqual(true);
        break;
    }
    //ExEnd
  });


  test.each([false,
    true])('ExportXhtmlTransitional', (showDoctypeDeclaration) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportXhtmlTransitional
    //ExFor:HtmlSaveOptions.htmlVersion
    //ExFor:HtmlVersion
    //ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    options.htmlVersion = aw.Saving.HtmlVersion.Xhtml;
    options.exportXhtmlTransitional = showDoctypeDeclaration;
    options.prettyFormat = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportXhtmlTransitional.html", options);

    // Our document will only contain a DOCTYPE declaration heading if we have set the "ExportXhtmlTransitional" flag to "true".
    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportXhtmlTransitional.html").toString();

    if (showDoctypeDeclaration)
      expect(outDocContents.includes(
        "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\r\n" +
        "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
        "<html xmlns=\"http://www.w3.org/1999/xhtml\">")).toEqual(true);
    else
      expect(outDocContents.includes("<html>")).toEqual(true);
    //ExEnd
  });


  test('EpubHeadings', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.navigationMapLevel
    //ExSummary:Shows how to filter headings that appear in the navigation panel of a saved Epub document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Every paragraph that we format using a "Heading" style can serve as a heading.
    // Each heading may also have a heading level, determined by the number of its heading style.
    // The headings below are of levels 1-3.
    builder.paragraphFormat.style = builder.document.styles.at("Heading 1");
    builder.writeln("Heading #1");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 2");
    builder.writeln("Heading #2");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 3");
    builder.writeln("Heading #3");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 1");
    builder.writeln("Heading #4");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 2");
    builder.writeln("Heading #5");
    builder.paragraphFormat.style = builder.document.styles.at("Heading 3");
    builder.writeln("Heading #6");

    // Epub readers typically create a table of contents for their documents.
    // Each paragraph with a "Heading" style in the document will create an entry in this table of contents.
    // We can use the "NavigationMapLevel" property to set a maximum heading level. 
    // The Epub reader will not add headings with a level above the one we specify to the contents table.
    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Epub);
    options.navigationMapLevel = 2;

    // Our document has six headings, two of which are above level 2.
    // The table of contents for this document will have four entries.
    doc.save(base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", options);
    //ExEnd

    TestUtil.docPackageFileContainsString("<navLabel><text>Heading #1</text></navLabel>", 
      base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
    TestUtil.docPackageFileContainsString("<navLabel><text>Heading #2</text></navLabel>", 
      base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
    TestUtil.docPackageFileContainsString("<navLabel><text>Heading #4</text></navLabel>", 
      base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");
    TestUtil.docPackageFileContainsString("<navLabel><text>Heading #5</text></navLabel>", 
      base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "toc.ncx");

    expect(() =>
    {
      TestUtil.docPackageFileContainsString("<navLabel><text>Heading #3</text></navLabel>", 
        base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
    }).toThrow("");

    expect(() =>
    {
      TestUtil.docPackageFileContainsString("<navLabel><text>Heading #6</text></navLabel>", 
        base.artifactsDir + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
    }).toThrow("");
  });


  test('Doc2EpubSaveOptions', () => {
    //ExStart
    //ExFor:DocumentSplitCriteria
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.#ctor
    //ExFor:HtmlSaveOptions.encoding
    //ExFor:HtmlSaveOptions.documentSplitCriteria
    //ExFor:HtmlSaveOptions.exportDocumentProperties
    //ExFor:HtmlSaveOptions.saveFormat
    //ExFor:SaveOptions
    //ExFor:SaveOptions.saveFormat
    //ExSummary:Shows how to use a specific encoding when saving a document to .epub.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Use a SaveOptions object to specify the encoding for a document that we will save.
    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.saveFormat = aw.SaveFormat.Epub;
    saveOptions.encoding = "utf-8";

    // By default, an output .epub document will have all its contents in one HTML part.
    // A split criterion allows us to segment the document into several HTML parts.
    // We will set the criteria to split the document into heading paragraphs.
    // This is useful for readers who cannot read HTML files more significant than a specific size.
    saveOptions.documentSplitCriteria = aw.Saving.DocumentSplitCriteria.HeadingParagraph;

    // Specify that we want to export document properties.
    saveOptions.exportDocumentProperties = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
    //ExEnd
  });


  test.each([false,
    true])('ContentIdUrls', (exportCidUrlsForMhtmlResources) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportCidUrlsForMhtmlResources
    //ExSummary:Shows how to enable content IDs for output MHTML documents.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // Setting this flag will replace "Content-Location" tags
    // with "Content-ID" tags for each resource from the input document.
    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Mhtml);
    options.exportCidUrlsForMhtmlResources = exportCidUrlsForMhtmlResources;
    options.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    options.exportFontResources = true;
    options.prettyFormat = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.ContentIdUrls.mht", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.ContentIdUrls.mht").toString();

    if (exportCidUrlsForMhtmlResources)
    {
      expect(outDocContents.includes("Content-ID: <document.html>")).toEqual(true);
      expect(outDocContents.includes("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />")).toEqual(true);
      expect(outDocContents.includes("@font-face { font-family:'Arial Black'; font-weight:bold; src:url('cid:arib=\r\nlk.ttf') }")).toEqual(true);
      expect(outDocContents.includes("<img src=3D\"cid:image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />")).toEqual(true);
    }
    else
    {
      expect(outDocContents.includes("Content-Location: document.html")).toEqual(true);
      expect(outDocContents.includes("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />")).toEqual(true);
      expect(outDocContents.includes("@font-face { font-family:'Arial Black'; font-weight:bold; src:url('ariblk.t=\r\ntf') }")).toEqual(true);
      expect(outDocContents.includes("<img src=3D\"image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />")).toEqual(true);
    }
    //ExEnd
  });


  test.each([false,
    true])('DropDownFormField', (exportDropDownFormFieldAsText) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportDropDownFormFieldAsText
    //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a document builder to insert a combo box with the value "Two" selected.
    builder.insertComboBox("MyComboBox", [ "One", "Two", "Three" ], 1);

    // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
    // control how saving the document to HTML treats drop-down combo boxes.
    // Setting it to "true" will convert each combo box into simple text
    // that displays the combo box's currently selected value, effectively freezing it.
    // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportDropDownFormFieldAsText = exportDropDownFormFieldAsText;    

    doc.save(base.artifactsDir + "HtmlSaveOptions.DropDownFormField.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.DropDownFormField.html").toString();

    if (exportDropDownFormFieldAsText)
      expect(outDocContents.includes("<span>Two</span>")).toBe(true);
    else
      expect(outDocContents.includes(
        "<select name=\"MyComboBox\">" +
          "<option>One</option>" +
          "<option selected=\"selected\">Two</option>" +
          "<option>Three</option>" +
        "</select>")).toBe(true);
    //ExEnd
  });


  test.each([false,
    true])('ExportImagesAsBase64', (exportImagesAsBase64) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportFontsAsBase64
    //ExFor:HtmlSaveOptions.exportImagesAsBase64
    //ExSummary:Shows how to save a .html document with images embedded inside it.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.HtmlSaveOptions();
    options.exportImagesAsBase64 = exportImagesAsBase64;
    options.prettyFormat = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportImagesAsBase64.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportImagesAsBase64.html").toString();

    expect(exportImagesAsBase64
      ? outDocContents.includes("<img src=\"data:image/png;base64")
      : outDocContents.includes("<img src=\"HtmlSaveOptions.exportImagesAsBase64.001.png\"")).toBe(true);
    //ExEnd
  });


  test('ExportFontsAsBase64', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportFontsAsBase64
    //ExFor:HtmlSaveOptions.exportImagesAsBase64
    //ExSummary:Shows how to embed fonts inside a saved HTML document.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.HtmlSaveOptions();
    options.exportFontsAsBase64 = true;
    options.cssStyleSheetType = aw.Saving.CssStyleSheetType.Embedded;
    options.prettyFormat = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportFontsAsBase64.html", options);
    //ExEnd
  });


  test.each([false,
    true])('ExportLanguageInformation', (exportLanguageInformation) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportLanguageInformation
    //ExSummary:Shows how to preserve language information when saving to .html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use the builder to write text while formatting it in different locales.
    builder.font.localeId = 1033; // en-US
    builder.writeln("Hello world!");

    builder.font.localeId = 2057; // en-GB
    builder.writeln("Hello again!");

    builder.font.localeId = 1049;// ru-RU
    builder.write("Привет, мир!");

    // When saving the document to HTML, we can pass a SaveOptions object
    // to either preserve or discard each formatted text's locale.
    // If we set the "ExportLanguageInformation" flag to "true",
    // the output HTML document will contain the locales in "lang" attributes of <span> tags.
    // If we set the "ExportLanguageInformation" flag to "false',
    // the text in the output HTML document will not contain any locale information.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportLanguageInformation = exportLanguageInformation;
    options.prettyFormat = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportLanguageInformation.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportLanguageInformation.html").toString();

    if (exportLanguageInformation)
    {
      expect(outDocContents.includes("<span>Hello world!</span>")).toEqual(true);
      expect(outDocContents.includes("<span lang=\"en-GB\">Hello again!</span>")).toEqual(true);
      expect(outDocContents.includes("<span lang=\"ru-RU\">Привет, мир!</span>")).toEqual(true);
    }
    else
    {
      expect(outDocContents.includes("<span>Hello world!</span>")).toEqual(true);
      expect(outDocContents.includes("<span>Hello again!</span>")).toEqual(true);
      expect(outDocContents.includes("<span>Привет, мир!</span>")).toEqual(true);
    }
    //ExEnd
  });


  test.each([aw.Saving.ExportListLabels.AsInlineText,
    aw.Saving.ExportListLabels.Auto,
    aw.Saving.ExportListLabels.ByHtmlTags])('List', (exportListLabels) => {
    //ExStart
    //ExFor:ExportListLabels
    //ExFor:HtmlSaveOptions.exportListLabels
    //ExSummary:Shows how to configure list exporting to HTML.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let list = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);
    builder.listFormat.list = list;
            
    builder.writeln("Default numbered list item 1.");
    builder.writeln("Default numbered list item 2.");
    builder.listFormat.listIndent();
    builder.writeln("Default numbered list item 3.");
    builder.listFormat.removeNumbers();

    list = doc.lists.add(aw.Lists.ListTemplate.OutlineHeadingsLegal);
    builder.listFormat.list = list;

    builder.writeln("Outline legal heading list item 1.");
    builder.writeln("Outline legal heading list item 2.");
    builder.listFormat.listIndent();
    builder.writeln("Outline legal heading list item 3.");
    builder.listFormat.listIndent();
    builder.writeln("Outline legal heading list item 4.");
    builder.listFormat.listIndent();
    builder.writeln("Outline legal heading list item 5.");
    builder.listFormat.removeNumbers();

    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide which HTML elements the document will use to represent lists.
    // Setting the "ExportListLabels" property to "ExportListLabels.AsInlineText"
    // will create lists by formatting spans.
    // Setting the "ExportListLabels" property to "ExportListLabels.Auto" will use the <p> tag
    // to build lists in cases when using the <ol> and <li> tags may cause loss of formatting.
    // Setting the "ExportListLabels" property to "ExportListLabels.ByHtmlTags"
    // will use <ol> and <li> tags to build all lists.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportListLabels = exportListLabels;

    doc.save(base.artifactsDir + "HtmlSaveOptions.list.html", options);
    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.list.html").toString();

    switch (exportListLabels)
    {
      case aw.Saving.ExportListLabels.AsInlineText:
        expect(outDocContents.includes(
          "<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">" +
            "<span style=\"-aw-import:ignore\">" +
              "<span>a.</span>" +
              "<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
            "</span>" +
            "<span>Default numbered list item 3.</span>" +
          "</p>")).toBe(true);

        expect(outDocContents.includes(
          "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
            "<span style=\"-aw-import:ignore\">" +
              "<span>2.1.1.1</span>" +
              "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
            "</span>" +
            "<span>Outline legal heading list item 5.</span>" +
          "</p>")).toBe(true);
        break;
      case aw.Saving.ExportListLabels.Auto:
        expect(outDocContents.includes(
          "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
              "<span>Default numbered list item 3.</span>" +
            "</li>" +
          "</ol>")).toBe(true);
        break;
      case aw.Saving.ExportListLabels.ByHtmlTags:
        expect(outDocContents.includes(
          "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
              "<span>Default numbered list item 3.</span>" +
            "</li>" +
          "</ol>")).toBe(true);
        break;
    }
    //ExEnd
  });


  test.each([false,
    true])('ExportPageMargins(%o)', (exportPageMargins) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportPageMargins
    //ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a builder to insert a shape with no wrapping.
    let shape = builder.insertShape(aw.Drawing.ShapeType.Cube, 200, 200);

    shape.relativeHorizontalPosition = aw.Drawing.RelativeHorizontalPosition.Page;
    shape.relativeVerticalPosition = aw.Drawing.RelativeVerticalPosition.Page;
    shape.wrapType = aw.Drawing.WrapType.None;

    // Negative shape position values may place the shape outside of page boundaries.
    // If we export this to HTML, the shape will appear truncated.
    shape.left = -150;

    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide whether to adjust the page to display out-of-bounds objects fully.
    // If we set the "ExportPageMargins" flag to "true", the shape will be fully visible in the output HTML.
    // If we set the "ExportPageMargins" flag to "false",
    // our document will display the shape truncated as we would see it in Microsoft Word.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportPageMargins = exportPageMargins;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportPageMargins.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportPageMargins.html").toString();

    if (exportPageMargins)
    {
      expect(outDocContents.includes("<style type=\"text/css\">div.Section_1 { margin:72pt }</style>")).toEqual(true);
      expect(outDocContents.includes("<div class=\"Section_1\"><p style=\"margin-top:0pt; margin-left:150pt; margin-bottom:0pt\">")).toEqual(true);
    }
    else
    {
      expect(outDocContents.includes("<style type=\"text/css\">")).toEqual(false);
      expect(outDocContents.includes("<div><p style=\"margin-top:0pt; margin-left:222pt; margin-bottom:0pt\">")).toEqual(true);
     }
    //ExEnd
  });


  test.each([false,
    true])('ExportPageSetup(%o)', (exportPageSetup) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportPageSetup
    //ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");

    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.topMargin = 36.0;
    pageSetup.bottomMargin = 36.0;
    pageSetup.paperSize = aw.PaperSize.A5;

    // When saving the document to HTML, we can pass a SaveOptions object
    // to decide whether to preserve or discard page setup settings.
    // If we set the "ExportPageSetup" flag to "true", the output HTML document will contain our page setup configuration.
    // If we set the "ExportPageSetup" flag to "false", the save operation will discard our page setup settings
    // for the first section, and both sections will look identical.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportPageSetup = exportPageSetup;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportPageSetup.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportPageSetup.html").toString();

    if (exportPageSetup)
    {
      expect(outDocContents).toEqual(expect.stringContaining("@page Section_1"));
      expect(outDocContents).toEqual(expect.stringContaining("<div class=\"Section_1\">"));
    }
    else
    {
      expect(outDocContents.includes("style type=\"text/css\">")).toEqual(false);
      expect(outDocContents.includes(
        "<div>" +
          "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
            "<span>Section 1</span>" +
          "</p>" +
        "</div>")).toBe(true);
    }
    //ExEnd
  });


  test.each([false,
    true])('RelativeFontSize', (exportRelativeFontSize) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportRelativeFontSize
    //ExSummary:Shows how to use relative font sizes when saving to .html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Default font size, ");
    builder.font.size = 24;
    builder.writeln("2x default font size,");
    builder.font.size = 96;
    builder.write("8x default font size");

    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine whether to use relative or absolute font sizes.
    // Set the "ExportRelativeFontSize" flag to "true" to declare font sizes
    // using the "em" measurement unit, which is a factor that multiplies the current font size. 
    // Set the "ExportRelativeFontSize" flag to "false" to declare font sizes
    // using the "pt" measurement unit, which is the font's absolute size in points.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportRelativeFontSize = exportRelativeFontSize;

    doc.save(base.artifactsDir + "HtmlSaveOptions.RelativeFontSize.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.RelativeFontSize.html").toString();

    if (exportRelativeFontSize)
    {
      expect(outDocContents.includes(
        "<body style=\"font-family:'Times New Roman'\">" +
          "<div>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
              "<span>Default font size, </span>" +
            "</p>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" +
              "<span>2x default font size,</span>" +
            "</p>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" +
              "<span>8x default font size</span>" +
            "</p>" +
          "</div>" +
        "</body>")).toBe(true);
    }
    else
    {
      expect(outDocContents.includes(
        "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
          "<div>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
              "<span>Default font size, </span>" +
            "</p>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" +
              "<span>2x default font size,</span>" +
            "</p>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" +
              "<span>8x default font size</span>" +
            "</p>" +
          "</div>" +
        "</body>")).toBe(true);
    }
    //ExEnd
  });


  test.each([false,
    true])('ExportShape', (exportShapesAsSvg) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportShapesAsSvg
    //ExSummary:Shows how to export shape as scalable vector graphics.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let textBox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100.0, 60.0);
    builder.moveTo(textBox.firstParagraph);
    builder.write("My text box");

    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation will export text box shapes.
    // If we set the "ExportTextBoxAsSvg" flag to "true",
    // the save operation will convert shapes with text into SVG objects.
    // If we set the "ExportTextBoxAsSvg" flag to "false",
    // the save operation will convert shapes with text into images.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportShapesAsSvg = exportShapesAsSvg;

    doc.save(base.artifactsDir + "HtmlSaveOptions.ExportTextBox.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.ExportTextBox.html").toString();

    if (exportShapesAsSvg)
    {
      expect(outDocContents.includes(
        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">")).toBe(true);
    }
    else
    {
      expect(outDocContents.includes(
        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
          "<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
          "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
        "</p>")).toBe(true);
    }
    //ExEnd
  });


  test.each([false,
    true])('RoundTripInformation', (exportRoundtripInformation) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportRoundtripInformation
    //ExSummary:Shows how to preserve hidden elements when converting to .html.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
    // or footnotes will be either removed or converted to plain text and effectively be lost.
    // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements.

    // When we save the document to HTML, we can pass a SaveOptions object to determine
    // how the saving operation will export document elements that HTML does not support or use,
    // such as hidden bookmarks and original shape positions.
    // If we set the "ExportRoundtripInformation" flag to "true", the save operation will preserve these elements.
    // If we set the "ExportRoundTripInformation" flag to "false", the save operation will discard these elements.
    // We will want to preserve such elements if we intend to load the saved HTML using Aspose.words,
    // as they could be of use once again.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportRoundtripInformation = exportRoundtripInformation;

    doc.save(base.artifactsDir + "HtmlSaveOptions.RoundTripInformation.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.RoundTripInformation.html").toString();
    doc = new aw.Document(base.artifactsDir + "HtmlSaveOptions.RoundTripInformation.html");

    if (exportRoundtripInformation)
    {
      expect(outDocContents.includes("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">")).toEqual(true);
      expect(outDocContents.includes("<span style=\"-aw-import:ignore\">&#xa0;</span>")).toEqual(true);

      expect(outDocContents.includes(
                    "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single #000000; " +
                    "-aw-border-left:0.5pt single #000000; -aw-border-right:6pt single #000000; -aw-border-top:0.5pt single #000000\">")).toEqual(true);

      expect(outDocContents.includes(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">")).toEqual(true);

      expect(outDocContents.includes(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" " +
                    "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />")).toEqual(true);


      expect(outDocContents.includes(
                    "<span>Page number </span>" +
                    "<span style=\"-aw-field-start:true\"></span>" +
                    "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                    "<span style=\"-aw-field-separator:true\"></span>" +
                    "<span>1</span>" +
                    "<span style=\"-aw-field-end:true\"></span>")).toEqual(true);

      expect([...doc.range.fields].filter(f => f.type == aw.Fields.FieldType.FieldPage).length).toEqual(1);
    }
    else
    {
      expect(outDocContents.includes("<div style=\"clear:both\">")).toEqual(true);
      expect(outDocContents.includes("<span>&#xa0;</span>")).toEqual(true);

      expect(outDocContents.includes(
                    "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                    "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">")).toEqual(true);
                
      expect(outDocContents.includes(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">")).toEqual(true);

      expect(outDocContents.includes(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" />")).toEqual(true);

      expect(outDocContents.includes(
                    "<span>Page number 1</span>")).toEqual(true);

      expect([...doc.range.fields].filter(f => f.type == aw.Fields.FieldType.FieldPage).length).toEqual(0);
    }
    //ExEnd
  });


  test.skip.each([false,
    true])('ExportTocPageNumbers(%o) - TODO: Failed on true.', (exportTocPageNumbers) => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportTocPageNumbers
    //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
    // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
    // and the page number that contains the heading on the right.
    let fieldToc = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();

    builder.paragraphFormat.style = builder.document.styles.at("Heading 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Entry 1");
    builder.writeln("Entry 2");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Entry 3");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Entry 4");
    fieldToc.updatePageNumbers();
    doc.updateFields();

    // HTML documents do not have pages. If we save this document to HTML,
    // the page numbers that our TOC displays will have no meaning.
    // When we save the document to HTML, we can pass a SaveOptions object to omit these page numbers from the TOC.
    // If we set the "ExportTocPageNumbers" flag to "true",
    // each TOC entry will display the heading, separator, and page number, preserving its appearance in Microsoft Word.
    // If we set the "ExportTocPageNumbers" flag to "false",
    // the save operation will omit both the separator and page number and leave the heading for each entry intact.
    let options = new aw.Saving.HtmlSaveOptions();
    options.exportTocPageNumbers = exportTocPageNumbers;

    doc.save(base.artifactsDir + "HtmlSaveOptions.exportTocPageNumbers.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.exportTocPageNumbers.html").toString();

    if (exportTocPageNumbers)
    {
      expect(outDocContents.includes(
                    "<span>Entry 1</span>" +
                    "<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                    "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                    "<span>2</span>" +
                    "</p>")).toEqual(true);
    }
    else
    {
      expect(outDocContents.includes(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                    "<span>Entry 2</span>" +
                    "</p>")).toEqual(true);
    }
    //ExEnd
  });


  test.each([0,
    1000000,
    2147483647])('FontSubsetting', (fontResourcesSubsettingSizeThreshold) => {
    //ExStart
    //ExFor:HtmlSaveOptions.fontResourcesSubsettingSizeThreshold
    //ExSummary:Shows how to work with font subsetting.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Hello world!");
    builder.font.name = "Times New Roman";
    builder.writeln("Hello world!");
    builder.font.name = "Courier New";
    builder.writeln("Hello world!");

    // When we save the document to HTML, we can pass a SaveOptions object configure font subsetting.
    // Suppose we set the "ExportFontResources" flag to "true" and also name a folder in the "FontsFolder" property.
    // In that case, the saving operation will create that folder and place a .ttf file inside
    // that folder for each font that our document uses.
    // Each .ttf file will contain that font's entire glyph set,
    // which may potentially result in a very large file that accompanies the document.
    // When we apply subsetting to a font, its exported raw data will only contain the glyphs that the document is
    // using instead of the entire glyph set. If the text in our document only uses a small fraction of a font's
    // glyph set, then subsetting will significantly reduce our output documents' size.
    // We can use the "FontResourcesSubsettingSizeThreshold" property to define a .ttf file size, in bytes.
    // If an exported font creates a size bigger file than that, then the save operation will apply subsetting to that font. 
    // Setting a threshold of 0 applies subsetting to all fonts,
    // and setting it to "int.MaxValue" effectively disables subsetting.
    let fontsFolder = base.artifactsDir + "HtmlSaveOptions.FontSubsetting.Fonts";

    let options = new aw.Saving.HtmlSaveOptions();
    options.exportFontResources = true;
    options.fontsFolder = fontsFolder;
    options.fontResourcesSubsettingSizeThreshold = fontResourcesSubsettingSizeThreshold;

    doc.save(base.artifactsDir + "HtmlSaveOptions.FontSubsetting.html", options);

    let fontFileNames = fs.readdirSync(fontsFolder).filter(s => s.endsWith(".ttf"));

    expect(fontFileNames.length).toEqual(3);

    for (let filename of fontFileNames)
    {
      // By default, the .ttf files for each of our three fonts will be over 700MB.
      // Subsetting will reduce them all to under 30MB.
      let fontFileInfo = fs.statSync(path.join(fontsFolder, filename));

      expect(fontFileInfo.size > 700000 || fontFileInfo.size < 30000).toEqual(true);
      expect(Math.max(fontResourcesSubsettingSizeThreshold, 30000) > fontFileInfo.size).toEqual(true);
    }
    //ExEnd
  });


  test.each([aw.Saving.HtmlMetafileFormat.Png,
    aw.Saving.HtmlMetafileFormat.Svg,
    aw.Saving.HtmlMetafileFormat.EmfOrWmf])('MetafileFormat', (htmlMetafileFormat) => {
    //ExStart
    //ExFor:HtmlMetafileFormat
    //ExFor:HtmlSaveOptions.metafileFormat
    //ExFor:HtmlLoadOptions.convertSvgToEmf
    //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
    let html = 
      `<html>
        <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
          <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
        </svg>
      </html>`

    // Use 'ConvertSvgToEmf' to turn back the legacy behavior
    // where all SVG images loaded from an HTML document were converted to EMF.
    // Now SVG images are loaded without conversion
    // if the MS Word version specified in load options supports SVG images natively.
    let loadOptions = new aw.Loading.HtmlLoadOptions();
    loadOptions.convertSvgToEmf = true;

    let doc = new aw.Document(Buffer.from(html, 'utf-8'), loadOptions);

    // This document contains a <svg> element in the form of text.
    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation handles this object.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
    // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
    let options = new aw.Saving.HtmlSaveOptions();
    options.metafileFormat = htmlMetafileFormat;

    doc.save(base.artifactsDir + "HtmlSaveOptions.metafileFormat.html", options);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.metafileFormat.html").toString();

    switch (htmlMetafileFormat)
    {
      case aw.Saving.HtmlMetafileFormat.Png:
        expect(outDocContents.includes(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.metafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>")).toEqual(true);
        break;
      case aw.Saving.HtmlMetafileFormat.Svg:
        expect(outDocContents.includes(
                        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">")).toEqual(true);
        break;
      case aw.Saving.HtmlMetafileFormat.EmfOrWmf:
        expect(outDocContents.includes(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.metafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>")).toEqual(true);
        break;
    }
    //ExEnd
  });


  test.each([aw.Saving.HtmlOfficeMathOutputMode.Image,
    aw.Saving.HtmlOfficeMathOutputMode.MathML,
    aw.Saving.HtmlOfficeMathOutputMode.Text])('OfficeMathOutputMode', (htmlOfficeMathOutputMode) => {
    //ExStart
    //ExFor:HtmlOfficeMathOutputMode
    //ExFor:HtmlSaveOptions.officeMathOutputMode
    //ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    // When we save the document to HTML, we can pass a SaveOptions object
    // to determine how the saving operation handles OfficeMath objects.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Image"
    // will render each OfficeMath object into an image.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.MathML"
    // will convert each OfficeMath object into MathML.
    // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Text"
    // will represent each OfficeMath formula using plain HTML text.
    let options = new aw.Saving.HtmlSaveOptions();
    options.officeMathOutputMode = htmlOfficeMathOutputMode;

    doc.save(base.artifactsDir + "HtmlSaveOptions.officeMathOutputMode.html", options);
    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.officeMathOutputMode.html").toString();

    switch (htmlOfficeMathOutputMode)
    {
      case aw.Saving.HtmlOfficeMathOutputMode.Image:
        expect(outDocContents.search('<p style="margin-top:0pt; margin-bottom:10pt">' + 
           '<img src="HtmlSaveOptions.officeMathOutputMode.001.png" width="163" height="19" alt="" style="vertical-align:middle; ' +
           '-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline" />' +
           '</p>')).not.toEqual(-1);
        break;
      case aw.Saving.HtmlOfficeMathOutputMode.MathML:
        expect(outDocContents.search(
          '<p style="margin-top:0pt; margin-bottom:10pt; text-align:center">' +
            '<math xmlns="http://www.w3.org/1998/Math/MathML">' +
              '<mi>i</mi>' +
              '<mo>[+]</mo>' +
              '<mi>b</mi>' +
              '<mo>-</mo>' +
              '<mi>c</mi>' +
              '<mo>≥</mo>' +
              '.*' +
            '</math>' +
          '</p>')).not.toEqual(-1);
        break;
      case aw.Saving.HtmlOfficeMathOutputMode.Text:
        expect(outDocContents.search('<p style="margin-top:0pt; margin-bottom:10pt; text-align:center">' +
           `<span style="font-family:'Cambria Math'">i[+]b-c≥iM[+]bM-cM </span>` +
           '</p>')).not.toEqual(-1);
        break;
    }
    //ExEnd
  });


  test.each([false,
    true])('ScaleImageToShapeSize', (scaleImageToShapeSize) => {
    //ExStart
    //ExFor:HtmlSaveOptions.scaleImageToShapeSize
    //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a shape which contains an image, and then make that shape considerably smaller than the image.
    let imageShape = builder.insertImage(base.imageDir + "Transparent background logo.png");
    imageShape.width = 50;
    imageShape.height = 50;

    // Saving a document that contains shapes with images to HTML will create an image file in the local file system
    // for each such shape. The output HTML document will use <image> tags to link to and display these images.
    // When we save the document to HTML, we can pass a SaveOptions object to determine
    // whether to scale all images that are inside shapes to the sizes of their shapes.
    // Setting the "ScaleImageToShapeSize" flag to "true" will shrink every image
    // to the size of the shape that contains it, so that no saved images will be larger than the document requires them to be.
    // Setting the "ScaleImageToShapeSize" flag to "false" will preserve these images' original sizes,
    // which will take up more space in exchange for preserving image quality.
    let options = new aw.Saving.HtmlSaveOptions();
    options.scaleImageToShapeSize = scaleImageToShapeSize;

    doc.save(base.artifactsDir + "HtmlSaveOptions.scaleImageToShapeSize.html", options);
    //ExEnd

    var testedImageLength = fs.statSync(base.artifactsDir + "HtmlSaveOptions.scaleImageToShapeSize.001.png").size;

    if (scaleImageToShapeSize)
      expect(testedImageLength < 6200).toEqual(true);
    else
      expect(testedImageLength < 16000).toEqual(true);
  });


  test('ImageFolder', () => {
    //ExStart
    //ExFor:HtmlSaveOptions
    //ExFor:HtmlSaveOptions.exportTextInputFormFieldAsText
    //ExFor:HtmlSaveOptions.imagesFolder
    //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let imagesDir = path.join(base.artifactsDir, "SaveHtmlWithOptions");

    if (fs.existsSync(imagesDir)) {
      fs.rmSync(imagesDir, { recursive: true, force: true });
    }

    fs.mkdirSync(imagesDir, { recursive: true });
    
    // Set an option to export form fields as plain text instead of HTML input elements.
    let options = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    options.exportTextInputFormFieldAsText = true;
    options.imagesFolder = imagesDir;
    
    doc.save(base.artifactsDir + "HtmlSaveOptions.SaveHtmlWithOptions.html", options);
    //ExEnd

    expect(fs.existsSync(base.artifactsDir + "HtmlSaveOptions.SaveHtmlWithOptions.html")).toEqual(true);
    expect(fs.readdirSync(imagesDir).length).toEqual(9);

    fs.rmSync(imagesDir, { recursive: true, force: true });
  });


  //ExStart
  //ExFor:ImageSavingArgs.CurrentShape
  //ExFor:ImageSavingArgs.Document
  //ExFor:ImageSavingArgs.ImageStream
  //ExFor:ImageSavingArgs.IsImageAvailable
  //ExFor:ImageSavingArgs.KeepImageStreamOpen
  //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
  test.skip('ImageSavingCallback - TODO: imageSavingCallback not supported', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // When we save the document to HTML, we can pass a SaveOptions object to designate a callback
    // to customize the image saving process.
    let options = new aw.Saving.HtmlSaveOptions();
    options.imageSavingCallback = new ImageShapePrinter();

    doc.save(base.artifactsDir + "HtmlSaveOptions.imageSavingCallback.html", options);
  });

/*
    /// <summary>
    /// Prints the properties of each image as the saving process saves it to an image file in the local file system
    /// during the exporting of a document to HTML.
    /// </summary>
  private class ImageShapePrinter : IImageSavingCallback
  {
    void aw.Saving.IImageSavingCallback.imageSaving(ImageSavingArgs args)
    {
      args.keepImageStreamOpen = false;
      expect(args.isImageAvailable).toEqual(true);

      console.log(`${args.document.originalFileName.split('\\').Last()} Image #${++mImageCount}`);

      let layoutCollector = new aw.Layout.LayoutCollector(args.document);

      console.log(`\tOn page:\t${layoutCollector.getStartPageIndex(args.currentShape)}`);
      console.log(`\tDimensions:\t${args.currentShape.bounds}`);
      console.log(`\tAlignment:\t${args.currentShape.verticalAlignment}`);
      console.log(`\tWrap type:\t${args.currentShape.wrapType}`);
      console.log(`Output filename:\t${args.imageFileName}\n`);
    }

    private int mImageCount;
  }
    //ExEnd
*/    

  test.each([true,
    false])('PrettyFormat', (usePrettyFormat) => {
    //ExStart
    //ExFor:SaveOptions.prettyFormat
    //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    let htmlOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    htmlOptions.prettyFormat = usePrettyFormat;

    doc.save(base.artifactsDir + "HtmlSaveOptions.prettyFormat.html", htmlOptions);

    // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
    let html = fs.readFileSync(base.artifactsDir + "HtmlSaveOptions.prettyFormat.html").toString();

    if (usePrettyFormat)
      expect(html).toEqual(
        "<html>\r\n" +
              "\t<head>\r\n" +
                "\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n" +
                "\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\r\n" +
                `\t\t<meta name=\"generator\" content=\"${aw.BuildVersionInfo.product} ${aw.BuildVersionInfo.version}\" />\r\n` +
                "\t\t<title>\r\n" +
                "\t\t</title>\r\n" +
              "\t</head>\r\n" +
              "\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\r\n" +
                "\t\t<div>\r\n" +
                  "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                    "\t\t\t\t<span>Hello world!</span>\r\n" +
                  "\t\t\t</p>\r\n" +
                  "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                    "\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\r\n" +
                  "\t\t\t</p>\r\n" +
                "\t\t</div>\r\n" +
              "\t</body>\r\n</html>");
    else
      expect(html).toEqual(
        "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
            "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
            `<meta name=\"generator\" content=\"${aw.BuildVersionInfo.product} ${aw.BuildVersionInfo.version}\" /><title></title></head>` +
            "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
            "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
            "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>");
    //ExEnd
  });


  //ExStart
  //ExFor:SaveOptions.ProgressCallback
  //ExFor:IDocumentSavingCallback
  //ExFor:IDocumentSavingCallback.Notify(DocumentSavingArgs)
  //ExFor:DocumentSavingArgs.EstimatedProgress
  //ExFor:DocumentSavingArgs
  //ExSummary:Shows how to manage a document while saving to html.
  test.skip.each([[aw.SaveFormat.Html, "html"],
    [aw.SaveFormat.Mhtml, "mhtml"],
    [aw.SaveFormat.Epub, "epub"]])('ProgressCallback - TODO: progressCallback not supported', (saveFormat, ext) => {
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // Following formats are supported: Html, Mhtml, Epub.
    let saveOptions = new aw.Saving.HtmlSaveOptions(saveFormat);
    saveOptions.progressCallback = new SavingProgressCallback();

    expect(() =>
      doc.save(base.artifactsDir + `HtmlSaveOptions.progressCallback.${ext}`, saveOptions)).toThrow("OperationCanceledException");
  });


  /*
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
    private const double MaxDuration = 0.1;
  }
    //ExEnd
  */

  test.each([aw.SaveFormat.Mobi,
    aw.SaveFormat.Azw3])('MobiAzw3DefaultEncoding', (saveFormat) => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.saveFormat = saveFormat;
    saveOptions.encoding = "us-ascii";

    let outputFileName = `${base.artifactsDir}HtmlSaveOptions.MobiDefaultEncoding${aw.FileFormatUtil.saveFormatToExtension(saveFormat)}`;
    doc.save(outputFileName);

    let encoding = TestUtil.getEncoding(outputFileName);
    expect(encoding).toEqual("UTF32");
  });


  test('HtmlReplaceBackslashWithYenSign', () => {
    //ExStart:HtmlReplaceBackslashWithYenSign
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:HtmlSaveOptions.replaceBackslashWithYenSign
    //ExSummary:Shows how to replace backslash characters with yen signs (Html).
    let doc = new aw.Document(base.myDir + "Korean backslash symbol.docx");

    // By default, Aspose.words mimics MS Word's behavior and doesn't replace backslash characters with yen signs in
    // generated HTML documents. However, previous versions of Aspose.words performed such replacements in certain
    // scenarios. This flag enables backward compatibility with previous versions of Aspose.words.
    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.replaceBackslashWithYenSign = true;

    doc.save(base.artifactsDir + "HtmlSaveOptions.replaceBackslashWithYenSign.html", saveOptions);
    //ExEnd:HtmlReplaceBackslashWithYenSign
  });


});
