// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const path = require('path');
const DocumentHelper = require('./DocumentHelper');

describe("ExHtmlFixedSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.skip('UseEncoding: WORDSNODEJS-96', () => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.encoding
    //ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello World!");

    // The default encoding is UTF-8. If we want to represent our document using a different encoding,
    // we can use a SaveOptions object to set a specific encoding.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.encoding = "ascii";

    expect(htmlFixedSaveOptions.encoding).toEqual("US-ASCII");

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
    //ExEnd

    expect(fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.UseEncoding.html", {"encoding": "ascii"}).includes("content=\"text/html; charset=us-ascii\"")).toBeTruthy()
  });


  test('GetEncoding', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();

    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.encoding = "utf-16";

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
  });


  test.each([true, false])('ExportEmbeddedCss(%o)', (exportEmbeddedCss) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.exportEmbeddedCss
    //ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // When we export a document to html, Aspose.words will also create a CSS stylesheet to format the document with.
    // Setting the "ExportEmbeddedCss" flag to "true" save the CSS stylesheet to a .css file,
    // and link to the file from the html document using a <link> element.
    // Setting the flag to "false" will embed the CSS stylesheet within the Html document,
    // which will create only one file instead of two.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.exportEmbeddedCss = exportEmbeddedCss

    const outPath = path.join(base.artifactsDir, "HtmlFixedSaveOptions.exportEmbeddedCss");
    if (fs.existsSync(outPath))
      fs.rmSync(outPath, {recursive: true, force: true} );
     
    doc.save(outPath + ".html", htmlFixedSaveOptions);
    let outDocContents = fs.readFileSync(outPath + ".html").toString();

    if (exportEmbeddedCss)
    {
      expect(outDocContents.includes("<style type=\"text/css\">")).toBeTruthy();
      expect(fs.existsSync(path.join(outPath, "styles.css"))).toBeFalsy();
    }
    else
    {
      expect(outDocContents.includes("<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions.exportEmbeddedCss/styles.css\" media=\"all\" />")).toBeTruthy();
      expect(fs.existsSync(path.join(outPath, "styles.css"))).toBeTruthy();
    }
    //ExEnd
  });


  test.each([true, false])('ExportEmbeddedFonts(%o)', (exportEmbeddedFonts) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.exportEmbeddedFonts
    //ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
    let doc = new aw.Document(base.myDir + "Embedded font.docx");

    // When we export a document with embedded fonts to .html,
    // Aspose.words can place the fonts in two possible locations.
    // Setting the "ExportEmbeddedFonts" flag to "true" will store the raw data for embedded fonts within the CSS stylesheet,
    // in the "url" property of the "@font-face" rule. This may create a huge CSS stylesheet file
    // and reduce the number of external files that this HTML conversion will create.
    // Setting this flag to "false" will create a file for each font.
    // The CSS stylesheet will link to each font file using the "url" property of the "@font-face" rule.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.exportEmbeddedFonts = exportEmbeddedFonts;

    const outPath = path.join(base.artifactsDir, "HtmlFixedSaveOptions.exportEmbeddedFonts");
    if (fs.existsSync(outPath))
      fs.rmSync(outPath, {recursive: true, force: true} );

    doc.save(outPath + ".html", htmlFixedSaveOptions);
    let outDocContents = fs.readFileSync(path.join(outPath, "styles.css")).toString();

    if (exportEmbeddedFonts)
    {
      const patternEmbedded = /@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local\('☺'\), url\(.+\) format\('woff'\); }/;
      expect(patternEmbedded.test(outDocContents)).toBeTruthy();
      expect(fs.readdirSync(outPath).filter(f => f.endsWith(".woff")).length).toEqual(0);
    }
    else
    {
      expect(outDocContents.includes("@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local('☺'), url('font001.woff') format('woff'); }")).toBeTruthy();
      expect(fs.readdirSync(outPath).filter(f => f.endsWith(".woff")).length).toEqual(2);
    }
    //ExEnd
  });


  test.each([true, false])('ExportEmbeddedImages(%o)', (exportImages) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.exportEmbeddedImages
    //ExSummary:Shows how to determine where to store images when exporting a document to Html.
    let doc = new aw.Document(base.myDir + "Images.docx");

    // When we export a document with embedded images to .html,
    // Aspose.words can place the images in two possible locations.
    // Setting the "ExportEmbeddedImages" flag to "true" will store the raw data
    // for all images within the output HTML document, in the "src" attribute of <image> tags.
    // Setting this flag to "false" will create an image file in the local file system for every image,
    // and store all these files in a separate folder.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.exportEmbeddedImages = exportImages;

    const outPath = path.join(base.artifactsDir, "HtmlFixedSaveOptions.exportEmbeddedImages");
    if (fs.existsSync(outPath))
      fs.rmSync(outPath, {recursive: true, force: true} );

    doc.save(outPath + ".html", htmlFixedSaveOptions);
    let outDocContents = fs.readFileSync(outPath + ".html").toString();

    if (exportImages)
    {
      expect(fs.existsSync(path.join(outPath, "image001.jpeg"))).toBeFalsy();
      expect(new RegExp("<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />").test(outDocContents)).toBeTruthy();
    }
    else
    {
      expect(fs.existsSync(path.join(outPath, "image001.jpeg"))).toBeTruthy();
      expect(outDocContents.includes("<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" " +
        "src=\"HtmlFixedSaveOptions.exportEmbeddedImages/image001.jpeg\" />")).toBeTruthy();
    }
    //ExEnd
  });


  test.each([true, false])('ExportEmbeddedSvgs(%o)', (exportSvgs) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.exportEmbeddedSvg
    //ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
    let doc = new aw.Document(base.myDir + "Images.docx");

    // When we export a document with SVG objects to .html,
    // Aspose.words can place these objects in two possible locations.
    // Setting the "ExportEmbeddedSvg" flag to "true" will embed all SVG object raw data
    // within the output HTML, inside <image> tags.
    // Setting this flag to "false" will create a file in the local file system for each SVG object.
    // The HTML will link to each file using the "data" attribute of an <object> tag.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.exportEmbeddedSvg = exportSvgs;

    const outPath = path.join(base.artifactsDir, "HtmlFixedSaveOptions.ExportEmbeddedSvgs");
    if (fs.existsSync(outPath))
      fs.rmSync(outPath, {recursive: true, force: true} );

    doc.save(outPath + ".html", htmlFixedSaveOptions);
    let outDocContents = fs.readFileSync(outPath + ".html").toString();

    if (exportSvgs)
    {
      expect(fs.existsSync(path.join(outPath, "svg001.svg"))).toBeFalsy();
      const pattern = /<image id="image004" xlink:href=.+\/>/;      
      expect(pattern.test(outDocContents)).toBeTruthy();
    }
    else
    {
      expect(fs.existsSync(path.join(outPath, "svg001.svg"))).toBeTruthy();
      const pattern = /<object type="image\/svg\+xml" data="HtmlFixedSaveOptions.ExportEmbeddedSvgs\/svg001\.svg"><\/object>/;
      expect(pattern.test(outDocContents)).toBeTruthy();
    }
    //ExEnd
  });


  test.each([true, false])('ExportFormFields(%o)', (exportFormFields) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.exportFormFields
    //ExSummary:Shows how to export form fields to Html.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertCheckBox("CheckBox", false, 15);

    // When we export a document with form fields to .html,
    // there are two ways in which Aspose.words can export form fields.
    // Setting the "ExportFormFields" flag to "true" will export them as interactive objects.
    // Setting this flag to "false" will display form fields as plain text.
    // This will freeze them at their current value, and prevent the reader of our HTML document
    // from being able to interact with them.
    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.exportFormFields = exportFormFields;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.exportFormFields.html", htmlFixedSaveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.exportFormFields.html").toString();

    if (exportFormFields)
    {
      expect(outDocContents.includes("<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
        "<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />")).toBeTruthy();
    }
    else
    {
      expect(outDocContents.includes("<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
        "<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"")).toBeTruthy();
    }
    //ExEnd
  });


  test('AddCssClassNamesPrefix', () => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.cssClassNamesPrefix
    //ExFor:aw.Saving.HtmlFixedSaveOptions.saveFontFaceCssSeparately
    //ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
    let doc = new aw.Document(base.myDir + "Bookmarks.docx");

    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.cssClassNamesPrefix = "myprefix";
    htmlFixedSaveOptions.saveFontFaceCssSeparately = true;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html").toString();

    expect(outDocContents.includes("<div class=\"myprefixdiv myprefixpage\" style=\"width:595.3pt; height:841.9pt;\">" +
      "<div class=\"myprefixdiv\" style=\"left:85.05pt; top:36pt; clip:rect(0pt,510.25pt,74.95pt,-85.05pt);\">" +
      "<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294.73pt; top:0.36pt; line-height:12.29pt;\">")).toBeTruthy();

    outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css").toString();

    expect(outDocContents.includes(".myprefixdiv { position:absolute; } " +
      ".myprefixspan { position:absolute; white-space:pre; color:#000000; font-size:12pt; }")).toBeTruthy();
    //ExEnd
  });


  test.each([aw.Saving.HtmlFixedPageHorizontalAlignment.Center,
    aw.Saving.HtmlFixedPageHorizontalAlignment.Left,
    aw.Saving.HtmlFixedPageHorizontalAlignment.Right])('HorizontalAlignment(%o)', (pageHorizontalAlignment) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.pageHorizontalAlignment
    //ExFor:HtmlFixedPageHorizontalAlignment
    //ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let htmlFixedSaveOptions = new aw.Saving.HtmlFixedSaveOptions();
    htmlFixedSaveOptions.pageHorizontalAlignment = pageHorizontalAlignment;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.horizontalAlignment.html", htmlFixedSaveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.horizontalAlignment/styles.css").toString();

    switch (pageHorizontalAlignment)
    {
      case aw.Saving.HtmlFixedPageHorizontalAlignment.Center:
        expect(outDocContents.includes(".awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }")).toBeTruthy();
        break;
      case aw.Saving.HtmlFixedPageHorizontalAlignment.Left:
        expect(outDocContents.includes(".awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }")).toBeTruthy();
        break;
      case aw.Saving.HtmlFixedPageHorizontalAlignment.Right:
        expect(outDocContents.includes(".awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }")).toBeTruthy();
        break;
    }
    //ExEnd
  });


  test('PageMargins', () => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.pageMargins
    //ExSummary:Shows how to adjust page margins when saving a document to HTML.
    let doc = new aw.Document(base.myDir + "Document.docx");

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.pageMargins = 15;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.pageMargins.html", saveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.pageMargins/styles.css").toString();

    expect(outDocContents.includes(".awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }")).toBeTruthy();
    //ExEnd
  });


  test('PageMarginsException', () => {
    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    expect(() => saveOptions.pageMargins = -1).toThrow("value");
  });


  test.each([false, true])('OptimizeGraphicsOutput(%o)', (optimizeOutput) => {
    //ExStart
    //ExFor:aw.Saving.HtmlFixedSaveOptions.optimizeOutput
    //ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.optimizeOutput = optimizeOutput;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);

    // The size of the optimized version of the document is almost a third of the size of the unoptimized document.
    expect(Math.abs(fs.statSync(base.artifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").size - (optimizeOutput ? 61889 : 191770))).toBeLessThanOrEqual(200);
    //ExEnd
  });


  test.each([false, true])('UsingMachineFonts(%o)', (useTargetMachineFonts) => {
    //ExStart
    //ExFor:ExportFontFormat
    //ExFor:aw.Saving.HtmlFixedSaveOptions.fontFormat
    //ExFor:aw.Saving.HtmlFixedSaveOptions.useTargetMachineFonts
    //ExSummary:Shows how use fonts only from the target machine when saving a document to HTML.
    let doc = new aw.Document(base.myDir + "Bullet points with alternative font.docx");

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.exportEmbeddedCss = true;
    saveOptions.useTargetMachineFonts = useTargetMachineFonts;
    saveOptions.fontFormat = aw.Saving.ExportFontFormat.Ttf;
    saveOptions.exportEmbeddedFonts = false;

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

    let outDocContents = fs.readFileSync(base.artifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html").toString();

    if (useTargetMachineFonts)
      expect(outDocContents.includes("@font-face")).toBeFalsy();
    else
      expect(outDocContents.includes("@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local('☺'), " +
    "url('HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf') format('truetype'); }")).toBeTruthy();
    //ExEnd
  });


  /*  //ExStart
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs
    //ExFor:ResourceSavingArgs.Document
    //ExFor:ResourceSavingArgs.ResourceFileName
    //ExFor:ResourceSavingArgs.ResourceFileUri
    //ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML.
  test('ResourceSavingCallback', () => {
    let doc = new aw.Document(base.myDir + "Bullet points with alternative font.docx");

    let callback = new FontSavingCallback();

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions
    {
      ResourceSavingCallback = callback
    };

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

    console.log(callback.getText());
    TestResourceSavingCallback(callback); //ExSkip
  });


  private class FontSavingCallback : IResourceSavingCallback
  {
      /// <summary>
      /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
      /// </summary>
    public void ResourceSaving(ResourceSavingArgs args)
    {
      mText.AppendLine(`Original document URI:\t${args.document.originalFileName}`);
      mText.AppendLine(`Resource being saved:\t${args.resourceFileName}`);
      mText.AppendLine(`Full uri after saving:\t${args.resourceFileUri}\n`);
    }

    public string GetText()
    {
      return mText.toString();
    }

    private readonly StringBuilder mText = new StringBuilder();
  }
    //ExEnd

  private void TestResourceSavingCallback(FontSavingCallback callback)
  {
    expect(callback.getText().includes("font001.woff")).toEqual(true);
    expect(callback.getText().includes("styles.css")).toEqual(true);
  }

    //ExStart
    //ExFor:HtmlFixedSaveOptions
    //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
    //ExFor:HtmlFixedSaveOptions.ResourcesFolder
    //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:HtmlFixedSaveOptions.SaveFormat
    //ExFor:HtmlFixedSaveOptions.ShowPageBorder
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
    //ExFor:ResourceSavingArgs.ResourceStream
    //ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML.
  test('HtmlFixedResourceFolder', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let callback = new ResourceUriPrinter();

    let options = new aw.Saving.HtmlFixedSaveOptions
    {
      SaveFormat = aw.SaveFormat.HtmlFixed,
      ExportEmbeddedImages = false,
      ResourcesFolder = base.artifactsDir + "HtmlFixedResourceFolder",
      ResourcesFolderAlias = base.artifactsDir + "HtmlFixedResourceFolderAlias",
      ShowPageBorder = false,
      ResourceSavingCallback = callback
    };

    // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder.
    // We must ensure the folder exists before the streams can put their resources into it.
    Directory.CreateDirectory(options.resourcesFolderAlias);

    doc.save(base.artifactsDir + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

    console.log(callback.getText());

    string[] resourceFiles = Directory.GetFiles(base.artifactsDir + "HtmlFixedResourceFolderAlias");

    expect(Directory.Exists(base.artifactsDir + "HtmlFixedResourceFolder")).toEqual(false);
    expect(resourceFiles.count(f => f.EndsWith(".jpeg") || f.EndsWith(".png") || f.EndsWith(".css"))).toEqual(6);
    TestHtmlFixedResourceFolder(callback); //ExSkip
  });


    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
    /// </summary>
  private class ResourceUriPrinter : IResourceSavingCallback
  {
    void aw.Saving.IResourceSavingCallback.resourceSaving(ResourceSavingArgs args)
    {
        // If we set a folder alias in the SaveOptions object, we will be able to print it from here.
      mText.AppendLine(`Resource #${++mSavedResourceCount} \"${args.resourceFileName}\"`);

      string extension = Path.GetExtension(args.resourceFileName);
      switch (extension)
      {
        case ".ttf":
        case ".woff":
        {
            // By default, 'ResourceFileUri' uses system folder for fonts.
            // To avoid problems in other platforms you must explicitly specify the path for the fonts.
          args.resourceFileUri = base.artifactsDir + Path.DirectorySeparatorChar + args.resourceFileName;
          break;
        }
      }

      mText.AppendLine("\t" + args.resourceFileUri);

        // If we have specified a folder in the "ResourcesFolderAlias" property,
        // we will also need to redirect each stream to put its resource in that folder.
      args.resourceStream = new FileStream(args.resourceFileUri, FileMode.create);
      args.keepResourceStreamOpen = false;
    }

    public string GetText()
    {
      return mText.toString();
    }

    private int mSavedResourceCount;
    private readonly StringBuilder mText = new StringBuilder();
  }
    //ExEnd

  private void TestHtmlFixedResourceFolder(ResourceUriPrinter callback)
  {
    expect(Regex.Matches(callback.getText(), "Resource #").Count).toEqual(16);
  }*/

});
