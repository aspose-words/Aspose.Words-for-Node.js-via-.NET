// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const TestUtil = require('./TestUtil');

describe("ExHtmlLoadOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.each([true, false])('SupportVml(%o)', (supportVml) => {
    //ExStart
    //ExFor:HtmlLoadOptions
    //ExFor:HtmlLoadOptions.#ctor
    //ExFor:HtmlLoadOptions.supportVml
    //ExSummary:Shows how to support conditional comments while loading an HTML document.
    let loadOptions = new aw.Loading.HtmlLoadOptions();

    // If the value is true, then we take VML code into account while parsing the loaded document.
    loadOptions.supportVml = supportVml;

    // This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
    // and a different PNG image within "<![if !vml]>" tags.
    // If we set the "SupportVml" flag to "true", then Aspose.words will load the JPEG.
    // If we set this flag to "false", then Aspose.words will only load the PNG.
    let doc = new aw.Document(base.myDir + "VML conditional.htm", loadOptions);

    if (supportVml)
      expect(doc.getShape(0, true).imageData.imageType).toEqual(aw.Drawing.ImageType.Jpeg);
    else
      expect(doc.getShape(0, true).imageData.imageType).toEqual(aw.Drawing.ImageType.Png);
    //ExEnd

    let imageShape = doc.getShape(0, true);

    if (supportVml)
      TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, imageShape);
    else
      TestUtil.verifyImageInShape(400, 400, aw.Drawing.ImageType.Png, imageShape);
  });


    /*//Commented
    //ExStart
    //ExFor:HtmlLoadOptions.WebRequestTimeout
    //ExSummary:Shows how to set a time limit for web requests when loading a document with external resources linked by URLs.
    [NonParallelizable]
  test('WebRequestTimeout', () => {
    // Create a new HtmlLoadOptions object and verify its timeout threshold for a web request.
    let options = new aw.Loading.HtmlLoadOptions();

    // When loading an Html document with resources externally linked by a web address URL,
    // Aspose.words will abort web requests that fail to fetch the resources within this time limit, in milliseconds.
    expect(options.webRequestTimeout).toEqual(100000);

    // Set a WarningCallback that will record all warnings that occur during loading.
    let warningCallback = new ListDocumentWarnings();
    options.warningCallback = warningCallback;

    // Load such a document and verify that a shape with image data has been created.
    // This linked image will require a web request to load, which will have to complete within our time limit.
    string html = $@"
      <html>
        <img src=""{base.imageUrl}"" alt=""Aspose logo"" style=""width:400px;height:400px;"">
      </html>
    ";

    // Set an unreasonable timeout limit and try load the document again.
    options.webRequestTimeout = 0;
    let doc = new aw.Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), options);
    expect(warningCallback.Warnings().Count).toEqual(2);

    // Values in the image exception cache are actual for 10s.
    // So, wait for 10+s to get the cache non-actual for other tests.
    // See internal WORDSNET-25999 for details.
    Thread.Sleep(10100);

    // A web request that fails to obtain an image within the time limit will still produce an image.
    // However, the image will be the red 'x' that commonly signifies missing images.
    let imageShape = (Shape)doc.getShape(0, true);
    expect(imageShape.imageData.imageBytes.length).toEqual(924);

    // We can also configure a custom callback to pick up any warnings from timed out web requests.
    expect(warningCallback.Warnings()[0].Source).toEqual(aw.WarningSource.Html);
    expect(warningCallback.Warnings()[0].WarningType).toEqual(aw.WarningType.DataLoss);
    expect(warningCallback.Warnings()[0].Description).toEqual(`Couldn't load a resource from \'${base.imageUrl}\'.`);

    expect(warningCallback.Warnings()[1].Source).toEqual(aw.WarningSource.Html);
    expect(warningCallback.Warnings()[1].WarningType).toEqual(aw.WarningType.DataLoss);
    expect(warningCallback.Warnings()[1].Description).toEqual("Image has been replaced with a placeholder.");

    doc.save(base.artifactsDir + "HtmlLoadOptions.webRequestTimeout.docx");
  });


    /// <summary>
    /// Stores all warnings that occur during a document loading operation in a List.
    /// </summary>
  private class ListDocumentWarnings : IWarningCallback
  {
    public void Warning(WarningInfo info)
    {
      mWarnings.add(info);
    }

    public List<WarningInfo> Warnings() { 
      return mWarnings;
    }

    private readonly List<WarningInfo> mWarnings = new aw.Lists.List<WarningInfo>();
  }
    //ExEnd*/

  /*test('LoadHtmlFixed', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
    saveOptions.saveFormat = aw.SaveFormat.HtmlFixed;

    doc.save(base.artifactsDir + "HtmlLoadOptions.fixed.html", saveOptions);

    let loadOptions = new aw.Loading.HtmlLoadOptions();

    let warningCallback = new ListDocumentWarnings();
    loadOptions.warningCallback = warningCallback;

    doc = new aw.Document(base.artifactsDir + "HtmlLoadOptions.fixed.html", loadOptions);
    expect(warningCallback.warnings().count).toEqual(1);

    expect(warningCallback.warnings()[0].source).toEqual(aw.WarningSource.Html);
    expect(warningCallback.warnings()[0].warningType).toEqual(aw.WarningType.MajorFormattingLoss);
    expect(warningCallback.warnings()[0].description).toEqual("The document is fixed-page HTML. Its structure may not be loaded correctly.");
  });*/


  /*[AotTests.IgnoreAot("CertificateHolder.Create and DigitalSignatureUtil.Sign are not used in AW.NET directly.")]
  test('EncryptHtml', () => {
    //ExStart
    //ExSummary:Shows how to encrypt an Html document.
    // Create and sign an encrypted HTML document from an encrypted .docx.
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    let signOptions = new aw.DigitalSignatures.SignOptions
    {
      Comments = "Comment",
      SignTime = Date.now(),
      DecryptionPassword = "docPassword"
    };

    string inputFileName = base.myDir + "Encrypted.docx";
    string outputFileName = base.myDir + "HtmlLoadOptions.EncryptedHtml.html";
    aw.DigitalSignatures.DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd
  });
  //EndCommented*/


  test('BaseUri', () => {
    //ExStart
    //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
    //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
    //ExFor:LoadOptions.loadFormat
    //ExFor:LoadFormat
    //ExSummary:Shows how to specify a base URI when opening an html document.
    // Suppose we want to load an .html document that contains an image linked by a relative URI
    // while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
    // We can provide a base URI using an HtmlLoadOptions object. 
    let loadOptions = new aw.Loading.HtmlLoadOptions(aw.LoadFormat.Html, "", base.imageDir);

    expect(loadOptions.loadFormat).toEqual(aw.LoadFormat.Html);

    let doc = new aw.Document(base.myDir + "Missing image.html", loadOptions);

    // While the image was broken in the input .html, our custom base URI helped us repair the link.
    let imageShape = doc.getChildNodes(aw.NodeType.Shape, true).at(0).asShape();
    expect(imageShape.isImage).toEqual(true);

    // This output document will display the image that was missing.
    doc.save(base.artifactsDir + "HtmlLoadOptions.baseUri.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "HtmlLoadOptions.baseUri.docx");

    expect(doc.getShape(0, true).imageData.imageBytes.length > 0).toBeTruthy();
  });


  test('GetSelectAsSdt', () => {
    //ExStart
    //ExFor:HtmlLoadOptions.preferredControlType
    //ExFor:HtmlControlType
    //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
    const html = String.raw`
      <html>
        <select name='ComboBox' size='1'>
          <option value='val1'>item1</option>
          <option value='val2'></option>
        </select>
      </html>`;

    let htmlLoadOptions = new aw.Loading.HtmlLoadOptions();
    htmlLoadOptions.preferredControlType = aw.Loading.HtmlControlType.StructuredDocumentTag;

    fs.writeFileSync(base.artifactsDir + "ExHtmlLoadOptions.GetSelectAsSdt.html", html, {encoding: "utf8"})
    let doc = new aw.Document(base.artifactsDir + "ExHtmlLoadOptions.GetSelectAsSdt.html", htmlLoadOptions);
    let nodes = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);

    let tag = nodes.at(0).asStructuredDocumentTag();
    //ExEnd

    expect(tag.listItems.count).toEqual(2);

    expect(tag.listItems.at(0).value).toEqual("val1");
    expect(tag.listItems.at(1).value).toEqual("val2");
  });


  test('GetInputAsFormField', () => {
    const html = String.raw`
      <html>
        <input type='text' value='Input value text' />
      </html>`;

    // By default, "HtmlLoadOptions.preferredControlType" value is "HtmlControlType.FormField".
    // So, we do not set this value.
    let htmlLoadOptions = new aw.Loading.HtmlLoadOptions();

    fs.writeFileSync(base.artifactsDir + "ExHtmlLoadOptions.GetInputAsFormField.html", html, {encoding: "utf8"})
    let doc = new aw.Document(base.artifactsDir + "ExHtmlLoadOptions.GetInputAsFormField.html", htmlLoadOptions);
    let nodes = doc.getChildNodes(aw.NodeType.FormField, true);

    expect(nodes.count).toEqual(1);

    let formField = nodes.at(0).asFormField();
    expect(formField.result).toEqual("Input value text");
  });


/*#if !WORDS_AOT
  test.each([true,
    false])('IgnoreNoscriptElements', (bool ignoreNoscriptElements) => {
    //ExStart
    //ExFor:HtmlLoadOptions.ignoreNoscriptElements
    //ExSummary:Shows how to ignore <noscript> HTML elements.
    const string html = @"
      <html>
      <head>
        <title>NOSCRIPT</title>
        <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
        <script type=""text/javascript"">
          alert(""Hello, world!"");
        </script>
      </head>
      <body>
      <noscript><p>Your browser does not support JavaScript!</p></noscript>
      </body>
      </html>";

    let htmlLoadOptions = new aw.Loading.HtmlLoadOptions();
    htmlLoadOptions.ignoreNoscriptElements = ignoreNoscriptElements;

    let doc = new aw.Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), htmlLoadOptions);
    doc.save(base.artifactsDir + "HtmlLoadOptions.ignoreNoscriptElements.pdf");
    //ExEnd
  });


  test.each([true,
    false])('UsePdfDocumentForIgnoreNoscriptElements', (bool ignoreNoscriptElements) => {
    IgnoreNoscriptElements(ignoreNoscriptElements);

    Aspose.pdf.document pdfDoc = new Aspose.pdf.document(base.artifactsDir + "HtmlLoadOptions.ignoreNoscriptElements.pdf");
    let textAbsorber = new TextAbsorber();
    textAbsorber.Visit(pdfDoc);

    expect(textAbsorber.text).toEqual(ignoreNoscriptElements ? "" : "Your browser does not support JavaScript!");
  });

#endif*/

  test.each([aw.Loading.BlockImportMode.Preserve,
    aw.Loading.BlockImportMode.Merge])('BlockImport(%o)', (blockImportMode) => {
    //ExStart
    //ExFor:HtmlLoadOptions.blockImportMode
    //ExFor:BlockImportMode
    //ExSummary:Shows how properties of block-level elements are imported from HTML-based documents.
    const html = String.raw`
    <html>
      <div style='border:dotted'>
        <div style='border:solid'>
          <p>paragraph 1</p>
          <p>paragraph 2</p>
        </div>
      </div>
    </html>`;
    const filename = base.artifactsDir + "ExHtmlLoadOptions.BlockImport.html"
    fs.writeFileSync(filename, html, {encoding: "utf8"})

    let loadOptions = new aw.Loading.HtmlLoadOptions();
    // Set the new mode of import HTML block-level elements.
    loadOptions.blockImportMode = blockImportMode;

    let doc = new aw.Document(filename, loadOptions);
    doc.save(base.artifactsDir + "HtmlLoadOptions.BlockImport.docx");
    //ExEnd
  });


  test('FontFaceRules', () => {
    //ExStart:FontFaceRules
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:HtmlLoadOptions.supportFontFaceRules
    //ExSummary:Shows how to load declared "@font-face" rules.
    let loadOptions = new aw.Loading.HtmlLoadOptions();
    loadOptions.supportFontFaceRules = true;
    let doc = new aw.Document(base.myDir + "Html with FontFace.html", loadOptions);

    expect(doc.fontInfos.at(0).name).toEqual("Squarish Sans CT Regular");
    //ExEnd:FontFaceRules
  });


});
