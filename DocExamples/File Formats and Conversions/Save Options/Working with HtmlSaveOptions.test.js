// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("WorkingWithHtmlSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('ExportRoundtripInformation', () => {
    //ExStart:ExportRoundtripInformation
    //GistId:e4b272992a7c8fafdd7ff42f8c2de379
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportRoundtripInformation = true;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.exportRoundtripInformation.html", saveOptions);
    //ExEnd:ExportRoundtripInformation
  });


  test('ExportFontsAsBase64', () => {
    //ExStart:ExportFontsAsBase64
    //GistId:e4b272992a7c8fafdd7ff42f8c2de379
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportFontsAsBase64 = true;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.exportFontsAsBase64.html", saveOptions);
    //ExEnd:ExportFontsAsBase64
  });


  test('ExportResources', () => {
    //ExStart:ExportResources
    //GistId:e4b272992a7c8fafdd7ff42f8c2de379
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External;
    saveOptions.exportFontResources = true;
    saveOptions.resourceFolder = base.artifactsDir + "Resources";
    saveOptions.resourceFolderAlias = "http://example.com/resources";

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.ExportResources.html", saveOptions);
    //ExEnd:ExportResources
  });


  test('ConvertMetafilesToPng', () => {
    //ExStart:ConvertMetafilesToPng
    const html =
      `<html>
        <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
          <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
        </svg>
      </html>`;

    // Use 'ConvertSvgToEmf' to turn back the legacy behavior
    // where all SVG images loaded from an HTML document were converted to EMF.
    // Now SVG images are loaded without conversion
    // if the MS Word version specified in load options supports SVG images natively.
    let loadOptions = new aw.Loading.HtmlLoadOptions();
    loadOptions.convertSvgToEmf = true;
    let doc = new aw.Document(Buffer.from(html, 'utf8'), loadOptions);

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.metafileFormat = aw.Saving.HtmlMetafileFormat.Png;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.convertMetafilesToPng.html", saveOptions);
    //ExEnd:ConvertMetafilesToPng
  });


  test('ConvertMetafilesToSvg', () => {
    //ExStart:ConvertMetafilesToSvg
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    builder.write("Here is an SVG image: ");
    builder.insertHtml(
      `<svg height='210' width='500'>
      <polygon points='100,10 40,198 190,78 10,78 160,198' 
        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
    </svg>`);

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.metafileFormat = aw.Saving.HtmlMetafileFormat.Svg;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.ConvertMetafilesToSvg.html", saveOptions);
    //ExEnd:ConvertMetafilesToSvg
  });


  test('AddCssClassNamePrefix', () => {
    //ExStart:AddCssClassNamePrefix
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.cssStyleSheetType = aw.Saving.CssStyleSheetType.External, CssClassNamePrefix = "pfx_";
            
    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.AddCssClassNamePrefix.html", saveOptions);
    //ExEnd:AddCssClassNamePrefix
  });


  test('ExportCidUrlsForMhtmlResources', () => {
    //ExStart:ExportCidUrlsForMhtmlResources
    let doc = new aw.Document(base.myDir + "Content-ID.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Mhtml);
    saveOptions.prettyFormat = true;
    saveOptions.exportCidUrlsForMhtmlResources = true;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.exportCidUrlsForMhtmlResources.mhtml", saveOptions);
    //ExEnd:ExportCidUrlsForMhtmlResources
  });


  test('ResolveFontNames', () => {
    //ExStart:ResolveFontNames
    let doc = new aw.Document(base.myDir + "Missing font.docx");

    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    saveOptions.prettyFormat = true;
    saveOptions.resolveFontNames = true;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.resolveFontNames.html", saveOptions);
    //ExEnd:ResolveFontNames
  });


  test('ExportTextInputFormFieldAsText', () => {
    //ExStart:ExportTextInputFormFieldAsText
    //GistId:03144d2d1bfafb75c89d385616fdf674
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let imagesDir = path.join(base.artifactsDir, "Images");

    // The folder specified needs to exist and should be empty.
    if (fs.existsSync(imagesDir))
      fs.rmSync(imagesDir, {recursive: true, force: true} );

    fs.mkdirSync(imagesDir, {recursive: true});
    
    // Set an option to export form fields as plain text, not as HTML input elements.
    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    saveOptions.exportTextInputFormFieldAsText = true;
    saveOptions.imagesFolder = imagesDir;

    doc.save(base.artifactsDir + "WorkingWithHtmlSaveOptions.exportTextInputFormFieldAsText.html", saveOptions);
    //ExEnd:ExportTextInputFormFieldAsText
  });

});
