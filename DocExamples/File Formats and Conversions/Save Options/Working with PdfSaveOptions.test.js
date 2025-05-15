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
const MemoryStream = require('memorystream');


describe("WorkingWithPdfSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('DisplayDocTitleInWindowTitlebar', () => {
    //ExStart:DisplayDocTitleInWindowTitlebar
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.displayDocTitle = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
    //ExEnd:DisplayDocTitleInWindowTitlebar
  });


  //ExStart:PdfRenderWarnings
  //GistId:f9c5250f94e595ea3590b3be679475ba
  test.skip('PdfRenderWarnings - TODO: warningCallback not supported yet', () => {
    let doc = new aw.Document(base.myDir + "WMF with image.docx");

    let metafileRenderingOptions = new aw.Saving.MetafileRenderingOptions();
    metafileRenderingOptions.emulateRasterOperations = false;
    metafileRenderingOptions.renderingMode = aw.Saving.MetafileRenderingMode.VectorWithFallback;

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.metafileRenderingOptions = metafileRenderingOptions;

    // If Aspose.words cannot correctly render some of the metafile records
    // to vector graphics then Aspose.words renders this metafile to a bitmap.
    let callback = new HandleDocumentWarnings();
    doc.warningCallback = callback;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

    // While the file saves successfully, rendering warnings that occurred during saving are collected here.
    for (let warningInfo of callback.mWarnings)
    {
      console.log(warningInfo.description);
    }
  });


/*  public class HandleDocumentWarnings : IWarningCallback
  {
      /// <summary>
      /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
      /// potential issue during document processing. The callback can be set to listen for warnings generated during
      /// document load and/or document save.
      /// </summary>
    public void Warning(WarningInfo info)
    {
        // For now type of warnings about unsupported metafile records changed
        // from DataLoss/UnexpectedContent to MinorFormattingLoss.
      if (info.warningType == aw.WarningType.MinorFormattingLoss)
      {
        console.log("Unsupported operation: " + info.description);
        mWarnings.warning(info);
      }
    }

    public WarningInfoCollection mWarnings = new aw.WarningInfoCollection();
  }
    //ExEnd:PdfRenderWarnings
*/

  test.skip('DigitallySignedPdfUsingCertificateHolder - TODO: CertificateHolder not supported yet', () => {
    //ExStart:DigitallySignedPdfUsingCertificateHolder
    //GistId:bdc15a6de6b25d9d4e66f2ce918fc01b
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Test Signed PDF.");

    /*let saveOptions = new aw.Saving.PdfSaveOptions();
      DigitalSignatureDetails = new aw.Saving.PdfDigitalSignatureDetails(
        aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw"), "reason", "location",
        Date.now())*/

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
    //ExEnd:DigitallySignedPdfUsingCertificateHolder
  });


  test('EmbeddedAllFonts', () => {
    //ExStart:EmbeddedAllFonts
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // The output PDF will be embedded with all fonts found in the document.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.embedFullFonts = true;
            
    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.EmbeddedAllFonts.pdf", saveOptions);
    //ExEnd:EmbeddedAllFonts
  });


  test('EmbeddedSubsetFonts', () => {
    //ExStart:EmbeddedSubsetFonts
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // The output PDF will contain subsets of the fonts in the document.
    // Only the glyphs used in the document are included in the PDF fonts.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.embedFullFonts = false;
            
    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.EmbeddedSubsetFonts.pdf", saveOptions);
    //ExEnd:EmbeddedSubsetFonts
  });


  test('DisableEmbedWindowsFonts', () => {
    //ExStart:DisableEmbedWindowsFonts
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // The output PDF will be saved without embedding standard windows fonts.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.fontEmbeddingMode = aw.Saving.PdfFontEmbeddingMode.EmbedNone;
            
    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
    //ExEnd:DisableEmbedWindowsFonts
  });


  test('SkipEmbeddedArialAndTimesRomanFonts', () => {
    //ExStart:SkipEmbeddedArialAndTimesRomanFonts
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.fontEmbeddingMode = aw.Saving.PdfFontEmbeddingMode.EmbedAll;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
    //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
  });


  test('AvoidEmbeddingCoreFonts', () => {
    //ExStart:AvoidEmbeddingCoreFonts
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.useCoreFonts = true;
            
    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
    //ExEnd:AvoidEmbeddingCoreFonts
  });


  test('EscapeUri', () => {
    //ExStart:EscapeUri
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    builder.insertHyperlink("Testlink",
      "https://www.google.com/search?q= aspose", false);

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.EscapeUri.pdf");
    //ExEnd:EscapeUri
  });


  test('ExportHeaderFooterBookmarks', () => {
    //ExStart:ExportHeaderFooterBookmarks
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Bookmarks in headers and footers.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.outlineOptions.defaultBookmarksOutlineLevel = 1;
    saveOptions.headerFooterBookmarksExportMode = aw.Saving.HeaderFooterBookmarksExportMode.First;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
    //ExEnd:ExportHeaderFooterBookmarks
  });


  test('EmulateRenderingToSizeOnPage', () => {
    //ExStart:EmulateRenderingToSizeOnPage
    let doc = new aw.Document(base.myDir + "WMF with text.docx");

    let metafileRenderingOptions = new aw.Saving.MetafileRenderingOptions();
    metafileRenderingOptions.emulateRenderingToSizeOnPage = false;

    // If Aspose.words cannot correctly render some of the metafile records to vector graphics
    // then Aspose.words renders this metafile to a bitmap.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.metafileRenderingOptions = metafileRenderingOptions;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.emulateRenderingToSizeOnPage.pdf", saveOptions);
    //ExEnd:EmulateRenderingToSizeOnPage
  });


  test('AdditionalTextPositioning', () => {
    //ExStart:AdditionalTextPositioning
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.additionalTextPositioning = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.additionalTextPositioning.pdf", saveOptions);
    //ExEnd:AdditionalTextPositioning
  });


  test('ConversionToPdf17', () => {
    //ExStart:ConversionToPdf17
    //GistId:38c6608baa855f951a4e117a721bdaae
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.compliance = aw.Saving.PdfCompliance.Pdf17;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
    //ExEnd:ConversionToPdf17
  });


  test('DownsamplingImages', () => {
    //ExStart:DownsamplingImages
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    // We can set a minimum threshold for downsampling.
    // This value will prevent the second image in the input document from being downsampled.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.downsampleOptions.resolution = 36;
    saveOptions.downsampleOptions.resolutionThreshold = 128;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
    //ExEnd:DownsamplingImages
  });


  test('OutlineOptions', () => {
    //ExStart:OutlineOptions
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.outlineOptions.headingsOutlineLevels = 3;
    saveOptions.outlineOptions.expandedOutlineLevels = 1;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.outlineOptions.pdf", saveOptions);
    //ExEnd:OutlineOptions
  });


  test('CustomPropertiesExport', () => {
    //ExStart:CustomPropertiesExport
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document();
    doc.customDocumentProperties.add("Company", "Aspose");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.customPropertiesExport = aw.Saving.PdfCustomPropertiesExport.Standard;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.customPropertiesExport.pdf", saveOptions);
    //ExEnd:CustomPropertiesExport
  });


  test('ExportDocumentStructure', () => {
    //ExStart:ExportDocumentStructure
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    // The file size will be increased and the structure will be visible in the "Content" navigation pane
    // of Adobe Acrobat Pro, while editing the .pdf.
    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.exportDocumentStructure = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.exportDocumentStructure.pdf", saveOptions);
    //ExEnd:ExportDocumentStructure
  });


  test('ImageCompression', () => {
    //ExStart:ImageCompression
    //GistId:d569206cfa68ce09d8f6c6e3de44c13e
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.imageCompression = aw.Saving.PdfImageCompression.Jpeg;
    saveOptions.preserveFormFields = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.imageCompression.pdf", saveOptions);

    let saveOptionsA2U = new aw.Saving.PdfSaveOptions();
    saveOptions.compliance = aw.Saving.PdfCompliance.PdfA2u;
    saveOptions.imageCompression = aw.Saving.PdfImageCompression.Jpeg;
    saveOptions.jpegQuality = 100; // Use JPEG compression at 50% quality to reduce file size.

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.ImageCompression_A2u.pdf", saveOptionsA2U);
    //ExEnd:ImageCompression
  });


  test('UpdateLastPrinted', () => {
    //ExStart:UpdateLastPrinted
    //GistId:03144d2d1bfafb75c89d385616fdf674
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.updateLastPrintedProperty = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.UpdateLastPrinted.pdf", saveOptions);
    //ExEnd:UpdateLastPrinted
  });


  test('Dml3DEffectsRendering', () => {
    //ExStart:Dml3DEffectsRendering
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.dml3DEffectsRenderingMode = aw.Saving.Dml3DEffectsRenderingMode.Advanced;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
    //ExEnd:Dml3DEffectsRendering
  });


  test('InterpolateImages', () => {
    //ExStart:SetImageInterpolation
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.interpolateImages = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.interpolateImages.pdf", saveOptions);
    //ExEnd:SetImageInterpolation
  });


  test('OptimizeOutput', () => {
    //ExStart:OptimizeOutput
    //GistId:38c6608baa855f951a4e117a721bdaae
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.optimizeOutput = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.optimizeOutput.pdf", saveOptions);
    //ExEnd:OptimizeOutput
  });


  test('UpdateScreenTip', () => {
    //ExStart:UpdateScreenTip
    //GistId:cf2b0536741aecc0d14447256bf060c4
    let doc = new aw.Document(base.myDir + "Table of contents.docx");

    var tocHyperLinks = Array.from(doc.range.fields).filter(f =>
      f.type == aw.Fields.FieldType.FieldHyperlink &&
      f.asFieldHyperlink().subAddress != null &&
      f.asFieldHyperlink().subAddress.startsWith("#_Toc")).map(f => f.asFieldHyperlink());

    for (let link of tocHyperLinks)
      link.screenTip = link.displayResult;

    let saveOptions = new aw.Saving.PdfSaveOptions();
    saveOptions.compliance = aw.Saving.PdfCompliance.PdfUa1;
    saveOptions.displayDocTitle = true;
    saveOptions.exportDocumentStructure = true;
    saveOptions.outlineOptions.headingsOutlineLevels = 3;
    saveOptions.outlineOptions.createMissingOutlineLevels = true;

    doc.save(base.artifactsDir + "WorkingWithPdfSaveOptions.UpdateScreenTip.pdf", saveOptions);
    //ExEnd:UpdateScreenTip
  });

});
