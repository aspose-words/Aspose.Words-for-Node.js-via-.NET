// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithDocumentOptionsAndSettings", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('OptimizeFor', () => {
    //ExStart:OptimizeFor
    //GistId:b6462c2505df4b8dd9946ac12ff637b7
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2016);

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.optimizeFor.docx");
    //ExEnd:OptimizeFor
  });


  test('ShowGrammaticalAndSpellingErrors', () => {
    //ExStart:ShowGrammaticalAndSpellingErrors
    let doc = new aw.Document(base.myDir + "Document.docx");

    doc.showGrammaticalErrors = true;
    doc.showSpellingErrors = true;

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.ShowGrammaticalAndSpellingErrors.docx");
    //ExEnd:ShowGrammaticalAndSpellingErrors
  });


  test('CleanupUnusedStylesAndLists', () => {
    //ExStart:CleanupUnusedStylesAndLists
    //GistId:c2ead2f41ca20b28eac045c61a41279e
    let doc = new aw.Document(base.myDir + "Unused styles.docx");

    // Combined with the built-in styles, the document now has eight styles.
    // A custom style is marked as "used" while there is any text within the document
    // formatted in that style. This means that the 4 styles we added are currently unused.
    console.log(`Count of styles before Cleanup: ${doc.styles.count}\n` +
            `Count of lists before Cleanup: ${doc.lists.count}`);

    // Cleans unused styles and lists from the document depending on given CleanupOptions. 
    let cleanupOptions = new aw.CleanupOptions();
    cleanupOptions.unusedLists = false;
    cleanupOptions.unusedStyles = true;
    doc.cleanup(cleanupOptions);

    console.log(`Count of styles after Cleanup was decreased: ${doc.styles.count}\n` +
            `Count of lists after Cleanup is the same: ${doc.lists.count}`);

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.CleanupUnusedStylesAndLists.docx");
    //ExEnd:CleanupUnusedStylesAndLists
  });


  test('CleanupDuplicateStyle', () => {
    //ExStart:CleanupDuplicateStyle
    //GistId:c2ead2f41ca20b28eac045c61a41279e
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Count of styles before Cleanup.
    console.log(doc.styles.count);

    // Cleans duplicate styles from the document.
    let options = new aw.CleanupOptions();
    options.duplicateStyle = true;
    doc.cleanup(options);

    // Count of styles after Cleanup was decreased.
    console.log(doc.styles.count);

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.CleanupDuplicateStyle.docx");
    //ExEnd:CleanupDuplicateStyle
  });


  test('ViewOptions', () => {
    //ExStart:SetViewOption
    //GistId:b6462c2505df4b8dd9946ac12ff637b7
    let doc = new aw.Document(base.myDir + "Document.docx");
            
    doc.viewOptions.viewType = aw.Settings.ViewType.PageLayout;
    doc.viewOptions.zoomPercent = 50;

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.viewOptions.docx");
    //ExEnd:SetViewOption
  });


  test('DocumentPageSetup', () => {
    //ExStart:DocumentPageSetup
    //GistId:b6462c2505df4b8dd9946ac12ff637b7
    let doc = new aw.Document(base.myDir + "Document.docx");

    // Set the layout mode for a section allowing to define the document grid behavior.
    // Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word
    // if any Asian language is defined as editing language.
    doc.firstSection.pageSetup.layoutMode = aw.SectionLayoutMode.Grid;
    doc.firstSection.pageSetup.charactersPerLine = 30;
    doc.firstSection.pageSetup.linesPerPage = 10;

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.DocumentPageSetup.docx");
    //ExEnd:DocumentPageSetup
  });


  test('AddEditingLanguage', () => {
    //ExStart:AddEditingLanguage
    //GistId:9298958b7a6872536299cd7e3f3ab24b
    let loadOptions = new aw.Loading.LoadOptions();
    // Set language preferences that will be used when document is loading.
    loadOptions.languagePreferences.addEditingLanguage(aw.Loading.EditingLanguage.Japanese);
            
    let doc = new aw.Document(base.myDir + "No default editing language.docx", loadOptions);
    //ExEnd:AddEditingLanguage

    let localeIdFarEast = doc.styles.defaultFont.localeIdFarEast;
    console.log(
      localeIdFarEast == aw.Loading.EditingLanguage.Japanese
        ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
        : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
  });


  test('SetRussianAsDefaultEditingLanguage', () => {
    //ExStart:SetRussianAsDefaultEditingLanguage
    //GistId:b6462c2505df4b8dd9946ac12ff637b7
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.languagePreferences.defaultEditingLanguage = aw.Loading.EditingLanguage.Russian;

    let doc = new aw.Document(base.myDir + "No default editing language.docx", loadOptions);

    let localeId = doc.styles.defaultFont.localeId;
    console.log(
      localeId == aw.Loading.EditingLanguage.Russian
        ? "The document either has no any language set in defaults or it was set to Russian originally."
        : "The document default language was set to another than Russian language originally, so it is not overridden.");
    //ExEnd:SetRussianAsDefaultEditingLanguage
  });


  test('PageSetupAndSectionFormatting', () => {
    //ExStart:PageSetupAndSectionFormatting
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.pageSetup.orientation = aw.Orientation.Landscape;
    builder.pageSetup.leftMargin = 50;
    builder.pageSetup.paperSize = aw.PaperSize.Paper10x14;

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.PageSetupAndSectionFormatting.docx");
    //ExEnd:PageSetupAndSectionFormatting
  });


  test('PageBorderProperties', () => {
    //ExStart:PageBorderProperties
    let doc = new aw.Document();

    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.borderAlwaysInFront = false;
    pageSetup.borderDistanceFrom = aw.PageBorderDistanceFrom.PageEdge;
    pageSetup.borderAppliesTo = aw.PageBorderAppliesTo.FirstPage;

    let border = pageSetup.borders.at(aw.BorderType.Top);
    border.lineStyle = aw.LineStyle.Single;
    border.lineWidth = 30;
    border.color = "#0000FF";
    border.distanceFromText = 0;

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.PageBorderProperties.docx");
    //ExEnd:PageBorderProperties
  });


  test('LineGridSectionLayoutMode', () => {
    //ExStart:LineGridSectionLayoutMode
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Enable pitching, and then use it to set the number of lines per page in this section.
    // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
    builder.pageSetup.layoutMode = aw.SectionLayoutMode.LineGrid;
    builder.pageSetup.linesPerPage = 15;

    builder.paragraphFormat.snapToGrid = true;

    for (let i = 0; i < 30; i++)
      builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

    doc.save(base.artifactsDir + "WorkingWithDocumentOptionsAndSettings.linesPerPage.docx");
    //ExEnd:LineGridSectionLayoutMode
  });

});
