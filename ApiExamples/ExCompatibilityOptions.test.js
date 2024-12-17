// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');


describe("ExCompatibilityOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  //ExStart
  //ExFor:Compatibility
  //ExFor:CompatibilityOptions
  //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
  //ExFor:Document.CompatibilityOptions
  //ExFor:MsWordVersion
  //ExSummary:Shows how to optimize the document for different versions of Microsoft Word.
  test('OptimizeFor', () => {
    let doc = new aw.Document();

    // This object contains an extensive list of flags unique to each document
    // that allow us to facilitate backward compatibility with older versions of Microsoft Word.
    let options = doc.compatibilityOptions;

    // Print the default settings for a blank document.
    console.log("\nDefault optimization settings:");
    printCompatibilityOptions(options);

    // We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
    doc.save(base.artifactsDir + "CompatibilityOptions.optimizeFor.DefaultSettings.docx");

    // We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2010);
    console.log("\nOptimized for Word 2010:");
    printCompatibilityOptions(options);

    doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);
    console.log("\nOptimized for Word 2000:");
    printCompatibilityOptions(options);
  });


  /// <summary>
  /// Groups all flags in a document's compatibility options object by state, then prints each group.
  /// </summary>
  function printCompatibilityOptions(options) {
    let enabledOptions = [];
    let disabledOptions = [];
    addOptionName(options.adjustLineHeightInTable, "AdjustLineHeightInTable", enabledOptions, disabledOptions);
    addOptionName(options.alignTablesRowByRow, "AlignTablesRowByRow", enabledOptions, disabledOptions);
    addOptionName(options.allowSpaceOfSameStyleInTable, "AllowSpaceOfSameStyleInTable", enabledOptions, disabledOptions);
    addOptionName(options.applyBreakingRules, "ApplyBreakingRules", enabledOptions, disabledOptions);
    addOptionName(options.autoSpaceLikeWord95, "AutoSpaceLikeWord95", enabledOptions, disabledOptions);
    addOptionName(options.autofitToFirstFixedWidthCell, "AutofitToFirstFixedWidthCell", enabledOptions, disabledOptions);
    addOptionName(options.balanceSingleByteDoubleByteWidth, "BalanceSingleByteDoubleByteWidth", enabledOptions, disabledOptions);
    addOptionName(options.cachedColBalance, "CachedColBalance", enabledOptions, disabledOptions);
    addOptionName(options.convMailMergeEsc, "ConvMailMergeEsc", enabledOptions, disabledOptions);
    addOptionName(options.disableOpenTypeFontFormattingFeatures, "DisableOpenTypeFontFormattingFeatures", enabledOptions, disabledOptions);
    addOptionName(options.displayHangulFixedWidth, "DisplayHangulFixedWidth", enabledOptions, disabledOptions);
    addOptionName(options.doNotAutofitConstrainedTables, "DoNotAutofitConstrainedTables", enabledOptions, disabledOptions);
    addOptionName(options.doNotBreakConstrainedForcedTable, "DoNotBreakConstrainedForcedTable", enabledOptions, disabledOptions);
    addOptionName(options.doNotBreakWrappedTables, "DoNotBreakWrappedTables", enabledOptions, disabledOptions);
    addOptionName(options.doNotExpandShiftReturn, "DoNotExpandShiftReturn", enabledOptions, disabledOptions);
    addOptionName(options.doNotLeaveBackslashAlone, "DoNotLeaveBackslashAlone", enabledOptions, disabledOptions);
    addOptionName(options.doNotSnapToGridInCell, "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
    addOptionName(options.doNotSuppressIndentation, "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
    addOptionName(options.doNotSuppressParagraphBorders, "DoNotSuppressParagraphBorders", enabledOptions, disabledOptions);
    addOptionName(options.doNotUseEastAsianBreakRules, "DoNotUseEastAsianBreakRules", enabledOptions, disabledOptions);
    addOptionName(options.doNotUseHTMLParagraphAutoSpacing, "DoNotUseHTMLParagraphAutoSpacing", enabledOptions, disabledOptions);
    addOptionName(options.doNotUseIndentAsNumberingTabStop, "DoNotUseIndentAsNumberingTabStop", enabledOptions, disabledOptions);
    addOptionName(options.doNotVertAlignCellWithSp, "DoNotVertAlignCellWithSp", enabledOptions, disabledOptions);
    addOptionName(options.doNotVertAlignInTxbx, "DoNotVertAlignInTxbx", enabledOptions, disabledOptions);
    addOptionName(options.doNotWrapTextWithPunct, "DoNotWrapTextWithPunct", enabledOptions, disabledOptions);
    addOptionName(options.footnoteLayoutLikeWW8, "FootnoteLayoutLikeWW8", enabledOptions, disabledOptions);
    addOptionName(options.forgetLastTabAlignment, "ForgetLastTabAlignment", enabledOptions, disabledOptions);
    addOptionName(options.growAutofit, "GrowAutofit", enabledOptions, disabledOptions);
    addOptionName(options.layoutRawTableWidth, "LayoutRawTableWidth", enabledOptions, disabledOptions);
    addOptionName(options.layoutTableRowsApart, "LayoutTableRowsApart", enabledOptions, disabledOptions);
    addOptionName(options.lineWrapLikeWord6, "LineWrapLikeWord6", enabledOptions, disabledOptions);
    addOptionName(options.mWSmallCaps, "MWSmallCaps", enabledOptions, disabledOptions);
    addOptionName(options.noColumnBalance, "NoColumnBalance", enabledOptions, disabledOptions);
    addOptionName(options.noExtraLineSpacing, "NoExtraLineSpacing", enabledOptions, disabledOptions);
    addOptionName(options.noLeading, "NoLeading", enabledOptions, disabledOptions);
    addOptionName(options.noSpaceRaiseLower, "NoSpaceRaiseLower", enabledOptions, disabledOptions);
    addOptionName(options.noTabHangInd, "NoTabHangInd", enabledOptions, disabledOptions);
    addOptionName(options.overrideTableStyleFontSizeAndJustification, "OverrideTableStyleFontSizeAndJustification", enabledOptions, disabledOptions);
    addOptionName(options.printBodyTextBeforeHeader, "PrintBodyTextBeforeHeader", enabledOptions, disabledOptions);
    addOptionName(options.printColBlack, "PrintColBlack", enabledOptions, disabledOptions);
    addOptionName(options.selectFldWithFirstOrLastChar, "SelectFldWithFirstOrLastChar", enabledOptions, disabledOptions);
    addOptionName(options.shapeLayoutLikeWW8, "ShapeLayoutLikeWW8", enabledOptions, disabledOptions);
    addOptionName(options.showBreaksInFrames, "ShowBreaksInFrames", enabledOptions, disabledOptions);
    addOptionName(options.spaceForUL, "SpaceForUL", enabledOptions, disabledOptions);
    addOptionName(options.spacingInWholePoints, "SpacingInWholePoints", enabledOptions, disabledOptions);
    addOptionName(options.splitPgBreakAndParaMark, "SplitPgBreakAndParaMark", enabledOptions, disabledOptions);
    addOptionName(options.subFontBySize, "SubFontBySize", enabledOptions, disabledOptions);
    addOptionName(options.suppressBottomSpacing, "SuppressBottomSpacing", enabledOptions, disabledOptions);
    addOptionName(options.suppressSpBfAfterPgBrk, "SuppressSpBfAfterPgBrk", enabledOptions, disabledOptions);
    addOptionName(options.suppressSpacingAtTopOfPage, "SuppressSpacingAtTopOfPage", enabledOptions, disabledOptions);
    addOptionName(options.suppressTopSpacing, "SuppressTopSpacing", enabledOptions, disabledOptions);
    addOptionName(options.suppressTopSpacingWP, "SuppressTopSpacingWP", enabledOptions, disabledOptions);
    addOptionName(options.swapBordersFacingPgs, "SwapBordersFacingPgs", enabledOptions, disabledOptions);
    addOptionName(options.swapInsideAndOutsideForMirrorIndentsAndRelativePositioning, "SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning", enabledOptions, disabledOptions);
    addOptionName(options.transparentMetafiles, "TransparentMetafiles", enabledOptions, disabledOptions);
    addOptionName(options.truncateFontHeightsLikeWP6, "TruncateFontHeightsLikeWP6", enabledOptions, disabledOptions);
    addOptionName(options.uICompat97To2003, "UICompat97To2003", enabledOptions, disabledOptions);
    addOptionName(options.ulTrailSpace, "UlTrailSpace", enabledOptions, disabledOptions);
    addOptionName(options.underlineTabInNumList, "UnderlineTabInNumList", enabledOptions, disabledOptions);
    addOptionName(options.useAltKinsokuLineBreakRules, "UseAltKinsokuLineBreakRules", enabledOptions, disabledOptions);
    addOptionName(options.useAnsiKerningPairs, "UseAnsiKerningPairs", enabledOptions, disabledOptions);
    addOptionName(options.useFELayout, "UseFELayout", enabledOptions, disabledOptions);
    addOptionName(options.useNormalStyleForList, "UseNormalStyleForList", enabledOptions, disabledOptions);
    addOptionName(options.usePrinterMetrics, "UsePrinterMetrics", enabledOptions, disabledOptions);
    addOptionName(options.useSingleBorderforContiguousCells, "UseSingleBorderforContiguousCells", enabledOptions, disabledOptions);
    addOptionName(options.useWord2002TableStyleRules, "UseWord2002TableStyleRules", enabledOptions, disabledOptions);
    addOptionName(options.useWord2010TableStyleRules, "UseWord2010TableStyleRules", enabledOptions, disabledOptions);
    addOptionName(options.useWord97LineBreakRules, "UseWord97LineBreakRules", enabledOptions, disabledOptions);
    addOptionName(options.wPJustification, "WPJustification", enabledOptions, disabledOptions);
    addOptionName(options.wPSpaceWidth, "WPSpaceWidth", enabledOptions, disabledOptions);
    addOptionName(options.wrapTrailSpaces, "WrapTrailSpaces", enabledOptions, disabledOptions);
    /*console.log("\tEnabled options:");
    for (let optionName of enabledOptions)
      console.log(`\t\t${optionName}`);
    console.log("\tDisabled options:");
    for (let optionName of disabledOptions)
      console.log(`\t\t${optionName}`);*/
  }

  function addOptionName(option, optionName, enabledOptions, disabledOptions)
  {
    if (option)
      enabledOptions.push(optionName);
    else
      disabledOptions.push(optionName);
  }
  //ExEnd

  test('Tables', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2002);

    expect(compatibilityOptions.adjustLineHeightInTable).toEqual(false);
    expect(compatibilityOptions.alignTablesRowByRow).toEqual(false);
    expect(compatibilityOptions.allowSpaceOfSameStyleInTable).toEqual(true);
    expect(compatibilityOptions.doNotAutofitConstrainedTables).toEqual(true);
    expect(compatibilityOptions.doNotBreakConstrainedForcedTable).toEqual(true);
    expect(compatibilityOptions.doNotBreakWrappedTables).toEqual(false);
    expect(compatibilityOptions.doNotSnapToGridInCell).toEqual(false);
    expect(compatibilityOptions.doNotUseHTMLParagraphAutoSpacing).toEqual(false);
    expect(compatibilityOptions.doNotVertAlignCellWithSp).toEqual(true);
    expect(compatibilityOptions.forgetLastTabAlignment).toEqual(false);
    expect(compatibilityOptions.growAutofit).toEqual(true);
    expect(compatibilityOptions.layoutRawTableWidth).toEqual(false);
    expect(compatibilityOptions.layoutTableRowsApart).toEqual(false);
    expect(compatibilityOptions.noColumnBalance).toEqual(false);
    expect(compatibilityOptions.overrideTableStyleFontSizeAndJustification).toEqual(false);
    expect(compatibilityOptions.useSingleBorderforContiguousCells).toEqual(false);
    expect(compatibilityOptions.useWord2002TableStyleRules).toEqual(true);
    expect(compatibilityOptions.useWord2010TableStyleRules).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.tables.docx");
  });


  test('Breaks', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.applyBreakingRules).toEqual(false);
    expect(compatibilityOptions.doNotUseEastAsianBreakRules).toEqual(true);
    expect(compatibilityOptions.showBreaksInFrames).toEqual(false);
    expect(compatibilityOptions.splitPgBreakAndParaMark).toEqual(true);
    expect(compatibilityOptions.useAltKinsokuLineBreakRules).toEqual(true);
    expect(compatibilityOptions.useWord97LineBreakRules).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.Breaks.docx");
  });


  test('Spacing', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.autoSpaceLikeWord95).toEqual(false);
    expect(compatibilityOptions.displayHangulFixedWidth).toEqual(true);
    expect(compatibilityOptions.noExtraLineSpacing).toEqual(false);
    expect(compatibilityOptions.noLeading).toEqual(false);
    expect(compatibilityOptions.noSpaceRaiseLower).toEqual(false);
    expect(compatibilityOptions.spaceForUL).toEqual(false);
    expect(compatibilityOptions.spacingInWholePoints).toEqual(false);
    expect(compatibilityOptions.suppressBottomSpacing).toEqual(false);
    expect(compatibilityOptions.suppressSpBfAfterPgBrk).toEqual(false);
    expect(compatibilityOptions.suppressSpacingAtTopOfPage).toEqual(false);
    expect(compatibilityOptions.suppressTopSpacing).toEqual(false);
    expect(compatibilityOptions.ulTrailSpace).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.spacing.docx");
  });


  test('WordPerfect', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.suppressTopSpacingWP).toEqual(false);
    expect(compatibilityOptions.truncateFontHeightsLikeWP6).toEqual(false);
    expect(compatibilityOptions.wpjustification).toEqual(false);
    expect(compatibilityOptions.wpspaceWidth).toEqual(false);
    expect(compatibilityOptions.wrapTrailSpaces).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.WordPerfect.docx");
  });


  test('Alignment', () => {
    let doc = new aw.Document();
            
    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.cachedColBalance).toEqual(true);
    expect(compatibilityOptions.doNotVertAlignInTxbx).toEqual(true);
    expect(compatibilityOptions.doNotWrapTextWithPunct).toEqual(true);
    expect(compatibilityOptions.noTabHangInd).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.alignment.docx");
  });


  test('Legacy', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.footnoteLayoutLikeWW8).toEqual(false);
    expect(compatibilityOptions.lineWrapLikeWord6).toEqual(false);
    expect(compatibilityOptions.mwsmallCaps).toEqual(false);
    expect(compatibilityOptions.shapeLayoutLikeWW8).toEqual(false);
    expect(compatibilityOptions.uicompat97To2003).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.legacy.docx");
  });


  test('List', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.underlineTabInNumList).toEqual(true);
    expect(compatibilityOptions.useNormalStyleForList).toEqual(true);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.list.docx");
  });


  test('Misc', () => {
    let doc = new aw.Document();

    let compatibilityOptions = doc.compatibilityOptions;
    compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2000);

    expect(compatibilityOptions.balanceSingleByteDoubleByteWidth).toEqual(false);
    expect(compatibilityOptions.convMailMergeEsc).toEqual(false);
    expect(compatibilityOptions.doNotExpandShiftReturn).toEqual(false);
    expect(compatibilityOptions.doNotLeaveBackslashAlone).toEqual(false);
    expect(compatibilityOptions.doNotSuppressParagraphBorders).toEqual(false);
    expect(compatibilityOptions.doNotUseIndentAsNumberingTabStop).toEqual(true);
    expect(compatibilityOptions.printBodyTextBeforeHeader).toEqual(false);
    expect(compatibilityOptions.printColBlack).toEqual(false);
    expect(compatibilityOptions.selectFldWithFirstOrLastChar).toEqual(true);
    expect(compatibilityOptions.subFontBySize).toEqual(false);
    expect(compatibilityOptions.swapBordersFacingPgs).toEqual(false);
    expect(compatibilityOptions.transparentMetafiles).toEqual(false);
    expect(compatibilityOptions.useAnsiKerningPairs).toEqual(true);
    expect(compatibilityOptions.useFELayout).toEqual(false);
    expect(compatibilityOptions.usePrinterMetrics).toEqual(false);

    // In the output document, these settings can be accessed in Microsoft Word via
    // File -> Options -> Advanced -> Compatibility options for...
    doc.save(base.artifactsDir + "CompatibilityOptions.misc.docx");
  });
});
