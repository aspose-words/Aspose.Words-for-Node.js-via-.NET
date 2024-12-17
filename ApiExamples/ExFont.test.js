// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const fs = require('fs');
const os = require("os");
const path = require("path");

describe("ExFont", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CreateFormattedRun', () => {
    //ExStart
    //ExFor:Document.#ctor
    //ExFor:Font
    //ExFor:aw.Font.name
    //ExFor:aw.Font.size
    //ExFor:aw.Font.highlightColor
    //ExFor:Run
    //ExFor:Run.#ctor(DocumentBase,String)
    //ExFor:aw.Story.firstParagraph
    //ExSummary:Shows how to format a run of text using its font property.
    let doc = new aw.Document();
    let run = new aw.Run(doc, "Hello world!");

    let font = run.font;
    font.name = "Courier New";
    font.size = 36;
    font.highlightColor = "#FFFF00";

    doc.firstSection.body.firstParagraph.appendChild(run);
    doc.save(base.artifactsDir + "Font.CreateFormattedRun.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.CreateFormattedRun.docx");
    run = doc.firstSection.body.firstParagraph.runs.at(0);

    expect(run.getText().trim()).toEqual("Hello world!");
    expect(run.font.name).toEqual("Courier New");
    expect(run.font.size).toEqual(36);
    expect(run.font.highlightColor).toEqual("#FFFF00");
  });


  test('Caps', () => {
    //ExStart
    //ExFor:aw.Font.allCaps
    //ExFor:aw.Font.smallCaps
    //ExSummary:Shows how to format a run to display its contents in capitals.
    let doc = new aw.Document();
    let para = doc.getParagraph(0, true);

    // There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
    // 1 -  Set the AllCaps flag to display all characters in regular capitals:
    let run = new aw.Run(doc, "all capitals");
    run.font.allCaps = true;
    para.appendChild(run);

    para = para.parentNode.appendChild(new aw.Paragraph(doc)).asParagraph();

    // 2 -  Set the SmallCaps flag to display all characters in small capitals:
    // If a character is lower case, it will appear in its upper case form
    // but will have the same height as the lower case (the font's x-height).
    // Characters that were in upper case originally will look the same.
    run = new aw.Run(doc, "Small Capitals");
    run.font.smallCaps = true;
    para.appendChild(run);

    doc.save(base.artifactsDir + "Font.caps.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.caps.docx");
    run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("all capitals");
    expect(run.font.allCaps).toEqual(true);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("Small Capitals");
    expect(run.font.smallCaps).toEqual(true);
  });


  test('GetDocumentFonts', () => {
    //ExStart
    //ExFor:FontInfoCollection
    //ExFor:aw.DocumentBase.fontInfos
    //ExFor:FontInfo
    //ExFor:aw.Fonts.FontInfo.name
    //ExFor:aw.Fonts.FontInfo.isTrueType
    //ExSummary:Shows how to print the details of what fonts are present in a document.
    let doc = new aw.Document(base.myDir + "Embedded font.docx");

    let allFonts = doc.fontInfos;
    expect(allFonts.count).toEqual(5);

    // Print all the used and unused fonts in the document.
    for (let i = 0; i < allFonts.count; i++) {
      console.log(`Font index #${i}`);
      console.log(`\tName: ${allFonts.at(i).name}`);
      console.log(`\tIs ${allFonts.at(i).isTrueType ? "" : "not "}a trueType font`);
    }
    //ExEnd
  });


  test('DefaultValuesEmbeddedFontsParameters', () => {
    let doc = new aw.Document();

    expect(doc.fontInfos.embedTrueTypeFonts).toEqual(false);
    expect(doc.fontInfos.embedSystemFonts).toEqual(false);
    expect(doc.fontInfos.saveSubsetFonts).toEqual(false);
  });


  test.each([false, true])('FontInfoCollection(embedAllFonts = %o)', (embedAllFonts) => {
    //ExStart
    //ExFor:FontInfoCollection
    //ExFor:aw.DocumentBase.fontInfos
    //ExFor:aw.Fonts.FontInfoCollection.embedTrueTypeFonts
    //ExFor:aw.Fonts.FontInfoCollection.embedSystemFonts
    //ExFor:aw.Fonts.FontInfoCollection.saveSubsetFonts
    //ExSummary:Shows how to save a document with embedded TrueType fonts.
    let doc = new aw.Document(base.myDir + "Document.docx");

    let fontInfos = doc.fontInfos;
    fontInfos.embedTrueTypeFonts = embedAllFonts;
    fontInfos.embedSystemFonts = embedAllFonts;
    fontInfos.saveSubsetFonts = embedAllFonts;

    doc.save(base.artifactsDir + "Font.FontInfoCollection.docx");
    //ExEnd

    let testedFileLength = fs.statSync(base.artifactsDir + "Font.FontInfoCollection.docx").size;

    if (embedAllFonts)
      expect(testedFileLength < 28000).toEqual(true);
    else
      expect(testedFileLength < 13000).toEqual(true);
  });


  test.each([
    [true, false, false], // Save a document with embedded TrueType fonts. System fonts are not included. Saves full versions of embedding fonts.
    [true, true, false],  // Save a document with embedded TrueType fonts. System fonts are included. Saves full versions of embedding fonts.
    [true, true, true],   // Save a document with embedded TrueType fonts. System fonts are included. Saves subset of embedding fonts.
    [true, false, true],  // Save a document with embedded TrueType fonts. System fonts are not included. Saves subset of embedding fonts.
    [false, false, false] // Remove embedded fonts from the saved document.
  ])('WorkWithEmbeddedFonts(embedTrueTypeFonts = %o, embedSystemFonts = %o, saveSubsetFonts = %o)', (embedTrueTypeFonts, embedSystemFonts, saveSubsetFonts) => {
    let doc = new aw.Document(base.myDir + "Document.docx");

    let fontInfos = doc.fontInfos;
    fontInfos.embedTrueTypeFonts = embedTrueTypeFonts;
    fontInfos.embedSystemFonts = embedSystemFonts;
    fontInfos.saveSubsetFonts = saveSubsetFonts;

    doc.save(base.artifactsDir + "Font.WorkWithEmbeddedFonts.docx");
  });


  test('StrikeThrough', () => {
    //ExStart
    //ExFor:aw.Font.strikeThrough
    //ExFor:aw.Font.doubleStrikeThrough
    //ExSummary:Shows how to add a line strikethrough to text.
    let doc = new aw.Document();
    let para = doc.getParagraph(0, true);

    let run = new aw.Run(doc, "Text with a single-line strikethrough.");
    run.font.strikeThrough = true;
    para.appendChild(run);

    para = para.parentNode.appendChild(new aw.Paragraph(doc)).asParagraph();

    run = new aw.Run(doc, "Text with a double-line strikethrough.");
    run.font.doubleStrikeThrough = true;
    para.appendChild(run);

    doc.save(base.artifactsDir + "Font.strikeThrough.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.strikeThrough.docx");

    run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Text with a single-line strikethrough.");
    expect(run.font.strikeThrough).toEqual(true);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("Text with a double-line strikethrough.");
    expect(run.font.doubleStrikeThrough).toEqual(true);
  });


  test('PositionSubscript', () => {
    //ExStart
    //ExFor:aw.Font.position
    //ExFor:aw.Font.subscript
    //ExFor:aw.Font.superscript
    //ExSummary:Shows how to format text to offset its position.
    let doc = new aw.Document();
    let para = doc.getParagraph(0, true);

    // Raise this run of text 5 points above the baseline.
    let run = new aw.Run(doc, "Raised text. ");
    run.font.position = 5;
    para.appendChild(run);

    // Lower this run of text 10 points below the baseline.
    run = new aw.Run(doc, "Lowered text. ");
    run.font.position = -10;
    para.appendChild(run);

    // Add a run of normal text.
    run = new aw.Run(doc, "Text in its default position. ");
    para.appendChild(run);

    // Add a run of text that appears as subscript.
    run = new aw.Run(doc, "Subscript. ");
    run.font.subscript = true;
    para.appendChild(run);

    // Add a run of text that appears as superscript.
    run = new aw.Run(doc, "Superscript.");
    run.font.superscript = true;
    para.appendChild(run);

    doc.save(base.artifactsDir + "Font.PositionSubscript.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.PositionSubscript.docx");
    run = doc.firstSection.body.firstParagraph.runs.at(0);

    expect(run.getText().trim()).toEqual("Raised text.");
    expect(run.font.position).toEqual(5);

    doc = new aw.Document(base.artifactsDir + "Font.PositionSubscript.docx");
    run = doc.firstSection.body.firstParagraph.runs.at(1);

    expect(run.getText().trim()).toEqual("Lowered text.");
    expect(run.font.position).toEqual(-10);

    run = doc.firstSection.body.firstParagraph.runs.at(3);

    expect(run.getText().trim()).toEqual("Subscript.");
    expect(run.font.subscript).toEqual(true);

    run = doc.firstSection.body.firstParagraph.runs.at(4);

    expect(run.getText().trim()).toEqual("Superscript.");
    expect(run.font.superscript).toEqual(true);
  });


  test('ScalingSpacing', () => {
    //ExStart
    //ExFor:aw.Font.scaling
    //ExFor:aw.Font.spacing
    //ExSummary:Shows how to set horizontal scaling and spacing for characters.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add run of text and increase character width to 150%.
    builder.font.scaling = 150;
    builder.writeln("Wide characters");

    // Add run of text and add 1pt of extra horizontal spacing between each character.
    builder.font.spacing = 1;
    builder.writeln("Expanded by 1pt");

    // Add run of text and bring characters closer together by 1pt.
    builder.font.spacing = -1;
    builder.writeln("Condensed by 1pt");

    doc.save(base.artifactsDir + "Font.ScalingSpacing.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.ScalingSpacing.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Wide characters");
    expect(run.font.scaling).toEqual(150);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("Expanded by 1pt");
    expect(run.font.spacing).toEqual(1);

    run = doc.firstSection.body.paragraphs.at(2).runs.at(0);

    expect(run.getText().trim()).toEqual("Condensed by 1pt");
    expect(run.font.spacing).toEqual(-1);
  });


  test('Italic', () => {
    //ExStart
    //ExFor:aw.Font.italic
    //ExSummary:Shows how to write italicized text using a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.size = 36;
    builder.font.italic = true;
    builder.writeln("Hello world!");

    doc.save(base.artifactsDir + "Font.italic.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.italic.docx");
    let run = doc.firstSection.body.firstParagraph.runs.at(0);

    expect(run.getText().trim()).toEqual("Hello world!");
    expect(run.font.italic).toEqual(true);
  });


  test('EngraveEmboss', () => {
    //ExStart
    //ExFor:aw.Font.emboss
    //ExFor:aw.Font.engrave
    //ExSummary:Shows how to apply engraving/embossing effects to text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.size = 36;
    builder.font.color = "#ADD8E6";

    // Below are two ways of using shadows to apply a 3D-like effect to the text.
    // 1 -  Engrave text to make it look like the letters are sunken into the page:
    builder.font.engrave = true;

    builder.writeln("This text is engraved.");

    // 2 -  Emboss text to make it look like the letters pop out of the page:
    builder.font.engrave = false;
    builder.font.emboss = true;

    builder.writeln("This text is embossed.");

    doc.save(base.artifactsDir + "Font.EngraveEmboss.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.EngraveEmboss.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("This text is engraved.");
    expect(run.font.engrave).toEqual(true);
    expect(run.font.emboss).toEqual(false);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("This text is embossed.");
    expect(run.font.engrave).toEqual(false);
    expect(run.font.emboss).toEqual(true);
  });


  test('Shadow', () => {
    //ExStart
    //ExFor:aw.Font.shadow
    //ExSummary:Shows how to create a run of text formatted with a shadow.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set the Shadow flag to apply an offset shadow effect,
    // making it look like the letters are floating above the page.
    builder.font.shadow = true;
    builder.font.size = 36;

    builder.writeln("This text has a shadow.");

    doc.save(base.artifactsDir + "Font.shadow.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.shadow.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("This text has a shadow.");
    expect(run.font.shadow).toEqual(true);
  });


  test('Outline', () => {
    //ExStart
    //ExFor:aw.Font.outline
    //ExSummary:Shows how to create a run of text formatted as outline.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set the Outline flag to change the text's fill color to white and
    // leave a thin outline around each character in the original color of the text. 
    builder.font.outline = true;
    builder.font.color = "#0000FF";
    builder.font.size = 36;

    builder.writeln("This text has an outline.");

    doc.save(base.artifactsDir + "Font.outline.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.outline.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("This text has an outline.");
    expect(run.font.outline).toEqual(true);
  });


  test('Hidden', () => {
    //ExStart
    //ExFor:aw.Font.hidden
    //ExSummary:Shows how to create a run of hidden text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
    // We will not see or highlight hidden text unless we enable the "Hidden text" option
    // found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
    // and we will be able to access this text programmatically.
    // It is not advised to use this method to hide sensitive information.
    builder.font.hidden = true;
    builder.font.size = 36;

    builder.writeln("This text will not be visible in the document.");

    doc.save(base.artifactsDir + "Font.hidden.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.hidden.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("This text will not be visible in the document.");
    expect(run.font.hidden).toEqual(true);
  });


  test('Kerning', () => {
    //ExStart
    //ExFor:aw.Font.kerning
    //ExSummary:Shows how to specify the font size at which kerning begins to take effect.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.font.name = "Arial Black";

    // Set the builder's font size, and minimum size at which kerning will take effect.
    // The font size falls below the kerning threshold, so the run bellow will not have kerning.
    builder.font.size = 18;
    builder.font.kerning = 24;

    builder.writeln("TALLY. (Kerning not applied)");

    // Set the kerning threshold so that the builder's current font size is above it.
    // Any text we add from this point will have kerning applied. The spaces between characters
    // will be adjusted, normally resulting in a slightly more aesthetically pleasing text run.
    builder.font.kerning = 12;

    builder.writeln("TALLY. (Kerning applied)");

    doc.save(base.artifactsDir + "Font.kerning.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.kerning.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("TALLY. (Kerning not applied)");
    expect(run.font.kerning).toEqual(24);
    expect(run.font.size).toEqual(18);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("TALLY. (Kerning applied)");
    expect(run.font.kerning).toEqual(12);
    expect(run.font.size).toEqual(18);
  });


  test('NoProofing', () => {
    //ExStart
    //ExFor:aw.Font.noProofing
    //ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
    // We can un-set the "NoProofing" flag to create a portion of text that
    // bypasses the spell checker while completely disabling it.
    builder.font.noProofing = true;

    builder.writeln("Proofing has been disabled, so these spelking errrs will not display red lines underneath.");

    doc.save(base.artifactsDir + "Font.noProofing.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.noProofing.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Proofing has been disabled, so these spelking errrs will not display red lines underneath.");
    expect(run.font.noProofing).toEqual(true);
  });


  test.skip('LocaleId: CultureInfo', () => {
    //ExStart
    //ExFor:aw.Font.localeId
    //ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If we set the font's locale to English and insert some Russian text,
    // the English locale spell checker will not recognize the text and detect it as a spelling error.
    builder.font.localeId = new CultureInfo("en-US", false).LCID;
    builder.writeln("Привет!");

    // Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
    builder.font.localeId = new CultureInfo("ru-RU", false).LCID;
    builder.writeln("Привет!");

    doc.save(base.artifactsDir + "Font.localeId.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.localeId.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Привет!");
    expect(run.font.localeId).toEqual(1033);

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("Привет!");
    expect(run.font.localeId).toEqual(1049);
  });


  test('Underlines', () => {
    //ExStart
    //ExFor:aw.Font.underline
    //ExFor:aw.Font.underlineColor
    //ExSummary:Shows how to configure the style and color of a text underline.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.underline = aw.Underline.Dotted;
    builder.font.underlineColor = "#FF0000";

    builder.writeln("Underlined text.");

    doc.save(base.artifactsDir + "Font.Underlines.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.Underlines.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Underlined text.");
    expect(run.font.underline).toEqual(aw.Underline.Dotted);
    expect(run.font.underlineColor).toEqual("#FF0000");
  });


  test('ComplexScript', () => {
    //ExStart
    //ExFor:aw.Font.complexScript
    //ExSummary:Shows how to add text that is always treated as complex script.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.complexScript = true;

    builder.writeln("Text treated as complex script.");

    doc.save(base.artifactsDir + "Font.complexScript.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.complexScript.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Text treated as complex script.");
    expect(run.font.complexScript).toEqual(true);
  });


  test('SparklingText', () => {
    //ExStart
    //ExFor:aw.Font.textEffect
    //ExSummary:Shows how to apply a visual effect to a run.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.size = 36;
    builder.font.textEffect = aw.TextEffect.SparkleText;

    builder.writeln("Text with a sparkle effect.");

    // Older versions of Microsoft Word only support font animation effects.
    doc.save(base.artifactsDir + "Font.SparklingText.doc");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.SparklingText.doc");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Text with a sparkle effect.");
    expect(run.font.textEffect).toEqual(aw.TextEffect.SparkleText);
  });


  test('ForegroundAndBackground', () => {
    //ExStart
    //ExFor:aw.Shading.foregroundPatternThemeColor
    //ExFor:aw.Shading.backgroundPatternThemeColor
    //ExFor:aw.Shading.foregroundTintAndShade
    //ExFor:aw.Shading.backgroundTintAndShade
    //ExSummary:Shows how to set foreground and background colors for shading texture.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shading = doc.firstSection.body.firstParagraph.paragraphFormat.shading;
    shading.texture = aw.TextureIndex.Texture12Pt5Percent;
    shading.foregroundPatternThemeColor = aw.Themes.ThemeColor.Dark1;
    shading.backgroundPatternThemeColor = aw.Themes.ThemeColor.Dark2;

    shading.foregroundTintAndShade = 0.5;
    shading.backgroundTintAndShade = -0.2;

    builder.font.border.color = "#008000";
    builder.font.border.lineWidth = 2.5;
    builder.font.border.lineStyle = aw.LineStyle.DashDotStroker;

    builder.writeln("Foreground and background pattern colors for shading texture.");

    doc.save(base.artifactsDir + "Font.ForegroundAndBackground.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.ForegroundAndBackground.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Foreground and background pattern colors for shading texture.");
    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.shading.foregroundPatternThemeColor).toEqual(aw.Themes.ThemeColor.Dark1);
    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.shading.backgroundPatternThemeColor).toEqual(aw.Themes.ThemeColor.Dark2);

    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.shading.foregroundTintAndShade).toBeCloseTo(0.5, 1);
    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.shading.backgroundTintAndShade).toBeCloseTo(-0.2, 1);
  });


  test('Shading', () => {
    //ExStart
    //ExFor:aw.Font.shading
    //ExSummary:Shows how to apply shading to text created by a document builder.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.color = "#FFFFFF";

    // One way to make the text created using our white font color visible
    // is to apply a background shading effect.
    let shading = builder.font.shading;
    shading.texture = aw.TextureIndex.TextureDiagonalUp;
    shading.backgroundPatternColor = "#FF4500";
    shading.foregroundPatternColor = "#00008B";

    builder.writeln("White text on an orange background with a two-tone texture.");

    doc.save(base.artifactsDir + "Font.shading.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.shading.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("White text on an orange background with a two-tone texture.");
    expect(run.font.color).toEqual("#FFFFFF");

    expect(run.font.shading.texture).toEqual(aw.TextureIndex.TextureDiagonalUp);
    expect(run.font.shading.backgroundPatternColor).toEqual("#FF4500");
    expect(run.font.shading.foregroundPatternColor).toEqual("#00008B");
  });


  test.skip('Bidi: CultureInfo', () => {
    //ExStart
    //ExFor:aw.Font.bidi
    //ExFor:aw.Font.nameBi
    //ExFor:aw.Font.sizeBi
    //ExFor:aw.Font.italicBi
    //ExFor:aw.Font.boldBi
    //ExFor:aw.Font.localeIdBi
    //ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Define a set of font settings for left-to-right text.
    builder.font.name = "Courier New";
    builder.font.size = 16;
    builder.font.italic = false;
    builder.font.bold = false;
    builder.font.localeId = new CultureInfo("en-US", false).LCID;

    // Define another set of font settings for right-to-left text.
    builder.font.nameBi = "Andalus";
    builder.font.sizeBi = 24;
    builder.font.italicBi = true;
    builder.font.boldBi = true;
    builder.font.localeIdBi = new CultureInfo("ar-AR", false).LCID;

    // We can use the Bidi flag to indicate whether the text we are about to add
    // with the document builder is right-to-left. When we add text with this flag set to true,
    // it will be formatted using the right-to-left set of font settings.
    builder.font.bidi = true;
    builder.write("مرحبًا");

    // Set the flag to false, and then add left-to-right text.
    // The document builder will format these using the left-to-right set of font settings.
    builder.font.bidi = false;
    builder.write(" Hello world!");

    doc.save(base.artifactsDir + "Font.bidi.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.bidi.docx");

    for (let run of doc.firstSection.body.paragraphs.at(0).runs)
    {
      switch (doc.firstSection.body.paragraphs.at(0).indexOf(run))
      {
        case 0:
          expect(run.getText().trim()).toEqual("مرحبًا");
          expect(run.font.bidi).toEqual(true);
          break;
        case 1:
          expect(run.getText().trim()).toEqual("Hello world!");
          expect(run.font.bidi).toEqual(false);
          break;
      }

      expect(run.font.localeId).toEqual(1033);
      expect(run.font.size).toEqual(16);
      expect(run.font.italic).toEqual(false);
      expect(run.font.bold).toEqual(false);
      expect(run.font.localeIdBi).toEqual(1025);
      expect(run.font.sizeBi).toEqual(24);
      expect(run.font.nameBi).toEqual("Andalus");
      expect(run.font.italicBi).toEqual(true);
      expect(run.font.boldBi).toEqual(true);
    }
  });


  test.skip('FarEast: CultureInfo', () => {
    //ExStart
    //ExFor:aw.Font.nameFarEast
    //ExFor:aw.Font.localeIdFarEast
    //ExSummary:Shows how to insert and format text in a Far East language.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Specify font settings that the document builder will apply to any text that it inserts.
    builder.font.name = "Courier New";
    builder.font.localeId = new CultureInfo("en-US", false).LCID;

    // Name "FarEast" equivalents for our font and locale.
    // If the builder inserts Asian characters with this Font configuration, then each run that contains
    // these characters will display them using the "FarEast" font/locale instead of the default.
    // This could be useful when a western font does not have ideal representations for Asian characters.
    builder.font.nameFarEast = "SimSun";
    builder.font.localeIdFarEast = new CultureInfo("zh-CN", false).LCID;

    // This text will be displayed in the default font/locale.
    builder.writeln("Hello world!");

    // Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
    builder.writeln("你好世界");

    doc.save(base.artifactsDir + "Font.farEast.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.farEast.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Hello world!");
    expect(run.font.localeId).toEqual(1033);
    expect(run.font.name).toEqual("Courier New");
    expect(run.font.localeIdFarEast).toEqual(2052);
    expect(run.font.nameFarEast).toEqual("SimSun");

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("你好世界");
    expect(run.font.localeId).toEqual(1033);
    expect(run.font.name).toEqual("SimSun");
    expect(run.font.localeIdFarEast).toEqual(2052);
    expect(run.font.nameFarEast).toEqual("SimSun");
  });


  test('NameAscii', () => {
    //ExStart
    //ExFor:aw.Font.nameAscii
    //ExFor:aw.Font.nameOther
    //ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Suppose a run that we use the builder to insert while using this font configuration
    // contains characters within the ASCII characters' range. In that case,
    // it will display those characters using this font.
    builder.font.nameAscii = "Calibri";

    // With no other font specified, the builder will also apply this font to all characters that it inserts.
    expect(builder.font.name).toEqual("Calibri");

    // Specify a font to use for all characters outside of the ASCII range.
    // Ideally, this font should have a glyph for each required non-ASCII character code.
    builder.font.nameOther = "Courier New";

    // Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range.
    // Each character will be displayed using either of the fonts, depending on.
    builder.writeln("Hello, Привет");

    doc.save(base.artifactsDir + "Font.nameAscii.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.nameAscii.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("Hello, Привет");
    expect(run.font.name).toEqual("Calibri");
    expect(run.font.nameAscii).toEqual("Calibri");
    expect(run.font.nameOther).toEqual("Courier New");
  });


  test('ChangeStyle', () => {
    //ExStart
    //ExFor:aw.Font.styleName
    //ExFor:aw.Font.styleIdentifier
    //ExFor:StyleIdentifier
    //ExSummary:Shows how to change the style of existing text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways of referencing styles.
    // 1 -  Using the style name:
    builder.font.styleName = "Emphasis";
    builder.writeln("Text originally in \"Emphasis\" style");

    // 2 -  Using a built-in style identifier:
    builder.font.styleIdentifier = aw.StyleIdentifier.IntenseEmphasis;
    builder.writeln("Text originally in \"Intense Emphasis\" style");

    // Convert all uses of one style to another,
    // using the above methods to reference old and new styles.
    for (let run of doc.getChildNodes(aw.NodeType.Run, true).toArray().map(node => node.asRun()))
    {
      if (run.font.styleName == "Emphasis")
        run.font.styleName = "Strong";

      if (run.font.styleIdentifier == aw.StyleIdentifier.IntenseEmphasis)
        run.font.styleIdentifier = aw.StyleIdentifier.Strong;
    }

    doc.save(base.artifactsDir + "Font.ChangeStyle.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.ChangeStyle.docx");
    let docRun = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(docRun.getText().trim()).toEqual("Text originally in \"Emphasis\" style");
    expect(docRun.font.styleIdentifier).toEqual(aw.StyleIdentifier.Strong);
    expect(docRun.font.styleName).toEqual("Strong");

    docRun = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(docRun.getText().trim()).toEqual("Text originally in \"Intense Emphasis\" style");
    expect(docRun.font.styleIdentifier).toEqual(aw.StyleIdentifier.Strong);
    expect(docRun.font.styleName).toEqual("Strong");
  });


  test('BuiltIn', () => {
    //ExStart
    //ExFor:aw.Style.builtIn
    //ExSummary:Shows how to differentiate custom styles from built-in styles.
    let doc = new aw.Document();

    // When we create a document using Microsoft Word, or programmatically using Aspose.words,
    // the document will come with a collection of styles to apply to its text to modify its appearance.
    // We can access these built-in styles via the document's "Styles" collection.
    // These styles will all have the "BuiltIn" flag set to "true".
    let style = doc.styles.at("Emphasis");

    expect(style.builtIn).toEqual(true);

    // Create a custom style and add it to the collection.
    // Custom styles such as this will have the "BuiltIn" flag set to "false". 
    style = doc.styles.add(aw.StyleType.Character, "MyStyle");
    style.font.color = "#000080";
    style.font.name = "Courier New";

    expect(style.builtIn).toEqual(false);
    //ExEnd
  });


  test('Style', () => {
    //ExStart
    //ExFor:aw.Font.style
    //ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a custom style and apply it to text created using a document builder.
    let style = doc.styles.add(aw.StyleType.Character, "MyStyle");
    style.font.color = "#FF0000";
    style.font.name = "Courier New";

    builder.font.styleName = "MyStyle";
    builder.write("This text is in a custom style.");

    // Iterate over every run and add a double underline to every custom style.
    for (let run of doc.getChildNodes(aw.NodeType.Run, true).toArray().map(node => node.asRun()))
    {
      let charStyle = run.font.style;

      if (!charStyle.builtIn)
        run.font.underline = aw.Underline.Double;
    }

    doc.save(base.artifactsDir + "Font.style.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.style.docx");
    let docRun = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(docRun.getText().trim()).toEqual("This text is in a custom style.");
    expect(docRun.font.styleName).toEqual("MyStyle");
    expect(docRun.font.style.builtIn).toEqual(false);
    expect(docRun.font.underline).toEqual(aw.Underline.Double);
  });


  test.skip('GetAvailableFonts: Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', () => {
    //ExStart
    //ExFor:PhysicalFontInfo
    //ExFor:aw.Fonts.FontSourceBase.getAvailableFonts
    //ExFor:aw.Fonts.PhysicalFontInfo.fontFamilyName
    //ExFor:aw.Fonts.PhysicalFontInfo.fullFontName
    //ExFor:aw.Fonts.PhysicalFontInfo.version
    //ExFor:aw.Fonts.PhysicalFontInfo.filePath
    //ExSummary:Shows how to list available fonts.
    // Configure Aspose.words to source fonts from a custom folder, and then print every available font.
    let folderFontSource = [new aw.Fonts.FolderFontSource(base.fontsDir, true)];

    for (let fontInfo of folderFontSource.at(0).getAvailableFonts())
    {
      console.log("FontFamilyName : {0}", fontInfo.fontFamilyName);
      console.log("FullFontName  : {0}", fontInfo.fullFontName);
      console.log("Version  : {0}", fontInfo.version);
      console.log("FilePath : {0}\n", fontInfo.filePath);
    }
    //ExEnd

    expect(Directory.EnumerateFiles(base.fontsDir, "*.*", SearchOption.AllDirectories).Count(f => f.EndsWith(".ttf") || f.EndsWith(".otf"))).toEqual(folderFontSource.at(0).getAvailableFonts().Count);
  });


  test.skip('SetFontAutoColor: WORDSNODEJS-123', () => {
    //ExStart
    //ExFor:aw.Font.autoColor
    //ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If a run's Font object does not specify text color, it will automatically
    // select either black or white depending on the background color's color.
    expect(builder.font.color).toEqual("#000000");

    // The default color for text is black. If the color of the background is dark, black text will be difficult to see.
    // To solve this problem, the AutoColor property will display this text in white.
    builder.font.shading.backgroundPatternColor = "#00008B";

    builder.writeln("The text color automatically chosen for this run is white.");

    expect(doc.firstSection.body.paragraphs.at(0).runs.at(0).font.autoColor).toEqual("#FFFFFF");

    // If we change the background to a light color, black will be a more
    // suitable text color than white so that the auto color will display it in black.
    builder.font.shading.backgroundPatternColor = "#ADD8E6";

    builder.writeln("The text color automatically chosen for this run is black.");

    expect(doc.firstSection.body.paragraphs.at(1).runs.at(0).font.autoColor).toEqual("#000000");

    doc.save(base.artifactsDir + "Font.SetFontAutoColor.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Font.SetFontAutoColor.docx");
    let run = doc.firstSection.body.paragraphs.at(0).runs.at(0);

    expect(run.getText().trim()).toEqual("The text color automatically chosen for this run is white.");
    expect(run.font.color).toEqual("#000000");
    expect(run.font.shading.backgroundPatternColor).toEqual("#00008B");

    run = doc.firstSection.body.paragraphs.at(1).runs.at(0);

    expect(run.getText().trim()).toEqual("The text color automatically chosen for this run is black.");
    expect(run.font.color).toEqual("#000000");
    expect(run.font.shading.backgroundPatternColor).toEqual("#ADD8E6");
  });


  /*  //ExStart
    //ExFor:Font.Hidden
    //ExFor:Paragraph.Accept
    //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
    //ExFor:DocumentVisitor.VisitFormField(FormField)
    //ExFor:DocumentVisitor.VisitTableEnd(Table)
    //ExFor:DocumentVisitor.VisitCellEnd(Cell)
    //ExFor:DocumentVisitor.VisitRowEnd(Row)
    //ExFor:DocumentVisitor.VisitSpecialChar(SpecialChar)
    //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
    //ExFor:DocumentVisitor.VisitShapeStart(Shape)
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:SpecialChar
    //ExFor:Node.Accept
    //ExFor:Paragraph.ParagraphBreakFont
    //ExFor:Table.Accept
    //ExSummary:Shows how to use a DocumentVisitor implementation to remove all hidden content from a document.
  test('RemoveHiddenContentFromDocument', () => {
    let doc = new aw.Document(base.myDir + "Hidden content.docx");
    expect(doc.getChildNodes(aw.NodeType.Paragraph, true).Count).toEqual(26);
    expect(doc.getChildNodes(aw.NodeType.Table, true).Count).toEqual(2);

    let hiddenContentRemover = new RemoveHiddenContentVisitor();

    // Below are three types of fields which can accept a document visitor,
    // which will allow it to visit the accepting node, and then traverse its child nodes in a depth-first manner.
    // 1 -  Paragraph node:
    let para = (Paragraph)doc.getChild(aw.NodeType.Paragraph, 4, true);
    para.accept(hiddenContentRemover);

    // 2 -  Table node:
    let table = doc.firstSection.body.tables.at(0);
    table.accept(hiddenContentRemover);

    // 3 -  Document node:
    doc.accept(hiddenContentRemover);

    doc.save(base.artifactsDir + "Font.RemoveHiddenContentFromDocument.docx");
    TestRemoveHiddenContent(new aw.Document(base.artifactsDir + "Font.RemoveHiddenContentFromDocument.docx")); //ExSkip
  });


    /// <summary>
    /// Removes all visited nodes marked as "hidden content".
    /// </summary>
  public class RemoveHiddenContentVisitor : DocumentVisitor
  {
      /// <summary>
      /// Called when a FieldStart node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldStart(FieldStart fieldStart)
    {
      if (fieldStart.font.hidden)
        fieldStart.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a FieldEnd node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
    {
      if (fieldEnd.font.hidden)
        fieldEnd.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a FieldSeparator node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
    {
      if (fieldSeparator.font.hidden)
        fieldSeparator.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Run node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitRun(Run run)
    {
      if (run.font.hidden)
        run.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Paragraph node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitParagraphStart(Paragraph paragraph)
    {
      if (paragraph.paragraphBreakFont.hidden)
        paragraph.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a FormField is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFormField(FormField formField)
    {
      if (formField.font.hidden)
        formField.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a GroupShape is encountered in the document.
      /// </summary>
    public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
    {
      if (groupShape.font.hidden)
        groupShape.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Shape is encountered in the document.
      /// </summary>
    public override VisitorAction VisitShapeStart(Shape shape)
    {
      if (shape.font.hidden)
        shape.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Comment is encountered in the document.
      /// </summary>
    public override VisitorAction VisitCommentStart(Comment comment)
    {
      if (comment.font.hidden)
        comment.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Footnote is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFootnoteStart(Footnote footnote)
    {
      if (footnote.font.hidden)
        footnote.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a SpecialCharacter is encountered in the document.
      /// </summary>
    public override VisitorAction VisitSpecialChar(SpecialChar specialChar)
    {
      if (specialChar.font.hidden)
        specialChar.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when visiting of a Table node is ended in the document.
      /// </summary>
    public override VisitorAction VisitTableEnd(Table table)
    {
        // The content inside table cells may have the hidden content flag, but the tables themselves cannot.
        // If this table had nothing but hidden content, this visitor would have removed all of it,
        // and there would be no child nodes left.
        // Thus, we can also treat the table itself as hidden content and remove it.
        // Tables which are empty but do not have hidden content will have cells with empty paragraphs inside,
        // which this visitor will not remove.
      if (!table.hasChildNodes)
        table.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when visiting of a Cell node is ended in the document.
      /// </summary>
    public override VisitorAction VisitCellEnd(Cell cell)
    {
      if (!cell.hasChildNodes && cell.parentNode != null)
        cell.remove();

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when visiting of a Row node is ended in the document.
      /// </summary>
    public override VisitorAction VisitRowEnd(Row row)
    {
      if (!row.hasChildNodes && row.parentNode != null)
        row.remove();

      return aw.VisitorAction.Continue;
    }
  }
    //ExEnd

  private void TestRemoveHiddenContent(Document doc)
  {
    expect(doc.getChildNodes(aw.NodeType.Paragraph, true).Count).toEqual(20);
    expect(doc.getChildNodes(aw.NodeType.Table, true).Count).toEqual(1);

    foreach (Node node in doc.getChildNodes(aw.NodeType.Any, true))
    {
      switch (node)
      {
        case FieldStart fieldStart:
          expect(fieldStart.font.hidden).toEqual(false);
          break;
        case FieldEnd fieldEnd:
          expect(fieldEnd.font.hidden).toEqual(false);
          break;
        case FieldSeparator fieldSeparator:
          expect(fieldSeparator.font.hidden).toEqual(false);
          break;
        case Run run:
          expect(run.font.hidden).toEqual(false);
          break;
        case Paragraph paragraph:
          expect(paragraph.paragraphBreakFont.hidden).toEqual(false);
          break;
        case FormField formField:
          expect(formField.font.hidden).toEqual(false);
          break;
        case GroupShape groupShape:
          expect(groupShape.font.hidden).toEqual(false);
          break;
        case Shape shape:
          expect(shape.font.hidden).toEqual(false);
          break;
        case Comment comment:
          expect(comment.font.hidden).toEqual(false);
          break;
        case Footnote footnote:
          expect(footnote.font.hidden).toEqual(false);
          break;
        case SpecialChar specialChar:
          expect(specialChar.font.hidden).toEqual(false);
          break;
      }
    }
  }*/

  test('DefaultFonts', () => {
    //ExStart
    //ExFor:aw.Fonts.FontInfoCollection.contains(String)
    //ExFor:aw.Fonts.FontInfoCollection.count
    //ExSummary:Shows info about the fonts that are present in the blank document.
    let doc = new aw.Document();

    // A blank document contains 3 default fonts. Each font in the document
    // will have a corresponding FontInfo object which contains details about that font.
    expect(doc.fontInfos.count).toEqual(3);

    expect(doc.fontInfos.contains("Times New Roman")).toEqual(true);
    expect(doc.fontInfos.at("Times New Roman").charset).toEqual(204);

    expect(doc.fontInfos.contains("Symbol")).toEqual(true);
    expect(doc.fontInfos.contains("Arial")).toEqual(true);
    //ExEnd
  });


  test('ExtractEmbeddedFont', () => {
    //ExStart
    //ExFor:EmbeddedFontFormat
    //ExFor:EmbeddedFontStyle
    //ExFor:aw.Fonts.FontInfo.getEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
    //ExFor:aw.Fonts.FontInfo.getEmbeddedFontAsOpenType(EmbeddedFontStyle)
    //ExFor:aw.Fonts.FontInfoCollection.item(Int32)
    //ExFor:aw.Fonts.FontInfoCollection.item(String)
    //ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
    let doc = new aw.Document(base.myDir + "Embedded font.docx");

    let embeddedFont = doc.fontInfos.at("Alte DIN 1451 Mittelschrift");
    let embeddedFontBytes = embeddedFont.getEmbeddedFont(aw.Fonts.EmbeddedFontFormat.OpenType, aw.Fonts.EmbeddedFontStyle.Regular);
    expect(embeddedFontBytes).not.toBe(null);

    fs.writeFileSync(base.artifactsDir + "Alte DIN 1451 Mittelschrift.ttf", Buffer.from(embeddedFontBytes));

    // Embedded font formats may be different in other formats such as .doc.
    // We need to know the correct format before we can extract the font.
    doc = new aw.Document(base.myDir + "Embedded font.doc");

    expect(doc.fontInfos.at("Alte DIN 1451 Mittelschrift").getEmbeddedFont(aw.Fonts.EmbeddedFontFormat.OpenType, aw.Fonts.EmbeddedFontStyle.Regular)).toBe(null);
    expect(doc.fontInfos.at("Alte DIN 1451 Mittelschrift").getEmbeddedFont(aw.Fonts.EmbeddedFontFormat.EmbeddedOpenType, aw.Fonts.EmbeddedFontStyle.Regular)).not.toBe(null);

    // Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
    embeddedFontBytes = doc.fontInfos.at("Alte DIN 1451 Mittelschrift").getEmbeddedFontAsOpenType(aw.Fonts.EmbeddedFontStyle.Regular);

    fs.writeFileSync(base.artifactsDir + "Alte DIN 1451 Mittelschrift.otf", Buffer.from(embeddedFontBytes));
    //ExEnd
  });


  test('GetFontInfoFromFile', () => {
    //ExStart
    //ExFor:FontFamily
    //ExFor:FontPitch
    //ExFor:aw.Fonts.FontInfo.altName
    //ExFor:aw.Fonts.FontInfo.charset
    //ExFor:aw.Fonts.FontInfo.family
    //ExFor:aw.Fonts.FontInfo.panose
    //ExFor:aw.Fonts.FontInfo.pitch
    //ExFor:aw.Fonts.FontInfoCollection.getEnumerator
    //ExSummary:Shows how to access and print details of each font in a document.
    let doc = new aw.Document(base.myDir + "Document.docx");

    for (let fontInfo of doc.fontInfos)
    {
      if (fontInfo != null)
      {
        console.log("Font name: " + fontInfo.name);

        // Alt names are usually blank.
        console.log("Alt name: " + fontInfo.altName);
        console.log("\t- Family: " + fontInfo.family);
        console.log("\t- " + (fontInfo.isTrueType ? "Is TrueType" : "Is not TrueType"));
        console.log("\t- Pitch: " + fontInfo.pitch);
        console.log("\t- Charset: " + fontInfo.charset);
        console.log("\t- Panose:");
        console.log("\t\tFamily Kind: " + fontInfo.panose.at(0));
        console.log("\t\tSerif Style: " + fontInfo.panose.at(1));
        console.log("\t\tWeight: " + fontInfo.panose.at(2));
        console.log("\t\tProportion: " + fontInfo.panose.at(3));
        console.log("\t\tContrast: " + fontInfo.panose.at(4));
        console.log("\t\tStroke Variation: " + fontInfo.panose.at(5));
        console.log("\t\tArm Style: " + fontInfo.panose.at(6));
        console.log("\t\tLetterform: " + fontInfo.panose.at(7));
        console.log("\t\tMidline: " + fontInfo.panose.at(8));
        console.log("\t\tX-Height: " + fontInfo.panose.at(9));
      }
    }
    //ExEnd

    expect([2, 15, 5, 2, 2, 2, 4, 3, 2, 4]).toEqual(doc.fontInfos.at("Calibri").panose);
    expect([2, 15, 3, 2, 2, 2, 4, 3, 2, 4]).toEqual(doc.fontInfos.at("Calibri Light").panose);
    expect([2, 2, 6, 3, 5, 4, 5, 2, 3, 4]).toEqual(doc.fontInfos.at("Times New Roman").panose);
  });


  test('LineSpacing', () => {
    //ExStart
    //ExFor:aw.Font.lineSpacing
    //ExSummary:Shows how to get a font's line spacing, in points.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set different fonts for the DocumentBuilder and verify their line spacing.
    builder.font.name = "Calibri";
    expect(builder.font.lineSpacing).toEqual(14.6484375);

    builder.font.name = "Times New Roman";
    expect(builder.font.lineSpacing).toEqual(13.798828125);
    //ExEnd
  });


  test('HasDmlEffect', () => {
    //ExStart
    //ExFor:aw.Font.hasDmlEffect(TextDmlEffect)
    //ExSummary:Shows how to check if a run displays a DrawingML text effect.
    let doc = new aw.Document(base.myDir + "DrawingML text effects.docx");

    let runs = doc.firstSection.body.firstParagraph.runs;

    expect(runs.at(0).font.hasDmlEffect(aw.TextDmlEffect.Shadow)).toEqual(true);
    expect(runs.at(1).font.hasDmlEffect(aw.TextDmlEffect.Shadow)).toEqual(true);
    expect(runs.at(2).font.hasDmlEffect(aw.TextDmlEffect.Reflection)).toEqual(true);
    expect(runs.at(3).font.hasDmlEffect(aw.TextDmlEffect.Effect3D)).toEqual(true);
    expect(runs.at(4).font.hasDmlEffect(aw.TextDmlEffect.Fill)).toEqual(true);
    //ExEnd
  });


  test.skip('CheckScanUserFontsFolder: Aspose.Words.Fonts.FontSourceBase.GetAvailableFonts() skipped', () => {
    let userProfile = os.homedir();
    let currentUserFontsFolder = path.join(userProfile, "AppData\\Local\\Microsoft\\Windows\\Fonts");
    console.log(currentUserFontsFolder);
    let currentUserFonts = fs.readdirSync(currentUserFontsFolder).map(file => file.endsWith(".ttf"));
    if (currentUserFonts.length != 0)
    {
      // On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
      // or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user.
      let systemFontSource = new aw.Fonts.SystemFontSource();
      Assert.NotNull(systemFontSource.getAvailableFonts()
          .FirstOrDefault(x => x.filePath.contains("\\AppData\\Local\\Microsoft\\Windows\\Fonts")),
        "Fonts did not install to the user font folder");
    }
  });


  test.each([aw.EmphasisMark.None,
    aw.EmphasisMark.OverComma,
    aw.EmphasisMark.OverSolidCircle,
    aw.EmphasisMark.OverWhiteCircle,
    aw.EmphasisMark.UnderSolidCircle])('SetEmphasisMark(emphasisMark = %o)', (emphasisMark) => {
    //ExStart
    //ExFor:EmphasisMark
    //ExFor:aw.Font.emphasisMark
    //ExSummary:Shows how to add additional character rendered above/below the glyph-character.
    let builder = new aw.DocumentBuilder();

    // Possible types of emphasis mark:
    // https://apireference.aspose.com/words/net/aspose.words/emphasismark
    builder.font.emphasisMark = emphasisMark; 

    builder.write("Emphasis text");
    builder.writeln();
    builder.font.clearFormatting();
    builder.write("Simple text");
 
    builder.document.save(base.artifactsDir + "Fonts.SetEmphasisMark.docx");
    //ExEnd
  });


  test.skip('ThemeFontsColors: WORDSNODEJS-123', () => {
    //ExStart
    //ExFor:aw.Font.themeFont
    //ExFor:aw.Font.themeFontAscii
    //ExFor:aw.Font.themeFontBi
    //ExFor:aw.Font.themeFontFarEast
    //ExFor:aw.Font.themeFontOther
    //ExFor:aw.Font.themeColor
    //ExFor:ThemeFont
    //ExFor:ThemeColor
    //ExSummary:Shows how to work with theme fonts and colors.
    let doc = new aw.Document();

    // Define fonts for languages uses by default.
    doc.theme.minorFonts.latin = "Algerian";
    doc.theme.minorFonts.eastAsian = "Aharoni";
    doc.theme.minorFonts.complexScript = "Andalus";

    let font = doc.styles.at("Normal").font;
    console.log(`Originally the Normal style theme color is: ${font.themeColor} and RGB color is: ${font.color}\n`);

    // We can use theme font and color instead of default values.
    font.themeFont = aw.Themes.ThemeFont.Minor;
    font.themeColor = aw.Themes.ThemeColor.Accent2;

    expect(font.themeFont).toEqual(aw.Themes.ThemeFont.Minor);
    expect(font.name).toEqual("Algerian");

    expect(font.themeFontAscii).toEqual(aw.Themes.ThemeFont.Minor);
    expect(font.nameAscii).toEqual("Algerian");

    expect(font.themeFontBi).toEqual(aw.Themes.ThemeFont.Minor);
    expect(font.nameBi).toEqual("Andalus");

    expect(font.themeFontFarEast).toEqual(aw.Themes.ThemeFont.Minor);
    expect(font.nameFarEast).toEqual("Aharoni");

    expect(font.themeFontOther).toEqual(aw.Themes.ThemeFont.Minor);
    expect(font.nameOther).toEqual("Algerian");

    expect(font.themeColor).toEqual(aw.Themes.ThemeColor.Accent2);
    expect(font.color).toEqual("#000000");

    // There are several ways of reset them font and color.
    // 1 -  By setting aw.Themes.ThemeFont.None/aw.Themes.ThemeColor.None:
    font.themeFont = aw.Themes.ThemeFont.None;
    font.themeColor = aw.Themes.ThemeColor.None;

    expect(font.themeFont).toEqual(aw.Themes.ThemeFont.None);
    expect(font.name).toEqual("Algerian");

    expect(font.themeFontAscii).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameAscii).toEqual("Algerian");

    expect(font.themeFontBi).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameBi).toEqual("Andalus");

    expect(font.themeFontFarEast).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameFarEast).toEqual("Aharoni");

    expect(font.themeFontOther).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameOther).toEqual("Algerian");

    expect(font.themeColor).toEqual(aw.Themes.ThemeColor.None);
    expect(font.color).toEqual("#000000");

    // 2 -  By setting non-theme font/color names:
    font.name = "Arial";
    font.color = "#0000FF";

    expect(font.themeFont).toEqual(aw.Themes.ThemeFont.None);
    expect(font.name).toEqual("Arial");

    expect(font.themeFontAscii).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameAscii).toEqual("Arial");

    expect(font.themeFontBi).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameBi).toEqual("Arial");

    expect(font.themeFontFarEast).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameFarEast).toEqual("Arial");

    expect(font.themeFontOther).toEqual(aw.Themes.ThemeFont.None);
    expect(font.nameOther).toEqual("Arial");

    expect(font.themeColor).toEqual(aw.Themes.ThemeColor.None);
    expect(font.color).toEqual("#0000FF");
    //ExEnd
  });


  test.skip('CreateThemedStyle: WORDSNODEJS-123', () => {
    //ExStart
    //ExFor:aw.Font.themeFont
    //ExFor:aw.Font.themeColor
    //ExFor:aw.Font.tintAndShade
    //ExFor:ThemeFont
    //ExFor:ThemeColor
    //ExSummary:Shows how to create and use themed style.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln();

    // Create some style with theme font properties.
    let style = doc.styles.add(aw.StyleType.Paragraph, "ThemedStyle");
    style.font.themeFont = aw.Themes.ThemeFont.Major;
    style.font.themeColor = aw.Themes.ThemeColor.Accent5;
    style.font.tintAndShade = 0.3;

    builder.paragraphFormat.styleName = "ThemedStyle";
    builder.writeln("Text with themed style");
    //ExEnd

    let run = (builder.currentParagraph.previousSibling.asParagraph()).firstChild.asRun();

    expect(run.font.themeFont).toEqual(aw.Themes.ThemeFont.Major);
    expect(run.font.name).toEqual("Times New Roman");

    expect(run.font.themeFontAscii).toEqual(aw.Themes.ThemeFont.Major);
    expect(run.font.nameAscii).toEqual("Times New Roman");

    expect(run.font.themeFontBi).toEqual(aw.Themes.ThemeFont.Major);
    expect(run.font.nameBi).toEqual("Times New Roman");

    expect(run.font.themeFontFarEast).toEqual(aw.Themes.ThemeFont.Major);
    expect(run.font.nameFarEast).toEqual("Times New Roman");

    expect(run.font.themeFontOther).toEqual(aw.Themes.ThemeFont.Major);
    expect(run.font.nameOther).toEqual("Times New Roman");

    expect(run.font.themeColor).toEqual(aw.Themes.ThemeColor.Accent5);
    expect(run.font.color).toEqual("#000000");
  });

});
