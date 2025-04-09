// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExParagraphFormat", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AsianTypographyProperties', () => {
    //ExStart
    //ExFor:ParagraphFormat.farEastLineBreakControl
    //ExFor:ParagraphFormat.wordWrap
    //ExFor:ParagraphFormat.hangingPunctuation
    //ExSummary:Shows how to set special properties for Asian typography. 
    let doc = new aw.Document(base.myDir + "Document.docx");

    let format = doc.firstSection.body.firstParagraph.paragraphFormat;
    format.farEastLineBreakControl = true;
    format.wordWrap = false;
    format.hangingPunctuation = true;

    doc.save(base.artifactsDir + "ParagraphFormat.AsianTypographyProperties.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.AsianTypographyProperties.docx");
    format = doc.firstSection.body.firstParagraph.paragraphFormat;

    expect(format.farEastLineBreakControl).toEqual(true);
    expect(format.wordWrap).toEqual(false);
    expect(format.hangingPunctuation).toEqual(true);
  });


  test.each([aw.DropCapPosition.Margin,
    aw.DropCapPosition.Normal,
    aw.DropCapPosition.None])('DropCap(%o)', (dropCapPosition) => {
    //ExStart
    //ExFor:DropCapPosition
    //ExSummary:Shows how to create a drop cap.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert one paragraph with a large letter that the text in the second and third paragraphs begins with.
    builder.font.size = 54;
    builder.writeln("L");

    builder.font.size = 18;
    builder.writeln("orem ipsum dolor sit amet, consectetur adipiscing elit, " +
      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");
    builder.writeln("Ut enim ad minim veniam, quis nostrud exercitation " +
      "ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    // Currently, the second and third paragraphs will appear underneath the first.
    // We can convert the first paragraph as a drop cap for the other paragraphs via its "ParagraphFormat" object.
    // Set the "DropCapPosition" property to "DropCapPosition.Margin" to place the drop cap
    // outside the left-hand side page margin if our text is left-to-right.
    // Set the "DropCapPosition" property to "DropCapPosition.Normal" to place the drop cap within the page margins
    // and to wrap the rest of the text around it.
    // "DropCapPosition.None" is the default state for all paragraphs.
    let format = doc.firstSection.body.firstParagraph.paragraphFormat;
    format.dropCapPosition = dropCapPosition;

    doc.save(base.artifactsDir + "ParagraphFormat.DropCap.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.DropCap.docx");

    expect(doc.firstSection.body.paragraphs.at(0).asParagraph().paragraphFormat.dropCapPosition).toEqual(dropCapPosition);
    expect(doc.firstSection.body.paragraphs.at(1).asParagraph().paragraphFormat.dropCapPosition).toEqual(aw.DropCapPosition.None);
  });

  test('LineSpacing', () => {
    //ExStart
    //ExFor:ParagraphFormat.lineSpacing
    //ExFor:ParagraphFormat.lineSpacingRule
    //ExFor:LineSpacingRule
    //ExSummary:Shows how to work with line spacing.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are three line spacing rules that we can define using the
    // paragraph's "LineSpacingRule" property to configure spacing between paragraphs.
    // 1 -  Set a minimum amount of spacing.
    // This will give vertical padding to lines of text of any size
    // that is too small to maintain the minimum line-height.
    builder.paragraphFormat.lineSpacingRule = aw.LineSpacingRule.AtLeast;
    builder.paragraphFormat.lineSpacing = 20;

    builder.writeln("Minimum line spacing of 20.");
    builder.writeln("Minimum line spacing of 20.");

    // 2 -  Set exact spacing.
    // Using font sizes that are too large for the spacing will truncate the text.
    builder.paragraphFormat.lineSpacingRule = aw.LineSpacingRule.Exactly;
    builder.paragraphFormat.lineSpacing = 5;

    builder.writeln("Line spacing of exactly 5.");
    builder.writeln("Line spacing of exactly 5.");

    // 3 -  Set spacing as a multiple of default line spacing, which is 12 points by default.
    // This kind of spacing will scale to different font sizes.
    builder.paragraphFormat.lineSpacingRule = aw.LineSpacingRule.Multiple;
    builder.paragraphFormat.lineSpacing = 18;

    builder.writeln("Line spacing of 1.5 default lines.");
    builder.writeln("Line spacing of 1.5 default lines.");

    doc.save(base.artifactsDir + "ParagraphFormat.lineSpacing.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.lineSpacing.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.AtLeast);
    expect(paragraphs.at(0).paragraphFormat.lineSpacing).toEqual(20.0);
    expect(paragraphs.at(1).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.AtLeast);
    expect(paragraphs.at(1).paragraphFormat.lineSpacing).toEqual(20.0);

    expect(paragraphs.at(2).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.Exactly);
    expect(paragraphs.at(2).paragraphFormat.lineSpacing).toEqual(5.0);
    expect(paragraphs.at(3).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.Exactly);
    expect(paragraphs.at(3).paragraphFormat.lineSpacing).toEqual(5.0);

    expect(paragraphs.at(4).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.Multiple);
    expect(paragraphs.at(4).paragraphFormat.lineSpacing).toEqual(18.0);
    expect(paragraphs.at(5).paragraphFormat.lineSpacingRule).toEqual(aw.LineSpacingRule.Multiple);
    expect(paragraphs.at(5).paragraphFormat.lineSpacing).toEqual(18.0);
  });

  test.each([false, true])('ParagraphSpacingAuto(%o)', (autoSpacing) => {
    //ExStart
    //ExFor:ParagraphFormat.spaceAfter
    //ExFor:ParagraphFormat.spaceAfterAuto
    //ExFor:ParagraphFormat.spaceBefore
    //ExFor:ParagraphFormat.spaceBeforeAuto
    //ExSummary:Shows how to set automatic paragraph spacing.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Apply a large amount of spacing before and after paragraphs that this builder will create.
    builder.paragraphFormat.spaceBefore = 24;
    builder.paragraphFormat.spaceAfter = 24;

    // Set these flags to "true" to apply automatic spacing,
    // effectively ignoring the spacing in the properties we set above.
    // Leave them as "false" will apply our custom paragraph spacing.
    builder.paragraphFormat.spaceAfterAuto = autoSpacing;
    builder.paragraphFormat.spaceBeforeAuto = autoSpacing;

    // Insert two paragraphs that will have spacing above and below them and save the document.
    builder.writeln("Paragraph 1.");
    builder.writeln("Paragraph 2.");

    doc.save(base.artifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.ParagraphSpacingAuto.docx");
    let format = doc.firstSection.body.paragraphs.at(0).asParagraph().paragraphFormat;

    expect(format.spaceBefore).toEqual(24.0);
    expect(format.spaceAfter).toEqual(24.0);
    expect(format.spaceAfterAuto).toEqual(autoSpacing);
    expect(format.spaceBeforeAuto).toEqual(autoSpacing);

    format = doc.firstSection.body.paragraphs.at(1).paragraphFormat;

    expect(format.spaceBefore).toEqual(24.0);
    expect(format.spaceAfter).toEqual(24.0);
    expect(format.spaceAfterAuto).toEqual(autoSpacing);
    expect(format.spaceBeforeAuto).toEqual(autoSpacing);
  });

  test.each([false, true])('ParagraphSpacingSameStyle(%o)', (noSpaceBetweenParagraphsOfSameStyle) => {
    //ExStart
    //ExFor:ParagraphFormat.spaceAfter
    //ExFor:ParagraphFormat.spaceBefore
    //ExFor:ParagraphFormat.noSpaceBetweenParagraphsOfSameStyle
    //ExSummary:Shows how to apply no spacing between paragraphs with the same style.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Apply a large amount of spacing before and after paragraphs that this builder will create.
    builder.paragraphFormat.spaceBefore = 24;
    builder.paragraphFormat.spaceAfter = 24;

    // Set the "NoSpaceBetweenParagraphsOfSameStyle" flag to "true" to apply
    // no spacing between paragraphs with the same style, which will group similar paragraphs.
    // Leave the "NoSpaceBetweenParagraphsOfSameStyle" flag as "false"
    // to evenly apply spacing to every paragraph.
    builder.paragraphFormat.noSpaceBetweenParagraphsOfSameStyle = noSpaceBetweenParagraphsOfSameStyle;

    builder.paragraphFormat.style = doc.styles.at("Normal");
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.paragraphFormat.style = doc.styles.at("Quote");
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.paragraphFormat.style = doc.styles.at("Normal");
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);
    builder.writeln(`Paragraph in the \"${builder.paragraphFormat.style.name}\" style.`);

    doc.save(base.artifactsDir + "ParagraphFormat.ParagraphSpacingSameStyle.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.ParagraphSpacingSameStyle.docx");

    for (let paragraph of doc.firstSection.body.paragraphs)
    {
      let format = paragraph.asParagraph().paragraphFormat;

      expect(format.spaceBefore).toEqual(24.0);
      expect(format.spaceAfter).toEqual(24.0);
      expect(format.noSpaceBetweenParagraphsOfSameStyle).toEqual(noSpaceBetweenParagraphsOfSameStyle);
    }
  });

  test('ParagraphOutlineLevel', () => {
    //ExStart
    //ExFor:ParagraphFormat.outlineLevel
    //ExFor:OutlineLevel
    //ExSummary:Shows how to configure paragraph outline levels to create collapsible text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BodyText" value.
    // Setting the property to one of the numbered values will show an arrow to the left
    // of the beginning of the paragraph.
    builder.paragraphFormat.outlineLevel = aw.OutlineLevel.Level1;
    builder.writeln("Paragraph outline level 1.");

    // Level 1 is the topmost level. If there is a paragraph with a lower level below a paragraph with a higher level,
    // collapsing the higher-level paragraph will collapse the lower level paragraph.
    builder.paragraphFormat.outlineLevel = aw.OutlineLevel.Level2;
    builder.writeln("Paragraph outline level 2.");

    // Two paragraphs of the same level will not collapse each other,
    // and the arrows do not collapse the paragraphs they point to.
    builder.paragraphFormat.outlineLevel = aw.OutlineLevel.Level3;
    builder.writeln("Paragraph outline level 3.");
    builder.writeln("Paragraph outline level 3.");

    // The default "BodyText" value is the lowest, which a paragraph of any level can collapse.
    builder.paragraphFormat.outlineLevel = aw.OutlineLevel.BodyText;
    builder.writeln("Paragraph at main text level.");

    doc.save(base.artifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.ParagraphOutlineLevel.docx");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.outlineLevel).toEqual(aw.OutlineLevel.Level1);
    expect(paragraphs.at(1).paragraphFormat.outlineLevel).toEqual(aw.OutlineLevel.Level2);
    expect(paragraphs.at(2).paragraphFormat.outlineLevel).toEqual(aw.OutlineLevel.Level3);
    expect(paragraphs.at(3).paragraphFormat.outlineLevel).toEqual(aw.OutlineLevel.Level3);
    expect(paragraphs.at(4).paragraphFormat.outlineLevel).toEqual(aw.OutlineLevel.BodyText);
  });


  test.each([false, true])('PageBreakBefore(%o)', (pageBreakBefore) => {
    //ExStart
    //ExFor:ParagraphFormat.pageBreakBefore
    //ExSummary:Shows how to create paragraphs with page breaks at the beginning.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set this flag to "true" to apply a page break to each paragraph's beginning
    // that the document builder will create under this ParagraphFormat configuration.
    // The first paragraph will not receive a page break.
    // Leave this flag as "false" to start each new paragraph on the same page
    // as the previous, provided there is sufficient space.
    builder.paragraphFormat.pageBreakBefore = pageBreakBefore;

    builder.writeln("Paragraph 1.");
    builder.writeln("Paragraph 2.");

    let layoutCollector = new aw.Layout.LayoutCollector(doc);
    let paragraphs = doc.firstSection.body.paragraphs;

    if (pageBreakBefore)
    {
      expect(layoutCollector.getStartPageIndex(paragraphs.at(0))).toEqual(1);
      expect(layoutCollector.getStartPageIndex(paragraphs.at(1))).toEqual(2);
    }
    else
    {
      expect(layoutCollector.getStartPageIndex(paragraphs.at(0))).toEqual(1);
      expect(layoutCollector.getStartPageIndex(paragraphs.at(1))).toEqual(1);
    }

    doc.save(base.artifactsDir + "ParagraphFormat.pageBreakBefore.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.pageBreakBefore.docx");
    paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.pageBreakBefore).toEqual(pageBreakBefore);
    expect(paragraphs.at(1).paragraphFormat.pageBreakBefore).toEqual(pageBreakBefore);
  });


  test.each([false, true])('WidowControl(%o)', (widowControl) => {
    //ExStart
    //ExFor:ParagraphFormat.widowControl
    //ExSummary:Shows how to enable widow/orphan control for a paragraph.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // When we write the text that does not fit onto one page, one line may spill over onto the next page.
    // The single line that ends up on the next page is called an "Orphan",
    // and the previous line where the orphan broke off is called a "Widow".
    // We can fix orphans and widows by rearranging text via font size, spacing, or page margins.
    // If we wish to preserve our document's dimensions, we can set this flag to "true"
    // to push widows onto the same page as their respective orphans. 
    // Leave this flag as "false" will leave widow/orphan pairs in text.
    // Every paragraph has this setting accessible in Microsoft Word via Home -> Paragraph -> Paragraph Settings
    // (button on bottom right hand corner of "Paragraph" tab) -> "Widow/Orphan control".
    builder.paragraphFormat.widowControl = widowControl; 

    // Insert text that produces an orphan and a widow.
    builder.font.size = 68;
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    doc.save(base.artifactsDir + "ParagraphFormat.widowControl.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.widowControl.docx");

    expect(doc.firstSection.body.paragraphs.at(0).paragraphFormat.widowControl).toEqual(widowControl);
  });


   test('LinesToDrop', () => {
    //ExStart
    //ExFor:ParagraphFormat.linesToDrop
    //ExSummary:Shows how to set the size of a drop cap.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Modify the "LinesToDrop" property to designate a paragraph as a drop cap,
    // which will turn it into a large capital letter that will decorate the next paragraph.
    // Give this property a value of 4 to give the drop cap the height of four text lines.
    builder.paragraphFormat.linesToDrop = 4;
    builder.writeln("H");

    // Reset the "LinesToDrop" property to 0 to turn the next paragraph into an ordinary paragraph.
    // The text in this paragraph will wrap around the drop cap.
    builder.paragraphFormat.linesToDrop = 0;
    builder.writeln("ello world!");

    doc.save(base.artifactsDir + "ParagraphFormat.linesToDrop.odt");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.linesToDrop.odt");
    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.at(0).paragraphFormat.linesToDrop).toEqual(4);
    expect(paragraphs.at(1).paragraphFormat.linesToDrop).toEqual(0);
  });

  test('ParagraphSpacingAndIndents', () => {
    //ExStart
    //ExFor:ParagraphFormat.characterUnitLeftIndent
    //ExFor:ParagraphFormat.characterUnitRightIndent
    //ExFor:ParagraphFormat.characterUnitFirstLineIndent
    //ExFor:ParagraphFormat.lineUnitBefore
    //ExFor:ParagraphFormat.lineUnitAfter
    //ExSummary:Shows how to change paragraph spacing and indents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let format = doc.firstSection.body.firstParagraph.paragraphFormat;

    // Below are five different spacing options, along with the properties that their configuration indirectly affects.
    // 1 -  Left indent:
    expect(0.0).toEqual(format.leftIndent);

    format.characterUnitLeftIndent = 10.0;

    expect(120.0).toEqual(format.leftIndent);

    // 2 -  Right indent:
    expect(0.0).toEqual(format.rightIndent);

    format.characterUnitRightIndent = -5.5;

    expect(-66.0).toEqual(format.rightIndent);

    // 3 -  Hanging indent:
    expect(0.0).toEqual(format.firstLineIndent);

    format.characterUnitFirstLineIndent = 20.3;

    expect(format.firstLineIndent).toBeCloseTo(243.59, 1);

    // 4 -  Line spacing before paragraphs:
    expect(0.0).toEqual(format.spaceBefore);

    format.lineUnitBefore = 5.1;

    expect(format.spaceBefore).toBeCloseTo(61.1, 1);

    // 5 -  Line spacing after paragraphs:
    expect(0.0).toEqual(format.spaceAfter);

    format.lineUnitAfter = 10.9;

    expect(format.spaceAfter).toBeCloseTo(130.8, 1);

    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
      "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder.write("测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试" +
      "文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    format = doc.firstSection.body.firstParagraph.paragraphFormat;

    expect(10.0).toEqual(format.characterUnitLeftIndent);
    expect(120.0).toEqual(format.leftIndent);
            
    expect(-5.5).toEqual(format.characterUnitRightIndent);
    expect(-66.0).toEqual(format.rightIndent);

    expect(20.3).toEqual(format.characterUnitFirstLineIndent);
    expect(format.firstLineIndent).toBeCloseTo(243.59, 1);

    expect(format.lineUnitBefore).toBeCloseTo(5.1, 1);
    expect(format.spaceBefore).toBeCloseTo(61.1, 1);

    expect(10.9).toEqual(format.lineUnitAfter);
    expect(format.spaceAfter).toBeCloseTo(130.8, 1);
  });


  test('ParagraphBaselineAlignment', () => {
    //ExStart
    //ExFor:BaselineAlignment
    //ExFor:ParagraphFormat.baselineAlignment
    //ExSummary:Shows how to set fonts vertical position on a line.
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let format = doc.firstSection.body.paragraphs.at(0).paragraphFormat;
    if (format.baselineAlignment == aw.BaselineAlignment.Auto)
    {
      format.baselineAlignment = aw.BaselineAlignment.Top;
    }

    doc.save(base.artifactsDir + "ParagraphFormat.ParagraphBaselineAlignment.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.ParagraphBaselineAlignment.docx");
    format = doc.firstSection.body.paragraphs.at(0).paragraphFormat;
    expect(format.baselineAlignment).toEqual(aw.BaselineAlignment.Top);
  });


  test('MirrorIndents', () => {
    //ExStart:MirrorIndents
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ParagraphFormat.mirrorIndents
    //ExSummary:Show how to make left and right indents the same.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let format = doc.firstSection.body.paragraphs.at(0).paragraphFormat;

    format.mirrorIndents = true;

    doc.save(base.artifactsDir + "ParagraphFormat.mirrorIndents.docx");
    //ExEnd:MirrorIndents

    doc = new aw.Document(base.artifactsDir + "ParagraphFormat.mirrorIndents.docx");
    format = doc.firstSection.body.paragraphs.at(0).paragraphFormat;

    expect(format.mirrorIndents).toEqual(true);
  });
});