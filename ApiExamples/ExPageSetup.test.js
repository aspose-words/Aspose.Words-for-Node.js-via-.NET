// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const fs = require('fs');


describe("ExPageSetup", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });
  
  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('ClearFormatting', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.pageSetup
    //ExFor:aw.DocumentBuilder.insertBreak
    //ExFor:aw.DocumentBuilder.document
    //ExFor:PageSetup
    //ExFor:aw.PageSetup.orientation
    //ExFor:aw.PageSetup.verticalAlignment
    //ExFor:aw.PageSetup.clearFormatting
    //ExFor:Orientation
    //ExFor:PageVerticalAlignment
    //ExFor:BreakType
    //ExSummary:Shows how to apply and revert page setup settings to sections in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Modify the page setup properties for the builder's current section and add text.
    builder.pageSetup.orientation = aw.Orientation.Landscape;
    builder.pageSetup.verticalAlignment = aw.PageVerticalAlignment.Center;
    builder.writeln("This is the first section, which landscape oriented with vertically centered text.");

    // If we start a new section using a document builder,
    // it will inherit the builder's current page setup properties.
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    expect(doc.sections.at(1).pageSetup.orientation).toEqual(aw.Orientation.Landscape);
    expect(doc.sections.at(1).pageSetup.verticalAlignment).toEqual(aw.PageVerticalAlignment.Center);

    // We can revert its page setup properties to their default values using the "ClearFormatting" method.
    builder.pageSetup.clearFormatting();

    expect(doc.sections.at(1).pageSetup.orientation).toEqual(aw.Orientation.Portrait);
    expect(doc.sections.at(1).pageSetup.verticalAlignment).toEqual(aw.PageVerticalAlignment.Top);

    builder.writeln("This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");

    doc.save(base.artifactsDir + "PageSetup.clearFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.clearFormatting.docx");

    expect(doc.sections.at(0).pageSetup.orientation).toEqual(aw.Orientation.Landscape);
    expect(doc.sections.at(0).pageSetup.verticalAlignment).toEqual(aw.PageVerticalAlignment.Center);

    expect(doc.sections.at(1).pageSetup.orientation).toEqual(aw.Orientation.Portrait);
    expect(doc.sections.at(1).pageSetup.verticalAlignment).toEqual(aw.PageVerticalAlignment.Top);
  });


  test.each([false,
    true])('DifferentFirstPageHeaderFooter', (differentFirstPageHeaderFooter) => {
    //ExStart
    //ExFor:aw.PageSetup.differentFirstPageHeaderFooter
    //ExSummary:Shows how to enable or disable primary headers/footers.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two types of header/footers.
    // 1 -  The "First" header/footer, which appears on the first page of the section.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderFirst);
    builder.writeln("First page header.");

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterFirst);
    builder.writeln("First page footer.");

    // 2 -  The "Primary" header/footer, which appears on every page in the section.
    // We can override the primary header/footer by a first and an even page header/footer. 
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.writeln("Primary header.");

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.writeln("Primary footer.");

    builder.moveToSection(0);
    builder.writeln("Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3.");

    // Each section has a "PageSetup" object that specifies page appearance-related properties
    // such as orientation, size, and borders.
    // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
    // Set the "DifferentFirstPageHeaderFooter" property to "false"
    // to make the first page display the primary header/footer.
    builder.pageSetup.differentFirstPageHeaderFooter = differentFirstPageHeaderFooter;

    doc.save(base.artifactsDir + "PageSetup.differentFirstPageHeaderFooter.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.differentFirstPageHeaderFooter.docx");

    expect(doc.firstSection.pageSetup.differentFirstPageHeaderFooter).toEqual(differentFirstPageHeaderFooter);
  });


  test.each([false,
    true])('OddAndEvenPagesHeaderFooter', (oddAndEvenPagesHeaderFooter) => {
    //ExStart
    //ExFor:aw.PageSetup.oddAndEvenPagesHeaderFooter
    //ExSummary:Shows how to enable or disable even page headers/footers.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two types of header/footers.
    // 1 -  The "Primary" header/footer, which appears on every page in the section.
    // We can override the primary header/footer by a first and an even page header/footer. 
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.writeln("Primary header.");

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.writeln("Primary footer.");

    // 2 -  The "Even" header/footer, which appears on every even page of this section.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderEven);
    builder.writeln("Even page header.");

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterEven);
    builder.writeln("Even page footer.");

    builder.moveToSection(0);
    builder.writeln("Page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Page 3.");

    // Each section has a "PageSetup" object that specifies page appearance-related properties
    // such as orientation, size, and borders.
    // Set the "OddAndEvenPagesHeaderFooter" property to "true"
    // to display the even page header/footer on even pages.
    // Set the "OddAndEvenPagesHeaderFooter" property to "false"
    // to display the primary header/footer on even pages.
    builder.pageSetup.oddAndEvenPagesHeaderFooter = oddAndEvenPagesHeaderFooter;

    doc.save(base.artifactsDir + "PageSetup.oddAndEvenPagesHeaderFooter.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.oddAndEvenPagesHeaderFooter.docx");

    expect(doc.firstSection.pageSetup.oddAndEvenPagesHeaderFooter).toEqual(oddAndEvenPagesHeaderFooter);
  });


  test('CharactersPerLine', () => {
    //ExStart
    //ExFor:aw.PageSetup.charactersPerLine
    //ExFor:aw.PageSetup.layoutMode
    //ExFor:SectionLayoutMode
    //ExSummary:Shows how to specify a for the number of characters that each line may have.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Enable pitching, and then use it to set the number of characters per line in this section.
    builder.pageSetup.layoutMode = aw.SectionLayoutMode.Grid;
    builder.pageSetup.charactersPerLine = 10;

    // The number of characters also depends on the size of the font.
    doc.styles.at("Normal").font.size = 20;

    expect(doc.firstSection.pageSetup.charactersPerLine).toEqual(8);

    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    doc.save(base.artifactsDir + "PageSetup.charactersPerLine.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.charactersPerLine.docx");

    expect(doc.firstSection.pageSetup.layoutMode).toEqual(aw.SectionLayoutMode.Grid);
    expect(doc.firstSection.pageSetup.charactersPerLine).toEqual(8);
  });


  test('LinesPerPage', () => {
    //ExStart
    //ExFor:aw.PageSetup.linesPerPage
    //ExFor:aw.PageSetup.layoutMode
    //ExFor:aw.ParagraphFormat.snapToGrid
    //ExFor:SectionLayoutMode
    //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Enable pitching, and then use it to set the number of lines per page in this section.
    // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
    builder.pageSetup.layoutMode = aw.SectionLayoutMode.LineGrid;
    builder.pageSetup.linesPerPage = 15;

    builder.paragraphFormat.snapToGrid = true;

    for (let i = 0; i < 30; i++)
      builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

    doc.save(base.artifactsDir + "PageSetup.linesPerPage.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.linesPerPage.docx");

    expect(doc.firstSection.pageSetup.layoutMode).toEqual(aw.SectionLayoutMode.LineGrid);
    expect(doc.firstSection.pageSetup.linesPerPage).toEqual(14);

    for (let paragraph of doc.firstSection.body.paragraphs.toArray())
      expect(paragraph.paragraphFormat.snapToGrid).toEqual(true);
  });


  test('SetSectionStart', () => {
    //ExStart
    //ExFor:SectionStart
    //ExFor:aw.PageSetup.sectionStart
    //ExFor:aw.Document.sections
    //ExSummary:Shows how to specify how a new section separates itself from the previous.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("This text is in section 1.");

    // Section break types determine how a new section separates itself from the previous section.
    // Below are five types of section breaks.
    // 1 -  Starts the next section on a new page:
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.writeln("This text is in section 2.");

    expect(doc.sections.at(1).pageSetup.sectionStart).toEqual(aw.SectionStart.NewPage);

    // 2 -  Starts the next section on the current page:
    builder.insertBreak(aw.BreakType.SectionBreakContinuous);
    builder.writeln("This text is in section 3.");

    expect(doc.sections.at(2).pageSetup.sectionStart).toEqual(aw.SectionStart.Continuous);

    // 3 -  Starts the next section on a new even page:
    builder.insertBreak(aw.BreakType.SectionBreakEvenPage);
    builder.writeln("This text is in section 4.");

    expect(doc.sections.at(3).pageSetup.sectionStart).toEqual(aw.SectionStart.EvenPage);

    // 4 -  Starts the next section on a new odd page:
    builder.insertBreak(aw.BreakType.SectionBreakOddPage);
    builder.writeln("This text is in section 5.");

    expect(doc.sections.at(4).pageSetup.sectionStart).toEqual(aw.SectionStart.OddPage);

    // 5 -  Starts the next section on a new column:
    let columns = builder.pageSetup.textColumns;
    columns.setCount(2);

    builder.insertBreak(aw.BreakType.SectionBreakNewColumn);
    builder.writeln("This text is in section 6.");

    expect(doc.sections.at(5).pageSetup.sectionStart).toEqual(aw.SectionStart.NewColumn);

    doc.save(base.artifactsDir + "PageSetup.SetSectionStart.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.SetSectionStart.docx");

    expect(doc.sections.at(0).pageSetup.sectionStart).toEqual(aw.SectionStart.NewPage);
    expect(doc.sections.at(1).pageSetup.sectionStart).toEqual(aw.SectionStart.NewPage);
    expect(doc.sections.at(2).pageSetup.sectionStart).toEqual(aw.SectionStart.Continuous);
    expect(doc.sections.at(3).pageSetup.sectionStart).toEqual(aw.SectionStart.EvenPage);
    expect(doc.sections.at(4).pageSetup.sectionStart).toEqual(aw.SectionStart.OddPage);
    expect(doc.sections.at(5).pageSetup.sectionStart).toEqual(aw.SectionStart.NewColumn);
  });


/*  [Ignore("Run only when the printer driver is installed")]
  test('DefaultPaperTray', () => {
    //ExStart
    //ExFor:aw.PageSetup.firstPageTray
    //ExFor:aw.PageSetup.otherPagesTray
    //ExSummary:Shows how to get all the sections in a document to use the default paper tray of the selected printer.
    let doc = new aw.Document();

    // Find the default printer that we will use for printing this document.
    // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
    let settings = new PrinterSettings();

    // The paper tray value stored in documents is printer specific.
    // This means the code below resets all page tray values to use the current printers default tray.
    // You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer.
    foreach (Section section in doc.sections.OfType<Section>())
    {
      section.pageSetup.firstPageTray = settings.DefaultPageSettings.PaperSource.RawKind;
      section.pageSetup.otherPagesTray = settings.DefaultPageSettings.PaperSource.RawKind;
    }
    //ExEnd

    foreach (Section section in DocumentHelper.saveOpen(doc).Sections.OfType<Section>())
    {
      expect(section.pageSetup.firstPageTray).toEqual(settings.DefaultPageSettings.PaperSource.RawKind);
      expect(section.pageSetup.otherPagesTray).toEqual(settings.DefaultPageSettings.PaperSource.RawKind);
    }
  });*/


/*[Ignore("Run only when the printer driver is installed")]
  test('PaperTrayForDifferentPaperType', () => {
    //ExStart
    //ExFor:aw.PageSetup.firstPageTray
    //ExFor:aw.PageSetup.otherPagesTray
    //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
    let doc = new aw.Document();

    // Find the default printer that we will use for printing this document.
    // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
    let settings = new PrinterSettings();

    // This is the tray we will use for pages in the "A4" paper size.
    int printerTrayForA4 = settings.PaperSources.at(0).RawKind;

    // This is the tray we will use for pages in the "Letter" paper size.
    int printerTrayForLetter = settings.PaperSources.at(1).RawKind;

    // Modify the PageSettings object of this section to get Microsoft Word to instruct the printer
    // to use one of the trays we identified above, depending on this section's paper size.
    foreach (Section section in doc.sections.OfType<Section>())
    {
      if (section.pageSetup.paperSize == Aspose.words.paperSize.letter)
      {
        section.pageSetup.firstPageTray = printerTrayForLetter;
        section.pageSetup.otherPagesTray = printerTrayForLetter;
      }
      else if (section.pageSetup.paperSize == Aspose.words.paperSize.a4)
      {
        section.pageSetup.firstPageTray = printerTrayForA4;
        section.pageSetup.otherPagesTray = printerTrayForA4;
      }
    }
    //ExEnd

    foreach (Section section in DocumentHelper.saveOpen(doc).Sections.OfType<Section>())
    {
      if (section.pageSetup.paperSize == Aspose.words.paperSize.letter)
      {
        expect(section.pageSetup.firstPageTray).toEqual(printerTrayForLetter);
        expect(section.pageSetup.otherPagesTray).toEqual(printerTrayForLetter);
      }
      else if (section.pageSetup.paperSize == Aspose.words.paperSize.a4)
      {
        expect(section.pageSetup.firstPageTray).toEqual(printerTrayForA4);
        expect(section.pageSetup.otherPagesTray).toEqual(printerTrayForA4);
      }
    }
  });*/


  test('PageMargins', () => {
    //ExStart
    //ExFor:ConvertUtil
    //ExFor:aw.ConvertUtil.inchToPoint
    //ExFor:PaperSize
    //ExFor:aw.PageSetup.paperSize
    //ExFor:aw.PageSetup.orientation
    //ExFor:aw.PageSetup.topMargin
    //ExFor:aw.PageSetup.bottomMargin
    //ExFor:aw.PageSetup.leftMargin
    //ExFor:aw.PageSetup.rightMargin
    //ExFor:aw.PageSetup.headerDistance
    //ExFor:aw.PageSetup.footerDistance
    //ExSummary:Shows how to adjust paper size, orientation, margins, along with other settings for a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.pageSetup.paperSize = aw.PaperSize.Legal;
    builder.pageSetup.orientation = aw.Orientation.Landscape;
    builder.pageSetup.topMargin = aw.ConvertUtil.inchToPoint(1.0);
    builder.pageSetup.bottomMargin = aw.ConvertUtil.inchToPoint(1.0);
    builder.pageSetup.leftMargin = aw.ConvertUtil.inchToPoint(1.5);
    builder.pageSetup.rightMargin = aw.ConvertUtil.inchToPoint(1.5);
    builder.pageSetup.headerDistance = aw.ConvertUtil.inchToPoint(0.2);
    builder.pageSetup.footerDistance = aw.ConvertUtil.inchToPoint(0.2);

    builder.writeln("Hello world!");

    doc.save(base.artifactsDir + "PageSetup.pageMargins.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.pageMargins.docx");

    expect(doc.firstSection.pageSetup.paperSize).toEqual(aw.PaperSize.Legal);
    expect(doc.firstSection.pageSetup.pageWidth).toEqual(1008.0);
    expect(doc.firstSection.pageSetup.pageHeight).toEqual(612.0);
    expect(doc.firstSection.pageSetup.orientation).toEqual(aw.Orientation.Landscape);
    expect(doc.firstSection.pageSetup.topMargin).toEqual(72.0);
    expect(doc.firstSection.pageSetup.bottomMargin).toEqual(72.0);
    expect(doc.firstSection.pageSetup.leftMargin).toEqual(108.0);
    expect(doc.firstSection.pageSetup.rightMargin).toEqual(108.0);
    expect(doc.firstSection.pageSetup.headerDistance).toEqual(14.4);
    expect(doc.firstSection.pageSetup.footerDistance).toEqual(14.4);
  });


  test('PaperSizes', () => {
    //ExStart
    //ExFor:PaperSize
    //ExFor:aw.PageSetup.paperSize
    //ExSummary:Shows how to set page sizes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // We can change the current page's size to a pre-defined size
    // by using the "PaperSize" property of this section's PageSetup object.
    builder.pageSetup.paperSize = aw.PaperSize.Tabloid;

    expect(builder.pageSetup.pageWidth).toEqual(792.0);
    expect(builder.pageSetup.pageHeight).toEqual(1224.0);

    builder.writeln(`This page is ${builder.pageSetup.pageWidth}x${builder.pageSetup.pageHeight}.`);

    // Each section has its own PageSetup object. When we use a document builder to make a new section,
    // that section's PageSetup object inherits all the previous section's PageSetup object's values.
    builder.insertBreak(aw.BreakType.SectionBreakEvenPage);

    expect(builder.pageSetup.paperSize).toEqual(aw.PaperSize.Tabloid);

    builder.pageSetup.paperSize = aw.PaperSize.A5;
    builder.writeln(`This page is ${builder.pageSetup.pageWidth}x${builder.pageSetup.pageHeight}.`);

    expect(builder.pageSetup.pageWidth).toEqual(419.55);
    expect(builder.pageSetup.pageHeight).toEqual(595.30);

    builder.insertBreak(aw.BreakType.SectionBreakEvenPage);

    // Set a custom size for this section's pages.
    builder.pageSetup.pageWidth = 620;
    builder.pageSetup.pageHeight = 480;

    expect(builder.pageSetup.paperSize).toEqual(aw.PaperSize.Custom);

    builder.writeln(`This page is ${builder.pageSetup.pageWidth}x${builder.pageSetup.pageHeight}.`);

    doc.save(base.artifactsDir + "PageSetup.PaperSizes.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.PaperSizes.docx");

    expect(doc.sections.at(0).pageSetup.paperSize).toEqual(aw.PaperSize.Tabloid);
    expect(doc.sections.at(0).pageSetup.pageWidth).toEqual(792.0);
    expect(doc.sections.at(0).pageSetup.pageHeight).toEqual(1224.0);
    expect(doc.sections.at(1).pageSetup.paperSize).toEqual(aw.PaperSize.A5);
    expect(doc.sections.at(1).pageSetup.pageWidth).toEqual(419.55);
    expect(doc.sections.at(1).pageSetup.pageHeight).toEqual(595.30);
    expect(doc.sections.at(2).pageSetup.paperSize).toEqual(aw.PaperSize.Custom);
    expect(doc.sections.at(2).pageSetup.pageWidth).toEqual(620.0);
    expect(doc.sections.at(2).pageSetup.pageHeight).toEqual(480.0);
  });


  test('ColumnsSameWidth', () => {
    //ExStart
    //ExFor:aw.PageSetup.textColumns
    //ExFor:TextColumnCollection
    //ExFor:aw.TextColumnCollection.spacing
    //ExFor:aw.TextColumnCollection.setCount
    //ExSummary:Shows how to create multiple evenly spaced columns in a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let columns = builder.pageSetup.textColumns;
    columns.spacing = 100;
    columns.setCount(2);

    builder.writeln("Column 1.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.writeln("Column 2.");

    doc.save(base.artifactsDir + "PageSetup.ColumnsSameWidth.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.ColumnsSameWidth.docx");

    expect(doc.firstSection.pageSetup.textColumns.spacing).toEqual(100.0);
    expect(doc.firstSection.pageSetup.textColumns.count).toEqual(2);
  });


  test('CustomColumnWidth', () => {
    //ExStart
    //ExFor:aw.TextColumnCollection.evenlySpaced
    //ExFor:aw.TextColumnCollection.item
    //ExFor:TextColumn
    //ExFor:aw.TextColumn.width
    //ExFor:aw.TextColumn.spaceAfter
    //ExSummary:Shows how to create unevenly spaced columns.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let pageSetup = builder.pageSetup;

    let columns = pageSetup.textColumns;
    columns.evenlySpaced = false;
    columns.setCount(2);

    // Determine the amount of room that we have available for arranging columns.
    var contentWidth = pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin;

    expect(contentWidth).toBeCloseTo(468, 1);

    // Set the first column to be narrow.
    let column = columns.at(0);
    column.width = 100;
    column.spaceAfter = 20;

    // Set the second column to take the rest of the space available within the margins of the page.
    column = columns.at(1);
    column.width = contentWidth - column.width - column.spaceAfter;

    builder.writeln("Narrow column 1.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.writeln("Wide column 2.");

    doc.save(base.artifactsDir + "PageSetup.CustomColumnWidth.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.CustomColumnWidth.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.textColumns.evenlySpaced).toEqual(false);
    expect(pageSetup.textColumns.count).toEqual(2);
    expect(pageSetup.textColumns.at(0).width).toEqual(100.0);
    expect(pageSetup.textColumns.at(0).spaceAfter).toEqual(20.0);
    expect(pageSetup.textColumns.at(1).width).toEqual(468);
    expect(pageSetup.textColumns.at(1).spaceAfter).toEqual(0.0);
  });


  test.each([false,
    true])('VerticalLineBetweenColumns', (lineBetween) => {
    //ExStart
    //ExFor:aw.TextColumnCollection.lineBetween
    //ExSummary:Shows how to separate columns with a vertical line.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Configure the current section's PageSetup object to divide the text into several columns.
    // Set the "LineBetween" property to "true" to put a dividing line between columns.
    // Set the "LineBetween" property to "false" to leave the space between columns blank.
    let columns = builder.pageSetup.textColumns;
    columns.lineBetween = lineBetween;
    columns.setCount(3);

    builder.writeln("Column 1.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.writeln("Column 2.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.writeln("Column 3.");

    doc.save(base.artifactsDir + "PageSetup.VerticalLineBetweenColumns.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.VerticalLineBetweenColumns.docx");

    expect(doc.firstSection.pageSetup.textColumns.lineBetween).toEqual(lineBetween);
  });


  test('LineNumbers', () => {
    //ExStart
    //ExFor:aw.PageSetup.lineStartingNumber
    //ExFor:aw.PageSetup.lineNumberDistanceFromText
    //ExFor:aw.PageSetup.lineNumberCountBy
    //ExFor:aw.PageSetup.lineNumberRestartMode
    //ExFor:aw.ParagraphFormat.suppressLineNumbers
    //ExFor:LineNumberRestartMode
    //ExSummary:Shows how to enable line numbering for a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // We can use the section's PageSetup object to display numbers to the left of the section's text lines.
    // This is the same behavior as a List object,
    // but it covers the entire section and does not modify the text in any way.
    // Our section will restart the numbering on each new page from 1 and display the number,
    // if it is a multiple of 3, at 50pt to the left of the line.
    let pageSetup = builder.pageSetup;
    pageSetup.lineStartingNumber = 1;
    pageSetup.lineNumberCountBy = 3;
    pageSetup.lineNumberRestartMode = aw.LineNumberRestartMode.RestartPage;
    pageSetup.lineNumberDistanceFromText = 50.0;

    for (let i = 1; i <= 25; i++)
      builder.writeln(`Line ${i}.`);

    // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
    // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
    // The section's line counter will also ignore this line, treat the next line as the 15th,
    // and continue the count from that point onward.
    doc.firstSection.body.paragraphs.at(14).paragraphFormat.suppressLineNumbers = true;

    doc.save(base.artifactsDir + "PageSetup.LineNumbers.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.LineNumbers.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.lineStartingNumber).toEqual(1);
    expect(pageSetup.lineNumberCountBy).toEqual(3);
    expect(pageSetup.lineNumberRestartMode).toEqual(aw.LineNumberRestartMode.RestartPage);
    expect(pageSetup.lineNumberDistanceFromText).toEqual(50.0);
  });


  test('PageBorderProperties', () => {
    //ExStart
    //ExFor:aw.Section.pageSetup
    //ExFor:aw.PageSetup.borderAlwaysInFront
    //ExFor:aw.PageSetup.borderDistanceFrom
    //ExFor:aw.PageSetup.borderAppliesTo
    //ExFor:PageBorderDistanceFrom
    //ExFor:PageBorderAppliesTo
    //ExFor:aw.Border.distanceFromText
    //ExSummary:Shows how to create a wide blue band border at the top of the first page.
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

    doc.save(base.artifactsDir + "PageSetup.PageBorderProperties.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.PageBorderProperties.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.borderAlwaysInFront).toEqual(false);
    expect(pageSetup.borderDistanceFrom).toEqual(aw.PageBorderDistanceFrom.PageEdge);
    expect(pageSetup.borderAppliesTo).toEqual(aw.PageBorderAppliesTo.FirstPage);

    border = pageSetup.borders.at(aw.BorderType.Top);

    expect(border.lineStyle).toEqual(aw.LineStyle.Single);
    expect(border.lineWidth).toEqual(30.0);
    expect(border.color).toEqual("#0000FF");
    expect(border.distanceFromText).toEqual(0.0);
  });


  test('PageBorders', () => {
    //ExStart
    //ExFor:aw.PageSetup.borders
    //ExFor:aw.Border.shadow
    //ExFor:aw.BorderCollection.lineStyle
    //ExFor:aw.BorderCollection.lineWidth
    //ExFor:aw.BorderCollection.color
    //ExFor:aw.BorderCollection.distanceFromText
    //ExFor:aw.BorderCollection.shadow
    //ExSummary:Shows how to create green wavy page border with a shadow.
    let doc = new aw.Document();
    let pageSetup = doc.sections.at(0).pageSetup;

    pageSetup.borders.lineStyle = aw.LineStyle.DoubleWave;
    pageSetup.borders.lineWidth = 2;
    pageSetup.borders.color = "#008000";
    pageSetup.borders.distanceFromText = 24;
    pageSetup.borders.shadow = true;

    doc.save(base.artifactsDir + "PageSetup.PageBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.PageBorders.docx");
    pageSetup = doc.firstSection.pageSetup;

    for (let border of pageSetup.borders)
    {
      expect(border.lineStyle).toEqual(aw.LineStyle.DoubleWave);
      expect(border.lineWidth).toEqual(2.0);
      expect(border.color).toEqual("#008000");
      expect(border.distanceFromText).toEqual(24.0);
      expect(border.shadow).toEqual(true);
    }
  });


  test('PageNumbering', () => {
    //ExStart
    //ExFor:aw.PageSetup.restartPageNumbering
    //ExFor:aw.PageSetup.pageStartingNumber
    //ExFor:aw.PageSetup.pageNumberStyle
    //ExFor:aw.DocumentBuilder.insertField(String, String)
    //ExSummary:Shows how to set up page numbering in a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Section 1, page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Section 1, page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Section 1, page 3.");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.writeln("Section 2, page 1.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Section 2, page 2.");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.writeln("Section 2, page 3.");

    // Move the document builder to the first section's primary header,
    // which every page in that section will display.
    builder.moveToSection(0);
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);

    // Insert a PAGE field, which will display the number of the current page.
    builder.write("Page ");
    builder.insertField("PAGE", "");

    // Configure the section to have the page count that PAGE fields display start from 5.
    // Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.restartPageNumbering = true;
    pageSetup.pageStartingNumber = 5;
    pageSetup.pageNumberStyle = aw.NumberStyle.UppercaseRoman;

    // Create another primary header for the second section, with another PAGE field.
    builder.moveToSection(1);
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.write(" - ");
    builder.insertField("PAGE", "");
    builder.write(" - ");

    // Configure the section to have the page count that PAGE fields display start from 10.
    // Also, configure all PAGE fields to display their page numbers using Arabic numbers.
    pageSetup = doc.sections.at(1).pageSetup;
    pageSetup.pageStartingNumber = 10;
    pageSetup.restartPageNumbering = true;
    pageSetup.pageNumberStyle = aw.NumberStyle.Arabic;

    doc.save(base.artifactsDir + "PageSetup.PageNumbering.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.PageNumbering.docx");
    pageSetup = doc.sections.at(0).pageSetup;

    expect(pageSetup.restartPageNumbering).toEqual(true);
    expect(pageSetup.pageStartingNumber).toEqual(5);
    expect(pageSetup.pageNumberStyle).toEqual(aw.NumberStyle.UppercaseRoman);

    pageSetup = doc.sections.at(1).pageSetup;

    expect(pageSetup.restartPageNumbering).toEqual(true);
    expect(pageSetup.pageStartingNumber).toEqual(10);
    expect(pageSetup.pageNumberStyle).toEqual(aw.NumberStyle.Arabic);
  });


  test('FootnoteOptions', () => {
    //ExStart
    //ExFor:aw.PageSetup.endnoteOptions
    //ExFor:aw.PageSetup.footnoteOptions
    //ExSummary:Shows how to configure options affecting footnotes/endnotes in a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Hello world!");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Footnote reference text.");

    // Configure all footnotes in the first section to restart the numbering from 1
    // at each new page and display themselves directly beneath the text on every page.
    let footnoteOptions = doc.sections.at(0).pageSetup.footnoteOptions;
    footnoteOptions.position = aw.Notes.FootnotePosition.BeneathText;
    footnoteOptions.restartRule = aw.Notes.FootnoteNumberingRule.RestartPage;
    footnoteOptions.startNumber = 1;

    builder.write(" Hello again.");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Endnote reference text.");

    // Configure all endnotes in the first section to maintain a continuous count throughout the section,
    // starting from 1. Also, set them all to appear collected at the end of the document.
    let endnoteOptions = doc.sections.at(0).pageSetup.endnoteOptions;
    endnoteOptions.position = aw.Notes.EndnotePosition.EndOfDocument;
    endnoteOptions.restartRule = aw.Notes.FootnoteNumberingRule.Continuous;
    endnoteOptions.startNumber = 1;

    doc.save(base.artifactsDir + "PageSetup.footnoteOptions.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.footnoteOptions.docx");
    footnoteOptions = doc.firstSection.pageSetup.footnoteOptions;

    expect(footnoteOptions.position).toEqual(aw.Notes.FootnotePosition.BeneathText);
    expect(footnoteOptions.restartRule).toEqual(aw.Notes.FootnoteNumberingRule.RestartPage);
    expect(footnoteOptions.startNumber).toEqual(1);

    endnoteOptions = doc.firstSection.pageSetup.endnoteOptions;

    expect(endnoteOptions.position).toEqual(aw.Notes.EndnotePosition.EndOfDocument);
    expect(endnoteOptions.restartRule).toEqual(aw.Notes.FootnoteNumberingRule.Continuous);
    expect(endnoteOptions.startNumber).toEqual(1);
  });


  test.each([false,
    true])('Bidi', (reverseColumns) => {
    //ExStart
    //ExFor:aw.PageSetup.bidi
    //ExSummary:Shows how to set the order of text columns in a section.
    let doc = new aw.Document();

    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.textColumns.setCount(3);

    let builder = new aw.DocumentBuilder(doc);
    builder.write("Column 1.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.write("Column 2.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.write("Column 3.");

    // Set the "Bidi" property to "true" to arrange the columns starting from the page's right side.
    // The order of the columns will match the direction of the right-to-left text.
    // Set the "Bidi" property to "false" to arrange the columns starting from the page's left side.
    // The order of the columns will match the direction of the left-to-right text.
    pageSetup.bidi = reverseColumns;

    doc.save(base.artifactsDir + "PageSetup.bidi.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.bidi.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.textColumns.count).toEqual(3);
    expect(pageSetup.bidi).toEqual(reverseColumns);
  });


  test('PageBorder', () => {
    //ExStart
    //ExFor:aw.PageSetup.borderSurroundsFooter
    //ExFor:aw.PageSetup.borderSurroundsHeader
    //ExSummary:Shows how to apply a border to the page and header/footer.
    let doc = new aw.Document();

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world! This is the main body text.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("This is the header.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.write("This is the footer.");
    builder.moveToDocumentEnd();

    // Insert a blue double-line border.
    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.borders.lineStyle = aw.LineStyle.Double;
    pageSetup.borders.color = "#0000FF";

    // A section's PageSetup object has "BorderSurroundsHeader" and "BorderSurroundsFooter" flags that determine
    // whether a page border surrounds the main body text, also includes the header or footer, respectively.
    // Set the "BorderSurroundsHeader" flag to "true" to surround the header with our border,
    // and then set the "BorderSurroundsFooter" flag to leave the footer outside of the border.
    pageSetup.borderSurroundsHeader = true;
    pageSetup.borderSurroundsFooter = false;

    doc.save(base.artifactsDir + "PageSetup.PageBorder.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.PageBorder.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.borderSurroundsHeader).toEqual(true);
    expect(pageSetup.borderSurroundsFooter).toEqual(false);
  });


  test('Gutter', () => {
    //ExStart
    //ExFor:aw.PageSetup.gutter
    //ExFor:aw.PageSetup.rtlGutter
    //ExFor:aw.PageSetup.multiplePages
    //ExSummary:Shows how to set gutter margins.
    let doc = new aw.Document();

    // Insert text that spans several pages.
    let builder = new aw.DocumentBuilder(doc);
    for (let i = 0; i < 6; i++)
    {
      builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
            "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
      builder.insertBreak(aw.BreakType.PageBreak);
    }

    // A gutter adds whitespaces to either the left or right page margin,
    // which makes up for the center folding of pages in a book encroaching on the page's layout.
    let pageSetup = doc.sections.at(0).pageSetup;

    // Determine how much space our pages have for text within the margins and then add an amount to pad a margin. 
    expect(pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin).toEqual(468);

    pageSetup.gutter = 100.0;

    // Set the "RtlGutter" property to "true" to place the gutter in a more suitable position for right-to-left text.
    pageSetup.rtlGutter = true;

    // Set the "MultiplePages" property to "MultiplePagesType.MirrorMargins" to alternate
    // the left/right page side position of margins every page.
    pageSetup.multiplePages = aw.Settings.MultiplePagesType.MirrorMargins;

    doc.save(base.artifactsDir + "PageSetup.gutter.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.gutter.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.gutter).toEqual(100.0);
    expect(pageSetup.rtlGutter).toEqual(true);
    expect(pageSetup.multiplePages).toEqual(aw.Settings.MultiplePagesType.MirrorMargins);
  });


  test('Booklet', () => {
    //ExStart
    //ExFor:aw.PageSetup.gutter
    //ExFor:aw.PageSetup.multiplePages
    //ExFor:aw.PageSetup.sheetsPerBooklet
    //ExSummary:Shows how to configure a document that can be printed as a book fold.
    let doc = new aw.Document();

    // Insert text that spans 16 pages.
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("My Booklet:");

    for (let i = 0; i < 15; i++)
    {
      builder.insertBreak(aw.BreakType.PageBreak);
      builder.write(`Booklet face #${i}`);
    }

    // Configure the first section's "PageSetup" property to print the document in the form of a book fold.
    // When we print this document on both sides, we can take the pages to stack them
    // and fold them all down the middle at once. The contents of the document will line up into a book fold.
    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.multiplePages = aw.Settings.MultiplePagesType.BookFoldPrinting;

    // We can only specify the number of sheets in multiples of 4.
    pageSetup.sheetsPerBooklet = 4;

    doc.save(base.artifactsDir + "PageSetup.Booklet.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.Booklet.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.multiplePages).toEqual(aw.Settings.MultiplePagesType.BookFoldPrinting);
    expect(pageSetup.sheetsPerBooklet).toEqual(4);
  });


  test('SetTextOrientation', () => {
    //ExStart
    //ExFor:aw.PageSetup.textOrientation
    //ExSummary:Shows how to set text orientation.
    let doc = new aw.Document();

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Set the "TextOrientation" property to "TextOrientation.Upward" to rotate all the text 90 degrees
    // to the right so that all left-to-right text now goes top-to-bottom.
    let pageSetup = doc.sections.at(0).pageSetup;
    pageSetup.textOrientation = aw.TextOrientation.Upward;

    doc.save(base.artifactsDir + "PageSetup.SetTextOrientation.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "PageSetup.SetTextOrientation.docx");
    pageSetup = doc.firstSection.pageSetup;

    expect(pageSetup.textOrientation).toEqual(aw.TextOrientation.Upward);
  });


  //ExStart
  //ExFor:PageSetup.SuppressEndnotes
  //ExFor:Body.ParentSection
  //ExSummary:Shows how to store endnotes at the end of each section, and modify their positions.
  test('SuppressEndnotes', () => {
    let doc = new aw.Document();
    doc.removeAllChildren();

    // By default, a document compiles all endnotes at its end. 
    expect(doc.endnoteOptions.position).toEqual(aw.Notes.EndnotePosition.EndOfDocument);

    // We use the "Position" property of the document's "EndnoteOptions" object
    // to collect endnotes at the end of each section instead. 
    doc.endnoteOptions.position = aw.Notes.EndnotePosition.EndOfSection;

    insertSectionWithEndnote(doc, "Section 1", "Endnote 1, will stay in section 1");
    insertSectionWithEndnote(doc, "Section 2", "Endnote 2, will be pushed down to section 3");
    insertSectionWithEndnote(doc, "Section 3", "Endnote 3, will stay in section 3");

    // While getting sections to display their respective endnotes, we can set the "SuppressEndnotes" flag
    // of a section's "PageSetup" object to "true" to revert to the default behavior and pass its endnotes
    // onto the next section.
    let pageSetup = doc.sections.at(1).pageSetup;
    pageSetup.suppressEndnotes = true;

    doc.save(base.artifactsDir + "PageSetup.suppressEndnotes.docx");
    testSuppressEndnotes(new aw.Document(base.artifactsDir + "PageSetup.suppressEndnotes.docx")); //ExSkip
  });


  /// <summary>
  /// Append a section with text and an endnote to a document.
  /// </summary>
  function insertSectionWithEndnote(doc, sectionBodyText, endnoteText)
  {
    let section = new aw.Section(doc);

    doc.appendChild(section);

    let body = new aw.Body(doc);
    section.appendChild(body);

    expect(body.parentNode.referenceEquals(section)).toEqual(true);

    let para = new aw.Paragraph(doc);
    body.appendChild(para);

    expect(para.parentNode.referenceEquals(body)).toEqual(true);

    let builder = new aw.DocumentBuilder(doc);
    builder.moveTo(para);
    builder.write(sectionBodyText);
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, endnoteText);
  }
    //ExEnd

  function testSuppressEndnotes(doc)
  {
    let pageSetup = doc.sections.at(1).pageSetup;
    expect(pageSetup.suppressEndnotes).toEqual(true);
  }

  test('ChapterPageSeparator', () => {
    //ExStart
    //ExFor:aw.PageSetup.headingLevelForChapter
    //ExFor:ChapterPageSeparator
    //ExFor:aw.PageSetup.chapterPageSeparator
    //ExSummary:Shows how to work with page chapters.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    let pageSetup = doc.firstSection.pageSetup;

    pageSetup.pageNumberStyle = aw.NumberStyle.UppercaseRoman;
    pageSetup.chapterPageSeparator = aw.ChapterPageSeparator.Colon;
    pageSetup.headingLevelForChapter = 1;
    //ExEnd
  });
});
