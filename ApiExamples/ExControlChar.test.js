// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');



describe("ExControlChar", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CarriageReturn', () => {
    //ExStart
    //ExFor:ControlChar
    //ExFor:ControlChar.cr
    //ExFor:Node.getText
    //ExSummary:Shows how to use control characters.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert paragraphs with text with DocumentBuilder.
    builder.writeln("Hello world!");
    builder.writeln("Hello again!");

    // Converting the document to text form reveals that control characters
    // represent some of the document's structural elements, such as page breaks.
    expect(doc.getText()).toEqual(`Hello world!${aw.ControlChar.cr}` +
                            `Hello again!${aw.ControlChar.cr}` +
                            aw.ControlChar.pageBreak);

    // When converting a document to string form,
    // we can omit some of the control characters with the Trim method.
    expect(doc.getText().trim()).toEqual(`Hello world!${aw.ControlChar.cr}` +
                            "Hello again!");
    //ExEnd
  });


  test('InsertControlChars', () => {
    //ExStart
    //ExFor:ControlChar.cell
    //ExFor:ControlChar.columnBreak
    //ExFor:ControlChar.crLf
    //ExFor:ControlChar.lf
    //ExFor:ControlChar.lineBreak
    //ExFor:ControlChar.lineFeed
    //ExFor:ControlChar.nonBreakingSpace
    //ExFor:ControlChar.pageBreak
    //ExFor:ControlChar.paragraphBreak
    //ExFor:ControlChar.sectionBreak
    //ExFor:ControlChar.cellChar
    //ExFor:ControlChar.columnBreakChar
    //ExFor:ControlChar.defaultTextInputChar
    //ExFor:ControlChar.fieldEndChar
    //ExFor:ControlChar.fieldStartChar
    //ExFor:ControlChar.fieldSeparatorChar
    //ExFor:ControlChar.lineBreakChar
    //ExFor:ControlChar.lineFeedChar
    //ExFor:ControlChar.nonBreakingHyphenChar
    //ExFor:ControlChar.nonBreakingSpaceChar
    //ExFor:ControlChar.optionalHyphenChar
    //ExFor:ControlChar.pageBreakChar
    //ExFor:ControlChar.paragraphBreakChar
    //ExFor:ControlChar.sectionBreakChar
    //ExFor:ControlChar.spaceChar
    //ExSummary:Shows how to add various control characters to a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add a regular space.
    builder.write("Before space." + aw.ControlChar.spaceChar + "After space.");

    // Add an NBSP, which is a non-breaking space.
    // Unlike the regular space, this space cannot have an automatic line break at its position.
    builder.write("Before space." + aw.ControlChar.nonBreakingSpace + "After space.");

    // Add a tab character.
    builder.write("Before tab." + aw.ControlChar.tab + "After tab.");

    // Add a line break.
    builder.write("Before line break." + aw.ControlChar.lineBreak + "After line break.");

    // Add a new line and starts a new paragraph.
    expect(doc.firstSection.body.getChildNodes(aw.NodeType.Paragraph, true).count).toEqual(1);
    builder.write("Before line feed." + aw.ControlChar.lineFeed + "After line feed.");
    expect(doc.firstSection.body.getChildNodes(aw.NodeType.Paragraph, true).count).toEqual(2);

    // The line feed character has two versions.
    expect(aw.ControlChar.lf).toEqual(aw.ControlChar.lineFeed);

    // Carriage returns and line feeds can be represented together by one character.
    expect(aw.ControlChar.cr + aw.ControlChar.lf).toEqual(aw.ControlChar.crLf);

    // Add a paragraph break, which will start a new paragraph.
    builder.write("Before paragraph break." + aw.ControlChar.paragraphBreak + "After paragraph break.");
    expect(doc.firstSection.body.getChildNodes(aw.NodeType.Paragraph, true).count).toEqual(3);

    // Add a section break. This does not make a new section or paragraph.
    expect(doc.sections.count).toEqual(1);
    builder.write("Before section break." + aw.ControlChar.sectionBreak + "After section break.");
    expect(doc.sections.count).toEqual(1);

    // Add a page break.
    builder.write("Before page break." + aw.ControlChar.pageBreak + "After page break.");

    // A page break is the same value as a section break.
    expect(aw.ControlChar.sectionBreak).toEqual(aw.ControlChar.pageBreak);

    // Insert a new section, and then set its column count to two.
    doc.appendChild(new aw.Section(doc));
    builder.moveToSection(1);
    builder.currentSection.pageSetup.textColumns.setCount(2);

    // We can use a control character to mark the point where text moves to the next column.
    builder.write("Text at end of column 1." + aw.ControlChar.columnBreak + "Text at beginning of column 2.");

    doc.save(base.artifactsDir + "ControlChar.InsertControlChars.docx");

    // There are char and string counterparts for most characters.
    expect(aw.ControlChar.cellChar).toEqual(aw.ControlChar.cell);
    expect(aw.ControlChar.nonBreakingSpaceChar).toEqual(aw.ControlChar.nonBreakingSpace);
    expect(aw.ControlChar.tabChar).toEqual(aw.ControlChar.tab);
    expect(aw.ControlChar.lineBreakChar).toEqual(aw.ControlChar.lineBreak);
    expect(aw.ControlChar.lineFeedChar).toEqual(aw.ControlChar.lineFeed);
    expect(aw.ControlChar.paragraphBreakChar).toEqual(aw.ControlChar.paragraphBreak);
    expect(aw.ControlChar.sectionBreakChar).toEqual(aw.ControlChar.sectionBreak);
    expect(aw.ControlChar.sectionBreakChar).toEqual(aw.ControlChar.pageBreak);
    expect(aw.ControlChar.columnBreakChar).toEqual(aw.ControlChar.columnBreak);
    //ExEnd
  });
});
