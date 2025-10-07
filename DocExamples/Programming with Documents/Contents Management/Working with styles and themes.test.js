// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithStylesAndThemes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AccessStyles', () => {
    //ExStart:AccessStyles
    //GistId:c6b0305cd373fae738c432637dd67ba5
    let doc = new aw.Document();

    let styleName = "";
    // Get styles collection from the document.
    let styles = doc.styles;
    for (let style of styles) {
      if (styleName == "") {
        styleName = style.name;
        console.log(styleName);
      } else {
        styleName = styleName + ", " + style.name;
        console.log(styleName);
      }
    }
    //ExEnd:AccessStyles
  });

  test('CopyStyles', () => {
    //ExStart:CopyStyles
    //GistId:c6b0305cd373fae738c432637dd67ba5
    let doc = new aw.Document();
    let target = new aw.Document(base.myDir + "Rendering.docx");

    target.copyStylesFromTemplate(doc);

    doc.save(base.artifactsDir + "WorkingWithStylesAndThemes.CopyStyles.docx");
    //ExEnd:CopyStyles
  });

  test('GetThemeProperties', () => {
    //ExStart:GetThemeProperties
    //GistId:c6b0305cd373fae738c432637dd67ba5
    let doc = new aw.Document();

    let theme = doc.theme;

    console.log(theme.majorFonts.latin);
    console.log(theme.minorFonts.eastAsian);
    console.log(theme.colors.Accent1);
    //ExEnd:GetThemeProperties
  });

  test('SetThemeProperties', () => {
    //ExStart:SetThemeProperties
    //GistId:c6b0305cd373fae738c432637dd67ba5
    let doc = new aw.Document();

    let theme = doc.theme;
    theme.minorFonts.latin = "Times New Roman";
    theme.colors.hyperlink = "#FFD700"; // Gold.
    //ExEnd:SetThemeProperties
  });

  test('InsertStyleSeparator', () => {
    //ExStart:InsertStyleSeparator
    //GistId:fc7e411a082bdf9bd715a4cf28552213
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let paraStyle = builder.document.styles.add(aw.StyleType.Paragraph, "MyParaStyle");
    paraStyle.font.bold = false;
    paraStyle.font.size = 8;
    paraStyle.font.name = "Arial";

    // Append text with "Heading 1" style.
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.write("Heading 1");
    builder.insertStyleSeparator();

    // Append text with another style.
    builder.paragraphFormat.styleName = paraStyle.name;
    builder.write("This is text with some other formatting ");

    doc.save(base.artifactsDir + "WorkingWithStylesAndThemes.InsertStyleSeparator.docx")
    //ExEnd:InsertStyleSeparator
  });

  test('CopyStyleDifferentDocument', () => {
    //ExStart:CopyStyleDifferentDocument
    //GistId:a79ed2d7052cbfbbbc1215708bb4ac4b
    let srcDoc = new aw.Document();

    // Create a custom style for the source document.
    let srcStyle = srcDoc.styles.add(aw.StyleType.Paragraph, "MyStyle");
    srcStyle.font.color = "#FF0000"; // Red.

    // Import the source document's custom style into the destination document.
    let dstDoc = new aw.Document();
    let newStyle = dstDoc.styles.addCopy(srcStyle);

    // The imported style has an appearance identical to its source style.
    console.log("Style name:", newStyle.name); // Should be "MyStyle".
    console.log("Style color:", newStyle.font.color); // Should be red color.
    //ExEnd:CopyStyleDifferentDocument
  });
});