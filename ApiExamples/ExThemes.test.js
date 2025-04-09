// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;


describe("ExThemes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CustomColorsAndFonts', () => {
    //ExStart
    //ExFor:Document.theme
    //ExFor:Theme
    //ExFor:Theme.colors
    //ExFor:Theme.majorFonts
    //ExFor:Theme.minorFonts
    //ExFor:ThemeColors
    //ExFor:ThemeColors.accent1
    //ExFor:ThemeColors.accent2
    //ExFor:ThemeColors.accent3
    //ExFor:ThemeColors.accent4
    //ExFor:ThemeColors.accent5
    //ExFor:ThemeColors.accent6
    //ExFor:ThemeColors.dark1
    //ExFor:ThemeColors.dark2
    //ExFor:ThemeColors.followedHyperlink
    //ExFor:ThemeColors.hyperlink
    //ExFor:ThemeColors.light1
    //ExFor:ThemeColors.light2
    //ExFor:ThemeFonts
    //ExFor:ThemeFonts.complexScript
    //ExFor:ThemeFonts.eastAsian
    //ExFor:ThemeFonts.latin
    //ExSummary:Shows how to set custom colors and fonts for themes.
    let doc = new aw.Document(base.myDir + "Theme colors.docx");

    // The "Theme" object gives us access to the document theme, a source of default fonts and colors.
    let theme = doc.theme;

    // Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
    theme.majorFonts.latin = "Courier New";
    theme.minorFonts.latin = "Agency FB";

    // Other languages may also have their custom fonts in this theme.
    expect(theme.majorFonts.complexScript).toEqual('');
    expect(theme.majorFonts.eastAsian).toEqual('');
    expect(theme.minorFonts.complexScript).toEqual('');
    expect(theme.minorFonts.eastAsian).toEqual('');

    // The "Colors" property contains the color palette from Microsoft Word,
    // which appears when changing shading or font color.
    // Apply custom colors to the color palette so we have easy access to them in Microsoft Word
    // when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
    // or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
    let colors = theme.colors;
    colors.dark1 = "#191970";
    colors.light1 = "#98FB98";
    colors.dark2 = "#4B0082";
    colors.light2 = "#F0E68C";

    colors.accent1 = "#FF4500";
    colors.accent2 = "#FFA07A";
    colors.accent3 = "#FFFF00";
    colors.accent4 = "#FFD700";
    colors.accent5 = "#8A2BE2";
    colors.accent6 = "#9400D3";

    // Apply custom colors to hyperlinks in their clicked and un-clicked states.
    colors.hyperlink = "#000000";
    colors.followedHyperlink = "#808080";

    doc.save(base.artifactsDir + "Themes.CustomColorsAndFonts.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Themes.CustomColorsAndFonts.docx");

    expect(doc.theme.colors.accent1).toEqual("#FF4500");
    expect(doc.theme.colors.dark1).toEqual("#191970");
    expect(doc.theme.colors.followedHyperlink).toEqual("#808080");
    expect(doc.theme.colors.hyperlink).toEqual("#000000");
    expect(doc.theme.colors.light1).toEqual("#98FB98");

    expect(doc.theme.majorFonts.complexScript).toEqual('');
    expect(doc.theme.majorFonts.eastAsian).toEqual('');
    expect(doc.theme.majorFonts.latin).toEqual("Courier New");

    expect(doc.theme.minorFonts.complexScript).toEqual('');
    expect(doc.theme.minorFonts.eastAsian).toEqual('');
    expect(doc.theme.minorFonts.latin).toEqual("Agency FB");
  });

});
