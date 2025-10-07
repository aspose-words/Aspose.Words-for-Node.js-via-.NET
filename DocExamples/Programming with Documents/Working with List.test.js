// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;

describe("WorkingWithList", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('RestartListAtEachSection', () => {
    //ExStart:RestartListAtEachSection
    //GistId:d8326242115a099a83c0072f78763ca2
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    doc.lists.add(aw.Lists.ListTemplate.NumberDefault);
    let list = doc.lists.at(0);
    list.isRestartAtEachSection = true;
    // The "IsRestartAtEachSection" property will only be applicable when
    // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
    let options = new aw.Saving.OoxmlSaveOptions();
    options.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional;
    builder.listFormat.list = list;
    builder.writeln("List item 1");
    builder.writeln("List item 2");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.writeln("List item 3");
    builder.writeln("List item 4");
    doc.save(base.artifactsDir + "WorkingWithList.RestartingDocumentList.docx", options);
    //ExEnd:RestartListAtEachSection
  });

  test('SpecifyListLevel', () => {
    //ExStart:SpecifyListLevel
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    // Create a numbered list based on one of the Microsoft Word list templates
    // and apply it to the document builder's current paragraph.
    builder.listFormat.list = doc.lists.add(aw.Lists.ListTemplate.NumberArabicDot);
    // There are nine levels in this list, let's try them all.
    for (let i = 0; i < 9; i++) {
      builder.listFormat.listLevelNumber = i;
      builder.writeln("Level " + i);
    }
    // Create a bulleted list based on one of the Microsoft Word list templates
    // and apply it to the document builder's current paragraph.
    builder.listFormat.list = doc.lists.add(aw.Lists.ListTemplate.BulletDiamonds);
    for (let i = 0; i < 9; i++) {
      builder.listFormat.listLevelNumber = i;
      builder.writeln("Level " + i);
    }
    // This is a way to stop list formatting.
    builder.listFormat.list = null;
    builder.document.save(base.artifactsDir + "WorkingWithList.SpecifyListLevel.docx");
    //ExEnd:SpecifyListLevel
  });

  test('RestartListNumber', () => {
    //ExStart:RestartListNumber
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    // Create a list based on a template.
    let list1 = doc.lists.add(aw.Lists.ListTemplate.NumberArabicParenthesis);
    list1.listLevels.at(0).font.color = "#FF0000";
    list1.listLevels.at(0).alignment = aw.Lists.ListLevelAlignment.Right;
    builder.writeln("List 1 starts below:");
    builder.listFormat.list = list1;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();
    // To reuse the first list, we need to restart numbering by creating a copy of the original list formatting.
    let list2 = doc.lists.addCopy(list1);
    // We can modify the new list in any way, including setting a new start number.
    list2.listLevels.at(0).startAt = 10;
    builder.writeln("List 2 starts below:");
    builder.listFormat.list = list2;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();
    builder.document.save(base.artifactsDir + "WorkingWithList.RestartListNumber.docx");
    //ExEnd:RestartListNumber
  });

});