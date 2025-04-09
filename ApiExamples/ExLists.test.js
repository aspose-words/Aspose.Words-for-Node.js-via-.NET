// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
const MemoryStream = require('memorystream');


describe("ExLists", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('ApplyDefaultBulletsAndNumbers', () => {
    //ExStart
    //ExFor:DocumentBuilder.listFormat
    //ExFor:ListFormat.applyNumberDefault
    //ExFor:ListFormat.applyBulletDefault
    //ExFor:ListFormat.listIndent
    //ExFor:ListFormat.listOutdent
    //ExFor:ListFormat.removeNumbers
    //ExFor:ListFormat.listLevelNumber
    //ExSummary:Shows how to create bulleted and numbered lists.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Aspose.words main advantages are:");

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Below are two types of lists that we can create with a document builder.
    // 1 -  A bulleted list:
    // This list will apply an indent and a bullet symbol ("•") before each paragraph.
    builder.listFormat.applyBulletDefault();
    builder.writeln("Great performance");
    builder.writeln("High reliability");
    builder.writeln("Quality code and working");
    builder.writeln("Wide variety of features");
    builder.writeln("Easy to understand API");

    // End the bulleted list.
    builder.listFormat.removeNumbers();

    builder.insertBreak(aw.BreakType.ParagraphBreak);
    builder.writeln("Aspose.words allows:");

    // 2 -  A numbered list:
    // Numbered lists create a logical order for their paragraphs by numbering each item.
    builder.listFormat.applyNumberDefault();

    // This paragraph is the first item. The first item of a numbered list will have a "1." as its list item symbol.
    builder.writeln("Opening documents from different formats:");

    expect(builder.listFormat.listLevelNumber).toEqual(0);

    // Call the "ListIndent" method to increase the current list level,
    // which will start a new self-contained list, with a deeper indent, at the current item of the first list level.
    builder.listFormat.listIndent();

    expect(builder.listFormat.listLevelNumber).toEqual(1);

    // These are the first three list items of the second list level, which will maintain a count
    // independent of the count of the first list level. According to the current list format,
    // they will have symbols of "a.", "b.", and "c.".
    builder.writeln("DOC");
    builder.writeln("PDF");
    builder.writeln("HTML");

    // Call the "ListOutdent" method to return to the previous list level.
    builder.listFormat.listOutdent();

    expect(builder.listFormat.listLevelNumber).toEqual(0);

    // These two paragraphs will continue the count of the first list level.
    // These items will have symbols of "2.", and "3."
    builder.writeln("Processing documents");
    builder.writeln("Saving documents in different formats:");

    // If we increase the list level to a level that we have added items to previously,
    // the nested list will be separate from the previous, and its numbering will start from the beginning. 
    // These list items will have symbols of "a.", "b.", "c.", "d.", and "e".
    builder.listFormat.listIndent();
    builder.writeln("DOC");
    builder.writeln("PDF");
    builder.writeln("HTML");
    builder.writeln("MHTML");
    builder.writeln("Plain text");

    // Outdent the list level again.
    builder.listFormat.listOutdent();
    builder.writeln("Doing many other things!");

    // End the numbered list.
    builder.listFormat.removeNumbers();

    doc.save(base.artifactsDir + "Lists.ApplyDefaultBulletsAndNumbers.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.ApplyDefaultBulletsAndNumbers.docx");

    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, doc.lists.at(1).listLevels.at(0));
    TestUtil.verifyListLevel("\u0001.", 54.0, aw.NumberStyle.LowercaseLetter, doc.lists.at(1).listLevels.at(1));
    TestUtil.verifyListLevel("\uf0b7", 18.0, aw.NumberStyle.Bullet, doc.lists.at(0).listLevels.at(0));
  });


  test('SpecifyListLevel', () => {
    //ExStart
    //ExFor:ListCollection
    //ExFor:List
    //ExFor:ListFormat
    //ExFor:ListFormat.isListItem
    //ExFor:ListFormat.listLevelNumber
    //ExFor:ListFormat.list
    //ExFor:ListTemplate
    //ExFor:DocumentBase.lists
    //ExFor:ListCollection.add(ListTemplate)
    //ExSummary:Shows how to work with list levels.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    expect(builder.listFormat.isListItem).toEqual(false);

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Below are two types of lists that we can create using a document builder.
    // 1 -  A numbered list:
    // Numbered lists create a logical order for their paragraphs by numbering each item.
    builder.listFormat.list = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);

    expect(builder.listFormat.isListItem).toEqual(true);

    // By setting the "ListLevelNumber" property, we can increase the list level
    // to begin a self-contained sub-list at the current list item.
    // The Microsoft Word list template called "NumberDefault" uses numbers to create list levels for the first list level.
    // Deeper list levels use letters and lowercase Roman numerals. 
    for (let i = 0; i < 9; i++)
    {
      builder.listFormat.listLevelNumber = i;
      builder.writeln("Level " + i);
    }

    // 2 -  A bulleted list:
    // This list will apply an indent and a bullet symbol ("•") before each paragraph.
    // Deeper levels of this list will use different symbols, such as "■" and "○".
    builder.listFormat.list = doc.lists.add(aw.Lists.ListTemplate.BulletDefault);

    for (let i = 0; i < 9; i++)
    {
      builder.listFormat.listLevelNumber = i;
      builder.writeln("Level " + i);
    }

    // We can disable list formatting to not format any subsequent paragraphs as lists by un-setting the "List" flag.
    builder.listFormat.list = null;

    expect(builder.listFormat.isListItem).toEqual(false);

    doc.save(base.artifactsDir + "Lists.SpecifyListLevel.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.SpecifyListLevel.docx");

    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, doc.lists.at(0).listLevels.at(0));
  });


  test('NestedLists', () => {
    //ExStart
    //ExFor:ListFormat.list
    //ExFor:ParagraphFormat.clearFormatting
    //ExFor:ParagraphFormat.dropCapPosition
    //ExFor:ParagraphFormat.isListItem
    //ExFor:Paragraph.isListItem
    //ExSummary:Shows how to nest a list inside another list.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create an outline list for the headings.
    let outlineList = doc.lists.add(aw.Lists.ListTemplate.OutlineNumbers);
    builder.listFormat.list = outlineList;
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.writeln("This is my Chapter 1");

    // Create a numbered list.
    let numberedList = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);
    builder.listFormat.list = numberedList;
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Normal;
    builder.writeln("Numbered list item 1.");

    // Every paragraph that comprises a list will have this flag.
    expect(builder.currentParagraph.isListItem).toEqual(true);
    expect(builder.paragraphFormat.isListItem).toEqual(true);

    // Create a bulleted list.
    let bulletedList = doc.lists.add(aw.Lists.ListTemplate.BulletDefault);
    builder.listFormat.list = bulletedList;
    builder.paragraphFormat.leftIndent = 72;
    builder.writeln("Bulleted list item 1.");
    builder.writeln("Bulleted list item 2.");
    builder.paragraphFormat.clearFormatting();

    // Revert to the numbered list.
    builder.listFormat.list = numberedList;
    builder.writeln("Numbered list item 2.");
    builder.writeln("Numbered list item 3.");

    // Revert to the outline list.
    builder.listFormat.list = outlineList;
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.writeln("This is my Chapter 2");

    builder.paragraphFormat.clearFormatting();

    builder.document.save(base.artifactsDir + "Lists.NestedLists.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.NestedLists.docx");

    TestUtil.verifyListLevel("\u0000)", 0.0, aw.NumberStyle.Arabic, doc.lists.at(0).listLevels.at(0));
    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, doc.lists.at(1).listLevels.at(0));
    TestUtil.verifyListLevel("\uf0b7", 18.0, aw.NumberStyle.Bullet, doc.lists.at(2).listLevels.at(0));
  });


  test('CreateCustomList', () => {
    //ExStart
    //ExFor:List
    //ExFor:List.listLevels
    //ExFor:ListFormat.listLevel
    //ExFor:ListLevelCollection
    //ExFor:ListLevelCollection.item
    //ExFor:ListLevel
    //ExFor:ListLevel.alignment
    //ExFor:ListLevel.font
    //ExFor:ListLevel.numberStyle
    //ExFor:ListLevel.startAt
    //ExFor:ListLevel.trailingCharacter
    //ExFor:ListLevelAlignment
    //ExFor:NumberStyle
    //ExFor:ListTrailingCharacter
    //ExFor:ListLevel.numberFormat
    //ExFor:ListLevel.numberPosition
    //ExFor:ListLevel.textPosition
    //ExFor:ListLevel.tabPosition
    //ExSummary:Shows how to apply custom list formatting to paragraphs when using DocumentBuilder.
    let doc = new aw.Document();

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create a list from a Microsoft Word template, and customize the first two of its list levels.
    let list = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);

    let listLevel = list.listLevels.at(0);
    listLevel.font.color = "#FF0000";
    listLevel.font.size = 24;
    listLevel.numberStyle = aw.NumberStyle.OrdinalText;
    listLevel.startAt = 21;
    listLevel.numberFormat = "\u0000";

    listLevel.numberPosition = -36;
    listLevel.textPosition = 144;
    listLevel.tabPosition = 144;

    listLevel = list.listLevels.at(1);
    listLevel.alignment = aw.Lists.ListLevelAlignment.Right;
    listLevel.numberStyle = aw.NumberStyle.Bullet;
    listLevel.font.name = "Wingdings";
    listLevel.font.color = "#0000FF";
    listLevel.font.size = 24;

    // This NumberFormat value will create star-shaped bullet list symbols.
    listLevel.numberFormat = "\uf0af";
    listLevel.trailingCharacter = aw.Lists.ListTrailingCharacter.Space;
    listLevel.numberPosition = 144;

    // Create paragraphs and apply both list levels of our custom list formatting to them.
    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.list = list;
    builder.writeln("The quick brown fox...");
    builder.writeln("The quick brown fox...");

    builder.listFormat.listIndent();
    builder.writeln("jumped over the lazy dog.");
    builder.writeln("jumped over the lazy dog.");

    builder.listFormat.listOutdent();
    builder.writeln("The quick brown fox...");

    builder.listFormat.removeNumbers();

    builder.document.save(base.artifactsDir + "Lists.CreateCustomList.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.CreateCustomList.docx");

    listLevel = doc.lists.at(0).listLevels.at(0);

    TestUtil.verifyListLevel("\u0000", -36.0, aw.NumberStyle.OrdinalText, listLevel);
    expect(listLevel.font.color).toEqual("#FF0000");
    expect(listLevel.font.size).toEqual(24.0);
    expect(listLevel.startAt).toEqual(21);

    listLevel = doc.lists.at(0).listLevels.at(1);

    TestUtil.verifyListLevel("\uf0af", 144.0, aw.NumberStyle.Bullet, listLevel);
    expect(listLevel.font.color).toEqual("#0000FF");
    expect(listLevel.font.size).toEqual(24.0);
    expect(listLevel.startAt).toEqual(1);
    expect(listLevel.trailingCharacter).toEqual(aw.Lists.ListTrailingCharacter.Space);
  });


  test('RestartNumberingUsingListCopy', () => {
    //ExStart
    //ExFor:List
    //ExFor:ListCollection
    //ExFor:ListCollection.add(ListTemplate)
    //ExFor:ListCollection.addCopy(List)
    //ExFor:ListLevel.startAt
    //ExFor:ListTemplate
    //ExSummary:Shows how to restart numbering in a list by copying a list.
    let doc = new aw.Document();

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // Create a list from a Microsoft Word template, and customize its first list level.
    let list1 = doc.lists.add(aw.Lists.ListTemplate.NumberArabicParenthesis);
    list1.listLevels.at(0).font.color = "#FF0000";
    list1.listLevels.at(0).alignment = aw.Lists.ListLevelAlignment.Right;

    // Apply our list to some paragraphs.
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("List 1 starts below:");
    builder.listFormat.list = list1;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();

    // We can add a copy of an existing list to the document's list collection
    // to create a similar list without making changes to the original.
    let list2 = doc.lists.addCopy(list1);
    list2.listLevels.at(0).font.color = "#0000FF";
    list2.listLevels.at(0).startAt = 10;

    // Apply the second list to new paragraphs.
    builder.writeln("List 2 starts below:");
    builder.listFormat.list = list2;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();

    doc.save(base.artifactsDir + "Lists.RestartNumberingUsingListCopy.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.RestartNumberingUsingListCopy.docx");

    list1 = doc.lists.at(0);
    TestUtil.verifyListLevel("\u0000)", 18.0, aw.NumberStyle.Arabic, list1.listLevels.at(0));
    expect(list1.listLevels.at(0).font.color).toEqual("#FF0000");
    expect(list1.listLevels.at(0).font.size).toEqual(10.0);
    expect(list1.listLevels.at(0).startAt).toEqual(1);

    list2 = doc.lists.at(1);
    TestUtil.verifyListLevel("\u0000)", 18.0, aw.NumberStyle.Arabic, list2.listLevels.at(0));
    expect(list2.listLevels.at(0).font.color).toEqual("#0000FF");
    expect(list2.listLevels.at(0).font.size).toEqual(10.0);
    expect(list2.listLevels.at(0).startAt).toEqual(10);
  });


  test('CreateAndUseListStyle', () => {
    //ExStart
    //ExFor:StyleCollection.add(StyleType,String)
    //ExFor:Style.list
    //ExFor:StyleType
    //ExFor:List.isListStyleDefinition
    //ExFor:List.isListStyleReference
    //ExFor:List.isMultiLevel
    //ExFor:List.style
    //ExFor:ListLevelCollection
    //ExFor:ListLevelCollection.count
    //ExFor:ListLevelCollection.item
    //ExFor:ListCollection.add(Style)
    //ExSummary:Shows how to create a list style and use it in a document.
    let doc = new aw.Document();

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    // We can contain an entire List object within a style.
    let listStyle = doc.styles.add(aw.StyleType.List, "MyListStyle");

    let list1 = listStyle.list;

    expect(list1.isListStyleDefinition).toEqual(true);
    expect(list1.isListStyleReference).toEqual(false);
    expect(list1.isMultiLevel).toEqual(true);
    expect(list1.style).toEqual(listStyle);

    // Change the appearance of all list levels in our list.
    for (let level of list1.listLevels)
    {
      level.font.name = "Verdana";
      level.font.color = "#0000FF";
      level.font.bold = true;
    }

    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Using list style first time:");

    // Create another list from a list within a style.
    let list2 = doc.lists.add(listStyle);

    expect(list2.isListStyleDefinition).toEqual(false);
    expect(list2.isListStyleReference).toEqual(true);
    expect(list2.style).toEqual(listStyle);

    // Add some list items that our list will format.
    builder.listFormat.list = list2;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();

    builder.writeln("Using list style second time:");

    // Create and apply another list based on the list style.
    let list3 = doc.lists.add(listStyle);
    builder.listFormat.list = list3;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();

    builder.document.save(base.artifactsDir + "Lists.CreateAndUseListStyle.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.CreateAndUseListStyle.docx");

    list1 = doc.lists.at(0);

    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, list1.listLevels.at(0));
    expect(list1.isListStyleDefinition).toEqual(true);
    expect(list1.isListStyleReference).toEqual(false);
    expect(list1.isMultiLevel).toEqual(true);
    expect(list1.listLevels.at(0).font.color).toEqual("#0000FF");
    expect(list1.listLevels.at(0).font.name).toEqual("Verdana");
    expect(list1.listLevels.at(0).font.bold).toEqual(true);

    list2 = doc.lists.at(1);

    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, list2.listLevels.at(0));
    expect(list2.isListStyleDefinition).toEqual(false);
    expect(list2.isListStyleReference).toEqual(true);
    expect(list2.isMultiLevel).toEqual(true);

    list3 = doc.lists.at(2);

    TestUtil.verifyListLevel("\u0000.", 18.0, aw.NumberStyle.Arabic, list3.listLevels.at(0));
    expect(list3.isListStyleDefinition).toEqual(false);
    expect(list3.isListStyleReference).toEqual(true);
    expect(list3.isMultiLevel).toEqual(true);
  });


  test('DetectBulletedParagraphs', () => {
    //ExStart
    //ExFor:Paragraph.listFormat
    //ExFor:ListFormat.isListItem
    //ExFor:CompositeNode.getText
    //ExFor:List.listId
    //ExSummary:Shows how to output all paragraphs in a document that are list items.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.applyNumberDefault();
    builder.writeln("Numbered list item 1");
    builder.writeln("Numbered list item 2");
    builder.writeln("Numbered list item 3");
    builder.listFormat.removeNumbers();

    builder.listFormat.applyBulletDefault();
    builder.writeln("Bulleted list item 1");
    builder.writeln("Bulleted list item 2");
    builder.writeln("Bulleted list item 3");
    builder.listFormat.removeNumbers();

    let nodes = [...doc.getChildNodes(aw.NodeType.Paragraph, true)];

    for (let node of nodes.filter(p => p.asParagraph().listFormat.isListItem))
    { 
      var para = node.asParagraph();
      console.log(`This paragraph belongs to list ID# ${para.listFormat.list.listId}, number style \"${para.listFormat.listLevel.numberStyle}\"`);
      console.log(`\t\"${para.getText().trim()}\"`);
    }
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(6);
  });


  test('RemoveBulletsFromParagraphs', () => {
    //ExStart
    //ExFor:ListFormat.removeNumbers
    //ExSummary:Shows how to remove list formatting from all paragraphs in the main text of a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.applyNumberDefault();
    builder.writeln("Numbered list item 1");
    builder.writeln("Numbered list item 2");
    builder.writeln("Numbered list item 3");
    builder.listFormat.removeNumbers();

    let paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();
    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(3);

    for (let n of paras)
      n.asParagraph().listFormat.removeNumbers();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(0);
    //ExEnd
  });


  test('ApplyExistingListToParagraphs', () => {
    //ExStart
    //ExFor:ListCollection.item(Int32)
    //ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Paragraph 1");
    builder.writeln("Paragraph 2");
    builder.write("Paragraph 3");

    let paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(0);

    doc.lists.add(aw.Lists.ListTemplate.NumberDefault);
    let list = doc.lists.at(0);

    for (let node of paras)
    {
      let paragraph = node.asParagraph();
      paragraph.listFormat.list = list;
      paragraph.listFormat.listLevelNumber = 2;
    }

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(3);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(3);
    expect(paras.filter(n => n.asParagraph().listFormat.listLevelNumber == 2).length).toEqual(3);
  });


  test('ApplyNewListToParagraphs', () => {
    //ExStart
    //ExFor:ListCollection.add(ListTemplate)
    //ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Paragraph 1");
    builder.writeln("Paragraph 2");
    builder.write("Paragraph 3");

    let paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(0);

    let list = doc.lists.add(aw.Lists.ListTemplate.NumberUppercaseLetterDot);

    for (let node of paras)
    {
      let paragraph = node.asParagraph();
      paragraph.listFormat.list = list;
      paragraph.listFormat.listLevelNumber = 1;
    }

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(3);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    expect(paras.filter(n => n.asParagraph().listFormat.isListItem).length).toEqual(3);
    expect(paras.filter(n => n.asParagraph().listFormat.listLevelNumber == 1).length).toEqual(3);
  });


  //ExStart
  //ExFor:ListTemplate
  //ExSummary:Shows how to create a document that contains all outline headings list templates.
  test('OutlineHeadingTemplates', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let list = doc.lists.add(aw.Lists.ListTemplate.OutlineHeadingsArticleSection);
    addOutlineHeadingParagraphs(builder, list, "Aspose.words Outline - \"Article Section\"");

    list = doc.lists.add(aw.Lists.ListTemplate.OutlineHeadingsLegal);
    addOutlineHeadingParagraphs(builder, list, "Aspose.words Outline - \"Legal\"");

    builder.insertBreak(aw.BreakType.PageBreak);

    list = doc.lists.add(aw.Lists.ListTemplate.OutlineHeadingsNumbers);
    addOutlineHeadingParagraphs(builder, list, "Aspose.words Outline - \"Numbers\"");

    list = doc.lists.add(aw.Lists.ListTemplate.OutlineHeadingsChapter);
    addOutlineHeadingParagraphs(builder, list, "Aspose.words Outline - \"Chapters\"");

    doc.save(base.artifactsDir + "Lists.OutlineHeadingTemplates.docx");
    testOutlineHeadingTemplates(new aw.Document(base.artifactsDir + "Lists.OutlineHeadingTemplates.docx")); //ExSkip
  });
  //ExEnd

  function addOutlineHeadingParagraphs(builder, list, title) {
    builder.paragraphFormat.clearFormatting();
    builder.writeln(title);
  
    for (let i = 0; i < 9; i++)
    {
      builder.listFormat.list = list;
      builder.listFormat.listLevelNumber = i;
  
      let styleName = "Heading " + (i + 1);
      builder.paragraphFormat.styleName = styleName;
      builder.writeln(styleName);
    }
  
    builder.listFormat.removeNumbers();
  }
  
  function testOutlineHeadingTemplates(doc) {
    let list = doc.lists.at(0); // Article section list template.
  
    TestUtil.verifyListLevel("Article \u0000.", 0.0, aw.NumberStyle.UppercaseRoman, list.listLevels.at(0));
    TestUtil.verifyListLevel("Section \u0000.\u0001", 0.0, aw.NumberStyle.LeadingZero, list.listLevels.at(1));
    TestUtil.verifyListLevel("(\u0002)", 14.4, aw.NumberStyle.LowercaseLetter, list.listLevels.at(2));
    TestUtil.verifyListLevel("(\u0003)", 36.0, aw.NumberStyle.LowercaseRoman, list.listLevels.at(3));
    TestUtil.verifyListLevel("\u0004)", 28.8, aw.NumberStyle.Arabic, list.listLevels.at(4));
    TestUtil.verifyListLevel("\u0005)", 36.0, aw.NumberStyle.LowercaseLetter, list.listLevels.at(5));
    TestUtil.verifyListLevel("\u0006)", 50.4, aw.NumberStyle.LowercaseRoman, list.listLevels.at(6));
    TestUtil.verifyListLevel("\u0007.", 50.4, aw.NumberStyle.LowercaseLetter, list.listLevels.at(7));
    TestUtil.verifyListLevel("\u0008.", 72.0, aw.NumberStyle.LowercaseRoman, list.listLevels.at(8));
  
    list = doc.lists.at(1); // Legal list template.
  
    TestUtil.verifyListLevel("\u0000", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(0));
    TestUtil.verifyListLevel("\u0000.\u0001", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(1));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(2));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(3));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(4));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(5));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(6));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\u0007", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(7));
    TestUtil.verifyListLevel("\u0000.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\u0007.\u0008", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(8));
  
    list = doc.lists.at(2); // Numbered list template.
  
    TestUtil.verifyListLevel("\u0000.", 0.0, aw.NumberStyle.UppercaseRoman, list.listLevels.at(0));
    TestUtil.verifyListLevel("\u0001.", 36.0, aw.NumberStyle.UppercaseLetter, list.listLevels.at(1));
    TestUtil.verifyListLevel("\u0002.", 72.0, aw.NumberStyle.Arabic, list.listLevels.at(2));
    TestUtil.verifyListLevel("\u0003)", 108.0, aw.NumberStyle.LowercaseLetter, list.listLevels.at(3));
    TestUtil.verifyListLevel("(\u0004)", 144.0, aw.NumberStyle.Arabic, list.listLevels.at(4));
    TestUtil.verifyListLevel("(\u0005)", 180.0, aw.NumberStyle.LowercaseLetter, list.listLevels.at(5));
    TestUtil.verifyListLevel("(\u0006)", 216.0, aw.NumberStyle.LowercaseRoman, list.listLevels.at(6));
    TestUtil.verifyListLevel("(\u0007)", 252.0, aw.NumberStyle.LowercaseLetter, list.listLevels.at(7));
    TestUtil.verifyListLevel("(\u0008)", 288.0, aw.NumberStyle.LowercaseRoman, list.listLevels.at(8));
  
    list = doc.lists.at(3); // Chapter list template.
  
    TestUtil.verifyListLevel("Chapter \u0000", 0.0, aw.NumberStyle.Arabic, list.listLevels.at(0));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(1));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(2));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(3));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(4));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(5));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(6));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(7));
    TestUtil.verifyListLevel("", 0.0, aw.NumberStyle.None, list.listLevels.at(8));
  }
  
  
  //ExStart
  //ExFor:ListCollection
  //ExFor:ListCollection.AddCopy(List)
  //ExSummary:Shows how to create a document with a sample of all the lists from another document.
  test('PrintOutAllLists', () => {
    let srcDoc = new aw.Document(base.myDir + "Rendering.docx");

    let dstDoc = new aw.Document();
    let builder = new aw.DocumentBuilder(dstDoc);

    for (let srcList of srcDoc.lists)
    {
      let dstList = dstDoc.lists.addCopy(srcList);
      addListSample(builder, dstList);
    }

    dstDoc.save(base.artifactsDir + "Lists.PrintOutAllLists.docx");
    testPrintOutAllLists(srcDoc, new aw.Document(base.artifactsDir + "Lists.PrintOutAllLists.docx")); //ExSkip
  });

  function addListSample(builder, list) {
    builder.writeln("Sample formatting of list with ListId:" + list.listId);
    builder.listFormat.list = list;
    for (let i = 0; i < list.listLevels.count; i++)
    {
      builder.listFormat.listLevelNumber = i;
      builder.writeln("Level " + i);
    }

    builder.listFormat.removeNumbers();
    builder.writeln();
  }
  //ExEnd

  function testPrintOutAllLists(listSourceDoc, outDoc) {
    for (let list of outDoc.lists)
      for (let i = 0; i < list.listLevels.count; i++)
      {
        let expectedListLevel = [...listSourceDoc.lists].find(l => l.listId == list.listId).listLevels.at(i);
        expect(list.listLevels.at(i).numberFormat).toEqual(expectedListLevel.numberFormat);
        expect(list.listLevels.at(i).numberPosition).toEqual(expectedListLevel.numberPosition);
        expect(list.listLevels.at(i).numberStyle).toEqual(expectedListLevel.numberStyle);
      }
  }


  test('ListDocument', () => {
    //ExStart
    //ExFor:ListCollection.document
    //ExFor:ListCollection.count
    //ExFor:ListCollection.item(Int32)
    //ExFor:ListCollection.getListByListId
    //ExFor:List.document
    //ExFor:List.listId
    //ExSummary:Shows how to verify owner document properties of lists.
    let doc = new aw.Document();

    let lists = doc.lists;
    expect(lists.document.referenceEquals(doc)).toEqual(true);

    let list = lists.add(aw.Lists.ListTemplate.BulletDefault);
    expect(list.document.referenceEquals(doc)).toBe(true);

    console.log("Current list count: " + lists.count);
    console.log("Is the first document list: " + (lists.at(0).equals(list)));
    console.log("ListId: " + list.listId);
    console.log("List is the same by ListId: " + (lists.getListByListId(1).equals(list)));
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    lists = doc.lists;
            
    expect(lists.document.referenceEquals(doc)).toBe(true);
    expect(lists.count).toEqual(1);
    expect(lists.at(0).listId).toEqual(1);
    expect(lists.getListByListId(1)).toEqual(lists.at(0));
  });


  test('CreateListRestartAfterHigher', () => {
    //ExStart
    //ExFor:ListLevel.numberStyle
    //ExFor:ListLevel.numberFormat
    //ExFor:ListLevel.isLegal
    //ExFor:ListLevel.restartAfterLevel
    //ExFor:ListLevel.linkedStyle
    //ExSummary:Shows advances ways of customizing list labels.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
    // We can create nested lists by increasing the indent level. 
    // We can begin and end a list by using a document builder's "ListFormat" property. 
    // Each paragraph that we add between a list's start and the end will become an item in the list.
    let list = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);

    // Level 1 labels will be formatted according to the "Heading 1" paragraph style and will have a prefix.
    // These will look like "Appendix A", "Appendix B"...
    list.listLevels.at(0).numberFormat = "Appendix \u0000";
    list.listLevels.at(0).numberStyle = aw.NumberStyle.UppercaseLetter;
    list.listLevels.at(0).linkedStyle = doc.styles.at("Heading 1");

    // Level 2 labels will display the current numbers of the first and the second list levels and have leading zeroes.
    // If the first list level is at 1, then the list labels from these will look like "Section (1.01)", "Section (1.02)"...
    list.listLevels.at(1).numberFormat = "Section (\u0000.\u0001)";
    list.listLevels.at(1).numberStyle = aw.NumberStyle.LeadingZero;

    // Note that the higher-level uses UppercaseLetter numbering.
    // We can set the "IsLegal" property to use Arabic numbers for the higher list levels.
    list.listLevels.at(1).isLegal = true;
    list.listLevels.at(1).restartAfterLevel = 0;

    // Level 3 labels will be upper case Roman numerals with a prefix and a suffix and will restart at each List level 1 item.
    // These list labels will look like "-I-", "-II-"...
    list.listLevels.at(2).numberFormat = "-\u0002-";
    list.listLevels.at(2).numberStyle = aw.NumberStyle.UppercaseRoman;
    list.listLevels.at(2).restartAfterLevel = 1;

    // Make labels of all list levels bold.
    for (let level of list.listLevels)
      level.font.bold = true;

    // Apply list formatting to the current paragraph.
    builder.listFormat.list = list;

    // Create list items that will display all three of our list levels.
    for (let n = 0; n < 2; n++)
    {
      for (let i = 0; i < 3; i++)
      {
        builder.listFormat.listLevelNumber = i;
        builder.writeln("Level " + i);
      }
    }

    builder.listFormat.removeNumbers();

    doc.save(base.artifactsDir + "Lists.CreateListRestartAfterHigher.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.CreateListRestartAfterHigher.docx");

    let listLevel = doc.lists.at(0).listLevels.at(0);

    TestUtil.verifyListLevel("Appendix \u0000", 18.0, aw.NumberStyle.UppercaseLetter, listLevel);
    expect(listLevel.isLegal).toEqual(false);
    expect(listLevel.restartAfterLevel).toEqual(-1);
    expect(listLevel.linkedStyle.name).toEqual("Heading 1");

    listLevel = doc.lists.at(0).listLevels.at(1);

    TestUtil.verifyListLevel("Section (\u0000.\u0001)", 54.0, aw.NumberStyle.LeadingZero, listLevel);
    expect(listLevel.isLegal).toEqual(true);
    expect(listLevel.restartAfterLevel).toEqual(0);
    expect(listLevel.linkedStyle).toBe(null);
  });


  test('GetListLabels', () => {
    //ExStart
    //ExFor:Document.updateListLabels()
    //ExFor:Node.toString(SaveFormat)
    //ExFor:ListLabel
    //ExFor:Paragraph.listLabel
    //ExFor:ListLabel.labelValue
    //ExFor:ListLabel.labelString
    //ExSummary:Shows how to extract the list labels of all paragraphs that are list items.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    doc.updateListLabels();

    let paras = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray();

    // Find if we have the paragraph list. In our document, our list uses plain Arabic numbers,
    // which start at three and ends at six.
    for (let node of paras.filter(p => p.asParagraph().listFormat.isListItem))
    {
      let paragraph = node.asParagraph();
      console.log(`List item paragraph #${paras.indexOf(paragraph)}`);

      // This is the text we get when getting when we output this node to text format.
      // This text output will omit list labels. Trim any paragraph formatting characters. 
      let paragraphText = paragraph.toString(aw.SaveFormat.Text).trim();
      console.log(`\tExported Text: ${paragraphText}`);

      let label = paragraph.listLabel;

      // This gets the position of the paragraph in the current level of the list. If we have a list with multiple levels,
      // this will tell us what position it is on that level.
      console.log(`\tNumerical Id: ${label.labelValue}`);

      // Combine them together to include the list label with the text in the output.
      console.log(`\tList label combined with text: ${label.labelString} ${paragraphText}`);
    }
    //ExEnd

    expect(paras.filter(p => p.asParagraph().listFormat.isListItem).length).toEqual(10);
  });


  test('CreatePictureBullet', () => {
    //ExStart
    //ExFor:ListLevel.createPictureBullet
    //ExFor:ListLevel.deletePictureBullet
    //ExSummary:Shows how to set a custom image icon for list item labels.
    let doc = new aw.Document();

    let list = doc.lists.add(aw.Lists.ListTemplate.BulletCircle);

    // Create a picture bullet for the current list level, and set an image from a local file system
    // as the icon that the bullets for this list level will display.
    list.listLevels.at(0).createPictureBullet();
    list.listLevels.at(0).imageData.setImage(base.imageDir + "Logo icon.ico");

    expect(list.listLevels.at(0).imageData.hasImage).toEqual(true);

    let builder = new aw.DocumentBuilder(doc);

    builder.listFormat.list = list;
    builder.writeln("Hello world!");
    builder.write("Hello again!");

    doc.save(base.artifactsDir + "Lists.createPictureBullet.docx");

    list.listLevels.at(0).deletePictureBullet();

    expect(list.listLevels.at(0).imageData).toBe(null);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Lists.createPictureBullet.docx");

    expect(doc.lists.at(0).listLevels.at(0).imageData.hasImage).toEqual(true);
  });


  test('GetCustomNumberStyleFormat', () => {
    //ExStart
    //ExFor:ListLevel.customNumberStyleFormat
    //ExFor:ListLevel.getEffectiveValue(Int32, NumberStyle, String)
    //ExSummary:Shows how to get the format for a list with the custom number style.
    let doc = new aw.Document(base.myDir + "List with leading zero.docx");

    let listLevel = doc.firstSection.body.paragraphs.at(0).listFormat.listLevel;

    let customNumberStyleFormat = '';

    if (listLevel.numberStyle == aw.NumberStyle.Custom)
      customNumberStyleFormat = listLevel.customNumberStyleFormat;

    expect(customNumberStyleFormat).toEqual("001, 002, 003, ...");

    // We can get value for the specified index of the list item.
    expect(aw.Lists.ListLevel.getEffectiveValue(4, aw.NumberStyle.LowercaseRoman, null)).toEqual("iv");
    expect(aw.Lists.ListLevel.getEffectiveValue(5, aw.NumberStyle.Custom, customNumberStyleFormat)).toEqual("005");
    //ExEnd

    expect(() => aw.Lists.ListLevel.getEffectiveValue(5, aw.NumberStyle.LowercaseRoman, customNumberStyleFormat))
      .toThrow("For this number style, the specified argument must be null or empty. (Parameter 'customNumberStyleFormat')");
    expect(() => aw.Lists.ListLevel.getEffectiveValue(5, aw.NumberStyle.Custom, null)).toThrow(
      "Specified argument must not be null or empty if the number style is custom. (Parameter 'customNumberStyleFormat')");
    expect(() => aw.Lists.ListLevel.getEffectiveValue(5, aw.NumberStyle.Custom, "....")).toThrow(
      "Unexpected custom number format style: '....'");
  });


  test('HasSameTemplate', () => {
    //ExStart
    //ExFor:List.hasSameTemplate(List)
    //ExSummary:Shows how to define lists with the same ListDefId.
    let doc = new aw.Document(base.myDir + "Different lists.docx");

    expect(doc.lists.at(0).hasSameTemplate(doc.lists.at(1))).toEqual(true);
    expect(doc.lists.at(1).hasSameTemplate(doc.lists.at(2))).toEqual(false);
    //ExEnd
  });


  test('SetCustomNumberStyleFormat', () => {
    //ExStart:SetCustomNumberStyleFormat
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:ListLevel.customNumberStyleFormat
    //ExSummary:Shows how to set customer number style format.
    let doc = new aw.Document(base.myDir + "List with leading zero.docx");

    doc.updateListLabels();

    let paras = doc.firstSection.body.paragraphs;
    expect(paras.at(0).listLabel.labelString).toEqual("001.");
    expect(paras.at(1).listLabel.labelString).toEqual("0001.");
    expect(paras.at(2).listLabel.labelString).toEqual("0002.");

    paras.at(1).listFormat.listLevel.customNumberStyleFormat = "001, 002, 003, ...";

    doc.updateListLabels();

    expect(paras.at(0).listLabel.labelString).toEqual("001.");
    expect(paras.at(1).listLabel.labelString).toEqual("001.");
    expect(paras.at(2).listLabel.labelString).toEqual("002.");
    //ExEnd:SetCustomNumberStyleFormat
  });

  
  test('AddSingleLevelList', () => {
    //ExStart:AddSingleLevelList
    //GistId:95fdae949cefbf2ce485acc95cccc495
    //ExFor:ListCollection.addSingleLevelList(ListTemplate)
    //ExSummary:Shows how to create a new single level list based on the predefined template.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let listCollection = doc.lists;

    // Creates the bulleted list from BulletCircle template.
    let bulletedList = listCollection.addSingleLevelList(aw.Lists.ListTemplate.BulletCircle);

    // Writes the bulleted list to the resulting document.
    builder.writeln("Bulleted list starts below:");
    builder.listFormat.list = bulletedList;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    builder.listFormat.removeNumbers();

    // Creates the numbered list from NumberUppercaseLetterDot template.
    let numberedList = listCollection.addSingleLevelList(aw.Lists.ListTemplate.NumberUppercaseLetterDot);

    // Writes the numbered list to the resulting document.
    builder.writeln("Numbered list starts below:");
    builder.listFormat.list = numberedList;
    builder.writeln("Item 1");
    builder.writeln("Item 2");

    doc.save(base.artifactsDir + "Lists.addSingleLevelList.docx");
    //ExEnd:AddSingleLevelList
  });


});
