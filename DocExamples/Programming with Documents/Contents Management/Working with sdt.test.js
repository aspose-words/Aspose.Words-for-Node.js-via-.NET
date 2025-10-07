// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithSdt", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('SdtCheckBox', () => {
    //ExStart:SdtCheckBox
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let sdtCheckBox = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Checkbox, aw.Markup.MarkupLevel.Inline);
    builder.insertNode(sdtCheckBox);

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtCheckBox.docx");
    //ExEnd:SdtCheckBox
  });

  test('CurrentStateOfCheckBox', () => {
    //ExStart:CurrentStateOfCheckBox
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    // Get the first content control from the document.
    let sdtCheckBox = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();

    if (sdtCheckBox.sdtType == aw.Markup.SdtType.Checkbox) {
      sdtCheckBox.checked = true;
    }

    doc.save(base.artifactsDir + "WorkingWithSdt.CurrentStateOfCheckBox.docx");
    //ExEnd:CurrentStateOfCheckBox
  });

  test('ModifySdt', () => {
    //ExStart:ModifySdt
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    let sdts = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);
    for (let i = 0; i < sdts.count; i++) {
      let sdt = sdts.at(i).asStructuredDocumentTag();

      if (sdt.sdtType == aw.Markup.SdtType.PlainText) {
        sdt.removeAllChildren();
        let para = sdt.appendChild(new aw.Paragraph(doc)).asParagraph();
        let run = new aw.Run(doc, "new text goes here");
        para.appendChild(run);
      } else if (sdt.sdtType == aw.Markup.SdtType.DropDownList) {
        let second_item = sdt.listItems.at(2);
        sdt.listItems.selectedValue = second_item;
      } else if (sdt.sdtType == aw.Markup.SdtType.Picture) {
        let shape = sdt.getChild(aw.NodeType.Shape, 0, true).asShape();
        if (shape.hasImage) {
          shape.imageData.setImage(base.imagesDir + "Watermark.png");
        }
      }
    }

    doc.save(base.artifactsDir + "WorkingWithSdt.ModifySdt.docx");
    //ExEnd:ModifySdt
  });

  test('SdtComboBox', () => {
    //ExStart:SdtComboBox
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document();

    let sdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.ComboBox, aw.Markup.MarkupLevel.Block);
    sdt.listItems.add(new aw.Markup.SdtListItem("Choose an item", "-1"));
    sdt.listItems.add(new aw.Markup.SdtListItem("Item 1", "1"));
    sdt.listItems.add(new aw.Markup.SdtListItem("Item 2", "2"));
    doc.firstSection.body.appendChild(sdt);

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtComboBox.docx");
    //ExEnd:SdtComboBox
  });

  test('SdtRichTextBox', () => {
    //ExStart:SdtRichTextBox
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document();

    let sdtRichText = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RichText, aw.Markup.MarkupLevel.Block);

    let para = new aw.Paragraph(doc);
    let run = new aw.Run(doc);
    run.text = "Hello World";
    run.font.color = "#008000"; // green
    para.runs.add(run);
    sdtRichText.getChildNodes(aw.NodeType.Any, false).add(para);
    doc.firstSection.body.appendChild(sdtRichText);

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtRichTextBox.docx");
    //ExEnd:SdtRichTextBox
  });

  test('SdtColor', () => {
    //ExStart:SdtColor
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    let sdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    sdt.color = "#FF0000"; // Red.

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtColor.docx");
    //ExEnd:SdtColor
  });

  test('ClearSdt', () => {
    //ExStart:ClearSdt
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    let sdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    sdt.clear();

    doc.save(base.artifactsDir + "WorkingWithSdt.ClearSdt.doc");
    //ExEnd:ClearSdt
  });

  test('BindSdtToCustomXmlPart', () => {
    //ExStart:BindSdtToCustomXmlPart
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document();
    let xmlPart = doc.customXmlParts.add(generateUuid(), "<root><text>Hello, World!</text></root>");

    let sdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Block);
    doc.firstSection.body.appendChild(sdt);

    sdt.xmlMapping.setMapping(xmlPart, "/root[1]/text[1]", "");

    doc.save(base.artifactsDir + "WorkingWithSdt.BindSdtToCustomXmlPart.doc");
    //ExEnd:BindSdtToCustomXmlPart
  });

  test('SdtStyle', () => {
    //ExStart:SdtStyle
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    let sdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    let style = doc.styles.at(aw.StyleIdentifier.Quote);
    sdt.style = style;

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtStyle.docx");
    //ExEnd:SdtStyle
  });

  test('RepeatingSectionMappedToCustomXmlPart', () => {
    //ExStart:RepeatingSectionMappedToCustomXmlPart
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let xmlPart = doc.customXmlParts.add("Books", `
        <books>
            <book>
                <title>Everyday Italian</title>
                <author>Giada De Laurentiis</author>
            </book>
            <book>
                <title>Harry Potter</title>
                <author>J K. Rowling</author>
            </book>
            <book>
                <title>Learning XML</title>
                <author>Erik T. Ray</author>
            </book>
       </books>`);

    let table = builder.startTable();

    builder.insertCell();
    builder.write("Title");

    builder.insertCell();
    builder.write("Author");

    builder.endRow();
    builder.endTable();

    let repeatingSectionSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RepeatingSection, aw.Markup.MarkupLevel.Row);
    repeatingSectionSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book", "");
    table.appendChild(repeatingSectionSdt);

    let repeatingSectionItemSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RepeatingSectionItem, aw.Markup.MarkupLevel.Row);
    repeatingSectionSdt.appendChild(repeatingSectionItemSdt);

    let row = new aw.Tables.Row(doc);
    repeatingSectionItemSdt.appendChild(row);

    let titleSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Cell);
    titleSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book[1]/title[1]", "");
    row.appendChild(titleSdt);

    let authorSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Cell);
    authorSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book[1]/author[1]", "");
    row.appendChild(authorSdt);

    doc.save(base.artifactsDir + "WorkingWithSdt.RepeatingSectionMappedToCustomXmlPart.docx");
    //ExEnd:RepeatingSectionMappedToCustomXmlPart
  });

  test('MultiSection', () => {
    //ExStart:MultiSection
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");

    let tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTagRangeStart, true);

    for (let i = 0; i < tags.count; i++) {
      let tag = tags.at(i);
      console.log(tag.asStructuredDocumentTagRangeStart().title);
    }
    //ExEnd:MultiSection
  });

  test('SdtRangeStartXmlMapping', () => {
    //ExStart:SdtRangeStartXmlMapping
    //GistId:625644238d5cac4a2215ccfe46030666
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");

    // Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
    let xmlPartId = generateUuid();
    let xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
    let xmlPart = doc.customXmlParts.add(xmlPartId, xmlPartContent);
    console.log(xmlPart.data);

    // Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
    let sdtRangeStart = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, true).asStructuredDocumentTagRangeStart();

    // If we set a mapping for our StructuredDocumentTag,
    // it will only display a part of the CustomXmlPart that the XPath points to.
    // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
    sdtRangeStart.xmlMapping.setMapping(xmlPart, "/root[1]/text[2]", null);

    doc.save(base.artifactsDir + "WorkingWithSdt.SdtRangeStartXmlMapping.docx");
    //ExEnd:SdtRangeStartXmlMapping
  });

  function generateUuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      let r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }
});