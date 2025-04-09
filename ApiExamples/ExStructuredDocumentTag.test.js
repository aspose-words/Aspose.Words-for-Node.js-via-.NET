// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const { Guid } = require('js-guid');


describe("ExStructuredDocumentTag", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('RepeatingSection', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.sdtType
    //ExFor:IStructuredDocumentTag.sdtType
    //ExSummary:Shows how to get the type of a structured document tag.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    var tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);

    expect(tags.at(0).asStructuredDocumentTag().sdtType).toEqual(aw.Markup.SdtType.RepeatingSection);
    expect(tags.at(1).asStructuredDocumentTag().sdtType).toEqual(aw.Markup.SdtType.RepeatingSectionItem);
    expect(tags.at(2).asStructuredDocumentTag().sdtType).toEqual(aw.Markup.SdtType.RichText);
    //ExEnd
  });


  test('FlatOpcContent', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.wordOpenXML
    //ExFor:IStructuredDocumentTag.wordOpenXML
    //ExSummary:Shows how to get XML contained within the node in the FlatOpc format.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    var tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);

    expect(tags.at(0).asStructuredDocumentTag().wordOpenXML.includes(
        "<pkg:part pkg:name=\"/docProps/app.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\">")).toEqual(true);
    //ExEnd
  });


  test('ApplyStyle', () => {
    //ExStart
    //ExFor:StructuredDocumentTag
    //ExFor:StructuredDocumentTag.nodeType
    //ExFor:StructuredDocumentTag.style
    //ExFor:StructuredDocumentTag.styleName
    //ExFor:StructuredDocumentTag.wordOpenXMLMinimal
    //ExFor:MarkupLevel
    //ExFor:SdtType
    //ExSummary:Shows how to work with styles for content control elements.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways to apply a style from the document to a structured document tag.
    // 1 -  Apply a style object from the document's style collection:
    let quoteStyle = doc.styles.at(aw.StyleIdentifier.Quote);
    let sdtPlainText = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);
    sdtPlainText.style = quoteStyle;

    // 2 -  Reference a style in the document by name:
    let sdtRichText = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RichText, aw.Markup.MarkupLevel.Inline);
    sdtRichText.styleName = "Quote";

    builder.insertNode(sdtPlainText);
    builder.insertNode(sdtRichText);

    expect(sdtPlainText.nodeType).toEqual(aw.NodeType.StructuredDocumentTag);

    let tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);

    for (let node of tags)
    {
      let sdt = node.asStructuredDocumentTag();

      console.log(sdt.wordOpenXMLMinimal);

      expect(sdt.style.styleIdentifier).toEqual(aw.StyleIdentifier.Quote);
      expect(sdt.styleName).toEqual("Quote");
    }
    //ExEnd
  });


  test('CheckBox', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.#ctor(DocumentBase, SdtType, MarkupLevel)
    //ExFor:StructuredDocumentTag.checked
    //ExFor:StructuredDocumentTag.setCheckedSymbol(Int32, String)
    //ExFor:StructuredDocumentTag.setUncheckedSymbol(Int32, String)
    //ExSummary:Show how to create a structured document tag in the form of a check box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let sdtCheckBox = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Checkbox, aw.Markup.MarkupLevel.Inline);
    sdtCheckBox.checked = true;

    // We can set the symbols used to represent the checked/unchecked state of a checkbox content control.
    sdtCheckBox.setCheckedSymbol(0x00A9, "Times New Roman");
    sdtCheckBox.setUncheckedSymbol(0x00AE, "Times New Roman");

    builder.insertNode(sdtCheckBox);

    doc.save(base.artifactsDir + "StructuredDocumentTag.checkBox.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.checkBox.docx");

    let tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true);

    expect(tags.at(0).asStructuredDocumentTag().checked).toEqual(true);
    expect(tags.at(0).asStructuredDocumentTag().xmlMapping.storeItemId).toEqual('');
  });


  test('Date', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.calendarType
    //ExFor:StructuredDocumentTag.dateDisplayFormat
    //ExFor:StructuredDocumentTag.dateDisplayLocale
    //ExFor:StructuredDocumentTag.dateStorageFormat
    //ExFor:StructuredDocumentTag.fullDate
    //ExFor:SdtCalendarType
    //ExFor:SdtDateStorageFormat
    //ExSummary:Shows how to prompt the user to enter a date with a structured document tag.
    let doc = new aw.Document();

    // Insert a structured document tag that prompts the user to enter a date.
    // In Microsoft Word, this element is known as a "Date picker content control".
    // When we click on the arrow on the right end of this tag in Microsoft Word,
    // we will see a pop up in the form of a clickable calendar.
    // We can use that popup to select a date that the tag will display.
    let sdtDate = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Date, aw.Markup.MarkupLevel.Inline);

    // Display the date, according to the Saudi Arabian Arabic locale.
    sdtDate.dateDisplayLocale = 1025;//CultureInfo.GetCultureInfo("ar-SA").LCID;

    // Set the format with which to display the date.
    sdtDate.dateDisplayFormat = "dd MMMM, yyyy";
    sdtDate.dateStorageFormat = aw.Markup.SdtDateStorageFormat.DateTime;

    // Display the date according to the Hijri calendar.
    sdtDate.calendarType = aw.Markup.SdtCalendarType.Hijri;

    // Before the user chooses a date in Microsoft Word, the tag will display the text "Click here to enter a date.".
    // According to the tag's calendar, set the "FullDate" property to get the tag to display a default date.
    sdtDate.fullDate = new Date(1440, 9, 20);

    let builder = new aw.DocumentBuilder(doc);
    builder.insertNode(sdtDate);

    doc.save(base.artifactsDir + "StructuredDocumentTag.date.docx");
    //ExEnd
  });


  test('PlainText', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.color
    //ExFor:StructuredDocumentTag.contentsFont
    //ExFor:StructuredDocumentTag.endCharacterFont
    //ExFor:StructuredDocumentTag.id
    //ExFor:StructuredDocumentTag.level
    //ExFor:StructuredDocumentTag.multiline
    //ExFor:IStructuredDocumentTag.tag
    //ExFor:StructuredDocumentTag.tag
    //ExFor:StructuredDocumentTag.title
    //ExFor:StructuredDocumentTag.removeSelfOnly
    //ExFor:StructuredDocumentTag.appearance
    //ExSummary:Shows how to create a structured document tag in a plain text box and modify its appearance.
    let doc = new aw.Document();

    // Create a structured document tag that will contain plain text.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);

    // Set the title and color of the frame that appears when you mouse over the structured document tag in Microsoft Word.
    tag.title = "My plain text";
    tag.color = "#FF00FF";

    // Set a tag for this structured document tag, which is obtainable
    // as an XML element named "tag", with the string below in its "@val" attribute.
    tag.tag = "MyPlainTextSDT";

    // Every structured document tag has a random unique ID.
    expect(tag.id > 0).toEqual(true);

    // Set the font for the text inside the structured document tag.
    tag.contentsFont.name = "Arial";

    // Set the font for the text at the end of the structured document tag.
    // Any text that we type in the document body after moving out of the tag with arrow keys will use this font.
    tag.endCharacterFont.name = "Arial Black";

    // By default, this is false and pressing enter while inside a structured document tag does nothing.
    // When set to true, our structured document tag can have multiple lines.

    // Set the "Multiline" property to "false" to only allow the contents
    // of this structured document tag to span a single line.
    // Set the "Multiline" property to "true" to allow the tag to contain multiple lines of content.
    tag.multiline = true;

    // Set the "Appearance" property to "SdtAppearance.Tags" to show tags around content.
    // By default structured document tag shows as BoundingBox. 
    tag.appearance = aw.Markup.SdtAppearance.Tags;

    let builder = new aw.DocumentBuilder(doc);
    builder.insertNode(tag);

    // Insert a clone of our structured document tag in a new paragraph.
    let tagClone = tag.clone(true).asStructuredDocumentTag();
    builder.insertParagraph();
    builder.insertNode(tagClone);

    // Use the "RemoveSelfOnly" method to remove a structured document tag, while keeping its contents in the document.
    tagClone.removeSelfOnly();

    doc.save(base.artifactsDir + "StructuredDocumentTag.plainText.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.plainText.docx");
    tag = doc.getSdt(0, true);

    expect(tag.title).toEqual("My plain text");
    expect(tag.color).toEqual("#FF00FF");
    expect(tag.tag).toEqual("MyPlainTextSDT");
    expect(tag.id > 0).toEqual(true);
    expect(tag.contentsFont.name).toEqual("Arial");
    expect(tag.endCharacterFont.name).toEqual("Arial Black");
    expect(tag.multiline).toEqual(true);
    expect(tag.appearance).toEqual(aw.Markup.SdtAppearance.Tags);
  });


  test.each([false,
    true])('IsTemporary', (isTemporary) => {
    //ExStart
    //ExFor:StructuredDocumentTag.isTemporary
    //ExSummary:Shows how to make single-use controls.
    let doc = new aw.Document();

    // Insert a plain text structured document tag,
    // which will act as a plain text form that the user may enter text into.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);

    // Set the "IsTemporary" property to "true" to make the structured document tag disappear and
    // assimilate its contents into the document after the user edits it once in Microsoft Word.
    // Set the "IsTemporary" property to "false" to allow the user to edit the contents
    // of the structured document tag any number of times.
    tag.isTemporary = isTemporary;

    let builder = new aw.DocumentBuilder(doc);
    builder.write("Please enter text: ");
    builder.insertNode(tag);

    // Insert another structured document tag in the form of a check box and set its default state to "checked".
    tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Checkbox, aw.Markup.MarkupLevel.Inline);
    tag.checked = true;

    // Set the "IsTemporary" property to "true" to make the check box become a symbol
    // once the user clicks on it in Microsoft Word.
    // Set the "IsTemporary" property to "false" to allow the user to click on the check box any number of times.
    tag.isTemporary = isTemporary;

    builder.write("\nPlease click the check box: ");
    builder.insertNode(tag);

    doc.save(base.artifactsDir + "StructuredDocumentTag.isTemporary.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.isTemporary.docx");

    expect(doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true).toArray()
      .filter(sdt => sdt.asStructuredDocumentTag().isTemporary == isTemporary).length).toEqual(2);
  });


  test.each([false,
    true])('PlaceholderBuildingBlock', (isShowingPlaceholderText) => {
    //ExStart
    //ExFor:StructuredDocumentTag.isShowingPlaceholderText
    //ExFor:IStructuredDocumentTag.isShowingPlaceholderText
    //ExFor:StructuredDocumentTag.placeholder
    //ExFor:StructuredDocumentTag.placeholderName
    //ExFor:IStructuredDocumentTag.placeholder
    //ExFor:IStructuredDocumentTag.placeholderName
    //ExSummary:Shows how to use a building block's contents as a custom placeholder text for a structured document tag. 
    let doc = new aw.Document();

    // Insert a plain text structured document tag of the "PlainText" type, which will function as a text box.
    // The contents that it will display by default are a "Click here to enter text." prompt.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);

    // We can get the tag to display the contents of a building block instead of the default text.
    // First, add a building block with contents to the glossary document.
    let glossaryDoc = doc.glossaryDocument;

    let substituteBlock = new aw.BuildingBlocks.BuildingBlock(glossaryDoc);
    substituteBlock.name = "Custom Placeholder";
    substituteBlock.appendChild(new aw.Section(glossaryDoc));
    substituteBlock.firstSection.appendChild(new aw.Body(glossaryDoc));
    substituteBlock.firstSection.body.appendParagraph("Custom placeholder text.");

    glossaryDoc.appendChild(substituteBlock);

    // Then, use the structured document tag's "PlaceholderName" property to reference that building block by name.
    tag.placeholderName = "Custom Placeholder";

    // If "PlaceholderName" refers to an existing block in the parent document's glossary document,
    // we will be able to verify the building block via the "Placeholder" property.
    expect(tag.placeholder.referenceEquals(substituteBlock)).toBe(true);

    // Set the "IsShowingPlaceholderText" property to "true" to treat the
    // structured document tag's current contents as placeholder text.
    // This means that clicking on the text box in Microsoft Word will immediately highlight all the tag's contents.
    // Set the "IsShowingPlaceholderText" property to "false" to get the
    // structured document tag to treat its contents as text that a user has already entered.
    // Clicking on this text in Microsoft Word will place the blinking cursor at the clicked location.
    tag.isShowingPlaceholderText = isShowingPlaceholderText;

    let builder = new aw.DocumentBuilder(doc);
    builder.insertNode(tag);

    doc.save(base.artifactsDir + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.PlaceholderBuildingBlock.docx");
    tag = doc.getSdt(0, true);
    substituteBlock = doc.glossaryDocument.getChild(aw.NodeType.BuildingBlock, 0, true).asBuildingBlock();

    expect(substituteBlock.name).toEqual("Custom Placeholder");
    expect(tag.isShowingPlaceholderText).toEqual(isShowingPlaceholderText);
    expect(tag.placeholder.referenceEquals(substituteBlock)).toBe(true);
    expect(tag.placeholderName).toEqual(substituteBlock.name);
  });

  test('Lock', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.lockContentControl
    //ExFor:StructuredDocumentTag.lockContents
    //ExFor:IStructuredDocumentTag.lockContentControl
    //ExFor:IStructuredDocumentTag.lockContents
    //ExSummary:Shows how to apply editing restrictions to structured document tags.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a plain text structured document tag, which acts as a text box that prompts the user to fill it in.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);

    // Set the "LockContents" property to "true" to prohibit the user from editing this text box's contents.
    tag.lockContents = true;
    builder.write("The contents of this structured document tag cannot be edited: ");
    builder.insertNode(tag);

    tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Inline);

    // Set the "LockContentControl" property to "true" to prohibit the user from
    // deleting this structured document tag manually in Microsoft Word.
    tag.lockContentControl = true;

    builder.insertParagraph();
    builder.write("This structured document tag cannot be deleted but its contents can be edited: ");
    builder.insertNode(tag);

    doc.save(base.artifactsDir + "StructuredDocumentTag.Lock.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.Lock.docx");
    tag = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();

    expect(tag.lockContents).toEqual(true);
    expect(tag.lockContentControl).toEqual(false);

    tag = doc.getChild(aw.NodeType.StructuredDocumentTag, 1, true).asStructuredDocumentTag();

    expect(tag.lockContents).toEqual(false);
    expect(tag.lockContentControl).toEqual(true);
  });


  test('ListItemCollection', () => {
    //ExStart
    //ExFor:SdtListItem
    //ExFor:SdtListItem.#ctor(String)
    //ExFor:SdtListItem.#ctor(String,String)
    //ExFor:SdtListItem.displayText
    //ExFor:SdtListItem.value
    //ExFor:SdtListItemCollection
    //ExFor:SdtListItemCollection.add(SdtListItem)
    //ExFor:SdtListItemCollection.clear
    //ExFor:SdtListItemCollection.count
    //ExFor:SdtListItemCollection.getEnumerator
    //ExFor:SdtListItemCollection.item(Int32)
    //ExFor:SdtListItemCollection.removeAt(Int32)
    //ExFor:SdtListItemCollection.selectedValue
    //ExFor:StructuredDocumentTag.listItems
    //ExSummary:Shows how to work with drop down-list structured document tags.
    let doc = new aw.Document();
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.DropDownList, aw.Markup.MarkupLevel.Block);
    doc.firstSection.body.appendChild(tag);

    // A drop-down list structured document tag is a form that allows the user to
    // select an option from a list by left-clicking and opening the form in Microsoft Word.
    // The "ListItems" property contains all list items, and each list item is an "SdtListItem".
    let listItems = tag.listItems;
    listItems.add(new aw.Markup.SdtListItem("Value 1"));

    expect(listItems.at(0).value).toEqual(listItems.at(0).displayText);

    // Add 3 more list items. Initialize these items using a different constructor to the first item
    // to display strings that are different from their values.
    listItems.add(new aw.Markup.SdtListItem("Item 2", "Value 2"));
    listItems.add(new aw.Markup.SdtListItem("Item 3", "Value 3"));
    listItems.add(new aw.Markup.SdtListItem("Item 4", "Value 4"));

    expect(listItems.count).toEqual(4);

    // The drop-down list is displaying the first item. Assign a different list item to the "SelectedValue" to display it.
    listItems.selectedValue = listItems.at(3);

    expect(listItems.selectedValue.value).toEqual("Value 4");

    // Enumerate over the collection and print each element.
    for (let listItem of listItems) {
      if (listItem != null)
        console.log(`List item: ${listItem.displayText}, value: ${listItem.value}`);
    }

    // Remove the last list item. 
    listItems.removeAt(3);

    expect(listItems.count).toEqual(3);

    // Since our drop-down control is set to display the removed item by default, give it an item to display which exists.
    listItems.selectedValue = listItems.at(1);

    doc.save(base.artifactsDir + "StructuredDocumentTag.ListItemCollection.docx");

    // Use the "Clear" method to empty the entire drop-down item collection at once.
    listItems.clear();

    expect(listItems.count).toEqual(0);
    //ExEnd
  });


  test('CreatingCustomXml', () => {
    //ExStart
    //ExFor:CustomXmlPart
    //ExFor:CustomXmlPart.clone
    //ExFor:CustomXmlPart.data
    //ExFor:CustomXmlPart.id
    //ExFor:CustomXmlPart.schemas
    //ExFor:CustomXmlPartCollection
    //ExFor:CustomXmlPartCollection.add(CustomXmlPart)
    //ExFor:CustomXmlPartCollection.add(String, String)
    //ExFor:CustomXmlPartCollection.clear
    //ExFor:CustomXmlPartCollection.clone
    //ExFor:CustomXmlPartCollection.count
    //ExFor:CustomXmlPartCollection.getById(String)
    //ExFor:CustomXmlPartCollection.getEnumerator
    //ExFor:CustomXmlPartCollection.item(Int32)
    //ExFor:CustomXmlPartCollection.removeAt(Int32)
    //ExFor:Document.customXmlParts
    //ExFor:StructuredDocumentTag.xmlMapping
    //ExFor:IStructuredDocumentTag.xmlMapping
    //ExFor:XmlMapping.setMapping(CustomXmlPart, String, String)
    //ExSummary:Shows how to create a structured document tag with custom XML data.
    let doc = new aw.Document();

    // Construct an XML part that contains data and add it to the document's collection.
    // If we enable the "Developer" tab in Microsoft Word,
    // we can find elements from this collection in the "XML Mapping Pane", along with a few default elements.
    let xmlPartId = `{${Guid.newGuid().toString()}}`;
    let xmlPartContent = "<root><text>Hello world!</text></root>";
    let xmlPart = doc.customXmlParts.add(xmlPartId, xmlPartContent);

    expect(Buffer.from(xmlPart.data).toString("ascii")).toEqual(xmlPartContent);
    expect(xmlPart.id).toEqual(xmlPartId);

    // Below are two ways to refer to XML parts.
    // 1 -  By an index in the custom XML part collection:
    expect(doc.customXmlParts.at(0)).toEqual(xmlPart);

    // 2 -  By GUID:
    expect(doc.customXmlParts.getById(xmlPartId)).toEqual(xmlPart);

    // Add an XML schema association.
    xmlPart.schemas.add("http://www.w3.org/2001/XMLSchema");

    // Clone a part, and then insert it into the collection.
    let xmlPartClone = xmlPart.clone();
    xmlPartClone.id = `{${Guid.newGuid().toString()}}`;
    doc.customXmlParts.add(xmlPartClone);

    expect(doc.customXmlParts.count).toEqual(2);

    // Iterate through the collection and print the contents of each part.
    for (let index = 0; index < doc.customXmlParts.count; ++index) {
      console.log(`XML part index ${index}, ID: ${doc.customXmlParts.at(index).id}`);
      console.log(`\tContent: ${Buffer.from(doc.customXmlParts.at(index).data).toString("utf8")}`);
    }

    // Use the "RemoveAt" method to remove the cloned part by index.
    doc.customXmlParts.removeAt(1);

    expect(doc.customXmlParts.count).toEqual(1);

    // Clone the XML parts collection, and then use the "Clear" method to remove all its elements at once.
    let customXmlParts = doc.customXmlParts.clone();
    customXmlParts.clear();

    // Create a structured document tag that will display our part's contents and insert it into the document body.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Block);
    tag.xmlMapping.setMapping(xmlPart, "/root[1]/text[1]", '');

    doc.firstSection.body.appendChild(tag);

    doc.save(base.artifactsDir + "StructuredDocumentTag.customXml.docx");
    //ExEnd

    expect(DocumentHelper.compareDocs(base.artifactsDir + "StructuredDocumentTag.customXml.docx",
      base.goldsDir + "StructuredDocumentTag.customXml Gold.docx")).toEqual(true);

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.customXml.docx");
    xmlPart = doc.customXmlParts.at(0);

    expect(Guid.isValid(xmlPart.id.slice(1, -1))).toBeTruthy();
    expect(Buffer.from(xmlPart.data).toString("utf8")).toEqual("<root><text>Hello world!</text></root>");
    expect(xmlPart.schemas.at(0)).toEqual("http://www.w3.org/2001/XMLSchema");

    tag = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    expect(tag.getText().trim()).toEqual("Hello world!");
    expect(tag.xmlMapping.xpath).toEqual("/root[1]/text[1]");
    expect(tag.xmlMapping.prefixMappings).toEqual('');
    expect(tag.xmlMapping.customXmlPart.dataChecksum).toEqual(xmlPart.dataChecksum);
  });


  test('DataChecksum', () => {
    //ExStart
    //ExFor:CustomXmlPart.dataChecksum
    //ExSummary:Shows how the checksum is calculated in a runtime.
    let doc = new aw.Document();

    let richText = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RichText, aw.Markup.MarkupLevel.Block);
    doc.firstSection.body.appendChild(richText);

    // The checksum is read-only and computed using the data of the corresponding custom XML data part.
    richText.xmlMapping.setMapping(doc.customXmlParts.add(Guid.newGuid().toString(),
      "<root><text>ContentControl</text></root>"), "/root/text", "");

    let checksum = richText.xmlMapping.customXmlPart.dataChecksum;
    console.log(checksum);

    richText.xmlMapping.setMapping(doc.customXmlParts.add(Guid.newGuid().toString(),
      "<root><text>Updated ContentControl</text></root>"), "/root/text", "");

    let updatedChecksum = richText.xmlMapping.customXmlPart.dataChecksum;
    console.log(updatedChecksum);

    // We changed the XmlPart of the tag, and the checksum was updated at runtime.
    expect(updatedChecksum).not.toEqual(checksum);
    //ExEnd
  });


  test('XmlMapping', () => {
    //ExStart
    //ExFor:XmlMapping
    //ExFor:XmlMapping.customXmlPart
    //ExFor:XmlMapping.delete
    //ExFor:XmlMapping.isMapped
    //ExFor:XmlMapping.prefixMappings
    //ExFor:XmlMapping.xPath
    //ExSummary:Shows how to set XML mappings for custom XML parts.
    let doc = new aw.Document();

    // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
    let xmlPartId = `{${Guid.newGuid().toString()}}`;
    let xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
    let xmlPart = doc.customXmlParts.add(xmlPartId, xmlPartContent);

    expect(Buffer.from(xmlPart.data).toString("utf8")).toEqual("<root><text>Text element #1</text><text>Text element #2</text></root>");

    // Create a structured document tag that will display the contents of our CustomXmlPart.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Block);

    // Set a mapping for our structured document tag. This mapping will instruct
    // our structured document tag to display a portion of the XML part's text contents that the XPath points to.
    // In this case, it will be contents of the the second "<text>" element of the first "<root>" element: "Text element #2".
    tag.xmlMapping.setMapping(xmlPart, "/root[1]/text[2]", "xmlns:ns='http://www.w3.org/2001/XMLSchema'");

    expect(tag.xmlMapping.isMapped).toEqual(true);
    expect(tag.xmlMapping.customXmlPart).toEqual(xmlPart);
    expect(tag.xmlMapping.xpath).toEqual("/root[1]/text[2]");
    expect(tag.xmlMapping.prefixMappings).toEqual("xmlns:ns='http://www.w3.org/2001/XMLSchema'");

    // Add the structured document tag to the document to display the content from our custom part.
    doc.firstSection.body.appendChild(tag);
    doc.save(base.artifactsDir + "StructuredDocumentTag.xmlMapping.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.xmlMapping.docx");
    xmlPart = doc.customXmlParts.at(0);

    expect(Guid.isValid(xmlPart.id.slice(1, -1))).toBeTruthy();
    expect(Buffer.from(xmlPart.data).toString("utf8")).toEqual("<root><text>Text element #1</text><text>Text element #2</text></root>");

    tag = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    expect(tag.getText().trim()).toEqual("Text element #2");
    expect(tag.xmlMapping.xpath).toEqual("/root[1]/text[2]");
    expect(tag.xmlMapping.prefixMappings).toEqual("xmlns:ns='http://www.w3.org/2001/XMLSchema'");
  });


  test('StructuredDocumentTagRangeStartXmlMapping', () => {
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart.xmlMapping
    //ExSummary:Shows how to set XML mappings for the range start of a structured document tag.
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");

    // Construct an XML part that contains text and add it to the document's CustomXmlPart collection.
    let xmlPartId = `{${Guid.newGuid().toString()}}`;
    let xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
    let xmlPart = doc.customXmlParts.add(xmlPartId, xmlPartContent);

    expect(Buffer.from(xmlPart.data).toString("utf8")).toEqual("<root><text>Text element #1</text><text>Text element #2</text></root>");

    // Create a structured document tag that will display the contents of our CustomXmlPart in the document.
    let sdtRangeStart = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, true).asStructuredDocumentTagRangeStart();

    // If we set a mapping for our structured document tag,
    // it will only display a portion of the CustomXmlPart that the XPath points to.
    // This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
    sdtRangeStart.xmlMapping.setMapping(xmlPart, "/root[1]/text[2]", null);

    doc.save(base.artifactsDir + "StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.StructuredDocumentTagRangeStartXmlMapping.docx");
    xmlPart = doc.customXmlParts.at(0);

    expect(Guid.isValid(xmlPart.id.slice(1, -1))).toBeTruthy();
    expect(Buffer.from(xmlPart.data).toString("utf8")).toEqual("<root><text>Text element #1</text><text>Text element #2</text></root>");

    sdtRangeStart = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, true).asStructuredDocumentTagRangeStart();
    expect(sdtRangeStart.xmlMapping.xpath).toEqual("/root[1]/text[2]");
  });


  test('CustomXmlSchemaCollection', () => {
    //ExStart
    //ExFor:CustomXmlSchemaCollection
    //ExFor:CustomXmlSchemaCollection.add(String)
    //ExFor:CustomXmlSchemaCollection.clear
    //ExFor:CustomXmlSchemaCollection.clone
    //ExFor:CustomXmlSchemaCollection.count
    //ExFor:CustomXmlSchemaCollection.getEnumerator
    //ExFor:CustomXmlSchemaCollection.indexOf(String)
    //ExFor:CustomXmlSchemaCollection.item(Int32)
    //ExFor:CustomXmlSchemaCollection.remove(String)
    //ExFor:CustomXmlSchemaCollection.removeAt(Int32)
    //ExSummary:Shows how to work with an XML schema collection.
    let doc = new aw.Document();

    let xmlPartId = `{${Guid.newGuid().toString()}}`;
    let xmlPartContent = "<root><text>Hello, World!</text></root>";
    let xmlPart = doc.customXmlParts.add(xmlPartId, xmlPartContent);

    // Add an XML schema association.
    xmlPart.schemas.add("http://www.w3.org/2001/XMLSchema");

    // Clone the custom XML part's XML schema association collection,
    // and then add a couple of new schemas to the clone.
    let schemas = xmlPart.schemas.clone();
    schemas.add("http://www.w3.org/2001/XMLSchema-instance");
    schemas.add("http://schemas.microsoft.com/office/2006/metadata/contentType");

    expect(schemas.count).toEqual(3);
    expect(schemas.indexOf("http://schemas.microsoft.com/office/2006/metadata/contentType")).toEqual(2);

    // Enumerate the schemas and print each element.
    for (let schema of schemas) {
      console.log(schema);
    }

    // Below are three ways of removing schemas from the collection.
    // 1 -  Remove a schema by index:
    schemas.removeAt(2);

    // 2 -  Remove a schema by value:
    schemas.remove("http://www.w3.org/2001/XMLSchema");

    // 3 -  Use the "Clear" method to empty the collection at once.
    schemas.clear();

    expect(schemas.count).toEqual(0);
    //ExEnd
  });


  test('CustomXmlPartStoreItemIdReadOnly', () => {
    //ExStart
    //ExFor:XmlMapping.storeItemId
    //ExSummary:Shows how to get the custom XML data identifier of an XML part.
    let doc = new aw.Document(base.myDir + "Custom XML part in structured document tag.docx");

    // Structured document tags have IDs in the form of GUIDs.
    let tag = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();

    expect(tag.xmlMapping.storeItemId).toEqual("{F3029283-4FF8-4DD2-9F31-395F19ACEE85}");
    //ExEnd
  });


  test('CustomXmlPartStoreItemIdReadOnlyNull', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let sdtCheckBox = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Checkbox, aw.Markup.MarkupLevel.Inline);
    sdtCheckBox.checked = true;

    builder.insertNode(sdtCheckBox);

    doc = DocumentHelper.saveOpen(doc);

    let sdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();
    console.log("The Id of your custom xml part is: " + sdt.xmlMapping.storeItemId);
  });


  test('ClearTextFromStructuredDocumentTags', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.clear
    //ExSummary:Shows how to delete contents of structured document tag elements.
    let doc = new aw.Document();

    // Create a plain text structured document tag, and then append it to the document.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Block);
    doc.firstSection.body.appendChild(tag);

    // This structured document tag, which is in the form of a text box, already displays placeholder text.
    expect(tag.getText().trim()).toEqual("Click here to enter text.");
    expect(tag.isShowingPlaceholderText).toEqual(true);

    // Create a building block with text contents.
    let glossaryDoc = doc.glossaryDocument;
    let substituteBlock = new aw.BuildingBlocks.BuildingBlock(glossaryDoc);
    substituteBlock.name = "My placeholder";
    substituteBlock.appendChild(new aw.Section(glossaryDoc));
    substituteBlock.firstSection.ensureMinimum();
    substituteBlock.firstSection.body.firstParagraph.appendChild(new aw.Run(glossaryDoc, "Custom placeholder text."));
    glossaryDoc.appendChild(substituteBlock);

    // Set the structured document tag's "PlaceholderName" property to our building block's name to get
    // the structured document tag to display the contents of the building block in place of the original default text.
    tag.placeholderName = "My placeholder";

    expect(tag.getText().trim()).toEqual("Custom placeholder text.");
    expect(tag.isShowingPlaceholderText).toEqual(true);

    // Edit the text of the structured document tag and hide the placeholder text.
    let run = tag.getChild(aw.NodeType.Run, 0, true).asRun();
    run.text = "New text.";
    tag.isShowingPlaceholderText = false;

    expect(tag.getText().trim()).toEqual("New text.");

    // Use the "Clear" method to clear this structured document tag's contents and display the placeholder again.
    tag.clear();

    expect(tag.isShowingPlaceholderText).toEqual(true);
    expect(tag.getText().trim()).toEqual("Custom placeholder text.");
    //ExEnd
  });


  test('AccessToBuildingBlockPropertiesFromDocPartObjSdt', () => {
    let doc = new aw.Document(base.myDir + "Structured document tags with building blocks.docx");

    let docPartObjSdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();

    expect(docPartObjSdt.sdtType).toEqual(aw.Markup.SdtType.DocPartObj);
    expect(docPartObjSdt.buildingBlockGallery).toEqual("Table of Contents");
  });


  test('AccessToBuildingBlockPropertiesFromPlainTextSdt', () => {
    let doc = new aw.Document(base.myDir + "Structured document tags with building blocks.docx");

    let plainTextSdt = doc.getChild(aw.NodeType.StructuredDocumentTag, 1, true).asStructuredDocumentTag();

    expect(plainTextSdt.sdtType).toEqual(aw.Markup.SdtType.PlainText);
    expect(() => {let _ = plainTextSdt.buildingBlockGallery; }).toThrow("BuildingBlockType is only accessible for BuildingBlockGallery SDT type.");
  });


  test('BuildingBlockCategories', () => {
    //ExStart
    //ExFor:StructuredDocumentTag.buildingBlockCategory
    //ExFor:StructuredDocumentTag.buildingBlockGallery
    //ExSummary:Shows how to insert a structured document tag as a building block, and set its category and gallery.
    let doc = new aw.Document();

    let buildingBlockSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.BuildingBlockGallery, aw.Markup.MarkupLevel.Block);
    buildingBlockSdt.buildingBlockCategory = "Built-in";
    buildingBlockSdt.buildingBlockGallery = "Table of Contents";

    doc.firstSection.body.appendChild(buildingBlockSdt);

    doc.save(base.artifactsDir + "StructuredDocumentTag.BuildingBlockCategories.docx");
    //ExEnd

    buildingBlockSdt = doc.firstSection.body.getChild(aw.NodeType.StructuredDocumentTag, 0, true).asStructuredDocumentTag();

    expect(buildingBlockSdt.sdtType).toEqual(aw.Markup.SdtType.BuildingBlockGallery);
    expect(buildingBlockSdt.buildingBlockGallery).toEqual("Table of Contents");
    expect(buildingBlockSdt.buildingBlockCategory).toEqual("Built-in");
  });


  test('UpdateSdtContent', () => {
    let doc = new aw.Document();

    // Insert a drop-down list structured document tag.
    let tag = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.DropDownList, aw.Markup.MarkupLevel.Block);
    tag.listItems.add(new aw.Markup.SdtListItem("Value 1"));
    tag.listItems.add(new aw.Markup.SdtListItem("Value 2"));
    tag.listItems.add(new aw.Markup.SdtListItem("Value 3"));

    // The drop-down list currently displays "Choose an item" as the default text.
    // Set the "SelectedValue" property to one of the list items to get the tag to
    // display that list item's value instead of the default text.
    tag.listItems.selectedValue = tag.listItems.at(1);

    doc.firstSection.body.appendChild(tag);

    doc.save(base.artifactsDir + "StructuredDocumentTag.UpdateSdtContent.pdf");
  });


  test('FillTableUsingRepeatingSectionItem', () => {
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to fill a table with data from in an XML part.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let xmlPart = doc.customXmlParts.add("Books",
      "<books>" +
        "<book>" +
          "<title>Everyday Italian</title>" +
          "<author>Giada De Laurentiis</author>" +
        "</book>" +
        "<book>" +
          "<title>The C Programming Language</title>" +
          "<author>Brian W. Kernighan, Dennis M. Ritchie</author>" +
        "</book>" +
        "<book>" +
          "<title>Learning XML</title>" +
          "<author>Erik T. Ray</author>" +
        "</book>" +
      "</books>");

    // Create headers for data from the XML content.
    let table = builder.startTable();
    builder.insertCell();
    builder.write("Title");
    builder.insertCell();
    builder.write("Author");
    builder.endRow();
    builder.endTable();

    // Create a table with a repeating section inside.
    let repeatingSectionSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RepeatingSection, aw.Markup.MarkupLevel.Row);
    repeatingSectionSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book", '');
    table.appendChild(repeatingSectionSdt);

    // Add repeating section item inside the repeating section and mark it as a row.
    // This table will have a row for each element that we can find in the XML document
    // using the "/books[1]/book" XPath, of which there are three.
    let repeatingSectionItemSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.RepeatingSectionItem, aw.Markup.MarkupLevel.Row);
    repeatingSectionSdt.appendChild(repeatingSectionItemSdt);

    let row = new aw.Tables.Row(doc);
    repeatingSectionItemSdt.appendChild(row);

    // Map XML data with created table cells for the title and author of each book.
    let titleSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Cell);
    titleSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book[1]/title[1]", '');
    row.appendChild(titleSdt);

    let authorSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Cell);
    authorSdt.xmlMapping.setMapping(xmlPart, "/books[1]/book[1]/author[1]", '');
    row.appendChild(authorSdt);

    doc.save(base.artifactsDir + "StructuredDocumentTag.repeatingSectionItem.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.repeatingSectionItem.docx");
    let tags = doc.getChildNodes(aw.NodeType.StructuredDocumentTag, true).toArray().map(node => node.asStructuredDocumentTag());

    expect(tags.at(0).xmlMapping.xpath).toEqual("/books[1]/book");
    expect(tags.at(0).xmlMapping.prefixMappings).toEqual('');

    expect(tags.at(1).xmlMapping.xpath).toEqual('');
    expect(tags.at(1).xmlMapping.prefixMappings).toEqual('');

    expect(tags.at(2).xmlMapping.xpath).toEqual("/books[1]/book[1]/title[1]");
    expect(tags.at(2).xmlMapping.prefixMappings).toEqual('');

    expect(tags.at(3).xmlMapping.xpath).toEqual("/books[1]/book[1]/author[1]");
    expect(tags.at(3).xmlMapping.prefixMappings).toEqual('');

    expect(doc.firstSection.body.tables.at(0).getText().trim()).toEqual("Title\u0007Author\u0007\u0007" +
                            "Everyday Italian\u0007Giada De Laurentiis\u0007\u0007" +
                            "The C Programming Language\u0007Brian W. Kernighan, Dennis M. Ritchie\u0007\u0007" +
                            "Learning XML\u0007Erik T. Ray\u0007\u0007");
  });


  test('CustomXmlPart', () => {
    let xmlString =
    "<?xml version=\"1.0\"?>" +
    "<Company>" +
      "<Employee id=\"1\">" +
        "<FirstName>John</FirstName>" +
        "<LastName>Doe</LastName>" +
      "</Employee>" +
      "<Employee id=\"2\">" +
        "<FirstName>Jane</FirstName>" +
        "<LastName>Doe</LastName>" +
      "</Employee>" +
    "</Company>";

    let doc = new aw.Document();

    // Insert the full XML document as a custom document part.
    // We can find the mapping for this part in Microsoft Word via "Developer" -> "XML Mapping Pane", if it is enabled.
    let xmlPart = doc.customXmlParts.add(`{${Guid.newGuid().toString()}}`, xmlString);

    // Create a structured document tag, which will use an XPath to refer to a single element from the XML.
    let sdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.PlainText, aw.Markup.MarkupLevel.Block);
    sdt.xmlMapping.setMapping(xmlPart, "Company//Employee[@id='2']/FirstName", "");

    // Add the StructuredDocumentTag to the document to display the element in the text.
    doc.firstSection.body.appendChild(sdt);
  });


  test('MultiSectionTags', () => {
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart
    //ExFor:IStructuredDocumentTag.id
    //ExFor:StructuredDocumentTagRangeStart.id
    //ExFor:StructuredDocumentTagRangeStart.title
    //ExFor:StructuredDocumentTagRangeStart.placeholderName
    //ExFor:StructuredDocumentTagRangeStart.isShowingPlaceholderText
    //ExFor:StructuredDocumentTagRangeStart.lockContentControl
    //ExFor:StructuredDocumentTagRangeStart.lockContents
    //ExFor:IStructuredDocumentTag.level
    //ExFor:StructuredDocumentTagRangeStart.level
    //ExFor:StructuredDocumentTagRangeStart.rangeEnd
    //ExFor:IStructuredDocumentTag.color
    //ExFor:StructuredDocumentTagRangeStart.color
    //ExFor:StructuredDocumentTagRangeStart.sdtType
    //ExFor:StructuredDocumentTagRangeStart.wordOpenXML
    //ExFor:StructuredDocumentTagRangeStart.tag
    //ExFor:StructuredDocumentTagRangeEnd
    //ExFor:StructuredDocumentTagRangeEnd.id
    //ExSummary:Shows how to get the properties of multi-section structured document tags.
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");

    let rangeStartTag = doc.getChildNodes(aw.NodeType.StructuredDocumentTagRangeStart, true).at(0).asStructuredDocumentTagRangeStart();
    let rangeEndTag = doc.getChildNodes(aw.NodeType.StructuredDocumentTagRangeEnd, true).at(0).asStructuredDocumentTagRangeEnd();

    expect(rangeEndTag.id).toEqual(rangeStartTag.id);
    expect(rangeStartTag.nodeType).toEqual(aw.NodeType.StructuredDocumentTagRangeStart);
    expect(rangeEndTag.nodeType).toEqual(aw.NodeType.StructuredDocumentTagRangeEnd);

    console.log("StructuredDocumentTagRangeStart values:");
    console.log(`\t|Id: ${rangeStartTag.id}`);
    console.log(`\t|Title: ${rangeStartTag.title}`);
    console.log(`\t|PlaceholderName: ${rangeStartTag.placeholderName}`);
    console.log(`\t|IsShowingPlaceholderText: ${rangeStartTag.isShowingPlaceholderText}`);
    console.log(`\t|LockContentControl: ${rangeStartTag.lockContentControl}`);
    console.log(`\t|LockContents: ${rangeStartTag.lockContents}`);
    console.log(`\t|Level: ${rangeStartTag.level}`);
    console.log(`\t|NodeType: ${rangeStartTag.nodeType}`);
    console.log(`\t|RangeEnd.NodeType: ${rangeStartTag.rangeEnd.nodeType}`);
    console.log(`\t|Color: ${rangeStartTag.color}`);
    console.log(`\t|SdtType: ${rangeStartTag.sdtType}`);
    console.log(`\t|FlatOpcContent: ${rangeStartTag.wordOpenXML}`);
    console.log(`\t|Tag: ${rangeStartTag.tag}\n`);

    console.log("StructuredDocumentTagRangeEnd values:");
    console.log(`\t|Id: ${rangeEndTag.id}`);
    console.log(`\t|NodeType: ${rangeEndTag.nodeType}`);
    //ExEnd
  });


  test('SdtChildNodes', () => {
    //ExStart
    //ExFor:StructuredDocumentTagRangeStart.getChildNodes(NodeType, bool)
    //ExSummary:Shows how to get child nodes of StructuredDocumentTagRangeStart.
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");
    let tag = doc.getChildNodes(aw.NodeType.StructuredDocumentTagRangeStart, true).at(0).asStructuredDocumentTagRangeStart();

    console.log("StructuredDocumentTagRangeStart values:");
    console.log(`\t|Child nodes count: ${tag.getChildNodes(aw.NodeType.Any, false).count}\n`);

    tag.getChildNodes(aw.NodeType.Any, false).toArray().forEach(node => { console.log(`\t|Child node type: ${node.nodeType}`); });
    tag.getChildNodes(aw.NodeType.Run, true).toArray().forEach(node => { console.log(`\t|Child node text: ${node.getText()}`); });
    //ExEnd
  });


  //ExStart
  //ExFor:StructuredDocumentTagRangeStart.#ctor(DocumentBase, SdtType)
  //ExFor:StructuredDocumentTagRangeEnd.#ctor(DocumentBase, int)
  //ExFor:StructuredDocumentTagRangeStart.RemoveSelfOnly
  //ExFor:StructuredDocumentTagRangeStart.RemoveAllChildren
  //ExSummary:Shows how to create/remove structured document tag and its content.
  test('SdtRangeExtendedMethods', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("StructuredDocumentTag element");

    let rangeStart = insertStructuredDocumentTagRanges(doc);

    // Removes ranged structured document tag, but keeps content inside.
    rangeStart.removeSelfOnly();

    rangeStart = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, false);
    expect(rangeStart).toEqual(null);

    let rangeEnd = doc.getChild(aw.NodeType.StructuredDocumentTagRangeEnd, 0, false);
    expect(rangeEnd).toEqual(null);
    expect(doc.getText().trim()).toEqual("StructuredDocumentTag element");

    rangeStart = insertStructuredDocumentTagRanges(doc);

    let paragraphNode = lastOrDefault(rangeStart);
    expect(paragraphNode).not.toBeNull();
    expect(paragraphNode.getText().trim()).toEqual("StructuredDocumentTag element");

    // Removes ranged structured document tag and content inside.
    rangeStart.removeAllChildren();

    paragraphNode = lastOrDefault(rangeStart);
    expect(paragraphNode).toBeNull();
  });

  function lastOrDefault(enumeratedObject) {
    let children = Array.from(enumeratedObject);
    return children.length == 0 ? null : children.at(-1);
  }

  function insertStructuredDocumentTagRanges(doc) {
    let rangeStart = new aw.Markup.StructuredDocumentTagRangeStart(doc, aw.Markup.SdtType.PlainText);
    let rangeEnd = new aw.Markup.StructuredDocumentTagRangeEnd(doc, rangeStart.id);

    doc.firstSection.body.insertBefore(rangeStart, doc.firstSection.body.firstParagraph);
    doc.lastSection.body.insertAfter(rangeEnd, doc.firstSection.body.firstParagraph);

    return rangeStart;
  }
  //ExEnd

  test('GetSdt', () => {
    //ExStart
    //ExFor:Range.structuredDocumentTags
    //ExFor:StructuredDocumentTagCollection.remove(int)
    //ExFor:StructuredDocumentTagCollection.removeAt(int)
    //ExSummary:Shows how to remove structured document tag.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    let structuredDocumentTags = doc.range.structuredDocumentTags;
    for (let sdt of structuredDocumentTags) {
      console.log(sdt.title);
    }

    let sdt = structuredDocumentTags.getById(1691867797);
    expect(sdt.id).toEqual(1691867797);

    expect(structuredDocumentTags.count).toEqual(5);
    // Remove the structured document tag by Id.
    structuredDocumentTags.remove(1691867797);
    // Remove the structured document tag at position 0.
    structuredDocumentTags.removeAt(0);
    expect(structuredDocumentTags.count).toEqual(3);
    //ExEnd
  });


  test('RangeSdt', () => {
    //ExStart
    //ExFor:StructuredDocumentTagCollection
    //ExFor:StructuredDocumentTagCollection.getById(int)
    //ExFor:StructuredDocumentTagCollection.getByTitle(String)
    //ExFor:IStructuredDocumentTag.isMultiSection
    //ExFor:IStructuredDocumentTag.title
    //ExSummary:Shows how to get structured document tag.
    let doc = new aw.Document(base.myDir + "Structured document tags by id.docx");

    // Get the structured document tag by Id.
    let sdt = doc.range.structuredDocumentTags.getById(1160505028);
    console.log(sdt.isMultiSection);
    console.log(sdt.title);

    // Get the structured document tag or ranged tag by Title.
    sdt = doc.range.structuredDocumentTags.getByTitle("Alias4");
    console.log(sdt.id);
    //ExEnd
  });


  test('SdtAtRowLevel', () => {
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to create group structured document tag at the Row level.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();

    // Create a Group structured document tag at the Row level.
    let groupSdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Group, aw.Markup.MarkupLevel.Row);
    table.appendChild(groupSdt);
    groupSdt.isShowingPlaceholderText = false;
    groupSdt.removeAllChildren();

    // Create a child row of the structured document tag.
    let row = new aw.Tables.Row(doc);
    groupSdt.appendChild(row);

    let cell = new aw.Tables.Cell(doc);
    row.appendChild(cell);

    builder.endTable();

    // Insert cell contents.
    cell.ensureMinimum();
    builder.moveTo(cell.lastParagraph);
    builder.write("Lorem ipsum dolor.");

    // Insert text after the table.
    builder.moveTo(table.nextSibling);
    builder.write("Nulla blandit nisi.");

    doc.save(base.artifactsDir + "StructuredDocumentTag.SdtAtRowLevel.docx");
    //ExEnd
  });


  test('IgnoreStructuredDocumentTags', () => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreStructuredDocumentTags
    //ExSummary:Shows how to ignore content of tags from replacement.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    // This paragraph contains SDT.
    let p = doc.firstSection.body.getChild(aw.NodeType.Paragraph, 2, true).asParagraph();
    let textToSearch = p.toString(aw.SaveFormat.Text).trim();

    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreStructuredDocumentTags = true;
    doc.range.replace(textToSearch, "replacement", options);

    doc.save(base.artifactsDir + "StructuredDocumentTag.ignoreStructuredDocumentTags.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "StructuredDocumentTag.ignoreStructuredDocumentTags.docx");
    expect(doc.getText().trim()).toEqual("This document contains Structured Document Tags with text inside them\r\rRepeatingSection\rRichText\rreplacement");
  });


  test('Citation', () => {
    //ExStart
    //ExFor:SdtType
    //ExSummary:Shows how to create a structured document tag of the Citation type.
    let doc = new aw.Document();

    let sdt = new aw.Markup.StructuredDocumentTag(doc, aw.Markup.SdtType.Citation, aw.Markup.MarkupLevel.Inline);
    let paragraph = doc.firstSection.body.firstParagraph;
    paragraph.appendChild(sdt);

    // Create a Citation field.
    let builder = new aw.DocumentBuilder(doc);
    builder.moveToParagraph(0, -1);
    builder.insertField(String.raw`CITATION Ath22 \l 1033 `, "(John Lennon, 2022)");

    // Move the field to the structured document tag.
    while (sdt.nextSibling != null)
      sdt.appendChild(sdt.nextSibling);

    doc.save(base.artifactsDir + "StructuredDocumentTag.citation.docx");
    //ExEnd
  });


  test('RangeStartWordOpenXmlMinimal', () => {
    //ExStart:RangeStartWordOpenXmlMinimal
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:StructuredDocumentTagRangeStart.wordOpenXMLMinimal
    //ExSummary:Shows how to get minimal XML contained within the node in the FlatOpc format.
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");
    let tag = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, true).asStructuredDocumentTagRangeStart();

    expect(tag.wordOpenXMLMinimal.includes("<pkg:part pkg:name=\"/docProps/app.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\">"))
      .toEqual(true);
    expect(tag.wordOpenXMLMinimal.includes("xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\"")).toEqual(false);
    //ExEnd:RangeStartWordOpenXmlMinimal
  });


  test('RemoveSelfOnly', () => {
    //ExStart:RemoveSelfOnly
    //GistId:e386727403c2341ce4018bca370a5b41
    //ExFor:IStructuredDocumentTag
    //ExFor:IStructuredDocumentTag.getChildNodes(NodeType, bool)
    //ExFor:IStructuredDocumentTag.removeSelfOnly
    //ExSummary:Shows how to remove structured document tag, but keeps content inside.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");

    // This collection provides a unified interface for accessing ranged and non-ranged structured tags. 
    let sdts = doc.range.structuredDocumentTags;
    expect(sdts.count).toEqual(5);

    // Here we can get child nodes from the common interface of ranged and non-ranged structured tags.
    for (let sdt of sdts)
    {
      if (sdt.getChildNodes(aw.NodeType.Any, false).count > 0)
        sdt.removeSelfOnly();
    }

    sdts = doc.range.structuredDocumentTags;
    expect(sdts.count).toEqual(0);
    //ExEnd:RemoveSelfOnly
  });


  test('Appearance', () => {
    //ExStart:Appearance
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:SdtAppearance
    //ExFor:StructuredDocumentTagRangeStart.appearance
    //ExFor:IStructuredDocumentTag.appearance
    //ExSummary:Shows how to show tag around content.
    let doc = new aw.Document(base.myDir + "Multi-section structured document tags.docx");
    let tag = doc.getChild(aw.NodeType.StructuredDocumentTagRangeStart, 0, true).asStructuredDocumentTagRangeStart();

    if (tag.appearance == aw.Markup.SdtAppearance.Hidden)
      tag.appearance = aw.Markup.SdtAppearance.Tags;
    //ExEnd:Appearance
  });


  test('InsertStructuredDocumentTag', () => {
    //ExStart:InsertStructuredDocumentTag
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:DocumentBuilder.insertStructuredDocumentTag(SdtType)
    //ExSummary:Shows how to simply insert structured document tag.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    let builder = new aw.DocumentBuilder(doc);

    builder.moveTo(doc.firstSection.body.paragraphs.at(3));
    // Note, that only following StructuredDocumentTag types are allowed for insertion:
    // SdtType.PlainText, SdtType.RichText, SdtType.Checkbox, SdtType.DropDownList,
    // SdtType.ComboBox, SdtType.Picture, SdtType.date.
    // Markup level of inserted StructuredDocumentTag will be detected automatically and depends on position being inserted at.
    // Added StructuredDocumentTag will inherit paragraph and font formatting from cursor position.
    let sdtPlain = builder.insertStructuredDocumentTag(aw.Markup.SdtType.PlainText);

    doc.save(base.artifactsDir + "StructuredDocumentTag.insertStructuredDocumentTag.docx");
    //ExEnd:InsertStructuredDocumentTag
  });


});
