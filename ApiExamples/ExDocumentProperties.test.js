// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const fs = require('fs');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');

/// <summary>
/// Counts the lines in a document.
/// Traverses the document's layout entities tree upon construction,
/// counting entities of the "Line" type that also contain real text.
/// </summary>
class LineCounter {
  #layoutEnumerator = null;
  #lineCount = 0;
  #scanningLineForRealText = false;

  constructor(doc) {
    this.#layoutEnumerator = new aw.Layout.LayoutEnumerator(doc);
    this.countLines();
  }

  getLineCount() {
    return this.#lineCount;
  }

  countLines() {
    do {
      if (this.#layoutEnumerator.type == aw.Layout.LayoutEntityType.Line) {
        this.#scanningLineForRealText = true;
      }
      if (this.#layoutEnumerator.moveFirstChild()) {
        if (this.#scanningLineForRealText && this.#layoutEnumerator.kind.startsWith("TEXT")) {
          this.#lineCount++;
          this.#scanningLineForRealText = false;
        }
        this.countLines();
        this.#layoutEnumerator.moveParent();
      }
    } while (this.#layoutEnumerator.moveNext());
  }
}


function testContent(doc) {
  let properties = doc.builtInDocumentProperties;

  expect(properties.pages).toEqual(6);

  expect(properties.words).toEqual(1035);
  expect(properties.characters).toEqual(6026);
  expect(properties.charactersWithSpaces).toEqual(7041);
  expect(properties.lines).toEqual(142);
  expect(properties.paragraphs).toEqual(29);
  expect(properties.bytes).toBeLessThan(15500 + 200);
  expect(properties.bytes).toBeGreaterThan(15500 - 200);
  expect(properties.template).toEqual(base.myDir.replace("\\\\", "\\") + "Business brochure.dotx");
  expect(properties.contentStatus).toEqual("Draft");
  expect(properties.contentType).toEqual('');
  expect(properties.linksUpToDate).toEqual(false);
}


describe("ExDocumentProperties", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  beforeEach(() => {
    base.setUnlimitedLicense();
  });


  test('BuiltIn', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties
    //ExFor:Document.builtInDocumentProperties
    //ExFor:Document.customDocumentProperties
    //ExFor:DocumentProperty
    //ExFor:DocumentProperty.name
    //ExFor:DocumentProperty.value
    //ExFor:DocumentProperty.type
    //ExSummary:Shows how to work with built-in document properties.
    let doc = new aw.Document(base.myDir + "Properties.docx");

    // The "Document" object contains some of its metadata in its members.
    //console.log(`Document filename:\n\t \"${doc.originalFileName}\"`);

    // The document also stores metadata in its built-in properties.
    // Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
    /*console.log("Built-in Properties:");
    for (let docProperty of doc.builtInDocumentProperties)
    {
      console.log(docProperty.name);
      console.log(`\tType:\t${docProperty.type}`);
      console.log(`\tValue:\t\"${docProperty.toString()}\"`);
    }*/
    //ExEnd

    expect(doc.builtInDocumentProperties.count).toEqual(31);
  });


  test('Custom', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.item(String)
    //ExFor:CustomDocumentProperties
    //ExFor:DocumentProperty.toString
    //ExFor:DocumentPropertyCollection.count
    //ExFor:DocumentPropertyCollection.item(int)
    //ExSummary:Shows how to work with custom document properties.
    let doc = new aw.Document(base.myDir + "Properties.docx");

    // Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
    // The document has a fixed list of built-in properties. The user creates all of the custom properties. 
    expect(doc.customDocumentProperties.at("CustomProperty").toString()).toEqual("Value of custom document property");

    doc.customDocumentProperties.add("CustomProperty2", "Value of custom document property #2");

    /*console.log("Custom Properties:");
    for (let customDocumentProperty of doc.customDocumentProperties)
    {
      console.log(customDocumentProperty.name);
      console.log(`\tType:\t${customDocumentProperty.type}`);
      console.log(`\tValue:\t\"${customDocumentProperty.toString()}\"`);
    }*/
    //ExEnd

    expect(doc.customDocumentProperties.count).toEqual(2);
  });


  test('Description', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.author
    //ExFor:BuiltInDocumentProperties.category
    //ExFor:BuiltInDocumentProperties.comments
    //ExFor:BuiltInDocumentProperties.keywords
    //ExFor:BuiltInDocumentProperties.subject
    //ExFor:BuiltInDocumentProperties.title
    //ExSummary:Shows how to work with built-in document properties in the "Description" category.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let properties = doc.builtInDocumentProperties;

    // Below are four built-in document properties that have fields that can display their values in the document body.
    // 1 -  "Author" property, which we can display using an AUTHOR field:
    properties.author = "John Doe";
    builder.write("Author:\t");
    builder.insertField(aw.Fields.FieldType.FieldAuthor, true);

    // 2 -  "Title" property, which we can display using a TITLE field:
    properties.title = "John's Document";
    builder.write("\nDoc title:\t");
    builder.insertField(aw.Fields.FieldType.FieldTitle, true);

    // 3 -  "Subject" property, which we can display using a SUBJECT field:
    properties.subject = "My subject";
    builder.write("\nSubject:\t");
    builder.insertField(aw.Fields.FieldType.FieldSubject, true);

    // 4 -  "Comments" property, which we can display using a COMMENTS field:
    properties.comments = `This is ${properties.author}'s document about ${properties.subject}`;
    builder.write("\nComments:\t\"");
    builder.insertField(aw.Fields.FieldType.FieldComments, true);
    builder.write("\"");

    // The "Category" built-in property does not have a field that can display its value.
    properties.category = "My category";

    // We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
    properties.keywords = "Tag 1; Tag 2; Tag 3";

    // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
    // The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
    doc.save(base.artifactsDir + "DocumentProperties.description.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentProperties.description.docx");

    properties = doc.builtInDocumentProperties;

    expect(properties.author).toEqual("John Doe");
    expect(properties.category).toEqual("My category");
    expect(properties.comments).toEqual(`This is ${properties.author}'s document about ${properties.subject}`);
    expect(properties.keywords).toEqual("Tag 1; Tag 2; Tag 3");
    expect(properties.subject).toEqual("My subject");
    expect(properties.title).toEqual("John's Document");
    expect(doc.getText().trim()).toEqual("Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r" +
                            "Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" +
                            "Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" +
                            "Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\"");
  });


  test('Origin', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.company
    //ExFor:BuiltInDocumentProperties.createdTime
    //ExFor:BuiltInDocumentProperties.lastPrinted
    //ExFor:BuiltInDocumentProperties.lastSavedBy
    //ExFor:BuiltInDocumentProperties.lastSavedTime
    //ExFor:BuiltInDocumentProperties.manager
    //ExFor:BuiltInDocumentProperties.nameOfApplication
    //ExFor:BuiltInDocumentProperties.revisionNumber
    //ExFor:BuiltInDocumentProperties.template
    //ExFor:BuiltInDocumentProperties.totalEditingTime
    //ExFor:BuiltInDocumentProperties.version
    //ExSummary:Shows how to work with document properties in the "Origin" category.
    // Open a document that we have created and edited using Microsoft Word.
    let doc = new aw.Document(base.myDir + "Properties.docx");
    let properties = doc.builtInDocumentProperties;

    // The following built-in properties contain information regarding the creation and editing of this document.
    // We can right-click this document in Windows Explorer and find
    // these properties via "Properties" -> "Details" -> "Origin" category.
    // Fields such as PRINTDATE and EDITTIME can display these values in the document body.
    /*console.log(`Created using ${properties.nameOfApplication}, on ${properties.createdTime}`);
    console.log(`Minutes spent editing: ${properties.totalEditingTime}`);
    console.log(`Date/time last printed: ${properties.lastPrinted}`);
    console.log(`Template document: ${properties.template}`);*/

    // We can also change the values of built-in properties.
    properties.company = "Doe Ltd.";
    properties.manager = "Jane Doe";
    properties.version = 5;
    properties.revisionNumber++;

    // Microsoft Word updates the following properties automatically when we save the document.
    // To use these properties with Aspose.words, we will need to set values for them manually.
    properties.lastSavedBy = "John Doe";
    properties.lastSavedTime = Date.now();

    // We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
    doc.save(base.artifactsDir + "DocumentProperties.origin.docx");
    //ExEnd

    properties = new aw.Document(base.artifactsDir + "DocumentProperties.origin.docx").builtInDocumentProperties;

    expect(properties.company).toEqual("Doe Ltd.");
    expect(properties.createdTime).toEqual(new Date(2006, 3, 25, 10, 10, 0));
    expect(properties.lastPrinted).toEqual(new Date(2019, 3, 21, 10, 0, 0));

    expect(properties.lastSavedBy).toEqual("John Doe");
    TestUtil.verifyDate(Date.now(), properties.lastSavedTime, 5000);
    expect(properties.manager).toEqual("Jane Doe");
    expect(properties.nameOfApplication).toEqual("Microsoft Office Word");
    expect(properties.revisionNumber).toEqual(12);
    expect(properties.template).toEqual("Normal");
    expect(properties.totalEditingTime).toEqual(8);
    expect(properties.version).toEqual(786432);
  });


  //ExStart
  //ExFor:BuiltInDocumentProperties.Bytes
  //ExFor:BuiltInDocumentProperties.Characters
  //ExFor:BuiltInDocumentProperties.CharactersWithSpaces
  //ExFor:BuiltInDocumentProperties.ContentStatus
  //ExFor:BuiltInDocumentProperties.ContentType
  //ExFor:BuiltInDocumentProperties.Lines
  //ExFor:BuiltInDocumentProperties.LinksUpToDate
  //ExFor:BuiltInDocumentProperties.Pages
  //ExFor:BuiltInDocumentProperties.Paragraphs
  //ExFor:BuiltInDocumentProperties.Words
  //ExSummary:Shows how to work with document properties in the "Content" category.
  test('Content', () => {
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");
    let properties = doc.builtInDocumentProperties;

    // By using built in properties,
    // we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
    // These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
    // If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
    // Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
    // Page count: The PageCount property shows the page count in real time and its value can be assigned to the Pages property

    // The "Pages" property stores the page count of the document. 
    expect(properties.pages).toEqual(6);

    // The "Words", "Characters", and "CharactersWithSpaces" built-in properties also display various document statistics,
    // but we need to call the "UpdateWordCount" method on the whole document before we can expect them to contain accurate values.
    expect(properties.words).toEqual(1054);
    expect(properties.characters).toEqual(6009);
    expect(properties.charactersWithSpaces).toEqual(7049);
    doc.updateWordCount();

    expect(properties.words).toEqual(1035);
    expect(properties.characters).toEqual(6026);
    expect(properties.charactersWithSpaces).toEqual(7041);

    // Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
    let lineCounter = new LineCounter(doc);
    properties.lines = lineCounter.getLineCount();

    expect(properties.lines).toEqual(142);

    // Assign the number of Paragraph nodes in the document to the "Paragraphs" built-in property.
    properties.paragraphs = doc.getChildNodes(aw.NodeType.Paragraph, true).count;
    expect(properties.paragraphs).toEqual(29);

    // Get an estimate of the file size of our document via the "Bytes" built-in property.
    expect(properties.bytes).toEqual(20310);

    // Set a different template for our document, and then update the "Template" built-in property manually to reflect this change.
    doc.attachedTemplate = base.myDir + "Business brochure.dotx";

    expect(properties.template).toEqual("Normal");

    properties.template = doc.attachedTemplate;

    // "ContentStatus" is a descriptive built-in property.
    properties.contentStatus = "Draft";

    // Upon saving, the "ContentType" built-in property will contain the MIME type of the output save format.
    expect(properties.contentType).toEqual('');

    // If the document contains links, and they are all up to date, we can set the "LinksUpToDate" property to "true".
    expect(properties.linksUpToDate).toEqual(false);

    doc.save(base.artifactsDir + "DocumentProperties.content.docx");
    testContent(new aw.Document(base.artifactsDir + "DocumentProperties.content.docx")); //ExSkip
  });
  //ExEnd


  test('Thumbnail', async () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.thumbnail
    //ExFor:DocumentProperty.toByteArray
    //ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
    // a reader that opens that document may display the image before the first page.
    let properties = doc.builtInDocumentProperties;

    let thumbnailBytes = base.loadFileToArray(base.imageDir + "Logo.jpg");
    properties.thumbnail = thumbnailBytes;

    doc.save(base.artifactsDir + "DocumentProperties.thumbnail.epub");

    // We can extract a document's thumbnail image and save it to the local file system.
    let thumbnail = doc.builtInDocumentProperties.thumbnail;
    fs.writeFileSync(base.artifactsDir + "DocumentProperties.thumbnail.gif", Buffer.from(thumbnail));
    //ExEnd

    await TestUtil.verifyImage(400, 400, base.artifactsDir + "DocumentProperties.thumbnail.gif");            
  });


  test('HyperlinkBase', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.hyperlinkBase
    //ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a relative hyperlink to a document in the local file system named "Document.docx".
    // Clicking on the link in Microsoft Word will open the designated document, if it is available.
    builder.insertHyperlink("Relative hyperlink", "Document.docx", false);

    // This link is relative. If there is no "Document.docx" in the same folder
    // as the document that contains this link, the link will be broken.
    expect(fs.existsSync(base.artifactsDir + "Document.docx")).toEqual(false);
    doc.save(base.artifactsDir + "DocumentProperties.hyperlinkBase.BrokenLink.docx");

    // The document we are trying to link to is in a different directory to the one we are planning to save the document in.
    // We could fix links like this by putting an absolute filename in each one. 
    // Alternatively, we could provide a base link that every hyperlink with a relative filename
    // will prepend to its link when we click on it. 
    let properties = doc.builtInDocumentProperties;
    properties.hyperlinkBase = base.myDir;

    expect(fs.existsSync(properties.hyperlinkBase + doc.range.fields.at(0).asFieldHyperlink().address)).toEqual(true);

    doc.save(base.artifactsDir + "DocumentProperties.hyperlinkBase.WorkingLink.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentProperties.hyperlinkBase.BrokenLink.docx");
    properties = doc.builtInDocumentProperties;

    expect(properties.hyperlinkBase).toEqual('');

    doc = new aw.Document(base.artifactsDir + "DocumentProperties.hyperlinkBase.WorkingLink.docx");
    properties = doc.builtInDocumentProperties;

    expect(properties.hyperlinkBase).toEqual(base.myDir);
    expect(fs.existsSync(properties.hyperlinkBase + doc.range.fields.at(0).asFieldHyperlink().address)).toEqual(true);
  });


  test.skip('HeadingPairs - TODO: object[] HeadingPairs not supported.', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.headingPairs
    //ExFor:BuiltInDocumentProperties.titlesOfParts
    //ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
    let doc = new aw.Document(base.myDir + "Heading pairs and titles of parts.docx");

    // We can find the combined values of these collections via
    // "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
    // The HeadingPairs property is a collection of <string, int> pairs that
    // determines how many document parts a heading spans across.
    let headingPairs = doc.builtInDocumentProperties.headingPairs;

    // The TitlesOfParts property contains the names of parts that belong to the above headings.
    let titlesOfParts = doc.builtInDocumentProperties.titlesOfParts;

    let headingPairsIndex = 0;
    let titlesOfPartsIndex = 0;
    /*while (headingPairsIndex < headingPairs.length)
    {
      console.log(`Parts for ${headingPairs.at(headingPairsIndex++)}:`);
      let partsCount = headingPairs.at(headingPairsIndex++);

      for (let i = 0; i < partsCount; i++)
        console.log(`\t\"${titlesOfParts.at(titlesOfPartsIndex++)}\"`);
    }*/
    //ExEnd

    // There are 6 array elements designating 3 heading/part count pairs
    expect(headingPairs.length).toEqual(6);
    expect(headingPairs.at(0)).toEqual("Title");
    expect(headingPairs.at(1)).toEqual("1");
    expect(headingPairs.at(2)).toEqual("Heading 1");
    expect(headingPairs.at(3)).toEqual("5");
    expect(headingPairs.at(4)).toEqual("Heading 2");
    expect(headingPairs.at(5)).toEqual("2");

    expect(titlesOfParts.length).toEqual(8);
    // "Title"
    expect(titlesOfParts.at(0)).toEqual("");
    // "Heading 1"
    expect(titlesOfParts.at(1)).toEqual("Part1");
    expect(titlesOfParts.at(2)).toEqual("Part2");
    expect(titlesOfParts.at(3)).toEqual("Part3");
    expect(titlesOfParts.at(4)).toEqual("Part4");
    expect(titlesOfParts.at(5)).toEqual("Part5");
    // "Heading 2"
    expect(titlesOfParts.at(6)).toEqual("Part6");
    expect(titlesOfParts.at(7)).toEqual("Part7");
  });


  test('Security', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.security
    //ExFor:DocumentSecurity
    //ExSummary:Shows how to use document properties to display the security level of a document.
    let doc = new aw.Document();

    expect(doc.builtInDocumentProperties.security).toEqual(aw.Properties.DocumentSecurity.None);

    // If we configure a document to be read-only, it will display this status using the "Security" built-in property.
    doc.writeProtection.readOnlyRecommended = true;
    doc.save(base.artifactsDir + "DocumentProperties.security.readOnlyRecommended.docx");

    expect(new aw.Document(base.artifactsDir + "DocumentProperties.security.readOnlyRecommended.docx")
      .builtInDocumentProperties.security).toEqual(aw.Properties.DocumentSecurity.ReadOnlyRecommended);

    // Write-protect a document, and then verify its security level.
    doc = new aw.Document();

    expect(doc.writeProtection.isWriteProtected).toEqual(false);

    doc.writeProtection.setPassword("MyPassword");

    expect(doc.writeProtection.validatePassword("MyPassword")).toEqual(true);
    expect(doc.writeProtection.isWriteProtected).toEqual(true);

    doc.save(base.artifactsDir + "DocumentProperties.security.readOnlyEnforced.docx");
            
    expect(new aw.Document(base.artifactsDir + "DocumentProperties.security.readOnlyEnforced.docx")
      .builtInDocumentProperties.security).toEqual(aw.Properties.DocumentSecurity.ReadOnlyEnforced);

    // "Security" is a descriptive property. We can edit its value manually.
    doc = new aw.Document();

    doc.protect(aw.ProtectionType.AllowOnlyComments, "MyPassword");
    doc.builtInDocumentProperties.security = aw.Properties.DocumentSecurity.ReadOnlyExceptAnnotations;
    doc.save(base.artifactsDir + "DocumentProperties.security.readOnlyExceptAnnotations.docx");

    expect(new aw.Document(base.artifactsDir + "DocumentProperties.security.readOnlyExceptAnnotations.docx")
      .builtInDocumentProperties.security).toEqual(aw.Properties.DocumentSecurity.ReadOnlyExceptAnnotations);
    //ExEnd
  });


  test('CustomNamedAccess', () => {
    //ExStart
    //ExFor:DocumentPropertyCollection.item(String)
    //ExFor:CustomDocumentProperties.add(String,DateTime)
    //ExFor:DocumentProperty.toDateTime
    //ExSummary:Shows how to create a custom document property which contains a date and time.
    let doc = new aw.Document();

    let date = new Date(2024, 6, 9);
    doc.customDocumentProperties.add("AuthorizationDate", date);

    //console.log(`Document authorized on ${doc.customDocumentProperties.at("AuthorizationDate")}`);
    //ExEnd

    TestUtil.verifyDate(date, 
      DocumentHelper.saveOpen(doc).customDocumentProperties.at("AuthorizationDate").toDateTime(), 
      1000);
  });


  test('LinkCustomDocumentPropertiesToBookmark', () => {
    //ExStart
    //ExFor:CustomDocumentProperties.addLinkToContent(String, String)
    //ExFor:DocumentProperty.isLinkToContent
    //ExFor:DocumentProperty.linkSource
    //ExSummary:Shows how to link a custom document property to a bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("MyBookmark");
    builder.write("Hello world!");
    builder.endBookmark("MyBookmark");

    // Link a new custom property to a bookmark. The value of this property
    // will be the contents of the bookmark that it references in the "LinkSource" member.
    let customProperties = doc.customDocumentProperties;
    let customProperty = customProperties.addLinkToContent("Bookmark", "MyBookmark");

    expect(customProperty.isLinkToContent).toEqual(true);
    expect(customProperty.linkSource).toEqual("MyBookmark");
    expect(customProperty.toString()).toEqual("Hello world!");

    doc.save(base.artifactsDir + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx");
    customProperty = doc.customDocumentProperties.at("Bookmark");

    expect(customProperty.isLinkToContent).toEqual(true);
    expect(customProperty.linkSource).toEqual("MyBookmark");
    expect(customProperty.toString()).toEqual("Hello world!");
  });


  test('DocumentPropertyCollection', () => {
    //ExStart
    //ExFor:CustomDocumentProperties.add(String,String)
    //ExFor:CustomDocumentProperties.add(String,Boolean)
    //ExFor:CustomDocumentProperties.add(String,int)
    //ExFor:CustomDocumentProperties.add(String,DateTime)
    //ExFor:CustomDocumentProperties.add(String,Double)
    //ExFor:DocumentProperty.type
    //ExFor:DocumentPropertyCollection
    //ExFor:DocumentPropertyCollection.clear
    //ExFor:DocumentPropertyCollection.contains(String)
    //ExFor:DocumentPropertyCollection.getEnumerator
    //ExFor:DocumentPropertyCollection.indexOf(String)
    //ExFor:DocumentPropertyCollection.removeAt(Int32)
    //ExFor:DocumentPropertyCollection.remove
    //ExFor:PropertyType
    //ExSummary:Shows how to work with a document's custom properties.
    let doc = new aw.Document();
    let properties = doc.customDocumentProperties;

    expect(properties.count).toEqual(0);

    // Custom document properties are key-value pairs that we can add to the document.
    properties.add("Authorized", true);
    properties.add("Authorized By", "John Doe");
    properties.add("Authorized Date", Date.now());
    properties.add("Authorized Revision", doc.builtInDocumentProperties.revisionNumber);
    properties.add("Authorized Amount", 123.45);

    // The collection sorts the custom properties in alphabetic order.
    expect(properties.indexOf("Authorized Amount")).toEqual(1);
    expect(properties.count).toEqual(5);

    // Print every custom property in the document.
    /*for (let p of properties) {
        console.log(`Name: \"${p.name}\"\n\tType: \"${p.type}\"\n\tValue: \"${p.toString()}\"`);
    }*/

    // Display the value of a custom property using a DOCPROPERTY field.
    let builder = new aw.DocumentBuilder(doc);
    let field = builder.insertField(" DOCPROPERTY \"Authorized By\"").asFieldDocProperty();
    field.update();

    expect(field.result).toEqual("John Doe");

    // We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
    doc.save(base.artifactsDir + "DocumentProperties.DocumentPropertyCollection.docx");

    // Below are three ways or removing custom properties from a document.
    // 1 -  Remove by index:
    properties.removeAt(1);

    expect(properties.contains("Authorized Amount")).toEqual(false);
    expect(properties.count).toEqual(4);

    // 2 -  Remove by name:
    properties.remove("Authorized Revision");

    expect(properties.contains("Authorized Revision")).toEqual(false);
    expect(properties.count).toEqual(3);

    // 3 -  Empty the entire collection at once:
    properties.clear();

    expect(properties.count).toEqual(0);
    //ExEnd
  });


  test('PropertyTypes', () => {
    //ExStart
    //ExFor:DocumentProperty.toBool
    //ExFor:DocumentProperty.toInt
    //ExFor:DocumentProperty.toDouble
    //ExFor:DocumentProperty.toString
    //ExFor:DocumentProperty.toDateTime
    //ExSummary:Shows various type conversion methods of custom document properties.
    let doc = new aw.Document();
    let properties = doc.customDocumentProperties;

    let authDate = new Date(2024, 9, 5);
    properties.add("Authorized", true);
    properties.add("Authorized By", "John Doe");
    properties.add("Authorized Date", authDate);
    properties.add("Authorized Revision", doc.builtInDocumentProperties.revisionNumber);
    properties.add("Authorized Amount", 123.45);

    expect(properties.at("Authorized").toBool()).toEqual(true);
    expect(properties.at("Authorized By").toString()).toEqual("John Doe");
    expect(properties.at("Authorized Date").toDateTime()).toEqual(authDate);
    expect(properties.at("Authorized Revision").toInt()).toEqual(1);
    expect(properties.at("Authorized Amount").type).toEqual(aw.Properties.PropertyType.Double);
    expect(properties.at("Authorized Amount").toDouble()).toEqual(123.45);
    //ExEnd
  });


  test('ExtendedProperties', () => {
    //ExStart:ExtendedProperties
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:BuiltInDocumentProperties.scaleCrop
    //ExFor:BuiltInDocumentProperties.sharedDocument
    //ExFor:BuiltInDocumentProperties.hyperlinksChanged
    //ExSummary:Shows how to get extended properties.
    let doc = new aw.Document(base.myDir + "Extended properties.docx");
    expect(doc.builtInDocumentProperties.scaleCrop).toEqual(true);
    expect(doc.builtInDocumentProperties.sharedDocument).toEqual(true);
    expect(doc.builtInDocumentProperties.hyperlinksChanged).toEqual(true);
    //ExEnd:ExtendedProperties
  });


});
