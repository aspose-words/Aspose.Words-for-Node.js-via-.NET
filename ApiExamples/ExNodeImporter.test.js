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


describe("ExNodeImporter", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test.each([false,
    true])('KeepSourceNumbering', (keepSourceNumbering) => {
    //ExStart
    //ExFor:ImportFormatOptions.keepSourceNumbering
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
    // Open a document with a custom list numbering scheme, and then clone it.
    // Since both have the same numbering format, the formats will clash if we import one document into the other.
    let srcDoc = new aw.Document(base.myDir + "Custom list numbering.docx");
    let dstDoc = srcDoc.clone();

    // When we import the document's clone into the original and then append it,
    // then the two lists with the same list format will join.
    // If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
    // that we append to the original will carry on the numbering of the list we append it to.
    // This will effectively merge the two lists into one.
    // If we set the "KeepSourceNumbering" flag to "true", then the document clone
    // list will preserve its original numbering, making the two lists appear as separate lists. 
    let importFormatOptions = new aw.ImportFormatOptions();
    importFormatOptions.keepSourceNumbering = keepSourceNumbering;

    let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KeepDifferentStyles, importFormatOptions);
    for (let paragraph of srcDoc.firstSection.body.paragraphs)
    {
      let importedNode = importer.importNode(paragraph, true);
      dstDoc.firstSection.body.appendChild(importedNode);
    }

    dstDoc.updateListLabels();

    if (keepSourceNumbering)
    {
      expect(dstDoc.firstSection.body.toString(aw.SaveFormat.Text).trim()).toEqual(
        "6. Item 1\r\n" +
        "7. Item 2 \r\n" +
        "8. Item 3\r\n" +
        "9. Item 4\r\n" +
        "6. Item 1\r\n" +
        "7. Item 2 \r\n" +
        "8. Item 3\r\n" +
        "9. Item 4");
    }
    else
    {
      expect(dstDoc.firstSection.body.toString(aw.SaveFormat.Text).trim()).toEqual(
        "6. Item 1\r\n" +
        "7. Item 2 \r\n" +
        "8. Item 3\r\n" +
        "9. Item 4\r\n" +
        "10. Item 1\r\n" +
        "11. Item 2 \r\n" +
        "12. Item 3\r\n" +
        "13. Item 4");
    }
    //ExEnd
  });


  //ExStart
  //ExFor:Paragraph.IsEndOfSection
  //ExFor:NodeImporter
  //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
  //ExFor:NodeImporter.ImportNode(Node, Boolean)
  //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
  test('InsertAtBookmark', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("InsertionPoint");
    builder.write("We will insert a document here: ");
    builder.endBookmark("InsertionPoint");

    let docToInsert = new aw.Document();
    builder = new aw.DocumentBuilder(docToInsert);

    builder.write("Hello world!");

    docToInsert.save(base.artifactsDir + "NodeImporter.InsertAtMergeField.docx");

    let bookmark = doc.range.bookmarks.at("InsertionPoint");
    insertDocument(bookmark.bookmarkStart.parentNode, docToInsert);

    expect(doc.getText().trim()).toEqual("We will insert a document here: " +
                            "\rHello world!");
  });


  /// <summary>
  /// Inserts the contents of a document after the specified node.
  /// </summary>
  function insertDocument(insertionDestination, docToInsert)
  {
    if (insertionDestination.nodeType == aw.NodeType.Paragraph || insertionDestination.nodeType == aw.NodeType.Table)
    {
      let destinationParent = insertionDestination.parentNode;

      let importer =
        new aw.NodeImporter(docToInsert, insertionDestination.document, aw.ImportFormatMode.KeepSourceFormatting);

        // Loop through all block-level nodes in the section's body,
        // then clone and insert every node that is not the last empty paragraph of a section.
      for (var srcSection of docToInsert.sections.toArray())
        for (let srcNode of srcSection.body)
        {
          if (srcNode.nodeType == aw.NodeType.Paragraph)
          {
            let para = srcNode.asParagraph();
            if (para.isEndOfSection && !para.hasChildNodes)
              continue;
          }

          let newNode = importer.importNode(srcNode, true);

          destinationParent.insertAfter(newNode, insertionDestination);
          insertionDestination = newNode;
        }
    }
    else
    {
      throw new Error("The destination node should be either a paragraph or table.");
    }
  }
  //ExEnd


  test.skip('InsertAtMergeField - TODO: WORDSNODEJS-119 - Add support of MailMerge', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("A document will appear here: ");
    builder.insertField(" MERGEFIELD Document_1 ");

    let subDoc = new aw.Document();
    builder = new aw.DocumentBuilder(subDoc);
    builder.write("Hello world!");

    subDoc.save(base.artifactsDir + "NodeImporter.InsertAtMergeField.docx");

    doc.mailMerge.fieldMergingCallback = new InsertDocumentAtMailMergeHandler();

    // The main document has a merge field in it called "Document_1".
    // Execute a mail merge using a data source that contains a local system filename
    // of the document that we wish to insert into the MERGEFIELD.
    doc.mailMerge.execute([ "Document_1" ],
      [ base.artifactsDir + "NodeImporter.InsertAtMergeField.docx" ]);

    expect(doc.getText().trim()).toEqual("A document will appear here: \r" +
                            "Hello world!");
  });

/*
    /// <summary>
    /// If the mail merge encounters a MERGEFIELD with a specified name,
    /// this handler treats the current value of a mail merge data source as a local system filename of a document.
    /// The handler will insert the document in its entirety into the MERGEFIELD instead of the current merge value.
    /// </summary>
  private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
  {
    void aw.MailMerging.IFieldMergingCallback.fieldMerging(FieldMergingArgs args)
    {
      if (args.documentFieldName == "Document_1")
      {
        let builder = new aw.DocumentBuilder(args.document);
        builder.moveToMergeField(args.documentFieldName);

        let subDoc = new aw.Document((string)args.fieldValue);

        InsertDocument(builder.currentParagraph, subDoc);

        if (!builder.currentParagraph.hasChildNodes)
          builder.currentParagraph.remove();

        args.text = null;
      }
    }

    void aw.MailMerging.IFieldMergingCallback.imageFieldMerging(ImageFieldMergingArgs args)
    {
        // Do nothing.
    }
  }
*/
});
