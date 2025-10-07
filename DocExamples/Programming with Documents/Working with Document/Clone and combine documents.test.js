// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithRevisions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CloneDocument', () => {
    //ExStart:CloneDocument
    //GistId:4140b2f1857750e685e7bf1b2d9ba8dd
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("This is the original document before applying the clone method");

    // Clone the document.
    let clone = doc.clone();

    // Edit the cloned document.
    builder = new aw.DocumentBuilder(clone);
    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");

    // This shows what is in the document originally. The document has two sections.
    console.log("Cloned document text:", clone.getText().trim());

    // Duplicate the last section and append the copy to the end of the document.
    let lastSectionIdx = clone.sections.count - 1;
    let newSection = clone.sections.at(lastSectionIdx).clone();
    clone.sections.add(newSection);

    // Check what the document contains after we changed it.
    console.log("Document text after changes:", clone.getText().trim());
    clone.save(base.artifactsDir + "CloneAndCombineDocuments.CloningDocument.docx");
    //ExEnd:CloneDocument
  });

  test('InsertDocumentAtBookmark', () => {
    //ExStart:InsertDocumentAtBookmark
    //GistId:814f45acd0c15059a9680cb661081d0f
    let mainDoc = new aw.Document(base.myDir + "Document insertion 1.docx");
    let subDoc = new aw.Document(base.myDir + "Document insertion 2.docx");

    let bookmark = mainDoc.range.bookmarks.at("insertionPlace");
    insertDocument(bookmark.bookmarkStart.parentNode, subDoc);

    mainDoc.save(base.artifactsDir + "CloneAndCombineDocuments.InsertDocumentAtBookmark.docx");
    //ExEnd:InsertDocumentAtBookmark
  });

  //ExStart:InsertDocumentAsNodes
  //GistId:814f45acd0c15059a9680cb661081d0f
  /// <summary>
  /// Inserts content of the external document after the specified node.
  /// Section breaks and section formatting of the inserted document are ignored.
  /// </summary>
  /// <param name="insertionDestination">Node in the destination document after which the content
  /// Should be inserted. This node should be a block level node (paragraph or table).</param>
  /// <param name="docToInsert">The document to insert.</param>
  function insertDocument(insertionDestination, docToInsert) {
    if (insertionDestination.nodeType == aw.NodeType.Paragraph || insertionDestination.nodeType == aw.NodeType.Table) {
      let destinationParent = insertionDestination.parentNode;

      let importer = new aw.NodeImporter(docToInsert, insertionDestination.document, aw.ImportFormatMode.KeepSourceFormatting);

      // Loop through all block-level nodes in the section's body,
      // then clone and insert every node that is not the last empty paragraph of a section.
      for (let srcSection of docToInsert.sections) {
        let bodyNodes = srcSection.asSection().body.getChildNodes(aw.NodeType.Any, false);

        for (let j = 0; j < bodyNodes.count; j++) {
          let srcNode = bodyNodes.at(j);

          if (srcNode.nodeType == aw.NodeType.Paragraph) {
            let para = srcNode;
            if (para.isEndOfSection && !para.hasChildNodes) {
              continue;
            }
          }

          let newNode = importer.importNode(srcNode, true);

          destinationParent.insertAfter(newNode, insertionDestination);
          insertionDestination = newNode;
        }
      }
    } else {
      throw new Error("The destination node should be either a paragraph or table.");
    }
  }
  //ExEnd:InsertDocumentAsNodes

  //ExStart:InsertDocumentWithSectionFormatting
  /// <summary>
  /// Inserts content of the external document after the specified node.
  /// </summary>
  /// <param name="insertAfterNode">Node in the destination document after which the content
  /// Should be inserted. This node should be a block level node (paragraph or table).</param>
  /// <param name="srcDoc">The document to insert.</param>
  function insertDocumentWithSectionFormatting(insertAfterNode, srcDoc) {
    if (insertAfterNode.nodeType != aw.NodeType.Paragraph &&
        insertAfterNode.nodeType != aw.NodeType.Table) {
      throw new Error("The destination node should be either a paragraph or table.");
    }

    let dstDoc = insertAfterNode.document;
    // To retain section formatting, split the current section into two at the marker node and then import the content
    // from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
    let currentSection = insertAfterNode.getAncestor(aw.NodeType.Section);

    // Don't clone the content inside the section, we just want the properties of the section retained.
    let cloneSection = currentSection.clone(false).asSection();
    // However, make sure the clone section has a body but no empty first paragraph.
    cloneSection.ensureMinimum();
    cloneSection.body.firstParagraph.remove();

    insertAfterNode.document.insertAfter(cloneSection, currentSection);

    // Append all nodes after the marker node to the new section. This will split the content at the section level at.
    // The marker so the sections from the other document can be inserted directly.
    let currentNode = insertAfterNode.nextSibling;
    while (currentNode != null) {
      let nextNode = currentNode.nextSibling;
      cloneSection.body.appendChild(currentNode);
      currentNode = nextNode;
    }

    // This object will be translating styles and lists during the import.
    let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.UseDestinationStyles);

    for (let i = 0; i < srcDoc.sections.count; i++) {
      let srcSection = srcDoc.sections.at(i);
      let newNode = importer.importNode(srcSection, true);

      dstDoc.insertAfter(newNode, currentSection);
      currentSection = newNode;
    }
  }
  //ExEnd:InsertDocumentWithSectionFormatting
});
