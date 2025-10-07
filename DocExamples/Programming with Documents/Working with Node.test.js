// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithNode", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('GetNodeType', () => {
    //ExStart:GetNodeType
    //GistId:581adffafc4abd2deaf7d140c4698990
    let doc = new aw.Document();
    let type = doc.nodeType;
    //ExEnd:GetNodeType
  });

  test('GetParentNode', () => {
    //ExStart:GetParentNode
    //GistId:581adffafc4abd2deaf7d140c4698990
    let doc = new aw.Document();
    // The section is the first child node of the document.
    let section = doc.firstChild;
    // The section's parent node is the document.
    console.log("Section parent is the document: " + (doc == section.parentNode));
    //ExEnd:GetParentNode
  });

  test('OwnerDocument', () => {
    //ExStart:OwnerDocument
    //GistId:581adffafc4abd2deaf7d140c4698990
    let doc = new aw.Document();

    // Creating a new node of any type requires a document passed into the constructor.
    let para = new aw.Paragraph(doc);
    // The new paragraph node does not yet have a parent.
    console.log("Paragraph has no parent node: " + (para.parentNode == null));
    // But the paragraph node knows its document.
    console.log("Both nodes' documents are the same: " + (para.document == doc));
    // The fact that a node always belongs to a document allows us to access and modify
    // properties that reference the document-wide data, such as styles or lists.
    para.paragraphFormat.styleName = "Heading 1";
    // Now add the paragraph to the main text of the first section.
    doc.firstSection.body.appendChild(para);

    // The paragraph node is now a child of the Body node.
    console.log("Paragraph has a parent node: " + (para.parentNode != null));
    //ExEnd:OwnerDocument
  });

  test('EnumerateChildNodes', () => {
    //ExStart:EnumerateChildNodes
    //GistId:581adffafc4abd2deaf7d140c4698990
    let doc = new aw.Document();
    let paragraph = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();

    let children = paragraph.getChildNodes(aw.NodeType.Any, false);
    for (let child of children) {
      // A paragraph may contain children of various types such as runs, shapes, and others.
      if (child.nodeType == aw.NodeType.Run) {
        let run = child.asRun();
        console.log(run.text);
      }
    }
    //ExEnd:EnumerateChildNodes
  });

  //ExStart:RecurseAllNodes
  //GistId:581adffafc4abd2deaf7d140c4698990
  test('RecurseAllNodes', () => {
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");
    // Invoke the recursive function that will walk the tree.
    traverseAllNodes(doc);
  });

  /// <summary>
  /// A simple function that will walk through all children of a specified node recursively
  /// and print the type of each node to the screen.
  /// </summary>
  function traverseAllNodes(parentNode) {
    // This is the most efficient way to loop through immediate children of a node.
    for (let childNode = parentNode.firstChild; childNode != null; childNode = childNode.nextSibling) {
      console.log(aw.Node.nodeTypeToString(childNode.nodeType));

      // Recurse into the node if it is a composite node.
      if (childNode.isComposite)
        traverseAllNodes(childNode);
    }
  }
  //ExEnd:RecurseAllNodes

  test('TypedAccess', () => {
    //ExStart:TypedAccess
    //GistId:581adffafc4abd2deaf7d140c4698990
    let doc = new aw.Document();

    let section = doc.firstSection;
    let body = section.body;
    // Quick typed access to all Table child nodes contained in the Body.
    let tables = body.tables;
    for (let table of tables) {
      // Quick typed access to the first row of the table.
      table.firstRow?.remove();
      // Quick typed access to the last row of the table.
      table.lastRow?.remove();
    }
    //ExEnd:TypedAccess
  });

  test('CreateAndAddParagraphNode', () => {
    //ExStart:CreateAndAddParagraphNode
    let doc = new aw.Document();

    let para = new aw.Paragraph(doc);

    let section = doc.lastSection;
    section.body.appendChild(para);
    //ExEnd:CreateAndAddParagraphNode
  });

});