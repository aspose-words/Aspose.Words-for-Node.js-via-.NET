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


describe("ExNode", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CloneCompositeNode', () => {
    //ExStart
    //ExFor:Node
    //ExFor:Node.clone
    //ExSummary:Shows how to clone a composite node.
    let doc = new aw.Document();
    let para = doc.firstSection.body.firstParagraph;
    para.appendChild(new aw.Run(doc, "Hello world!"));

    // Below are two ways of cloning a composite node.
    // 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
    let cloneWithChildren = para.clone(true).asParagraph();

    expect(cloneWithChildren.hasChildNodes).toEqual(true);
    expect(cloneWithChildren.getText().trim()).toEqual("Hello world!");

    // 2 -  Create a clone of a node just by itself without any children.
    let cloneWithoutChildren = para.clone(false).asParagraph();

    expect(cloneWithoutChildren.hasChildNodes).toEqual(false);
    expect(cloneWithoutChildren.getText().trim()).toEqual('');
    //ExEnd
  });


  test('GetParentNode', () => {
    //ExStart
    //ExFor:Node.parentNode
    //ExSummary:Shows how to access a node's parent node.
    let doc = new aw.Document();
    let para = doc.firstSection.body.firstParagraph;

    // Append a child Run node to the document's first paragraph.
    let run = new aw.Run(doc, "Hello world!");
    para.appendChild(run);

    // The paragraph is the parent node of the run node. We can trace this lineage
    // all the way to the document node, which is the root of the document's node tree.
    expect(run.parentNode.referenceEquals(para)).toEqual(true);
    expect(para.parentNode.referenceEquals(doc.firstSection.body)).toEqual(true);
    expect(doc.firstSection.body.parentNode.referenceEquals(doc.firstSection)).toEqual(true);
    expect(doc.firstSection.parentNode.referenceEquals(doc)).toEqual(true);
    //ExEnd
  });


  test('OwnerDocument', () => {
    //ExStart
    //ExFor:Node.document
    //ExFor:Node.parentNode
    //ExSummary:Shows how to create a node and set its owning document.
    let doc = new aw.Document();
    let para = new aw.Paragraph(doc);
    para.appendChild(new aw.Run(doc, "Hello world!"));

    // We have not yet appended this paragraph as a child to any composite node.
    expect(para.parentNode).toBe(null);

    // If a node is an appropriate child node type of another composite node,
    // we can attach it as a child only if both nodes have the same owner document.
    // The owner document is the document we passed to the node's constructor.
    // We have not attached this paragraph to the document, so the document does not contain its text.
    expect(doc.referenceEquals(para.document)).toEqual(true);
    expect(doc.getText().trim()).toEqual('');

    // Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
    para.paragraphFormat.style = doc.styles.at("Heading 1");

    // Add this node to the document, and then verify its contents.
    doc.firstSection.body.appendChild(para);

    expect(para.parentNode.referenceEquals(doc.firstSection.body)).toEqual(true);
    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd

    expect(para.document.referenceEquals(doc)).toEqual(true);
    expect(para.parentNode).not.toBe(null);
  });


  test('ChildNodesEnumerate', () => {
    //ExStart
    //ExFor:Node
    //ExFor:Node.customNodeId
    //ExFor:NodeType
    //ExFor:CompositeNode
    //ExFor:CompositeNode.getChild
    //ExFor:CompositeNode.getChildNodes(NodeType, bool)
    //ExFor:NodeCollection.count
    //ExFor:NodeCollection.item
    //ExSummary:Shows how to traverse through a composite node's collection of child nodes.
    let doc = new aw.Document();

    // Add two runs and one shape as child nodes to the first paragraph of this document.
    let paragraph = doc.getParagraph(0, true);
    paragraph.appendChild(new aw.Run(doc, "Hello world! "));

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Rectangle);
    shape.width = 200;
    shape.height = 200;
    // Note that the 'CustomNodeId' is not saved to an output file and exists only during the node lifetime.
    shape.customNodeId = 100;
    shape.wrapType = aw.Drawing.WrapType.Inline;
    paragraph.appendChild(shape);

    paragraph.appendChild(new aw.Run(doc, "Hello again!"));

    // Iterate through the paragraph's collection of immediate children,
    // and print any runs or shapes that we find within.
    let children = paragraph.getChildNodes(aw.NodeType.Any, false);

    expect(paragraph.getChildNodes(aw.NodeType.Any, false).count).toEqual(3);

    for (let child of children)
      switch (child.nodeType)
      {
        case aw.NodeType.Run:
          console.log("Run contents:");
          console.log(`\t\"${child.getText().trim()}\"`);
          break;
        case aw.NodeType.Shape:
          let childShape = child.asShape();
          console.log("Shape:");
          console.log(`\t${childShape.shapeType}, ${childShape.width}x${childShape.height}`);
          expect(shape.customNodeId).toEqual(100);
          break;
      }
    //ExEnd

    expect(paragraph.getChild(aw.NodeType.Run, 0, true).nodeType).toEqual(aw.NodeType.Run);
    expect(doc.getText().trim()).toEqual("Hello world! Hello again!");
  });


  //ExStart
  //ExFor:Node.NextSibling
  //ExFor:CompositeNode.FirstChild
  //ExFor:Node.IsComposite
  //ExFor:CompositeNode.IsComposite
  //ExFor:Node.NodeTypeToString
  //ExFor:Paragraph.NodeType
  //ExFor:Table.NodeType
  //ExFor:Node.NodeType
  //ExFor:Footnote.NodeType
  //ExFor:FormField.NodeType
  //ExFor:SmartTag.NodeType
  //ExFor:Cell.NodeType
  //ExFor:Row.NodeType
  //ExFor:Document.NodeType
  //ExFor:Comment.NodeType
  //ExFor:Run.NodeType
  //ExFor:Section.NodeType
  //ExFor:SpecialChar.NodeType
  //ExFor:Shape.NodeType
  //ExFor:FieldEnd.NodeType
  //ExFor:FieldSeparator.NodeType
  //ExFor:FieldStart.NodeType
  //ExFor:BookmarkStart.NodeType
  //ExFor:CommentRangeEnd.NodeType
  //ExFor:BuildingBlock.NodeType
  //ExFor:GlossaryDocument.NodeType
  //ExFor:BookmarkEnd.NodeType
  //ExFor:GroupShape.NodeType
  //ExFor:CommentRangeStart.NodeType
  //ExSummary:Shows how to traverse a composite node's tree of child nodes.
  test('RecurseChildren', () => {
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    // Any node that can contain child nodes, such as the document itself, is composite.
    expect(doc.isComposite).toEqual(true);

    // Invoke the recursive function that will go through and print all the child nodes of a composite node.
    traverseAllNodes(doc, 0);
  });


  /// <summary>
  /// Recursively traverses a node tree while printing the type of each node
  /// with an indent depending on depth as well as the contents of all inline nodes.
  /// </summary>
  function traverseAllNodes(parentNode, depth)
  {
    for (let childNode = parentNode.firstChild; childNode != null; childNode = childNode.nextSibling)
    {
      console.log(`${'\t'.repeat(depth)}${aw.Node.nodeTypeToString(childNode.nodeType)}`);

      // Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
      if (childNode.isComposite)
      {
        traverseAllNodes(childNode.asCompositeNode(), depth + 1);
      }
      else
      {
        var text = childNode.getText().trim();
        if (text !== undefined) {
          console.log(` - \"${text}\"`);
        }
      }
    }
  }
  //ExEnd


  test('RemoveNodes', () => {

    //ExStart
    //ExFor:Node
    //ExFor:Node.nodeType
    //ExFor:Node.remove
    //ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(2);

    let curNode = doc.firstSection.body.firstChild;

    while (curNode != null)
    {
      // Save the next sibling node as a variable in case we want to move to it after deleting this node.
      let nextNode = curNode.nextSibling;

      // A section body can contain Paragraph and Table nodes.
      // If the node is a Table, remove it from the parent.
      if (curNode.nodeType == aw.NodeType.Table)
        curNode.remove();

      curNode = nextNode;
    }

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(0);
    //ExEnd
  });


  test('EnumNextSibling', () => {
    //ExStart
    //ExFor:CompositeNode.firstChild
    //ExFor:Node.nextSibling
    //ExFor:Node.nodeTypeToString
    //ExFor:Node.nodeType
    //ExSummary:Shows how to use a node's NextSibling property to enumerate through its immediate children.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    for (let node = doc.firstSection.body.firstChild; node != null; node = node.nextSibling)
    {
      console.log(`Node type: ${aw.Node.nodeTypeToString(node.nodeType)}`);

      let contents = node.getText().trim();
      console.log(contents == '' ? "This node contains no text" : `Contents: \"${node.getText().trim()}\"`);
    }
    //ExEnd
  });


  test('TypedAccess', () => {

    //ExStart
    //ExFor:Story.tables
    //ExFor:Table.firstRow
    //ExFor:Table.lastRow
    //ExFor:TableCollection
    //ExSummary:Shows how to remove the first and last rows of all tables in a document.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let tables = doc.firstSection.body.tables.toArray();

    expect(tables[0].rows.count).toEqual(5);
    expect(tables[1].rows.count).toEqual(4);

    for (var table of tables)
    {
      table.firstRow?.remove();
      table.lastRow?.remove();
    }

    expect(tables[0].rows.count).toEqual(3);
    expect(tables[1].rows.count).toEqual(2);
    //ExEnd
  });


  test('RemoveChild', () => {
    //ExStart
    //ExFor:CompositeNode.lastChild
    //ExFor:Node.previousSibling
    //ExFor:CompositeNode.removeChild``1(``0)
    //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
            
    builder.writeln("Section 1 text.");
    builder.insertBreak(aw.BreakType.SectionBreakContinuous);
    builder.writeln("Section 2 text.");

    // Both sections are siblings of each other.
    let lastSection = doc.lastChild.asSection();
    let firstSection = lastSection.previousSibling.asSection();

    // Remove a section based on its sibling relationship with another section.
    if (lastSection.previousSibling != null)
      doc.removeChild(firstSection);

    // The section we removed was the first one, leaving the document with only the second.
    expect(doc.getText().trim()).toEqual("Section 2 text.");
    //ExEnd
  });


  test('SelectCompositeNodes', () => {
    //ExStart
    //ExFor:CompositeNode.selectSingleNode
    //ExFor:CompositeNode.selectNodes
    //ExFor:NodeList.getEnumerator
    //ExFor:NodeList.toArray
    //ExSummary:Shows how to select certain nodes by using an XPath expression.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    // This expression will extract all paragraph nodes,
    // which are descendants of any table node in the document.
    let nodeList = doc.selectNodes("//Table//Paragraph");

    // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table.
    var index = 0;

    for (var n of nodeList)
      console.log(`Table paragraph index ${index++}, contents: \"${n.asParagraph().getText().trim()}\"`);

    // This expression will select any paragraphs that are direct children of any Body node in the document.
    nodeList = doc.selectNodes("//Body/Paragraph");

    // We can treat the list as an array.
    expect(nodeList.toArray().length).toEqual(4);

    // Use SelectSingleNode to select the first result of the same expression as above.
    let node = doc.selectSingleNode("//Body/Paragraph");

    expect(node.nodeType).toEqual(aw.NodeType.Paragraph);
    //ExEnd
  });


  test('TestNodeIsInsideField', () => {
    //ExStart
    //ExFor:CompositeNode.selectNodes
    //ExSummary:Shows how to use an XPath expression to test whether a node is inside a field.
    let doc = new aw.Document(base.myDir + "Mail merge destination - Northwind employees.docx");

    // The NodeList that results from this XPath expression will contain all nodes we find inside a field.
    // However, FieldStart and FieldEnd nodes can be on the list if there are nested fields in the path.
    // Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
    let resultList =
      doc.selectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]").toArray();

    // Check if the specified run is one of the nodes that are inside the field.
    console.log(`Contents of the first Run node that's part of a field: ${resultList.find(n => n.nodeType == aw.NodeType.Run).asRun().getText().trim()}`);
    //ExEnd
  });


  test('CreateAndAddParagraphNode', () => {
    let doc = new aw.Document();

    let para = new aw.Paragraph(doc);

    let section = doc.lastSection;
    section.body.appendChild(para);
  });


  test('RemoveSmartTagsFromCompositeNode', () => {
    //ExStart
    //ExFor:CompositeNode.removeSmartTags
    //ExSummary:Removes all smart tags from descendant nodes of a composite node.
    let doc = new aw.Document(base.myDir + "Smart tags.doc");

    expect(doc.getChildNodes(aw.NodeType.SmartTag, true).count).toEqual(8);

    doc.removeSmartTags();

    expect(doc.getChildNodes(aw.NodeType.SmartTag, true).count).toEqual(0);
    //ExEnd
  });


  test('GetIndexOfNode', () => {
    //ExStart
    //ExFor:CompositeNode.indexOf
    //ExSummary:Shows how to get the index of a given child node from its parent.
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let body = doc.firstSection.body;

    // Retrieve the index of the last paragraph in the body of the first section.
    expect(body.getChildNodes(aw.NodeType.Any, false).indexOf(body.lastParagraph)).toEqual(24);
    //ExEnd
  });


  test('ConvertNodeToHtmlWithDefaultOptions', () => {
    //ExStart
    //ExFor:Node.toString(SaveFormat)
    //ExFor:Node.toString(SaveOptions)
    //ExSummary:Exports the content of a node to String in HTML format.
    let doc = new aw.Document(base.myDir + "Document.docx");

    let node = doc.lastSection.body.lastParagraph;

    // When we call the ToString method using the html SaveFormat overload,
    // it converts the node's contents to their raw html representation.
    expect(node.toString(aw.SaveFormat.Html)).toEqual("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">" +
                            "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                            "</p>");

    // We can also modify the result of this conversion using a SaveOptions object.
    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.exportRelativeFontSize = true;

    expect(node.toString(saveOptions)).toEqual("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">" +
                            "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                            "</p>");
    //ExEnd
  });


  test('TypedNodeCollectionToArray', () => {
    //ExStart
    //ExFor:ParagraphCollection.toArray
    //ExSummary:Shows how to create an array from a NodeCollection.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    let paras = doc.firstSection.body.paragraphs.toArray();

    expect(paras.length).toEqual(22);
    //ExEnd
  });


  test('NodeEnumerationHotRemove', () => {
    //ExStart
    //ExFor:ParagraphCollection.toArray
    //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("The first paragraph");
    builder.writeln("The second paragraph");
    builder.writeln("The third paragraph");
    builder.writeln("The fourth paragraph");

    // Remove a node from the collection in the middle of an enumeration.
    for (var para of doc.firstSection.body.paragraphs.toArray())
      if (para.range.text.includes("third"))
        para.remove();

    expect(doc.getText().includes("The third paragraph")).toEqual(false);
    //ExEnd
  });


  //ExStart
  //ExFor:NodeChangingAction
  //ExFor:NodeChangingArgs.Action
  //ExFor:NodeChangingArgs.NewParent
  //ExFor:NodeChangingArgs.OldParent
  //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
  test.skip('NodeChangingCallback - TODO: WORDSNODEJS-120 - Add support of doc.nodeChangingCallback', () => {
    let doc = new aw.Document();
    doc.nodeChangingCallback = new NodeChangingPrinter();

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");
    builder.startTable();
    builder.insertCell();
    builder.write("Cell 1");
    builder.insertCell();
    builder.write("Cell 2");
    builder.endTable();

    builder.insertImage(base.imageDir + "Logo.jpg");            

    builder.currentParagraph.parentNode.removeAllChildren();
  });


/*    /// <summary>
    /// Prints every node insertion/removal as it takes place in the document.
    /// </summary>
  private class NodeChangingPrinter : INodeChangingCallback
  {
    void aw.INodeChangingCallback.nodeInserting(NodeChangingArgs args)
    {
      expect(args.action).toEqual(aw.NodeChangingAction.Insert);
      expect(args.oldParent).toEqual(null);
    }

    void aw.INodeChangingCallback.nodeInserted(NodeChangingArgs args)
    {
      expect(args.action).toEqual(aw.NodeChangingAction.Insert);
      expect(args.newParent).not.toBe(null);

      console.log("Inserted node:");
      console.log(`\tType:\t${args.node.nodeType}`);

      if (args.node.getText().trim() != "")
      {
        console.log(`\tText:\t\"${args.node.getText().trim()}\"`);
      }

      console.log(`\tHash:\t${args.node.getHashCode()}`);
      console.log(`\tParent:\t${args.newParent.nodeType} (${args.newParent.getHashCode()})`);
    }

    void aw.INodeChangingCallback.nodeRemoving(NodeChangingArgs args)
    {
      expect(args.action).toEqual(aw.NodeChangingAction.Remove);
    }

    void aw.INodeChangingCallback.nodeRemoved(NodeChangingArgs args)
    {
      expect(args.action).toEqual(aw.NodeChangingAction.Remove);
      expect(args.newParent).toBe(null);

      console.log(`Removed node: ${args.node.nodeType} (${args.node.getHashCode()})`);
    }
  }
    //ExEnd
*/


  test('NodeCollection', () => {
    //ExStart
    //ExFor:NodeCollection.contains(Node)
    //ExFor:NodeCollection.insert(Int32,Node)
    //ExFor:NodeCollection.remove(Node)
    //ExSummary:Shows how to work with a NodeCollection.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add text to the document by inserting Runs using a DocumentBuilder.
    builder.write("Run 1. ");
    builder.write("Run 2. ");

    // Every invocation of the "Write" method creates a new Run,
    // which then appears in the parent Paragraph's RunCollection.
    let runs = doc.firstSection.body.firstParagraph.runs;

    expect(runs.count).toEqual(2);

    // We can also insert a node into the RunCollection manually.
    let newRun = new aw.Run(doc, "Run 3. ");
    runs.insert(3, newRun);

    expect(runs.contains(newRun)).toEqual(true);
    expect(doc.getText().trim()).toEqual("Run 1. Run 2. Run 3.");

    // Access individual runs and remove them to remove their text from the document.
    let run = runs.at(1);
    runs.remove(run);

    expect(doc.getText().trim()).toEqual("Run 1. Run 3.");
    expect(run).not.toBe(null);
    expect(runs.contains(run)).toEqual(false);
    //ExEnd
  });


  test('NodeList', () => {
    //ExStart
    //ExFor:NodeList.count
    //ExFor:NodeList.item(Int32)
    //ExSummary:Shows how to use XPaths to navigate a NodeList.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert some nodes with a DocumentBuilder.
    builder.writeln("Hello world!");

    builder.startTable();
    builder.insertCell();
    builder.write("Cell 1");
    builder.insertCell();
    builder.write("Cell 2");
    builder.endTable();

    builder.insertImage(base.imageDir + "Logo.jpg");            

    // Our document contains three Run nodes.
    var nodeList = doc.selectNodes("//Run").toArray();

    expect(nodeList.length).toEqual(3);
    expect(nodeList.find(n => n.getText().trim() == "Hello world!")).not.toEqual(null);
    expect(nodeList.find(n => n.getText().trim() == "Cell 1")).not.toEqual(null);
    expect(nodeList.find(n => n.getText().trim() == "Cell 2")).not.toEqual(null);

    // Use a double forward slash to select all Run nodes
    // that are indirect descendants of a Table node, which would be the runs inside the two cells we inserted.
    nodeList = doc.selectNodes("//Table//Run").toArray();

    expect(nodeList.length).toEqual(2);
    expect(nodeList.find(n => n.getText().trim() == "Cell 1")).not.toEqual(null);
    expect(nodeList.find(n => n.getText().trim() == "Cell 2")).not.toEqual(null);

    // Access the shape that contains the image we inserted.
    nodeList = doc.selectNodes("//Shape").toArray();

    expect(nodeList.length).toEqual(1);

    let shape = nodeList.at(0).asShape();
    expect(shape.hasImage).toEqual(true);
    //ExEnd
  });
});
