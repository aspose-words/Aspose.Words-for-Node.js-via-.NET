// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');
const DocumentHelper = require('./DocumentHelper');

describe("ExEditableRange", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CreateAndRemove', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.startEditableRange
    //ExFor:aw.DocumentBuilder.endEditableRange
    //ExFor:EditableRange
    //ExFor:aw.EditableRange.editableRangeEnd
    //ExFor:aw.EditableRange.editableRangeStart
    //ExFor:aw.EditableRange.id
    //ExFor:aw.EditableRange.remove
    //ExFor:aw.EditableRangeEnd.editableRangeStart
    //ExFor:aw.EditableRangeEnd.id
    //ExFor:aw.EditableRangeEnd.nodeType
    //ExFor:aw.EditableRangeStart.editableRange
    //ExFor:aw.EditableRangeStart.id
    //ExFor:aw.EditableRangeStart.nodeType
    //ExSummary:Shows how to work with an editable range.
    let doc = new aw.Document();
    doc.protect(aw.ProtectionType.ReadOnly, "MyPassword");

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
            " we cannot edit this paragraph without the password.");

    // Editable ranges allow us to leave parts of protected documents open for editing.
    let editableRangeStart = builder.startEditableRange();
    builder.writeln("This paragraph is inside an editable range, and can be edited.");
    let editableRangeEnd = builder.endEditableRange();

    // A well-formed editable range has a start node, and end node.
    // These nodes have matching IDs and encompass editable nodes.
    let editableRange = editableRangeStart.editableRange;

    expect(editableRange.id).toEqual(editableRangeStart.id);
    expect(editableRange.id).toEqual(editableRangeEnd.id);
            
    // Different parts of the editable range link to each other.
    expect(editableRange.editableRangeStart.id).toEqual(editableRangeStart.id);
    expect(editableRangeEnd.editableRangeStart.id).toEqual(editableRangeStart.id);
    expect(editableRangeStart.editableRange.id).toEqual(editableRange.id);
    expect(editableRange.editableRangeEnd.id).toEqual(editableRangeEnd.id);

    // We can access the node types of each part like this. The editable range itself is not a node,
    // but an entity which consists of a start, an end, and their enclosed contents.
    expect(editableRangeStart.nodeType).toEqual(aw.NodeType.EditableRangeStart);
    expect(editableRangeEnd.nodeType).toEqual(aw.NodeType.EditableRangeEnd);

    builder.writeln("This paragraph is outside the editable range, and cannot be edited.");

    doc.save(base.artifactsDir + "EditableRange.CreateAndRemove.docx");

    // Remove an editable range. All the nodes that were inside the range will remain intact.
    editableRange.remove();
    //ExEnd

    expect(doc.getText().trim()).toEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph is inside an editable range, and can be edited.\r" +
                            "This paragraph is outside the editable range, and cannot be edited.");
    expect(doc.getChildNodes(aw.NodeType.EditableRangeStart, true).count).toEqual(0);

    doc = new aw.Document(base.artifactsDir + "EditableRange.CreateAndRemove.docx");

    expect(doc.protectionType).toEqual(aw.ProtectionType.ReadOnly);
    expect(doc.getText().trim()).toEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph is inside an editable range, and can be edited.\r" +
                            "This paragraph is outside the editable range, and cannot be edited.");

    //editableRange = ((EditableRangeStart)doc.getChild(aw.NodeType.EditableRangeStart, 0, true)).EditableRange;
    editableRange = doc.getEditableRangeStart(0, true).editableRange;

    TestUtil.verifyEditableRange(0, '', aw.EditorType.Unspecified, editableRange);
  });


  test('Nested', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.startEditableRange
    //ExFor:aw.DocumentBuilder.endEditableRange(EditableRangeStart)
    //ExFor:aw.EditableRange.editorGroup
    //ExSummary:Shows how to create nested editable ranges.
    let doc = new aw.Document();
    doc.protect(aw.ProtectionType.ReadOnly, "MyPassword");

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world! Since we have set the document's protection level to read-only, " +
            "we cannot edit this paragraph without the password.");
             
    // Create two nested editable ranges.
    let outerEditableRangeStart = builder.startEditableRange();
    builder.writeln("This paragraph inside the outer editable range and can be edited.");

    let innerEditableRangeStart = builder.startEditableRange();
    builder.writeln("This paragraph inside both the outer and inner editable ranges and can be edited.");

    // Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
    // When we want to end an editable range in this situation,
    // we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
    builder.endEditableRange(innerEditableRangeStart);

    builder.writeln("This paragraph inside the outer editable range and can be edited.");

    builder.endEditableRange(outerEditableRangeStart);

    builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.");

    // If a region of text has two overlapping editable ranges with specified groups,
    // the combined group of users excluded by both groups are prevented from editing it.
    outerEditableRangeStart.editableRange.editorGroup = aw.EditorType.Everyone;
    innerEditableRangeStart.editableRange.editorGroup = aw.EditorType.Contributors;

    doc.save(base.artifactsDir + "EditableRange.Nested.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "EditableRange.Nested.docx");

    expect(doc.getText().trim()).toEqual("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                            "This paragraph inside the outer editable range and can be edited.\r" +
                            "This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
                            "This paragraph inside the outer editable range and can be edited.\r" +
                            "This paragraph is outside any editable ranges, and cannot be edited.");

    let editableRange = doc.getEditableRangeStart(0, true).editableRange;

    TestUtil.verifyEditableRange(0, '', aw.EditorType.Everyone, editableRange);

    editableRange = doc.getEditableRangeStart(1, true).editableRange;

    TestUtil.verifyEditableRange(1, '', aw.EditorType.Contributors, editableRange);
  });


  /*  //ExStart

    //ExFor:EditableRange
    //ExFor:EditableRange.EditorGroup
    //ExFor:EditableRange.SingleUser
    //ExFor:EditableRangeEnd
    //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
    //ExFor:EditableRangeStart
    //ExFor:EditableRangeStart.Accept(DocumentVisitor)
    //ExFor:EditorType
    //ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
  test('Visitor', () => {
    let doc = new aw.Document();
    doc.protect(aw.ProtectionType.ReadOnly, "MyPassword");

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
            " we cannot edit this paragraph without the password.");

    // When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
    // There are two mutually exclusive ways to narrow down the list of allowed editors.
    // 1 -  Specify a user:
    let editableRange = builder.startEditableRange().EditableRange;
    editableRange.singleUser = "john.doe@myoffice.com";
    builder.writeln(`This paragraph is inside the first editable range, can only be edited by ${editableRange.singleUser}.`);
    builder.endEditableRange();

    expect(editableRange.editorGroup).toEqual(aw.EditorType.Unspecified);

    // 2 -  Specify a group that allowed users are associated with:
    editableRange = builder.startEditableRange().EditableRange;
    editableRange.editorGroup = aw.EditorType.Administrators;
    builder.writeln(`This paragraph is inside the first editable range, can only be edited by ${editableRange.editorGroup}.`);
    builder.endEditableRange();

    expect(editableRange.singleUser).toEqual('');

    builder.writeln("This paragraph is outside the editable range, and cannot be edited by anybody.");

    // Print details and contents of every editable range in the document.
    let editableRangePrinter = new EditableRangePrinter();

    doc.accept(editableRangePrinter);

    console.log(editableRangePrinter.ToText());
  });


    /// <summary>
    /// Collects properties and contents of visited editable ranges in a string.
    /// </summary>
  public class EditableRangePrinter : DocumentVisitor
  {
    public EditableRangePrinter()
    {
      mBuilder = new StringBuilder();
    }

    public string ToText()
    {
      return mBuilder.toString();
    }

    public void Reset()
    {
      mBuilder.clear();
      mInsideEditableRange = false;
    }

      /// <summary>
      /// Called when an EditableRangeStart node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
    {
      mBuilder.AppendLine(" -- Editable range found! -- ");
      mBuilder.AppendLine("\tID:\t\t" + editableRangeStart.id);
      if (editableRangeStart.editableRange.singleUser == '')
        mBuilder.AppendLine("\tGroup:\t" + editableRangeStart.editableRange.editorGroup);
      else
        mBuilder.AppendLine("\tUser:\t" + editableRangeStart.editableRange.singleUser);
      mBuilder.AppendLine("\tContents:");

      mInsideEditableRange = true;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when an EditableRangeEnd node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
    {
      mBuilder.AppendLine(" -- End of editable range --\n");

      mInsideEditableRange = false;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
      /// </summary>
    public override VisitorAction VisitRun(Run run)
    {
      if (mInsideEditableRange) mBuilder.AppendLine("\t\"" + run.text + "\"");

      return aw.VisitorAction.Continue;
    }

    private bool mInsideEditableRange;
    private readonly StringBuilder mBuilder;
  }
    //ExEnd*/

  test('IncorrectStructureException', () => {
    let doc = new aw.Document();

    let builder = new aw.DocumentBuilder(doc);

    // Assert that isn't valid structure for the current document.
    expect(() => builder.endEditableRange()).toThrow("EndEditableRange can not be called before StartEditableRange.");

    builder.startEditableRange();
  });


  test('IncorrectStructureDoNotAdded', () => {
    let doc = DocumentHelper.createDocumentFillWithDummyText();
    let builder = new aw.DocumentBuilder(doc);

    let startRange1 = builder.startEditableRange();

    builder.writeln("EditableRange_1_1");
    builder.writeln("EditableRange_1_2");

    startRange1.editableRange.editorGroup = aw.EditorType.Everyone;
    doc = DocumentHelper.saveOpen(doc);

    // Assert that it's not valid structure and editable ranges aren't added to the current document.
    let startNodes = doc.getChildNodes(aw.NodeType.EditableRangeStart, true);
    expect(startNodes.count).toEqual(0);

    let endNodes = doc.getChildNodes(aw.NodeType.EditableRangeEnd, true);
    expect(endNodes.count).toEqual(0);
  });


});
