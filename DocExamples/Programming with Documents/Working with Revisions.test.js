// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithRevisions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('AcceptRevisions', () => {
    //ExStart:AcceptAllRevisions
    //GistId:e8d71fde166d275d0fc9471c56c3ad39
    let doc = new aw.Document();
    let body = doc.firstSection.body;
    let para = body.firstParagraph;

    // Add text to the first paragraph, then add two more paragraphs.
    para.appendChild(new aw.Run(doc, "Paragraph 1. "));
    body.appendParagraph("Paragraph 2. ");
    body.appendParagraph("Paragraph 3. ");

    // We have three paragraphs, none of which registered as any type of revision
    // If we add/remove any content in the document while tracking revisions,
    // they will be displayed as such in the document and can be accepted/rejected.
    doc.startTrackRevisions("John Doe", new Date());

    // This paragraph is a revision and will have the according "IsInsertRevision" flag set.
    para = body.appendParagraph("Paragraph 4. ");
    expect(para.isInsertRevision).toBe(true);

    // Get the document's paragraph collection and remove a paragraph.
    let paragraphs = body.paragraphs;
    expect(paragraphs.count).toBe(4);
    para = paragraphs.at(2);
    para.remove();

    // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
    // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
    expect(paragraphs.count).toBe(4);
    expect(para.isDeleteRevision).toBe(true);

    // The delete revision paragraph is removed once we accept changes.
    doc.acceptAllRevisions();
    expect(paragraphs.count).toBe(3);
    expect(para.parentNode).toBeNull();

    // Stopping the tracking of revisions makes this text appear as normal text.
    // Revisions are not counted when the document is changed.
    doc.stopTrackRevisions();

    // Save the document.
    doc.save(base.artifactsDir + "WorkingWithRevisions.AcceptRevisions.docx");
    //ExEnd:AcceptAllRevisions
  });

  test('GetRevisionTypes', () => {
    //ExStart:GetRevisionTypes
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    let paragraphs = doc.firstSection.body.paragraphs;
    for (let i = 0; i < paragraphs.count; i++) {
      if (paragraphs.at(i).isMoveFromRevision)
        console.log("The paragraph {0} has been moved (deleted).", i);
      if (paragraphs.at(i).isMoveToRevision)
        console.log("The paragraph {0} has been moved (inserted).", i);
    }
    //ExEnd:GetRevisionTypes
  });

  test('GetRevisionGroups', () => {
    //ExStart:GetRevisionGroups
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    for (let group of doc.revisions.groups) {
      console.log("{0}, {1}:", group.author, group.revisionType);
      console.log(group.text);
    }
    //ExEnd:GetRevisionGroups
  });

  test('RemoveCommentsInPdf', () => {
    //ExStart:RemoveCommentsInPDF
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    // Do not render the comments in PDF.
    doc.layoutOptions.commentDisplayMode = aw.Layout.CommentDisplayMode.Hide;

    doc.save(base.artifactsDir + "WorkingWithRevisions.RemoveCommentsInPdf.pdf");
    //ExEnd:RemoveCommentsInPDF
  });

  test('ShowRevisionsInBalloons', () => {
    //ExStart:ShowRevisionsInBalloons
    //GistId:ce015d9bade4e0294485ffb47462ded4
    //ExStart:SetMeasurementUnit
    //ExStart:SetRevisionBarsPosition
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    // Renders insert revisions inline, delete and format revisions in balloons.
    doc.layoutOptions.revisionOptions.showInBalloons = aw.Layout.ShowInBalloons.FormatAndDelete;
    doc.layoutOptions.revisionOptions.measurementUnit = aw.MeasurementUnits.Inches;
    // Renders revision bars on the right side of a page.
    doc.layoutOptions.revisionOptions.revisionBarsPosition = aw.Drawing.HorizontalAlignment.Right;

    doc.save(base.artifactsDir + "WorkingWithRevisions.ShowRevisionsInBalloons.pdf");
    //ExEnd:SetRevisionBarsPosition
    //ExEnd:SetMeasurementUnit
    //ExEnd:ShowRevisionsInBalloons
  });

  test('GetRevisionGroupDetails', () => {
    //ExStart:GetRevisionGroupDetails
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    for (let revision of doc.revisions) {
      let groupText = revision.group != null
          ? "Revision group text: " + revision.group.text
          : "Revision has no group";

      console.log("Type: " + revision.revisionType);
      console.log("Author: " + revision.author);
      console.log("Date: " + revision.dateTime);
      console.log("Revision text: " + revision.parentNode.toString(aw.SaveFormat.Text));
      console.log(groupText);
    }
    //ExEnd:GetRevisionGroupDetails
  });

  test('AccessRevisedVersion', () => {
    //ExStart:AccessRevisedVersion
    let doc = new aw.Document(base.myDir + "Revisions.docx");
    doc.updateListLabels();

    // Switch to the revised version of the document.
    doc.revisionsView = aw.RevisionsView.Final;

    for (let revision of doc.revisions) {
      if (revision.parentNode.nodeType == aw.NodeType.Paragraph) {
        let paragraph = revision.parentNode;
        if (paragraph.isListItem) {
          console.log(paragraph.listLabel.labelString);
          console.log(paragraph.listFormat.listLevel);
        }
      }
    }
    //ExEnd:AccessRevisedVersion
  });

  test('MoveNodeInTrackedDocument', () => {
    //ExStart:MoveNodeInTrackedDocument
    //GistId:e8d71fde166d275d0fc9471c56c3ad39
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Paragraph 1");
    builder.writeln("Paragraph 2");
    builder.writeln("Paragraph 3");
    builder.writeln("Paragraph 4");
    builder.writeln("Paragraph 5");
    builder.writeln("Paragraph 6");
    let body = doc.firstSection.body;
    console.log("Paragraph count: {0}", body.paragraphs.count);

    // Start tracking revisions.
    doc.startTrackRevisions("Author", new Date(2020, 11, 23, 14, 0, 0));

    // Generate revisions when moving a node from one location to another.
    let node = body.paragraphs.at(3);
    let endNode = body.paragraphs.at(5).nextSibling;
    let referenceNode = body.paragraphs.at(0);
    while (node != endNode && node != null) {
      let nextNode = node.nextSibling;
      body.insertBefore(node, referenceNode);
      node = nextNode;
    }

    // Stop the process of tracking revisions.
    doc.stopTrackRevisions();

    // There are 3 additional paragraphs in the move-from range.
    console.log("Paragraph count: {0}", body.paragraphs.count);
    doc.save(base.artifactsDir + "WorkingWithRevisions.MoveNodeInTrackedDocument.docx");
    //ExEnd:MoveNodeInTrackedDocument
  });

  test('ShapeRevision', () => {
    //ExStart:ShapeRevision
    //GistId:e8d71fde166d275d0fc9471c56c3ad39
    let doc = new aw.Document();

    // Insert an inline shape without tracking revisions.
    expect(doc.trackRevisions).toBe(false);
    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Cube);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.width = 100.0;
    shape.height = 100.0;
    doc.firstSection.body.firstParagraph.appendChild(shape);

    // Start tracking revisions and then insert another shape.
    doc.startTrackRevisions("John Doe");
    shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.Sun);
    shape.wrapType = aw.Drawing.WrapType.Inline;
    shape.width = 100.0;
    shape.height = 100.0;
    doc.firstSection.body.firstParagraph.appendChild(shape);

    // Get the document's shape collection which includes just the two shapes we added.
    let shapes = Array.from(doc.getChildNodes(aw.NodeType.Shape, true));
    expect(shapes.length).toBe(2);

    // Remove the first shape.
    let shape0 = shapes.at(0).asShape();
    shape0.remove();

    // Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
    expect(shape0.shapeType).toBe(aw.Drawing.ShapeType.Cube);
    expect(shape0.isDeleteRevision).toBe(true);

    // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
    let shape1 = shapes.at(1).asShape();
    expect(shape1.shapeType).toBe(aw.Drawing.ShapeType.Sun);
    expect(shape1.isInsertRevision).toBe(true);

    // The document has one shape that was moved, but shape move revisions will have two instances of that shape.
    // One will be the shape at its arrival destination and the other will be the shape at its original location.
    doc = new aw.Document(base.myDir + "Revision shape.docx");

    shapes = Array.from(doc.getChildNodes(aw.NodeType.Shape, true));
    expect(shapes.length).toBe(2);

    // This is the move to revision, also the shape at its arrival destination.
    shape0 = shapes.at(0).asShape();
    expect(shape0.isMoveFromRevision).toBe(false);
    expect(shape0.isMoveToRevision).toBe(true);

    // This is the move from revision, which is the shape at its original location.
    shape1 = shapes.at(1).asShape();
    expect(shape1.isMoveFromRevision).toBe(true);
    expect(shape1.isMoveToRevision).toBe(false);
    //ExEnd:ShapeRevision
  });

});