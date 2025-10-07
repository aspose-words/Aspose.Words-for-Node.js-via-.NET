// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("RemoveContent", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('RemovePageBreaks', () => {
    //ExStart:OpenDocument
    //GistId:1d626c7186a318d22d022dc96dd91d55
    let doc = new aw.Document(base.myDir + "Document.docx");
    //ExEnd:OpenDocument

    // In Aspose.words section breaks are represented as separate Section nodes in the document.
    // To remove these separate sections, the sections are combined.
    removePageBreaks(doc);
    removeSectionBreaks(doc);

    doc.save(base.artifactsDir + "RemoveContent.RemovePageBreaks.docx");
  });


  //ExStart:RemovePageBreaks
  function removePageBreaks(doc) {
    let paragraphs = doc.getChildNodes(aw.NodeType.Paragraph, true);

    for (let item of paragraphs) {
      let para = item.asParagraph();
      // If the paragraph has a page break before the set, then clear it.
      if (para.paragraphFormat.pageBreakBefore)
        para.paragraphFormat.pageBreakBefore = false;

      // Check all runs in the paragraph for page breaks and remove them.
      for (let node of para.runs) {
        var run = node.asRun();
        if (run.text.includes(aw.ControlChar.pageBreak))
          run.text = run.text.replace(aw.ControlChar.pageBreak, '');
      }
    }
  }
  //ExEnd:RemovePageBreaks


  //ExStart:RemoveSectionBreaks
  //GistId:1afca4d3da7cb4240fb91c3d93d8c30d
  function removeSectionBreaks(doc) {
    // Loop through all sections starting from the section that precedes the last one and moving to the first section.
    for (let i = doc.sections.count - 2; i >= 0; i--) {
      // Copy the content of the current section to the beginning of the last section.
      doc.lastSection.prependContent(doc.sections.at(i));
      // Remove the copied section.
      doc.sections.at(i).remove();
    }
  }
  //ExEnd:RemoveSectionBreaks


  test('RemoveFooters', () => {
    //ExStart:RemoveFooters
    //GistId:84cab3a22008f041ee6c1e959da09949
    let doc = new aw.Document(base.myDir + "Header and footer types.docx");

    for (let node of doc.sections) {
      var section = node.asSection();
      // Up to three different footers are possible in a section (for first, even and odd pages)
      // we check and delete all of them.
      let footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterFirst);
      footer?.remove();

      // Primary footer is the footer used for odd pages.
      footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary);
      footer?.remove();

      footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterEven);
      footer?.remove();
    }

    doc.save(base.artifactsDir + "RemoveContent.RemoveFooters.docx");
    //ExEnd:RemoveFooters
  });


  //ExStart:RemoveToc
  //GistId:db118a3e1559b9c88355356df9d7ea10
  test('RemoveToc', () => {
    let doc = new aw.Document(base.myDir + "Table of contents.docx");

    // Remove the first table of contents from the document.
    removeTableOfContents(doc, 0);

    doc.save(base.artifactsDir + "RemoveContent.RemoveToc.doc");
  });


  /// <summary>
  /// Removes the specified table of contents field from the document.
  /// </summary>
  /// <param name="doc">The document to remove the field from.</param>
  /// <param name="index">The zero-based index of the TOC to remove.</param>
  function removeTableOfContents(doc, index) {
    // Store the FieldStart nodes of TOC fields in the document for quick access.
    let fieldStarts = [];
    // This is a list to store the nodes found inside the specified TOC. They will be removed at the end of this method.
    let nodeList = [];

    for (var node of doc.getChildNodes(aw.NodeType.FieldStart, true)) {
      let start = node.asFieldStart();
      if (start.fieldType == aw.Fields.FieldType.FieldTOC) {
        fieldStarts.push(start);
      }
    }

    // Ensure the TOC specified by the passed index exists.
    if (index > fieldStarts.count - 1)
      throw new Error("TOC index is out of range");

    let isRemoving = true;

    let currentNode = fieldStarts.at(index);
    while (isRemoving) {
      // It is safer to store these nodes and delete them all at once later.
      nodeList.push(currentNode);
      currentNode = currentNode.nextPreOrder(doc);

      // Once we encounter a FieldEnd node of type FieldTOC,
      // we know we are at the end of the current TOC and stop here.
      if (currentNode.nodeType == aw.NodeType.FieldEnd) {
        let fieldEnd = currentNode.asFieldEnd();
        if (fieldEnd.fieldType == aw.Fields.FieldType.FieldTOC)
          isRemoving = false;
      }
    }

    for (let node of nodeList) {
      node.remove();
    }
  }
  //ExEnd:RemoveToc

});