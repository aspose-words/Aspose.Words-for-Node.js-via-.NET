// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("SplitDocument", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('ByHeadings', () => {
    //ExStart:SplitDocumentByHeadings
    //GistId:e4b272992a7c8fafdd7ff42f8c2de379
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let options = new aw.Saving.HtmlSaveOptions();
    // Split a document into smaller parts, in this instance split by heading.
    options.documentSplitCriteria = aw.Saving.DocumentSplitCriteria.HeadingParagraph;

    doc.save(base.artifactsDir + "SplitDocument.ByHeadings.epub", options);
    //ExEnd:SplitDocumentByHeadings
  });

  test('BySectionsHtml', () => {
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    //ExStart:SplitDocumentBySectionsHtml
    //GistId:76d266d6b68098eb4e95e484fe85538e
    let options = new aw.Saving.HtmlSaveOptions();
    options.documentSplitCriteria = aw.Saving.DocumentSplitCriteria.SectionBreak;
    //ExEnd:SplitDocumentBySectionsHtml

    doc.save(base.artifactsDir + "SplitDocument.BySections.html", options);
  });

  test('BySections', () => {
    //ExStart:SplitDocumentBySections
    //GistId:76d266d6b68098eb4e95e484fe85538e
    let doc = new aw.Document(base.myDir + "Big document.docx");
    for (let i = 0; i < doc.sections.count; i++) {
      // Split a document into smaller parts, in this instance, split by section.
      let section = doc.sections.at(i).clone();

      let newDoc = new aw.Document();
      newDoc.sections.clear();

      let newSection = newDoc.importNode(section, true);
      newDoc.sections.add(newSection);

      // Save each section as a separate document.
      newDoc.save(base.artifactsDir + `SplitDocument.BySections_${i}.docx`);
    }
    //ExEnd:SplitDocumentBySections
  });

  test('PageByPage', () => {
    //ExStart:SplitDocumentPageByPage
    //GistId:76d266d6b68098eb4e95e484fe85538e
    let doc = new aw.Document(base.myDir + "Big document.docx");

    let pageCount = doc.pageCount;

    for (let page = 0; page < pageCount; page++) {
      // Save each page as a separate document.
      let extractedPage = doc.extractPages(page, 1);
      extractedPage.save(base.artifactsDir + `SplitDocument.PageByPage_${page + 1}.docx`);
    }
    //ExEnd:SplitDocumentPageByPage

    mergeDocuments();
  });

  //ExStart:MergeSplitDocuments
  //GistId:76d266d6b68098eb4e95e484fe85538e
  function mergeDocuments() {
    // Find documents using for merge.
    let files = fs.readdirSync(base.artifactsDir)
        .filter(file => file.match(/^SplitDocument\.PageByPage_.*\.docx$/))
        .map(file => ({
          name: file,
          fullPath: path.join(base.artifactsDir, file),
          stats: fs.statSync(path.join(base.artifactsDir, file))
        }))
        .sort((a, b) => a.stats.birthtime - b.stats.birthtime);

    let sourceDocumentPath = path.join(base.artifactsDir, "SplitDocument.PageByPage_1.docx");

    // Open the first part of the resulting document.
    let sourceDoc = new aw.Document(sourceDocumentPath);

    // Create a new resulting document.
    let mergedDoc = new aw.Document();
    let mergedDocBuilder = new aw.DocumentBuilder(mergedDoc);

    // Merge document parts one by one.
    for (let file of files) {
      if (file.fullPath === sourceDocumentPath)
        continue;
      mergedDocBuilder.moveToDocumentEnd();
      mergedDocBuilder.insertDocument(sourceDoc, aw.ImportFormatMode.KeepSourceFormatting);
      sourceDoc = new aw.Document(file.fullPath);
    }

    mergedDoc.save(base.artifactsDir + "SplitDocument.MergeDocuments.docx");
  }
  //ExEnd:MergeSplitDocuments

  test('ByPageRange', () => {
    //ExStart:SplitDocumentByPageRange
    //GistId:76d266d6b68098eb4e95e484fe85538e
    let doc = new aw.Document(base.myDir + "Big document.docx");

    // Get part of the document.
    let extractedPages = doc.extractPages(3, 6);
    extractedPages.save(base.artifactsDir + "SplitDocument.ByPageRange.docx");
    //ExEnd:SplitDocumentByPageRange
  });
});
