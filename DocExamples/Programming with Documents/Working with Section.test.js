// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithSection", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('AddSection', () => {
    //ExStart:AddSection
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello1");
    builder.writeln("Hello2");

    let sectionToAdd = new aw.Section(doc);
    doc.sections.add(sectionToAdd);
    //ExEnd:AddSection
  });

  test('DeleteSection', () => {
    //ExStart:DeleteSection
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello1");
    doc.appendChild(new aw.Section(doc));
    builder.writeln("Hello2");
    doc.appendChild(new aw.Section(doc));

    doc.sections.removeAt(0);
    //ExEnd:DeleteSection

  });

  test('DeleteAllSections', () => {
    //ExStart:DeleteAllSections
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello1");
    doc.appendChild(new aw.Section(doc));
    builder.writeln("Hello2");
    doc.appendChild(new aw.Section(doc));

    doc.sections.clear();
    //ExEnd:DeleteAllSections
  });

  test('AppendSectionContent', () => {
    //ExStart:AppendSectionContent
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 3");

    let section = doc.sections.at(2);

    // Insert the contents of the first section to the beginning of the third section.
    let sectionToPrepend = doc.sections.at(0);
    section.prependContent(sectionToPrepend);

    // Insert the contents of the second section to the end of the third section.
    let sectionToAppend = doc.sections.at(1);
    section.appendContent(sectionToAppend);
    //ExEnd:AppendSectionContent
  });

  test('CloneSection', () => {
    //ExStart:CloneSection
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document(base.myDir + "Document.docx");
    let cloneSection = doc.sections.at(0).clone();
    //ExEnd:CloneSection
  });

  test('CopySection', () => {
    //ExStart:CopySection
    //GistId:5331edc61a2137fd92565f1e0c953887
    let srcDoc = new aw.Document(base.myDir + "Document.docx");
    let dstDoc = new aw.Document();

    let sourceSection = srcDoc.sections.at(0);
    let newSection = dstDoc.importNode(sourceSection, true).asSection();
    dstDoc.sections.add(newSection);

    dstDoc.save(base.artifactsDir + "WorkingWithSection.CopySection.docx");
    //ExEnd:CopySection
  });

  test('DeleteHeaderFooterContent', () => {
    //ExStart:DeleteHeaderFooterContent
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document(base.myDir + "Document.docx");

    let section = doc.sections.at(0);
    section.clearHeadersFooters();
    //ExEnd:DeleteHeaderFooterContent
  });

  test('DeleteHeaderFooterShapes', () => {
    //ExStart:DeleteHeaderFooterShapes
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document(base.myDir + "Document.docx");

    let section = doc.sections.at(0);
    section.deleteHeaderFooterShapes();
    //ExEnd:DeleteHeaderFooterShapes
  });

  test('DeleteSectionContent', () => {
    //ExStart:DeleteSectionContent
    let doc = new aw.Document(base.myDir + "Document.docx");

    let section = doc.sections.at(0);
    section.clearContent();
    //ExEnd:DeleteSectionContent
  });

  test('ModifyPageSetupInAllSections', () => {
    //ExStart:ModifyPageSetupInAllSections
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Section 1");
    doc.appendChild(new aw.Section(doc));
    builder.writeln("Section 2");
    doc.appendChild(new aw.Section(doc));
    builder.writeln("Section 3");
    doc.appendChild(new aw.Section(doc));
    builder.writeln("Section 4");

    // It is important to understand that a document can contain many sections,
    // and each section has its page setup. In this case, we want to modify them all.
    for (let section of doc) {
      section.asSection().pageSetup.paperSize = aw.PaperSize.Letter;
    }

    doc.save(base.artifactsDir + "WorkingWithSection.ModifyPageSetupInAllSections.doc");
    //ExEnd:ModifyPageSetupInAllSections
  });

  test('SectionsAccessByIndex', () => {
    //ExStart:SectionsAccessByIndex
    let doc = new aw.Document(base.myDir + "Document.docx");

    let section = doc.sections.at(0);
    section.pageSetup.leftMargin = 90; // 3.17 cm
    section.pageSetup.rightMargin = 90; // 3.17 cm
    section.pageSetup.topMargin = 72; // 2.54 cm
    section.pageSetup.bottomMargin = 72; // 2.54 cm
    section.pageSetup.headerDistance = 35.4; // 1.25 cm
    section.pageSetup.footerDistance = 35.4; // 1.25 cm
    section.pageSetup.textColumns.spacing = 35.4; // 1.25 cm
    //ExEnd:SectionsAccessByIndex
  });

  test('SectionChildNodes', () => {
    //ExStart:SectionChildNodes
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("Primary header");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.write("Primary footer");

    let section = doc.firstSection;

    // A Section is a composite node and can contain child nodes,
    // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
    for (let node of section) {
      switch (node.nodeType) {
        case aw.NodeType.Body: {
          let body = node.asBody();

          console.log("Body:");
          console.log(`\t"${body.getText().trim()}"`);
          break;
        }
        case aw.NodeType.HeaderFooter: {
          let headerFooter = node.asHeaderFooter();

          console.log(`HeaderFooter type: ${headerFooter.headerFooterType}:`);
          console.log(`\t"${headerFooter.getText().trim()}"`);
          break;
        }
        default: {
          throw new Error("Unexpected node type in a section.");
        }
      }
    }
    //ExEnd:SectionChildNodes
  });

  test('EnsureMinimum', () => {
    //ExStart:EnsureMinimum
    //GistId:5331edc61a2137fd92565f1e0c953887
    let doc = new aw.Document();

    // If we add a new section like this, it will not have a body, or any other child nodes.
    doc.sections.add(new aw.Section(doc));
    // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
    doc.lastSection.ensureMinimum();

    doc.sections.at(0).body.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));
    //ExEnd:EnsureMinimum
  });

});