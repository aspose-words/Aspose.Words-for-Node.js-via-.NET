// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExSection", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Protect', () => {
    //ExStart
    //ExFor:Document.protect(ProtectionType)
    //ExFor:ProtectionType
    //ExFor:Section.protectedForForms
    //ExSummary:Shows how to turn off protection for a section.
    let doc = new aw.Document();

    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Section 1. Hello world!");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    builder.writeln("Section 2. Hello again!");
    builder.write("Please enter text here: ");
    builder.insertTextInput("TextInput1", aw.Fields.TextFormFieldType.Regular, "", "Placeholder text", 0);

    // Apply write protection to every section in the document.
    doc.protect(aw.ProtectionType.AllowOnlyFormFields);

    // Turn off write protection for the first section.
    doc.sections.at(0).protectedForForms = false;

    // In this output document, we will be able to edit the first section freely,
    // and we will only be able to edit the contents of the form field in the second section.
    doc.save(base.artifactsDir + "Section.protect.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Section.protect.docx");

    expect(doc.sections.at(0).protectedForForms).toEqual(false);
    expect(doc.sections.at(1).protectedForForms).toEqual(true);
  });


  test('AddRemove', () => {
    //ExStart
    //ExFor:Document.sections
    //ExFor:Section.clone
    //ExFor:SectionCollection
    //ExFor:NodeCollection.removeAt(Int32)
    //ExSummary:Shows how to add and remove sections in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");

    expect(doc.getText().trim()).toEqual("Section 1\u000cSection 2");

    // Delete the first section from the document.
    doc.sections.removeAt(0);

    expect(doc.getText().trim()).toEqual("Section 2");

    // Append a copy of what is now the first section to the end of the document.
    let lastSectionIdx = doc.sections.count - 1;
    let newSection = doc.sections.at(lastSectionIdx).clone();
    doc.sections.add(newSection);

    expect(doc.getText().trim()).toEqual("Section 2\u000cSection 2");
    //ExEnd
  });


  test('FirstAndLast', () => {
    //ExStart
    //ExFor:Document.firstSection
    //ExFor:Document.lastSection
    //ExSummary:Shows how to create a new section with a document builder.
    let doc = new aw.Document();

    // A blank document contains one section by default,
    // which contains child nodes that we can edit.
    expect(doc.sections.count).toEqual(1);

    // Use a document builder to add text to the first section.
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    // Create a second section by inserting a section break.
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    expect(doc.sections.count).toEqual(2);

    // Each section has its own page setup settings.
    // We can split the text in the second section into two columns.
    // This will not affect the text in the first section.
    doc.lastSection.pageSetup.textColumns.setCount(2);
    builder.writeln("Column 1.");
    builder.insertBreak(aw.BreakType.ColumnBreak);
    builder.writeln("Column 2.");

    expect(doc.firstSection.pageSetup.textColumns.count).toEqual(1);
    expect(doc.lastSection.pageSetup.textColumns.count).toEqual(2);

    doc.save(base.artifactsDir + "Section.create.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Section.create.docx");

    expect(doc.firstSection.pageSetup.textColumns.count).toEqual(1);
    expect(doc.lastSection.pageSetup.textColumns.count).toEqual(2);
  });


  test('CreateManually', () => {
    //ExStart
    //ExFor:Node.getText
    //ExFor:CompositeNode.removeAllChildren
    //ExFor:CompositeNode.appendChild``1(``0)
    //ExFor:Section
    //ExFor:Section.#ctor
    //ExFor:Section.pageSetup
    //ExFor:PageSetup.sectionStart
    //ExFor:PageSetup.paperSize
    //ExFor:SectionStart
    //ExFor:PaperSize
    //ExFor:Body
    //ExFor:Body.#ctor
    //ExFor:Paragraph
    //ExFor:Paragraph.#ctor
    //ExFor:Paragraph.paragraphFormat
    //ExFor:ParagraphFormat
    //ExFor:ParagraphFormat.styleName
    //ExFor:ParagraphFormat.alignment
    //ExFor:ParagraphAlignment
    //ExFor:Run
    //ExFor:Run.#ctor(DocumentBase)
    //ExFor:Run.text
    //ExFor:Inline.font
    //ExSummary:Shows how to construct an Aspose.words document by hand.
    let doc = new aw.Document();

    // A blank document contains one section, one body and one paragraph.
    // Call the "RemoveAllChildren" method to remove all those nodes,
    // and end up with a document node with no children.
    doc.removeAllChildren();

    // This document now has no composite child nodes that we can add content to.
    // If we wish to edit it, we will need to repopulate its node collection.
    // First, create a new section, and then append it as a child to the root document node.
    let section = new aw.Section(doc);
    doc.appendChild(section);

    // Set some page setup properties for the section.
    section.pageSetup.sectionStart = aw.SectionStart.NewPage;
    section.pageSetup.paperSize = aw.PaperSize.Letter;

    // A section needs a body, which will contain and display all its contents
    // on the page between the section's header and footer.
    let body = new aw.Body(doc);
    section.appendChild(body);

    // Create a paragraph, set some formatting properties, and then append it as a child to the body.
    let para = new aw.Paragraph(doc);

    para.paragraphFormat.styleName = "Heading 1";
    para.paragraphFormat.alignment = aw.ParagraphAlignment.Center;

    body.appendChild(para);

    // Finally, add some content to do the document. Create a run,
    // set its appearance and contents, and then append it as a child to the paragraph.
    let run = new aw.Run(doc);
    run.text = "Hello World!";
    run.font.color = "#FF0000";
    para.appendChild(run);

    expect(doc.getText().trim()).toEqual("Hello World!");

    doc.save(base.artifactsDir + "Section.CreateManually.docx");
    //ExEnd
  });


  test('EnsureMinimum', () => {
    //ExStart
    //ExFor:NodeCollection.add
    //ExFor:Section.ensureMinimum
    //ExFor:SectionCollection.item(Int32)
    //ExSummary:Shows how to prepare a new section node for editing.
    let doc = new aw.Document();

    // A blank document comes with a section, which has a body, which in turn has a paragraph.
    // We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
    expect(doc.getChild(aw.NodeType.Any, 0, true).nodeType).toEqual(aw.NodeType.Section);
    expect(doc.sections.at(0).getChild(aw.NodeType.Any, 0, true).nodeType).toEqual(aw.NodeType.Body);
    expect(doc.sections.at(0).body.getChild(aw.NodeType.Any, 0, true).nodeType).toEqual(aw.NodeType.Paragraph);

    // If we add a new section like this, it will not have a body, or any other child nodes.
    doc.sections.add(new aw.Section(doc));

    expect(doc.sections.at(1).getChildNodes(aw.NodeType.Any, true).count).toEqual(0);

    // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
    doc.lastSection.ensureMinimum();

    expect(doc.sections.at(1).getChild(aw.NodeType.Any, 0, true).nodeType).toEqual(aw.NodeType.Body);
    expect(doc.sections.at(1).body.getChild(aw.NodeType.Any, 0, true).nodeType).toEqual(aw.NodeType.Paragraph);

    doc.sections.at(0).body.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));

    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('BodyEnsureMinimum', () => {
    //ExStart
    //ExFor:Section.body
    //ExFor:Body.ensureMinimum
    //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
    let doc = new aw.Document();

    // A blank document contains one section, one body and one paragraph.
    // Call the "RemoveAllChildren" method to remove all those nodes,
    // and end up with a document node with no children.
    doc.removeAllChildren();

    // This document now has no composite child nodes that we can add content to.
    // If we wish to edit it, we will need to repopulate its node collection.
    // First, create a new section, and then append it as a child to the root document node.
    let section = new aw.Section(doc);
    doc.appendChild(section);

    // A section needs a body, which will contain and display all its contents
    // on the page between the section's header and footer.
    let body = new aw.Body(doc);
    section.appendChild(body);

    // This body has no children, so we cannot add runs to it yet.
    expect(doc.firstSection.body.getChildNodes(aw.NodeType.Any, true).count).toEqual(0);

    // Call the "EnsureMinimum" to make sure that this body contains at least one empty paragraph. 
    body.ensureMinimum();

    // Now, we can add runs to the body, and get the document to display them.
    body.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));

    expect(doc.getText().trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('BodyChildNodes', () => {
    //ExStart
    //ExFor:Body.nodeType
    //ExFor:HeaderFooter.nodeType
    //ExFor:Document.firstSection
    //ExSummary:Shows how to iterate through the children of a composite node.
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
    for (let node of section)
    {
      switch (node.nodeType)
      {
        case aw.NodeType.Body:
        {
          let body = node.asBody();

          console.log("Body:");
          console.log(`\t\"${body.getText().trim()}\"`);
          break;
        }
        case aw.NodeType.HeaderFooter:
        {
          let headerFooter = node.asHeaderFooter();

          console.log(`HeaderFooter type: ${headerFooter.headerFooterType}:`);
          console.log(`\t\"${headerFooter.getText().trim()}\"`);
          break;
        }
        default:
        {
          throw new Error("Unexpected node type in a section.");
        }
      }
    }
    //ExEnd
  });


  test('Clear', () => {
    //ExStart
    //ExFor:NodeCollection.clear
    //ExSummary:Shows how to remove all sections from a document.
    let doc = new aw.Document(base.myDir + "Document.docx");

    // This document has one section with a few child nodes containing and displaying all the document's contents.
    expect(doc.sections.count).toEqual(1);
    expect(doc.sections.at(0).getChildNodes(aw.NodeType.Any, true).count).toEqual(17);
    expect(doc.getText().trim()).toEqual("Hello World!\r\rHello Word!\r\r\rHello World!");

    // Clear the collection of sections, which will remove all of the document's children.
    doc.sections.clear();

    expect(doc.getChildNodes(aw.NodeType.Any, true).count).toEqual(0);
    expect(doc.getText().trim()).toEqual('');
    //ExEnd
  });


  test('PrependAppendContent', () => {
    //ExStart
    //ExFor:Section.appendContent
    //ExFor:Section.prependContent
    //ExSummary:Shows how to append the contents of a section to another section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 3");

    let section = doc.sections.at(2);

    expect(section.getText()).toEqual("Section 3" + aw.ControlChar.sectionBreak);

    // Insert the contents of the first section to the beginning of the third section.
    let sectionToPrepend = doc.sections.at(0);
    section.prependContent(sectionToPrepend);

    // Insert the contents of the second section to the end of the third section.
    let sectionToAppend = doc.sections.at(1);
    section.appendContent(sectionToAppend);

    // The "PrependContent" and "AppendContent" methods did not create any new sections.
    expect(doc.sections.count).toEqual(3);
    expect(section.getText()).toEqual("Section 1" + aw.ControlChar.paragraphBreak +
      "Section 3" + aw.ControlChar.paragraphBreak +
      "Section 2" + aw.ControlChar.sectionBreak);
    //ExEnd
  });


  test('ClearContent', () => {
    //ExStart
    //ExFor:Section.clearContent
    //ExSummary:Shows how to clear the contents of a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Hello world!");

    expect(doc.getText().trim()).toEqual("Hello world!");
    expect(doc.firstSection.body.paragraphs.count).toEqual(1);

    // Running the "ClearContent" method will remove all the section contents
    // but leave a blank paragraph to add content again.
    doc.firstSection.clearContent();

    expect(doc.getText().trim()).toEqual('');
    expect(doc.firstSection.body.paragraphs.count).toEqual(1);
    //ExEnd
  });


  test('ClearHeadersFooters', () => {
    //ExStart
    //ExFor:Section.clearHeadersFooters
    //ExSummary:Shows how to clear the contents of all headers and footers in a section.
    let doc = new aw.Document();

    expect(doc.firstSection.headersFooters.count).toEqual(0);

    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.writeln("This is the primary header.");
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.writeln("This is the primary footer.");

    expect(doc.firstSection.headersFooters.count).toEqual(2);

    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getText().trim()).toEqual("This is the primary header.");
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary).getText().trim()).toEqual("This is the primary footer.");

    // Empty all the headers and footers in this section of all their contents.
    // The headers and footers themselves will still be present but will have nothing to display.
    doc.firstSection.clearHeadersFooters();

    expect(doc.firstSection.headersFooters.count).toEqual(2);

    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getText().trim()).toEqual('');
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary).getText().trim()).toEqual('');
    //ExEnd
  });


  test('DeleteHeaderFooterShapes', () => {
    //ExStart
    //ExFor:Section.deleteHeaderFooterShapes
    //ExSummary:Shows how to remove all shapes from all headers footers in a section.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a primary header with a shape.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.insertShape(aw.Drawing.ShapeType.Rectangle, 100, 100);

    // Create a primary footer with an image.
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.insertImage(base.imageDir + "Logo icon.ico");

    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getChildNodes(aw.NodeType.Shape, true).count).toEqual(1);
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary).getChildNodes(aw.NodeType.Shape, true).count).toEqual(1);

    // Remove all shapes from the headers and footers in the first section.
    doc.firstSection.deleteHeaderFooterShapes();

    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary).getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary).getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);
    //ExEnd
  });


  test('SectionsCloneSection', () => {
    let doc = new aw.Document(base.myDir + "Document.docx");
    let cloneSection = doc.sections.at(0).clone();
  });


  test('SectionsImportSection', () => {
    let srcDoc = new aw.Document(base.myDir + "Document.docx");
    let dstDoc = new aw.Document();

    let sourceSection = srcDoc.sections.at(0);
    let newSection = dstDoc.importNode(sourceSection, true).asSection();
    dstDoc.sections.add(newSection);
  });


  test('MigrateFrom2XImportSection', () => {
    let srcDoc = new aw.Document();
    let dstDoc = new aw.Document();

    let sourceSection = srcDoc.sections.at(0);
    let newSection = dstDoc.importNode(sourceSection, true).asSection();
    dstDoc.sections.add(newSection);
  });


  test('ModifyPageSetupInAllSections', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");

    // It is important to understand that a document can contain many sections,
    // and each section has its page setup. In this case, we want to modify them all.
    for (let node of doc.getChildNodes(aw.NodeType.Section, true))
    {
      let section = node.asSection();
      section.pageSetup.paperSize = aw.PaperSize.Letter;
    }

    doc.save(base.artifactsDir + "Section.ModifyPageSetupInAllSections.doc");
  });


  test.skip('CultureInfoPageSetupDefaults - we don\'t support of setting culture of current thread in .NET code now', () => {
    //Thread.currentThread.CurrentCulture = new CultureInfo("en-us");
    let docEn = new aw.Document();

    // Assert that page defaults comply with current culture info.
    let sectionEn = docEn.sections.at(0).asSection();
    expect(sectionEn.pageSetup.leftMargin).toEqual(72.0);
    expect(sectionEn.pageSetup.rightMargin).toEqual(72.0);
    expect(sectionEn.pageSetup.topMargin).toEqual(72.0);
    expect(sectionEn.pageSetup.bottomMargin).toEqual(72.0);
    expect(sectionEn.pageSetup.headerDistance).toEqual(36.0);
    expect(sectionEn.pageSetup.footerDistance).toEqual(36.0);
    expect(sectionEn.pageSetup.textColumns.spacing).toEqual(36.0);

    // Change the culture and assert that the page defaults are changed.
    //Thread.currentThread.CurrentCulture = new CultureInfo("de-de");

    let docDe = new aw.Document();

    let sectionDe = docDe.sections.at(0).asSection();
    expect(sectionDe.pageSetup.leftMargin).toEqual(70.85);
    expect(sectionDe.pageSetup.rightMargin).toEqual(70.85);
    expect(sectionDe.pageSetup.topMargin).toEqual(70.85);
    expect(sectionDe.pageSetup.bottomMargin).toEqual(56.7);
    expect(sectionDe.pageSetup.headerDistance).toEqual(35.4);
    expect(sectionDe.pageSetup.footerDistance).toEqual(35.4);
    expect(sectionDe.pageSetup.textColumns.spacing).toEqual(35.4);

    // Change page defaults.
    sectionDe.pageSetup.leftMargin = 90; // 3.17 cm
    sectionDe.pageSetup.rightMargin = 90; // 3.17 cm
    sectionDe.pageSetup.topMargin = 72; // 2.54 cm
    sectionDe.pageSetup.bottomMargin = 72; // 2.54 cm
    sectionDe.pageSetup.headerDistance = 35.4; // 1.25 cm
    sectionDe.pageSetup.footerDistance = 35.4; // 1.25 cm
    sectionDe.pageSetup.textColumns.spacing = 35.4; // 1.25 cm

    docDe = DocumentHelper.saveOpen(docDe);

    let sectionDeAfter = docDe.sections.at(0).asSection();
    expect(sectionDeAfter.pageSetup.leftMargin).toEqual(90.0);
    expect(sectionDeAfter.pageSetup.rightMargin).toEqual(90.0);
    expect(sectionDeAfter.pageSetup.topMargin).toEqual(72.0);
    expect(sectionDeAfter.pageSetup.bottomMargin).toEqual(72.0);
    expect(sectionDeAfter.pageSetup.headerDistance).toEqual(35.4);
    expect(sectionDeAfter.pageSetup.footerDistance).toEqual(35.4);
    expect(sectionDeAfter.pageSetup.textColumns.spacing).toEqual(35.4);
  });


  test('PreserveWatermarks', () => {
    //ExStart:PreserveWatermarks
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:Section.clearHeadersFooters(bool)
    //ExSummary:Shows how to clear the contents of header and footer with or without a watermark.
    let doc = new aw.Document(base.myDir + "Header and footer types.docx");

    // Add a plain text watermark.
    doc.watermark.setText("Aspose Watermark");

    // Make sure the headers and footers have content.
    let headersFooters = doc.firstSection.headersFooters;
    expect(headersFooters.at(aw.HeaderFooterType.HeaderFirst).getText().trim()).toEqual("First header");
    expect(headersFooters.at(aw.HeaderFooterType.HeaderEven).getText().trim()).toEqual("Second header");
    expect(headersFooters.at(aw.HeaderFooterType.HeaderPrimary).getText().trim()).toEqual("Third header");
    expect(headersFooters.at(aw.HeaderFooterType.FooterFirst).getText().trim()).toEqual("First footer");
    expect(headersFooters.at(aw.HeaderFooterType.FooterEven).getText().trim()).toEqual("Second footer");
    expect(headersFooters.at(aw.HeaderFooterType.FooterPrimary).getText().trim()).toEqual("Third footer");

    // Removes all header and footer content except watermarks.
    doc.firstSection.clearHeadersFooters(true);

    headersFooters = doc.firstSection.headersFooters;
    expect(headersFooters.at(aw.HeaderFooterType.HeaderFirst).getText().trim()).toEqual("");
    expect(headersFooters.at(aw.HeaderFooterType.HeaderEven).getText().trim()).toEqual("");
    expect(headersFooters.at(aw.HeaderFooterType.HeaderPrimary).getText().trim()).toEqual("");
    expect(headersFooters.at(aw.HeaderFooterType.FooterFirst).getText().trim()).toEqual("");
    expect(headersFooters.at(aw.HeaderFooterType.FooterEven).getText().trim()).toEqual("");
    expect(headersFooters.at(aw.HeaderFooterType.FooterPrimary).getText().trim()).toEqual("");
    expect(doc.watermark.type).toEqual(aw.WatermarkType.Text);

    // Removes all header and footer content including watermarks.
    doc.firstSection.clearHeadersFooters(false);
    expect(doc.watermark.type).toEqual(aw.WatermarkType.None);
    //ExEnd:PreserveWatermarks
  });


});
