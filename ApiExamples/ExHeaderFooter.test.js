// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExHeaderFooter", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Create', () => {
    //ExStart
    //ExFor:HeaderFooter
    //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
    //ExFor:HeaderFooter.headerFooterType
    //ExFor:HeaderFooter.isHeader
    //ExFor:HeaderFooterCollection
    //ExFor:Paragraph.isEndOfHeaderFooter
    //ExFor:Paragraph.parentSection
    //ExFor:Paragraph.parentStory
    //ExFor:Story.appendParagraph
    //ExSummary:Shows how to create a header and a footer.
    let doc = new aw.Document();

    // Create a header and append a paragraph to it. The text in that paragraph
    // will appear at the top of every page of this section, above the main body text.
    let header = new aw.HeaderFooter(doc, aw.HeaderFooterType.HeaderPrimary);
    doc.firstSection.headersFooters.add(header);

    let para = header.appendParagraph("My header.");

    expect(header.isHeader).toEqual(true);
    expect(para.isEndOfHeaderFooter).toEqual(true);

    // Create a footer and append a paragraph to it. The text in that paragraph
    // will appear at the bottom of every page of this section, below the main body text.
    let footer = new aw.HeaderFooter(doc, aw.HeaderFooterType.FooterPrimary);
    doc.firstSection.headersFooters.add(footer);

    para = footer.appendParagraph("My footer.");

    expect(footer.isHeader).toEqual(false);
    expect(para.isEndOfHeaderFooter).toEqual(true);

    expect(para.parentStory.referenceEquals(footer)).toEqual(true);
    expect(para.parentSection.referenceEquals(footer.parentSection)).toEqual(true);
    expect(header.parentSection.referenceEquals(footer.parentSection)).toEqual(true);

    doc.save(base.artifactsDir + "HeaderFooter.create.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "HeaderFooter.create.docx");

    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderPrimary)
      .range.text.includes("My header.")).toEqual(true);
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary)
      .range.text.includes("My footer.")).toEqual(true);
  });


  test('Link', () => {
    //ExStart
    //ExFor:HeaderFooter.isLinkedToPrevious
    //ExFor:HeaderFooterCollection.item(Int32)
    //ExFor:HeaderFooterCollection.linkToPrevious(HeaderFooterType,Boolean)
    //ExFor:HeaderFooterCollection.linkToPrevious(Boolean)
    //ExFor:HeaderFooter.parentSection
    //ExSummary:Shows how to link headers and footers between sections.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Section 1");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 2");
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);
    builder.write("Section 3");

    // Move to the first section and create a header and a footer. By default,
    // the header and the footer will only appear on pages in the section that contains them.
    builder.moveToSection(0);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("This is the header, which will be displayed in sections 1 and 2.");

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.write("This is the footer, which will be displayed in sections 1, 2 and 3.");

    // We can link a section's headers/footers to the previous section's headers/footers
    // to allow the linking section to display the linked section's headers/footers.
    doc.sections.at(1).headersFooters.linkToPrevious(true);

    // Each section will still have its own header/footer objects. When we link sections,
    // the linking section will display the linked section's header/footers while keeping its own.
    expect(doc.sections.at(0).headersFooters.at(0).referenceEquals(
      doc.sections.at(1).headersFooters.at(0))).toEqual(false);
    expect(doc.sections.at(0).headersFooters.at(0).parentSection.referenceEquals(
      doc.sections.at(1).headersFooters.at(0).parentSection)).toEqual(false);

      // Link the headers/footers of the third section to the headers/footers of the second section.
    // The second section already links to the first section's header/footers,
    // so linking to the second section will create a link chain.
    // The first, second, and now the third sections will all display the first section's headers.
    doc.sections.at(2).headersFooters.linkToPrevious(true);

    // We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
    doc.sections.at(2).headersFooters.linkToPrevious(false);

    // We can also select only a specific type of header/footer to link using this method.
    // The third section now will have the same footer as the second and first sections, but not the header.
    doc.sections.at(2).headersFooters.linkToPrevious(aw.HeaderFooterType.FooterPrimary, true);

    // The first section's header/footers cannot link themselves to anything because there is no previous section.
    expect(doc.sections.at(0).headersFooters.count).toEqual(2);
    let count = doc.sections.at(0).headersFooters.toArray().filter((hf) => !hf.isLinkedToPrevious).length;
    expect(count).toEqual(2);

    // All the second section's header/footers are linked to the first section's headers/footers.
    expect(doc.sections.at(1).headersFooters.count).toEqual(6);
    count = doc.sections.at(1).headersFooters.toArray().filter((hf) => hf.isLinkedToPrevious).length;
    expect(count).toEqual(6);

    // In the third section, only the footer is linked to the first section's footer via the second section.
    expect(doc.sections.at(2).headersFooters.count).toEqual(6);
    count = doc.sections.at(2).headersFooters.toArray().filter((hf) => !hf.isLinkedToPrevious).length;
    expect(count).toEqual(5);
    expect(doc.sections.at(2).headersFooters.at(3).isLinkedToPrevious).toEqual(true);

    doc.save(base.artifactsDir + "HeaderFooter.link.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "HeaderFooter.link.docx");

    expect(doc.sections.at(0).headersFooters.count).toEqual(2);
    expect(doc.sections.at(1).headersFooters.count).toEqual(0);
    expect(doc.sections.at(2).headersFooters.count).toEqual(5);
  });


  test('RemoveFooters', () => {
    //ExStart
    //ExFor:Section.headersFooters
    //ExFor:HeaderFooterCollection
    //ExFor:HeaderFooterCollection.item(HeaderFooterType)
    //ExFor:HeaderFooter
    //ExSummary:Shows how to delete all footers from a document.
    let doc = new aw.Document(base.myDir + "Header and footer types.docx");

    // Iterate through each section and remove footers of every kind.
    for (let section of doc.sections.toArray())
    {
      // There are three kinds of footer and header types.
      // 1 -  The "First" header/footer, which only appears on the first page of a section.
      let footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterFirst);
      footer.remove();

      // 2 -  The "Primary" header/footer, which appears on odd pages.
      footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary);
      footer.remove();

      // 3 -  The "Even" header/footer, which appears on even pages. 
      footer = section.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterEven);
      footer.remove();

      let count = section.headersFooters.toArray().filter((hf) => !hf.isHeader).length;
      expect(count).toEqual(0);
    }

    doc.save(base.artifactsDir + "HeaderFooter.RemoveFooters.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "HeaderFooter.RemoveFooters.docx");

    expect(doc.sections.count).toEqual(1);
    expect(doc.firstSection.headersFooters.toArray().filter((hf) => !hf.isHeader).length).toEqual(0);
    expect(doc.firstSection.headersFooters.toArray().filter((hf) => hf.isHeader).length).toEqual(3);
  });


  test('ExportMode', () => {
    //ExStart
    //ExFor:HtmlSaveOptions.exportHeadersFootersMode
    //ExFor:ExportHeadersFootersMode
    //ExSummary:Shows how to omit headers/footers when saving a document to HTML.
    let doc = new aw.Document(base.myDir + "Header and footer types.docx");

    // This document contains headers and footers. We can access them via the "HeadersFooters" collection.
    expect(doc.firstSection.headersFooters.getByHeaderFooterType(aw.HeaderFooterType.HeaderFirst).getText().trim()).toEqual("First header");

    // Formats such as .html do not split the document into pages, so headers/footers will not function the same way
    // they would when we open the document as a .docx using Microsoft Word.
    // If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
    // We can use a SaveOptions object to omit headers/footers while converting to html.
    let saveOptions = new aw.Saving.HtmlSaveOptions(aw.SaveFormat.Html);
    saveOptions.exportHeadersFootersMode = aw.Saving.ExportHeadersFootersMode.None;

    doc.save(base.artifactsDir + "HeaderFooter.ExportMode.html", saveOptions);

    // Open our saved document and verify that it does not contain the header's text
    doc = new aw.Document(base.artifactsDir + "HeaderFooter.ExportMode.html");

    expect(doc.range.text.includes("First header")).toEqual(false);
    //ExEnd
  });


  test('ReplaceText', () => {
    //ExStart
    //ExFor:Document.firstSection
    //ExFor:Section.headersFooters
    //ExFor:HeaderFooterCollection.item(HeaderFooterType)
    //ExFor:HeaderFooter
    //ExFor:Range.replace(String, String, FindReplaceOptions)
    //ExSummary:Shows how to replace text in a document's footer.
    let doc = new aw.Document(base.myDir + "Footer.docx");

    let headersFooters = doc.firstSection.headersFooters;
    let footer = headersFooters.getByHeaderFooterType(aw.HeaderFooterType.FooterPrimary);

    let options = new aw.Replacing.FindReplaceOptions();
    options.matchCase = false;
    options.findWholeWordsOnly = false;

    let currentYear = new Date().getYear();
    footer.range.replace("(C) 2006 Aspose Pty Ltd.", `Copyright (C) ${currentYear} by Aspose Pty Ltd.`, options);

    doc.save(base.artifactsDir + "HeaderFooter.ReplaceText.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "HeaderFooter.ReplaceText.docx");

    expect(doc.range.text.includes(`Copyright (C) ${currentYear} by Aspose Pty Ltd.`)).toEqual(true);
  });


/* TODO IReplacingCallback not supported    
  //ExStart
  //ExFor:IReplacingCallback
  //ExFor:PageSetup.DifferentFirstPageHeaderFooter
  //ExFor:FindReplaceOptions.#ctor(IReplacingCallback)
  //ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
  test.each([false,
    true])('Order', (differentFirstPageHeaderFooter) => {
    let doc = new aw.Document(base.myDir + "Header and footer types.docx");

    let firstPageSection = doc.firstSection;

    let logger = new ReplaceLog();
    let options = new aw.Replacing.FindReplaceOptions(logger);

    // Using a different header/footer for the first page will affect the search order.
    firstPageSection.pageSetup.differentFirstPageHeaderFooter = differentFirstPageHeaderFooter;
    doc.range.replace(new Regex("(header|footer)"), "", options);

    if (differentFirstPageHeaderFooter)
      expect(logger.text.replace("\r", "")).toEqual("First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n");
    else
      expect(logger.text.replace("\r", "")).toEqual("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n");
  });


    /// <summary>
    /// During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
    /// in the state it is in before the replacement takes place.
    /// This will display the order in which the text replacement operation traverses nodes.
    /// </summary>
  private class ReplaceLog : IReplacingCallback
  {
    public ReplaceAction Replacing(ReplacingArgs args)
    {
      mTextBuilder.AppendLine(args.matchNode.getText());
      return aw.Replacing.ReplaceAction.Skip;
    }

    internal string Text => mTextBuilder.toString();

    private readonly StringBuilder mTextBuilder = new StringBuilder();
  }
    //ExEnd
*/
});
