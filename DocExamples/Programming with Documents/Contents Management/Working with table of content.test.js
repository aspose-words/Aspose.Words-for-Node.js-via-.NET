// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithTableOfContent", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('ChangeStyleOfTocLevel', () => {
        //ExStart:ChangeStyleOfTocLevel
        //GistId:e0ccef8441be6a8e2de5810acdefd25a
        let doc = new aw.Document();
        // Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.styles.at(aw.StyleIdentifier.Toc1).font.bold = true;
        //ExEnd:ChangeStyleOfTocLevel
    });

    test('ChangeTocTabStops', () => {
        //ExStart:ChangeTocTabStops
        //GistId:e0ccef8441be6a8e2de5810acdefd25a
        let doc = new aw.Document(base.myDir + "Table of contents.docx");
        let paragraphs = doc.getChildNodes(aw.NodeType.Paragraph, true);

        for (let i = 0; i < paragraphs.count; i++) {
            let para = paragraphs.at(i).asParagraph();

            // Check if this paragraph is formatted using the TOC result based styles.
            // This is any style between TOC and TOC9.
            if (para.paragraphFormat.style.styleIdentifier >= aw.StyleIdentifier.Toc1 &&
                para.paragraphFormat.style.styleIdentifier <= aw.StyleIdentifier.Toc9) {

                // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                let tab = para.paragraphFormat.tabStops.at(0);

                // Remove the old tab from the collection.
                para.paragraphFormat.tabStops.removeByPosition(tab.position);

                // Insert a new tab using the same properties but at a modified position.
                // We could also change the separators used (dots) by passing a different Leader type.
                para.paragraphFormat.tabStops.add(tab.position - 50, tab.alignment, tab.leader);
            }
        }
        doc.save(base.artifactsDir + "WorkingWithTableOfContent.ChangeTocTabStops.docx");
        //ExEnd:ChangeTocTabStops
    });

    test('ExtractToc', () => {
        //ExStart:ExtractToc
        //GistId:e0ccef8441be6a8e2de5810acdefd25a
        let doc = new aw.Document(base.myDir + "Table of contents.docx");
        let fields = doc.range.fields;

        for (let field of fields) {
            if (field.type === aw.Fields.FieldType.FieldHyperlink) {
                let hyperlink = field;
                if (hyperlink.subAddress != null && hyperlink.subAddress.startsWith("_Toc")) {
                    let tocItem = field.start.getAncestor(aw.NodeType.Paragraph);
                    console.log(tocItem.toString(aw.SaveFormat.Text).trim());
                    console.log("------------------");
                    if (tocItem != null) {
                        let bm = doc.range.bookmarks.at(hyperlink.subAddress);
                        let pointer = bm.bookmarkStart.getAncestor(aw.NodeType.Paragraph);
                        console.log(pointer.toString(aw.SaveFormat.Text));
                    }
                }
            }
        }
        //ExEnd:ExtractToc
    });
});