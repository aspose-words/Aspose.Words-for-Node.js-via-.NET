// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("JoinAndAppendDocuments", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });


    test('SimpleAppendDocument', () => {
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Append the source document to the destination document using no extra options.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.SimpleAppendDocument.docx");
    });

    test('AppendDocument', () => {
        //ExStart:AppendDocumentManually
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Loop through all sections in the source document.
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        for (let srcSection of srcDoc.sections) {
            // Because we are copying a section from one document to another,
            // it is required to import the Section node into the destination document.
            // This adjusts any document-specific references to styles, lists, etc.
            //
            // Importing a node creates a copy of the original node, but the copy
            // is ready to be inserted into the destination document.
            let dstSection = dstDoc.importNode(srcSection, true, aw.ImportFormatMode.KeepSourceFormatting);

            // Now the new section node can be appended to the destination document.
            dstDoc.appendChild(dstSection);
        }

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.AppendDocument.docx");
        //ExEnd:AppendDocumentManually
    });

    test('AppendDocumentToBlank', () => {
        //ExStart:AppendDocumentToBlank
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document();

        // The destination document is not empty, often causing a blank page to appear before the appended document.
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc.removeAllChildren();
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.AppendDocumentToBlank.docx");
        //ExEnd:AppendDocumentToBlank
    });

    test('AppendWithImportFormatOptions', () => {
        //ExStart:AppendWithImportFormatOptions
        let srcDoc = new aw.Document(base.myDir + "Document source with list.docx");
        let dstDoc = new aw.Document(base.myDir + "Document destination with list.docx");

        // Specify that if numbering clashes in source and destination documents,
        // then numbering from the source document will be used.
        let options = new aw.ImportFormatOptions();
        options.keepSourceNumbering = true;

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles, options);
        //ExEnd:AppendWithImportFormatOptions
    });

    test('ConvertNumPageFields', () => {
        //ExStart:ConvertNumPageFields
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Restart the page numbering on the start of the source document.
        srcDoc.firstSection.pageSetup.restartPageNumbering = true;
        srcDoc.firstSection.pageSetup.pageStartingNumber = 1;

        // Append the source document to the end of the destination document.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        // After joining the documents the NUMPAGE fields will now display the total number of pages which
        // is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
        convertNumPageFieldsToPageRef(dstDoc);

        // This needs to be called in order to update the new fields with page numbers.
        dstDoc.updatePageLayout();

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.ConvertNumPageFields.docx");
        //ExEnd:ConvertNumPageFields
    });

    //ExStart:ConvertNumPageFieldsToPageRef
    function convertNumPageFieldsToPageRef(doc) {
        // This is the prefix for each bookmark, which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        let bookmarkPrefix = "_SubDocumentEnd";
        let numPagesFieldName = "NUMPAGES";
        let pageRefFieldName = "PAGEREF";

        // Defines the number of page restarts encountered and, therefore,
        // the number of "sub" documents found within this document.
        let subDocumentCount = 0;

        let builder = new aw.DocumentBuilder(doc);

        for (let section of doc.sections) {
            section = section.asSection();
            // This section has its page numbering restarted to treat this as the start of a sub-document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section.pageSetup.restartPageNumbering) {
                // Don't do anything if this is the first section of the document.
                // This part of the code will insert the bookmark marking the end of the previous sub-document so,
                // therefore, it does not apply to the first section in the document.
                if (!base.compareNodes(section, doc.firstSection)) {
                    // Get the previous section and the last node within the body of that section.
                    let prevSection = section.previousSibling.asSection();
                    let lastNode = prevSection.body.lastChild;

                    builder.moveTo(lastNode);

                    // This bookmark represents the end of the sub-document.
                    builder.startBookmark(bookmarkPrefix + subDocumentCount);
                    builder.endBookmark(bookmarkPrefix + subDocumentCount);

                    // Increase the sub-document count to insert the correct bookmarks.
                    subDocumentCount++;
                }
            }

            // The last section needs the ending bookmark to signal that it is the end of the current sub-document.
            if (base.compareNodes(section, doc.lastSection)) {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                let lastNode = doc.lastSection.body.lastChild;

                builder.moveTo(lastNode);
                builder.startBookmark(bookmarkPrefix + subDocumentCount);
                builder.endBookmark(bookmarkPrefix + subDocumentCount);
            }

            // Iterate through each NUMPAGES field in the section and replace it with a PAGEREF field
            // referring to the bookmark of the current sub-document. This bookmark is positioned at the end
            // of the sub-document but does not exist yet. It is inserted when a section with restart page numbering
            // or the last section is encountered.
            let nodes = section.getChildNodes(aw.NodeType.FieldStart, true);

            for (let fieldStart of nodes) {
                fieldStart = fieldStart.asFieldStart();
                if (fieldStart.fieldType === aw.Fields.FieldType.FieldNumPages) {
                    let fieldCode = getFieldCode(fieldStart);
                    // Since the NUMPAGES field does not take any additional parameters,
                    // we can assume the field's remaining part. Code after the field name is the switches.
                    // We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    let fieldSwitches = fieldCode.replace(numPagesFieldName, "").trim();

                    // Inserting the new field directly at the FieldStart node of the original field will cause
                    // the new field not to pick up the original field's formatting. To counter this,
                    // insert the field just before the original field if a previous run cannot be found,
                    // we are forced to use the FieldStart node.
                    let previousNode = fieldStart.previousSibling || fieldStart;

                    // Insert a PAGEREF field at the same position as the field.
                    builder.moveTo(previousNode);

                    let newField = builder.insertField(
                        ` ${pageRefFieldName} ${bookmarkPrefix}${subDocumentCount} ${fieldSwitches} `);

                    // The field will be inserted before the referenced node. Move the node before the field instead.
                    previousNode.parentNode.insertBefore(previousNode, newField.start);

                    // Remove the original NUMPAGES field from the document.
                    removeNumPageField(fieldStart);
                }
            }
        }
    }
    //ExEnd:ConvertNumPageFieldsToPageRef

    //ExStart:RemoveNumPageField
    function removeNumPageField(fieldStart) {
        let isRemoving = true;

        let currentNode = fieldStart;
        while (currentNode !== null && isRemoving) {
            if (currentNode.nodeType === aw.NodeType.FieldEnd)
                isRemoving = false;

            let nextNode = currentNode.nextPreOrder(currentNode.document);
            currentNode.remove();
            currentNode = nextNode;
        }
    }

    function getFieldCode(fieldStart) {
        let builder = [];

        for (let node = fieldStart;
             node !== null && node.nodeType !== aw.NodeType.FieldSeparator &&
             node.nodeType !== aw.NodeType.FieldEnd;
             node = node.nextPreOrder(node.document)) {
            // Use text only of Run nodes to avoid duplication.
            if (node.nodeType === aw.NodeType.Run)
                builder.push(node.getText());
        }

        return builder.join('');
    }
    //ExEnd:RemoveNumPageField

    test('DifferentPageSetup', () => {
        //ExStart:DifferentPageSetup
        //GistId:814f45acd0c15059a9680cb661081d0f
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.Continuous;

        // Restart the page numbering on the start of the source document.
        srcDoc.firstSection.pageSetup.restartPageNumbering = true;
        srcDoc.firstSection.pageSetup.pageStartingNumber = 1;

        // To ensure this does not happen when the source document has different page setup settings, make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document,
        // this will need to be repeated for those sections.
        srcDoc.firstSection.pageSetup.pageWidth = dstDoc.lastSection.pageSetup.pageWidth;
        srcDoc.firstSection.pageSetup.pageHeight = dstDoc.lastSection.pageSetup.pageHeight;
        srcDoc.firstSection.pageSetup.orientation = dstDoc.lastSection.pageSetup.orientation;

        // Iterate through all sections in the source document.
        let paragraphs = srcDoc.getChildNodes(aw.NodeType.Paragraph, true);
        for (let para of paragraphs) {
            para = para.asParagraph();
            para.paragraphFormat.keepWithNext = true;
        }

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.DifferentPageSetup.docx");
        //ExEnd:DifferentPageSetup
    });

    test('JoinContinuous', () => {
        //ExStart:JoinContinuous
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Make the document appear straight after the destination documents content.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.Continuous;
        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.JoinContinuous.docx");
        //ExEnd:JoinContinuous
    });

    test('JoinNewPage', () => {
        //ExStart:JoinNewPage
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Set the appended document to start on a new page.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.NewPage;
        // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.JoinNewPage.docx");
        //ExEnd:JoinNewPage
    });

    test('KeepSourceFormatting', () => {
        //ExStart:KeepSourceFormatting
        //GistId:814f45acd0c15059a9680cb661081d0f
        let dstDoc = new aw.Document();
        dstDoc.firstSection.body.appendParagraph("Destination document text. ");

        let srcDoc = new aw.Document();
        srcDoc.firstSection.body.appendParagraph("Source document text. ");

        // Append the source document to the destination document.
        // Pass format mode to retain the original formatting of the source document when importing it.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.KeepSourceFormatting.docx");
        //ExEnd:KeepSourceFormatting
    });

    test('KeepSourceTogether', () => {
        //ExStart:KeepSourceTogether
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Document destination with list.docx");

        // Set the source document to appear straight after the destination document's content.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.Continuous;

        let paragraphs = srcDoc.getChildNodes(aw.NodeType.Paragraph, true);
        for (let para of paragraphs) {
            para = para.asParagraph();
            para.paragraphFormat.keepWithNext = true;
        }

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.KeepSourceTogether.docx");
        //ExEnd:KeepSourceTogether
    });

    test('ListKeepSourceFormatting', () => {
        //ExStart:ListKeepSourceFormatting
        let srcDoc = new aw.Document(base.myDir + "Document source with list.docx");
        let dstDoc = new aw.Document(base.myDir + "Document destination with list.docx");

        // Append the content of the document so it flows continuously.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.Continuous;

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.ListKeepSourceFormatting.docx");
        //ExEnd:ListKeepSourceFormatting
    });

    test('ListUseDestinationStyles', () => {
        //ExStart:ListUseDestinationStyles
        let srcDoc = new aw.Document(base.myDir + "Document source with list.docx");
        let dstDoc = new aw.Document(base.myDir + "Document destination with list.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.Continuous;

        // Keep track of the lists that are created.
        let newLists = new Map();

        let paragraphs = srcDoc.getChildNodes(aw.NodeType.Paragraph, true);
        for (let para of paragraphs) {
            para = para.asParagraph();
            if (para.isListItem) {
                let listId = para.listFormat.list.listId;

                // Check if the destination document contains a list with this ID already. If it does, then this may
                // cause the two lists to run together. Create a copy of the list in the source document instead.
                if (dstDoc.lists.getListByListId(listId) !== null) {
                    let currentList;
                    // A newly copied list already exists for this ID, retrieve the stored list,
                    // and use it on the current paragraph.
                    if (newLists.has(listId)) {
                        currentList = newLists.get(listId);
                    } else {
                        // Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc.lists.addCopy(para.listFormat.list);
                        newLists.set(listId, currentList);
                    }

                    // Set the list of this paragraph to the copied list.
                    para.listFormat.list = currentList;
                }
            }
        }

        // Append the source document to end of the destination document.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.ListUseDestinationStyles.docx");
        //ExEnd:ListUseDestinationStyles
    });

    test('RestartPageNumbering', () => {
        //ExStart:RestartPageNumbering
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.NewPage;
        srcDoc.firstSection.pageSetup.restartPageNumbering = true;

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.RestartPageNumbering.docx");
        //ExEnd:RestartPageNumbering
    });

    test('UpdatePageLayout', () => {
        //ExStart:UpdatePageLayout
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // If the destination document is rendered to PDF, image etc.
        // or UpdatePageLayout is called before the source document. Is appended,
        // then any changes made after will not be reflected in the rendered output
        dstDoc.updatePageLayout();

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again, the appended document will not appear in the output of the next rendering.
        dstDoc.updatePageLayout();

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.UpdatePageLayout.docx");
        //ExEnd:UpdatePageLayout
    });

    test('UseDestinationStyles', () => {
        //ExStart:UseDestinationStyles
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Append the source document using the styles of the destination document.
        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.UseDestinationStyles.docx");
        //ExEnd:UseDestinationStyles
    });

    test('SmartStyleBehavior', () => {
        //ExStart:SmartStyleBehavior
        let dstDoc = new aw.Document();
        let builder = new aw.DocumentBuilder(dstDoc);

        let myStyle = builder.document.styles.add(aw.StyleType.Paragraph, "MyStyle");
        myStyle.font.size = 14;
        myStyle.font.name = "Courier New";
        myStyle.font.color = "#0000FF"; // Blue

        builder.paragraphFormat.styleName = myStyle.name;
        builder.writeln("Hello world!");

        // Clone the document and edit the clone's "MyStyle" style, so it is a different color than that of the original.
        // If we insert the clone into the original document, the two styles with the same name will cause a clash.
        let srcDoc = dstDoc.clone();
        srcDoc.styles.at("MyStyle").font.color = "#FF0000"; // Red

        // When we enable SmartStyleBehavior and use the KeepSourceFormatting import format mode,
        // Aspose.Words will resolve style clashes by converting source document styles.
        // with the same names as destination styles into direct paragraph attributes.
        let options = new aw.ImportFormatOptions();
        options.smartStyleBehavior = true;

        builder.insertDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.SmartStyleBehavior.docx");
        //ExEnd:SmartStyleBehavior
    });

    test('InsertDocument', () => {
        //ExStart:InsertDocumentWithBuilder
        //GistId:814f45acd0c15059a9680cb661081d0f
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");
        let builder = new aw.DocumentBuilder(dstDoc);

        builder.moveToDocumentEnd();
        builder.insertBreak(aw.BreakType.PageBreak);

        builder.insertDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);
        builder.document.save(base.artifactsDir + "JoinAndAppendDocuments.InsertDocument.docx");
        //ExEnd:InsertDocumentWithBuilder
    });

    test('InsertDocumentInline', () => {
        //ExStart:InsertDocumentInlineWithBuilder
        //GistId:814f45acd0c15059a9680cb661081d0f
        let srcDoc = new aw.DocumentBuilder();
        srcDoc.write("[src content]");

        // Create destination document.
        let dstDoc = new aw.DocumentBuilder();
        dstDoc.write("Before ");
        dstDoc.insertNode(new aw.BookmarkStart(dstDoc.document, "src_place"));
        dstDoc.insertNode(new aw.BookmarkEnd(dstDoc.document, "src_place"));
        dstDoc.write(" after");

        console.log(dstDoc.document.getText().trimEnd()); // Should output: "Before  after"

        // Insert source document into destination inline.
        dstDoc.moveToBookmark("src_place");
        dstDoc.insertDocumentInline(srcDoc.document, aw.ImportFormatMode.UseDestinationStyles, new aw.ImportFormatOptions());

        console.log(dstDoc.document.getText().trimEnd());
        // ExEnd:InsertDocumentInlineWithBuilder
    });

    test('KeepSourceNumbering', () => {
        //ExStart:KeepSourceNumbering
        let srcDoc = new aw.Document(base.myDir + "List source.docx");
        let dstDoc = new aw.Document(base.myDir + "List destination.docx");

        let options = new aw.ImportFormatOptions();
        // If there is a clash of list styles, apply the list format of the source document.
        // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
        // Set the "KeepSourceNumbering" property to "true" import all clashing
        // list style numbering with the same appearance that it had in the source document.
        options.keepSourceNumbering = true;

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);
        dstDoc.updateListLabels();

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.KeepSourceNumbering.docx");
        //ExEnd:KeepSourceNumbering
    });

    test('IgnoreTextBoxes', () => {
        //ExEnd:KeepSourceNumbering
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Keep the source text boxes formatting when importing.
        let importFormatOptions = new aw.ImportFormatOptions();
        importFormatOptions.ignoreTextBoxes = false;

        let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KeepSourceFormatting,
            importFormatOptions);

        let srcParas = srcDoc.firstSection.body.paragraphs;
        for (let srcPara of srcParas) {
            let importedNode = importer.importNode(srcPara, true);
            dstDoc.firstSection.body.appendChild(importedNode);
        }

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.IgnoreTextBoxes.docx");
        //ExEnd:IgnoreTextBoxes
    });

    test('IgnoreHeaderFooter', () => {
        //ExStart:IgnoreHeaderFooter
        let srcDocument = new aw.Document(base.myDir + "Document source.docx");
        let dstDocument = new aw.Document(base.myDir + "Northwind traders.docx");

        let importFormatOptions = new aw.ImportFormatOptions();
        importFormatOptions.ignoreHeaderFooter = false;

        dstDocument.appendDocument(srcDocument, aw.ImportFormatMode.KeepSourceFormatting, importFormatOptions);

        dstDocument.save(base.artifactsDir + "JoinAndAppendDocuments.IgnoreHeaderFooter.docx");
        //ExEnd:IgnoreHeaderFooter
    });

    test('LinkHeadersFooters', () => {
        //ExStart:LinkHeadersFooters
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Set the appended document to appear on a new page.
        srcDoc.firstSection.pageSetup.sectionStart = aw.SectionStart.NewPage;
        // Link the headers and footers in the source document to the previous section.
        // This will override any headers or footers already found in the source document.
        srcDoc.firstSection.headersFooters.linkToPrevious(true);

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.LinkHeadersFooters.docx");
        //ExEnd:LinkHeadersFooters
    });

    test('RemoveSourceHeadersFooters', () => {
        //ExStart:RemoveSourceHeadersFooters
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Remove the headers and footers from each of the sections in the source document.
        for (let section of srcDoc.sections) {
            section = section.asSection();
            section.clearHeadersFooters();
        }

        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        // document. This should set to false to avoid this behavior.
        srcDoc.firstSection.headersFooters.linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.RemoveSourceHeadersFooters.docx");
        //ExEnd:RemoveSourceHeadersFooters
    });

    test('UnlinkHeadersFooters', () => {
        //ExStart:UnlinkHeadersFooters
        let srcDoc = new aw.Document(base.myDir + "Document source.docx");
        let dstDoc = new aw.Document(base.myDir + "Northwind traders.docx");

        // Unlink the headers and footers in the source document to stop this
        // from continuing the destination document's headers and footers.
        srcDoc.firstSection.headersFooters.linkToPrevious(false);

        dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);

        dstDoc.save(base.artifactsDir + "JoinAndAppendDocuments.UnlinkHeadersFooters.docx");
        //ExEnd:UnlinkHeadersFooters
    });
});
