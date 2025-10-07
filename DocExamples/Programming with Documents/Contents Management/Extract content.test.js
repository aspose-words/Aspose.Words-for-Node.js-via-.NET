// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const ExtractContentHelper = require('./Extract content helper').ExtractContentHelper;

describe("ExtractContent", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('ExtractContentBetweenBlockLevelNodes', () => {
        //ExStart:ExtractContentBetweenBlockLevelNodes
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        let startPara = doc.lastSection.getChild(aw.NodeType.Paragraph, 2, true).asParagraph();
        let endTable = doc.lastSection.getChild(aw.NodeType.Table, 0, true).asTable();

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        let extractedNodes = ExtractContentHelper.extractContent(startPara, endTable, true, false);
        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);

        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenBlockLevelNodes.docx");
        //ExEnd:ExtractContentBetweenBlockLevelNodes
    });

    test('ExtractContentBetweenBookmark', () => {
        //ExStart:ExtractContentBetweenBookmark
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        let bookmark = doc.range.bookmarks.at("Bookmark1");
        let bookmarkStart = bookmark.bookmarkStart;
        let bookmarkEnd = bookmark.bookmarkEnd;

        // Firstly, extract the content between these nodes, including the bookmark.
        let extractedNodesInclusive = ExtractContentHelper.extractContent(bookmarkStart, bookmarkEnd, true, true);

        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesInclusive);
        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenBookmark.IncludingBookmark.docx");

        // Secondly, extract the content between these nodes this time without including the bookmark.
        let extractedNodesExclusive = ExtractContentHelper.extractContent(bookmarkStart, bookmarkEnd, false, true);

        dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesExclusive);
        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenBookmark.WithoutBookmark.docx");
        //ExEnd:ExtractContentBetweenBookmark
    });

    test('ExtractContentBetweenCommentRange', () => {
        //ExStart:ExtractContentBetweenCommentRange
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        let commentStart = doc.getChild(aw.NodeType.CommentRangeStart, 0, true).asCommentRangeStart();
        let commentEnd = doc.getChild(aw.NodeType.CommentRangeEnd, 0, true).asCommentRangeEnd();

        // Firstly, extract the content between these nodes including the comment as well.
        let extractedNodesInclusive = ExtractContentHelper.extractContent(commentStart, commentEnd, true, true);
        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesInclusive);

        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenCommentRange.IncludingComment.docx");

        // Secondly, extract the content between these nodes without the comment.
        let extractedNodesExclusive = ExtractContentHelper.extractContent(commentStart, commentEnd, false, true);
        dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesExclusive);

        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenCommentRange.WithoutComment.docx");
        //ExEnd:ExtractContentBetweenCommentRange
    });

    test('ExtractContentBetweenParagraphs', () => {
        //ExStart:ExtractContentBetweenParagraphs
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        let startPara = doc.firstSection.body.getChild(aw.NodeType.Paragraph, 6, true).asParagraph();
        let endPara = doc.firstSection.body.getChild(aw.NodeType.Paragraph, 10, true).asParagraph();

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        let extractedNodes = ExtractContentHelper.extractContent(startPara, endPara, true, true);
        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);

        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenParagraphs.docx");
        //ExEnd:ExtractContentBetweenParagraphs
    });

    test('ExtractContentBetweenParagraphStyles', () => {
        //ExStart:ExtractContentBetweenParagraphStyles
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        // Gather a list of the paragraphs using the respective heading styles.
        let parasStyleHeading1 = paragraphsByStyleName(doc, "Heading 1");
        let parasStyleHeading3 = paragraphsByStyleName(doc, "Heading 3");

        // Use the first instance of the paragraphs with those styles.
        let startPara = parasStyleHeading1.at(0);
        let endPara = parasStyleHeading3.at(0);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        let extractedNodes = ExtractContentHelper.extractContent(startPara, endPara, false, true);
        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);

        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentBetweenParagraphStyles.docx");
        //ExEnd:ExtractContentBetweenParagraphStyles
    });

    //ExStart:ParagraphsByStyleName
    //GistId:433f5122fe18fdc24a406528b70b0020
    function paragraphsByStyleName(doc, styleName) {
        // Create an array to collect paragraphs of the specified style.
        let paragraphsWithStyle = [];
        let paragraphs = doc.getChildNodes(aw.NodeType.Paragraph, true);

        // Look through all paragraphs to find those with the specified style.
        for (let paragraph of paragraphs) {
            paragraph = paragraph.asParagraph();
            if (paragraph.paragraphFormat.style.name === styleName)
                paragraphsWithStyle.push(paragraph);
        }

        return paragraphsWithStyle;
    }
    //ExEnd:ParagraphsByStyleName

    test('ExtractContentBetweenRuns', () => {
        //ExStart:ExtractContentBetweenRuns
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");

        let para = doc.getChild(aw.NodeType.Paragraph, 7, true).asParagraph();
        let startRun = para.runs.at(1);
        let endRun = para.runs.at(4);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        let extractedNodes = ExtractContentHelper.extractContent(startRun, endRun, true, false);
        for (let extractedNode of extractedNodes)
            console.log(extractedNode.toString(aw.SaveFormat.Text));
        //ExEnd:ExtractContentBetweenRuns
    });

    test('ExtractContentUsingField', () => {
        //ExStart:ExtractContentUsingField
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Extract content.docx");
        let builder = new aw.DocumentBuilder(doc);
        // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.moveToMergeField("Fullname", false, false);

        // The builder cursor should be positioned at the start of the field.
        let startField = builder.currentNode;
        let endPara = doc.firstSection.getChild(aw.NodeType.Paragraph, 5, true).asParagraph();
        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        let extractedNodes = ExtractContentHelper.extractContent(startField, endPara, false, true);

        let dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);
        dstDoc.save(base.artifactsDir + "ExtractContent.ExtractContentUsingField.docx");
        //ExEnd:ExtractContentUsingField
    });

    test('SimpleExtractText', () => {
        //ExStart:SimpleExtractText
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.insertField("MERGEFIELD Field");

        // When converted to text it will not retrieve fields code or special characters,
        // but will still contain some natural formatting characters such as paragraph markers etc.
        // This is the same as "viewing" the document as if it was opened in a text editor.
        console.log("Convert to text result: " + doc.toString(aw.SaveFormat.Text));
        //ExEnd:SimpleExtractText
    });

    test('ExtractPrintText', () => {
        //ExStart:ExtractText
        //GistId:1693b4ac01f19ec81c9618649b62acb8
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.
        console.log("Contents of the table: ");
        console.log(table.range.text);
        //ExEnd:ExtractText

        //ExStart:PrintTextRangeRowAndTable
        //GistId:1693b4ac01f19ec81c9618649b62acb8
        console.log("\nContents of the row: ");
        console.log(table.rows.at(1).range.text);

        console.log("\nContents of the cell: ");
        console.log(table.lastRow.lastCell.range.text);
        //ExEnd:PrintTextRangeRowAndTable
    });

    test('ExtractImages', () => {
        //ExStart:ExtractImages
        //GistId:433f5122fe18fdc24a406528b70b0020
        let doc = new aw.Document(base.myDir + "Images.docx");

        let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
        let imageIndex = 0;

        for (let shape of shapes) {
            shape = shape.asShape();
            if (shape.hasImage) {
                let imageFileName =
                    `Image.ExportImages.${imageIndex}_${aw.FileFormatUtil.imageTypeToExtension(shape.imageData.imageType)}`;

                // Note, if you have only an image (not a shape with a text and the image),
                // you can use shape.GetShapeRenderer().Save(...) method to save the image.
                shape.imageData.save(base.artifactsDir + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd:ExtractImages
    });

    test('ExtractContentBasedOnStyles', () => {
        //ExStart:ExtractContentBasedOnStyles
        //GistId:c6b0305cd373fae738c432637dd67ba5
        let doc = new aw.Document(base.myDir + "Styles.docx");

        let paragraphs = paragraphsByStyleName(doc, "Heading 1");
        console.log(`Paragraphs with "Heading 1" styles (${paragraphs.length}):`);

        for (let paragraph of paragraphs)
            console.log(paragraph.toString(aw.SaveFormat.Text));

        let runs = runsByStyleName(doc, "Intense Emphasis");
        console.log(`\nRuns with "Intense Emphasis" styles (${runs.length}):`);

        for (let run of runs)
            console.log(run.range.text);
        //ExEnd:ExtractContentBasedOnStyles
    });

    //ExStart:RunsByStyleName
    //GistId:c6b0305cd373fae738c432637dd67ba5
    function runsByStyleName(doc, styleName) {
        let runsWithStyle = [];
        let runs = doc.getChildNodes(aw.NodeType.Run, true);

        for (let run of runs) {
            run = run.asRun();
            if (run.font.style.name === styleName)
                runsWithStyle.push(run);
        }

        return runsWithStyle;
    }
    //ExEnd:RunsByStyleName

});