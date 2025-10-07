// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithBookmarks", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('AccessBookmarks', () => {
        //ExStart:AccessBookmarks
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        let doc = new aw.Document(base.myDir + "Bookmarks.docx");

        // By index:
        let bookmark1 = doc.range.bookmarks.at(0);
        // By name:
        let bookmark2 = doc.range.bookmarks.at("MyBookmark3");
        //ExEnd:AccessBookmarks
    });

    test('UpdateBookmarkData', () => {
        //ExStart:UpdateBookmarkData
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        let doc = new aw.Document(base.myDir + "Bookmarks.docx");

        let bookmark = doc.range.bookmarks.at("MyBookmark1");

        let name = bookmark.name;
        let text = bookmark.text;

        bookmark.name = "RenamedBookmark";
        bookmark.text = "This is a new bookmarked text.";
        //ExEnd:UpdateBookmarkData
    });

    test('BookmarkTableColumns', () => {
        //ExStart:BookmarkTable
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.startTable();

        builder.insertCell();

        builder.startBookmark("MyBookmark");

        builder.write("This is row 1 cell 1");

        builder.insertCell();
        builder.write("This is row 1 cell 2");

        builder.endRow();

        builder.insertCell();
        builder.writeln("This is row 2 cell 1");

        builder.insertCell();
        builder.writeln("This is row 2 cell 2");

        builder.endRow();
        builder.endTable();

        builder.endBookmark("MyBookmark");
        //ExEnd:BookmarkTable

        //ExStart:BookmarkTableColumns
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        for (let i = 0; i < doc.range.bookmarks.count; i++) {
            let bookmark = doc.range.bookmarks.at(i);
            console.log("Bookmark: " + bookmark.name + (bookmark.isColumn ? " (Column)" : ""));

            if (bookmark.isColumn) {
                let row = bookmark.bookmarkStart.getAncestor(aw.NodeType.Row);
                if (bookmark.firstColumn < row.cells.count) {
                    console.log(row.cells.get(bookmark.firstColumn).getText().trimEnd(aw.ControlChar.CELL, ''));
                }
            }
        }
        //ExEnd:BookmarkTableColumns
    });

    test('CopyBookmarkedText', () => {
        let srcDoc = new aw.Document(base.myDir + "Bookmarks.docx");

        // This is the bookmark whose content we want to copy.
        let srcBookmark = srcDoc.range.bookmarks.at("MyBookmark1");
        // We will be adding to this document.
        let dstDoc = new aw.Document();
        // Let's say we will be appended to the end of the body of the last section.
        let dstNode = dstDoc.lastSection.body;
        // If you import multiple times without a single context, it will result in many styles created.
        let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KeepSourceFormatting);

        appendBookmarkedText(importer, srcBookmark, dstNode);

        dstDoc.save(base.artifactsDir + "WorkingWithBookmarks.CopyBookmarkedText.docx");
    });

    function appendBookmarkedText(importer, srcBookmark, dstNode) {
        // This is the paragraph that contains the beginning of the bookmark.
        let startPara = srcBookmark.bookmarkStart.parentNode.asParagraph();

        // This is the paragraph that contains the end of the bookmark.
        let endPara = srcBookmark.bookmarkEnd.parentNode.asParagraph();

        if (startPara == null || endPara == null) {
            throw new Error("Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");
        }

        // Limit ourselves to a reasonably simple scenario.
        if (!base.compareNodes(startPara.parentNode, endPara.parentNode)) {
            throw new Error("Start and end paragraphs have different parents, cannot handle this scenario yet.");
        }

        // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        // therefore the node at which we stop is one after the end paragraph.
        let endNode = endPara.nextSibling;

        curNode = startPara
        while (!base.compareNodes(curNode, endNode)) {
            // This creates a copy of the current node and imports it (makes it valid) in the context
            // of the destination document. Importing means adjusting styles and list identifiers correctly.
            let newNode = importer.importNode(curNode, true);
            dstNode.appendChild(newNode);
            curNode = curNode.nextSibling;
        }
    }

    test('CreateBookmark', () => {
        //ExStart:CreateBookmark
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.startBookmark("My Bookmark");
        builder.writeln("Text inside a bookmark.");

        builder.startBookmark("Nested Bookmark");
        builder.writeln("Text inside a NestedBookmark.");
        builder.endBookmark("Nested Bookmark");

        builder.writeln("Text after Nested Bookmark.");
        builder.endBookmark("My Bookmark");

        let options = new aw.Saving.PdfSaveOptions();
        options.outlineOptions.bookmarksOutlineLevels.add("My Bookmark", 1);
        options.outlineOptions.bookmarksOutlineLevels.add("Nested Bookmark", 2);

        doc.save(base.artifactsDir + "WorkingWithBookmarks.CreateBookmark.pdf", options);
        //ExEnd:CreateBookmark
    });

    test('ShowHideBookmarks', () => {
        //ExStart:ShowHideBookmarks
        //GistId:6b8a885f5544cddd9bc77edb3ad18692
        let doc = new aw.Document(base.myDir + "Bookmarks.docx");

        ShowHideBookmarkedContent(doc, "MyBookmark1", true);

        doc.save(base.artifactsDir + "WorkingWithBookmarks.ShowHideBookmarks.docx");
        //ExEnd:ShowHideBookmarks
    });

    //ExStart:ShowHideBookmarkedContent
    //GistId:6b8a885f5544cddd9bc77edb3ad18692
    function ShowHideBookmarkedContent(doc, bookmarkName, isHidden) {
        let bm = doc.range.bookmarks.at(bookmarkName);
        let currentNode = bm.bookmarkStart;
        while (currentNode != null && currentNode.nodeType != aw.NodeType.BookmarkEnd) {
            if (currentNode.nodeType == aw.NodeType.Run) {
                let run = currentNode.asRun();
                run.font.hidden = isHidden;
            }
            currentNode = currentNode.nextSibling;
        }
    }
    //ExEnd:ShowHideBookmarkedContent

    test('UntangleRowBookmarks', () => {
        let doc = new aw.Document(base.myDir + "Table column bookmarks.docx");

        // This performs the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        untangle(doc);
        // Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        deleteRowByBookmark(doc, "ROW2");

        // This is just to check that the other bookmark was not damaged.
        if (doc.range.bookmarks.at("ROW1").bookmarkEnd == null) {
            throw new Error("Wrong, the end of the bookmark was deleted.");
        }

        doc.save(base.artifactsDir + "WorkingWithBookmarks.UntangleRowBookmarks.docx");
    });

    function untangle(doc) {
        for (let bookmark of doc.range.bookmarks) {
            // Get the parent row of both the bookmark and bookmark end node.
            let row1 = bookmark.bookmarkStart.getAncestor(aw.NodeType.Row);
            let row2 = bookmark.bookmarkEnd.getAncestor(aw.NodeType.Row);

            // If both rows are found okay, and the bookmark start and end are contained in adjacent rows,
            // move the bookmark end node to the end of the last paragraph in the top row's last cell.
            if (row1 != null && row2 != null && base.compareNodes(row1.nextSibling, row2)) {
                row1.asRow().lastCell.lastParagraph.appendChild(bookmark.bookmarkEnd);
            }
        }
    }

    function deleteRowByBookmark(doc, bookmarkName) {
        let bookmark = doc.range.bookmarks.at(bookmarkName);

        if (bookmark != null) {
            let row = bookmark.bookmarkStart.getAncestor(aw.NodeType.Row);
            if (row != null) {
                row.remove();
            }
        }
    }

});