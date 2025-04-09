// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

/// <summary>
/// Create a document with a given number of bookmarks.
/// </summary>
function CreateDocumentWithBookmarks(numberOfBookmarks)
{
  let doc = new aw.Document();
  let builder = new aw.DocumentBuilder(doc);
  for (let i = 1; i <= numberOfBookmarks; i++)
  {
    let bookmarkName = "MyBookmark_" + i;
    builder.write("Text before bookmark.");
    builder.startBookmark(bookmarkName);
    builder.write(`Text inside ${bookmarkName}.`);
    builder.endBookmark(bookmarkName);
    builder.writeln("Text after bookmark.");
  }
  return doc;
}

describe("ExBookmarks", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Insert', () => {
    //ExStart
    //ExFor:aw.Bookmark.name
    //ExSummary:Shows how to insert a bookmark.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A valid bookmark has a name, a BookmarkStart, and a BookmarkEnd node.
    // Any whitespace in the names of bookmarks will be converted to underscores if we open the saved document with Microsoft Word. 
    // If we highlight the bookmark's name in Microsoft Word via Insert -> Links -> Bookmark, and press "Go To",
    // the cursor will jump to the text enclosed between the BookmarkStart and BookmarkEnd nodes.
    builder.startBookmark("My Bookmark");
    builder.write("Contents of MyBookmark.");
    builder.endBookmark("My Bookmark");

    // Bookmarks are stored in this collection.
    expect(doc.range.bookmarks.at(0).name).toEqual("My Bookmark");

    doc.save(base.artifactsDir + "Bookmarks.insert.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Bookmarks.insert.docx");

    expect(doc.range.bookmarks.at(0).name).toEqual("My Bookmark");
  });


  //ExStart
  //ExFor:Bookmark
  //ExFor:Bookmark.Name
  //ExFor:Bookmark.Text
  //ExFor:Bookmark.BookmarkStart
  //ExFor:Bookmark.BookmarkEnd
  //ExFor:BookmarkStart
  //ExFor:BookmarkStart.#ctor
  //ExFor:BookmarkEnd
  //ExFor:BookmarkEnd.#ctor
  //ExFor:BookmarkStart.Accept(DocumentVisitor)
  //ExFor:BookmarkEnd.Accept(DocumentVisitor)
  //ExFor:BookmarkStart.Bookmark
  //ExFor:BookmarkStart.GetText
  //ExFor:BookmarkStart.Name
  //ExFor:BookmarkEnd.Name
  //ExFor:BookmarkCollection
  //ExFor:BookmarkCollection.Item(Int32)
  //ExFor:BookmarkCollection.Item(String)
  //ExFor:BookmarkCollection.GetEnumerator
  //ExFor:Range.Bookmarks
  //ExFor:DocumentVisitor.VisitBookmarkStart 
  //ExFor:DocumentVisitor.VisitBookmarkEnd
  //ExSummary:Shows how to add bookmarks and update their contents.
  test('CreateUpdateAndPrintBookmarks', () => {
    // Create a document with three bookmarks, then use a custom document visitor implementation to print their contents.
    let doc = CreateDocumentWithBookmarks(3);
    let bookmarks = doc.range.bookmarks;
    expect(bookmarks.count).toEqual(3);

    // Bookmarks can be accessed in the bookmark collection by index or name, and their names can be updated.
    bookmarks.at(0).name = `${bookmarks.at(0).name}_NewName`;
    bookmarks.at("MyBookmark_2").text = `Updated text contents of ${bookmarks.at(1).name}`;
  });
  //ExEnd

  test('TableColumnBookmarks', () => {
    //ExStart
    //ExFor:aw.Bookmark.isColumn
    //ExFor:aw.Bookmark.firstColumn
    //ExFor:aw.Bookmark.lastColumn
    //ExSummary:Shows how to get information about table column bookmarks.
    var doc = new aw.Document(base.myDir + "Table column bookmarks.doc");

    for (let bookmark of doc.range.bookmarks)
    {
      console.log(`Bookmark: ${bookmark.name}${(bookmark.isColumn ? " (Column)" : "")}`);
    }
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    let firstTableColumnBookmark = doc.range.bookmarks.at("FirstTableColumnBookmark");
    let secondTableColumnBookmark = doc.range.bookmarks.at("SecondTableColumnBookmark");

    expect(firstTableColumnBookmark.isColumn).toEqual(true);
    expect(firstTableColumnBookmark.firstColumn).toEqual(1);
    expect(firstTableColumnBookmark.lastColumn).toEqual(3);

    expect(secondTableColumnBookmark.isColumn).toEqual(true);
    expect(secondTableColumnBookmark.firstColumn).toEqual(0);
    expect(secondTableColumnBookmark.lastColumn).toEqual(3);
  });


  test('Remove', () => {
    //ExStart
    //ExFor:BookmarkCollection.clear
    //ExFor:BookmarkCollection.count
    //ExFor:BookmarkCollection.remove(Bookmark)
    //ExFor:BookmarkCollection.remove(String)
    //ExFor:BookmarkCollection.removeAt
    //ExFor:Bookmark.remove
    //ExSummary:Shows how to remove bookmarks from a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert five bookmarks with text inside their boundaries.
    for (let i = 1; i <= 5; i++)
    {
      let bookmarkName = "MyBookmark_" + i;

      builder.startBookmark(bookmarkName);
      builder.write(`Text inside ${bookmarkName}.`);
      builder.endBookmark(bookmarkName);
      builder.insertBreak(aw.BreakType.ParagraphBreak);
    }

    // This collection stores bookmarks.
    let bookmarks = doc.range.bookmarks;

    expect(bookmarks.count).toEqual(5);

    // There are several ways of removing bookmarks.
    // 1 -  Calling the bookmark's Remove method:
    bookmarks.at("MyBookmark_1").remove();

    for (let b of bookmarks)
      expect(b.Name).not.toEqual("MyBookmark_1");

    // 2 -  Passing the bookmark to the collection's Remove method:
    let bookmark = doc.range.bookmarks.at(0);
    doc.range.bookmarks.remove(bookmark);

    for (let b of bookmarks)
      expect(b.Name).not.toEqual("MyBookmark_2");
            
    // 3 -  Removing a bookmark from the collection by name:
    doc.range.bookmarks.remove("MyBookmark_3");

    for (let b of bookmarks)
      expect(b.Name).not.toEqual("MyBookmark_3");

    // 4 -  Removing a bookmark at an index in the bookmark collection:
    doc.range.bookmarks.removeAt(0);

    for (let b of bookmarks)
      expect(b.Name).not.toEqual("MyBookmark_4");

    // We can clear the entire bookmark collection.
    bookmarks.clear();

    // The text that was inside the bookmarks is still present in the document.
    expect(bookmarks.count).toEqual(0);
    expect(doc.getText().trim()).toEqual("Text inside MyBookmark_1.\r" +
            "Text inside MyBookmark_2.\r" +
            "Text inside MyBookmark_3.\r" +
            "Text inside MyBookmark_4.\r" +
            "Text inside MyBookmark_5.");
    //ExEnd
  });
});
