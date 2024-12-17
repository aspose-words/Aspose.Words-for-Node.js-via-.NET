// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
var moment = require('moment');


describe("ExComment", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('AddCommentWithReply', () => {
    //ExStart
    //ExFor:Comment
    //ExFor:aw.Comment.setText(String)
    //ExFor:aw.Comment.addReply(String, String, DateTime, String)
    //ExSummary:Shows how to add a comment to a document, and then reply to it.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let comment = new aw.Comment(doc, "John Doe", "J.D.", Date.now());
    comment.setText("My comment.");
            
    // Place the comment at a node in the document's body.
    // This comment will show up at the location of its paragraph,
    // outside the right-side margin of the page, and with a dotted line connecting it to its paragraph.
    builder.currentParagraph.appendChild(comment);

    // Add a reply, which will show up under its parent comment.
    comment.addReply("Joe Bloggs", "J.B.", Date.now(), "New reply");

    // Comments and replies are both Comment nodes.
    expect(doc.getChildNodes(aw.NodeType.Comment, true).count).toEqual(2);

    // Comments that do not reply to other comments are "top-level". They have no ancestor comments.
    expect(comment.ancestor).toBe(null);

    // Replies have an ancestor top-level comment.
    expect(comment.replies.at(0).ancestor).toEqual(comment);

    doc.save(base.artifactsDir + "Comment.AddCommentWithReply.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Comment.AddCommentWithReply.docx");
    let docComment = doc.getComment(0, true);

    expect(docComment.count).toEqual(1);
    expect(comment.replies.count).toEqual(1);

    expect(docComment.getText()).toEqual("\u0005My comment.\r");
    expect(docComment.replies.at(0).getText()).toEqual("\u0005New reply\r");
  });


  test('PrintAllComments', () => {
    //ExStart
    //ExFor:aw.Comment.ancestor
    //ExFor:aw.Comment.author
    //ExFor:aw.Comment.replies
    //ExFor:aw.CompositeNode.getChildNodes(NodeType, Boolean)
    //ExSummary:Shows how to print all of a document's comments and their replies.
    let doc = new aw.Document(base.myDir + "Comments.docx");

    let comments = [...doc.getChildNodes(aw.NodeType.Comment, true)];
    expect(comments.length).toEqual(12);

    // If a comment has no ancestor, it is a "top-level" comment as opposed to a reply-type comment.
    // Print all top-level comments along with any replies they may have.
    for (var node of comments.filter(n => n.ancestor == null))
    {
      let comment = node.asComment();
      console.log("Top-level comment:");
      console.log(`\t\"${comment.getText().trim()}\", by ${comment.author}`);
      console.log(`Has ${comment.replies.count} replies`);
      for (let commentReply of comment.replies)
      {
        console.log(`\t\"${commentReply.getText().trim()}\", by ${commentReply.author}`);
      }
      console.log();
    }
    //ExEnd
  });


  test('RemoveCommentReplies', () => {
    //ExStart
    //ExFor:aw.Comment.removeAllReplies
    //ExFor:aw.Comment.removeReply(Comment)
    //ExFor:aw.CommentCollection.item(Int32)
    //ExSummary:Shows how to remove comment replies.
    let doc = new aw.Document();

    let comment = new aw.Comment(doc, "John Doe", "J.D.", Date.now());
    comment.setText("My comment.");

    doc.firstSection.body.firstParagraph.appendChild(comment);
            
    comment.addReply("Joe Bloggs", "J.B.", Date.now(), "New reply");
    comment.addReply("Joe Bloggs", "J.B.", Date.now(), "Another reply");

    expect(comment.replies.count).toEqual(2);

    // Below are two ways of removing replies from a comment.
    // 1 -  Use the "RemoveReply" method to remove replies from a comment individually:
    comment.removeReply(comment.replies.at(0));

    expect(comment.replies.count).toEqual(1);

    // 2 -  Use the "RemoveAllReplies" method to remove all replies from a comment at once:
    comment.removeAllReplies();

    expect(comment.replies.count).toEqual(0);
    //ExEnd
  });


  test('Done', () => {
    //ExStart
    //ExFor:aw.Comment.done
    //ExFor:CommentCollection
    //ExSummary:Shows how to mark a comment as "done".
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Helo world!");

    // Insert a comment to point out an error. 
    let comment = new aw.Comment(doc, "John Doe", "J.D.", Date.now());
    comment.setText("Fix the spelling error!");
    doc.firstSection.body.firstParagraph.appendChild(comment);

    // Comments have a "Done" flag, which is set to "false" by default. 
    // If a comment suggests that we make a change within the document,
    // we can apply the change, and then also set the "Done" flag afterwards to indicate the correction.
    expect(comment.done).toEqual(false);

    doc.firstSection.body.firstParagraph.runs.at(0).text = "Hello world!";
    comment.done = true;

    // Comments that are "done" will differentiate themselves
    // from ones that are not "done" with a faded text color.
    comment = new aw.Comment(doc, "John Doe", "J.D.", Date.now());
    comment.setText("Add text to this paragraph.");
    builder.currentParagraph.appendChild(comment);

    doc.save(base.artifactsDir + "Comment.done.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Comment.done.docx");
    comment = doc.getComment(0, true);

    expect(comment.done).toEqual(true);
    expect(comment.getText().trim()).toEqual("\u0005Fix the spelling error!");
    expect(doc.firstSection.body.firstParagraph.runs.at(0).text).toEqual("Hello world!");
  });

/* TODO DocumentVisitor not supported
  //ExStart
  //ExFor:Comment.Done
  //ExFor:Comment.#ctor(DocumentBase)
  //ExFor:Comment.Accept(DocumentVisitor)
  //ExFor:Comment.DateTime
  //ExFor:Comment.Id
  //ExFor:Comment.Initial
  //ExFor:CommentRangeEnd
  //ExFor:CommentRangeEnd.#ctor(DocumentBase,Int32)
  //ExFor:CommentRangeEnd.Accept(DocumentVisitor)
  //ExFor:CommentRangeEnd.Id
  //ExFor:CommentRangeStart
  //ExFor:CommentRangeStart.#ctor(DocumentBase,Int32)
  //ExFor:CommentRangeStart.Accept(DocumentVisitor)
  //ExFor:CommentRangeStart.Id
  //ExSummary:Shows how print the contents of all comments and their comment ranges using a document visitor.
  test('CreateCommentsAndPrintAllInfo', () => {
    let doc = new aw.Document();
            
    let newComment = new aw.Comment(doc);
    newComment.author = "VDeryushev";
    newComment.initial = "VD",
    newComment.dateTime = Date.now();

    newComment.setText("Comment regarding text.");

    // Add text to the document, warp it in a comment range, and then add your comment.
    let para = doc.firstSection.body.firstParagraph;
    para.appendChild(new aw.CommentRangeStart(doc, newComment.id));
    para.appendChild(new aw.Run(doc, "Commented text."));
    para.appendChild(new aw.CommentRangeEnd(doc, newComment.id));
    para.appendChild(newComment); 
            
    // Add two replies to the comment.
    newComment.addReply("John Doe", "JD", Date.now(), "New reply.");
    newComment.addReply("John Doe", "JD", Date.now(), "Another reply.");

    printAllCommentInfo(doc.getChildNodes(aw.NodeType.Comment, true));
  });


  /// <summary>
  /// Iterates over every top-level comment and prints its comment range, contents, and replies.
  /// </summary>
  function printAllCommentInfo(comments)
  {
    let commentVisitor = new CommentInfoPrinter();

      // Iterate over all top-level comments. Unlike reply-type comments, top-level comments have no ancestor.
    foreach (Comment comment in comments.Where(c => ((Comment)c).Ancestor == null))
    {
        // First, visit the start of the comment range.
      let commentRangeStart = (CommentRangeStart)comment.previousSibling.previousSibling.previousSibling;
      commentRangeStart.accept(commentVisitor);

        // Then, visit the comment, and any replies that it may have.
      comment.accept(commentVisitor);

      for (let reply of comment.replies)
        reply.accept(commentVisitor);

        // Finally, visit the end of the comment range, and then print the visitor's text contents.
      let commentRangeEnd = (CommentRangeEnd)comment.previousSibling;
      commentRangeEnd.accept(commentVisitor);

      console.log(commentVisitor.getText());
    }
  }

    /// <summary>
    /// Prints information and contents of all comments and comment ranges encountered in the document.
    /// </summary>
  public class CommentInfoPrinter : DocumentVisitor
  {
    public CommentInfoPrinter()
    {
      mBuilder = new StringBuilder();
      mVisitorIsInsideComment = false;
    }

      /// <summary>
      /// Gets the plain text of the document that was accumulated by the visitor.
      /// </summary>
    public string GetText()
    {
      return mBuilder.toString();
    }

      /// <summary>
      /// Called when a Run node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitRun(Run run)
    {
      if (mVisitorIsInsideComment) IndentAndAppendLine("[Run] \"" + run.text + "\"");

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a CommentRangeStart node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
    {
      IndentAndAppendLine("[Comment range start] ID: " + commentRangeStart.id);
      mDocTraversalDepth++;
      mVisitorIsInsideComment = true;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a CommentRangeEnd node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
    {
      mDocTraversalDepth--;
      IndentAndAppendLine("[Comment range end] ID: " + commentRangeEnd.id + "\n");
      mVisitorIsInsideComment = false;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a Comment node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitCommentStart(Comment comment)
    {
      IndentAndAppendLine(
        `[Comment start] For comment range ID ${comment.id}, By ${comment.author} on ${comment.dateTime}`);
      mDocTraversalDepth++;
      mVisitorIsInsideComment = true;


      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when the visiting of a Comment node is ended in the document.
      /// </summary>
    public override VisitorAction VisitCommentEnd(Comment comment)
    {
      mDocTraversalDepth--;
      IndentAndAppendLine("[Comment end]");
      mVisitorIsInsideComment = false;

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
      /// </summary>
      /// <param name="text"></param>
    private void IndentAndAppendLine(string text)
    {
      for (let i = 0; i < mDocTraversalDepth; i++)
      {
        mBuilder.append("|  ");
      }

      mBuilder.AppendLine(text);
    }

    private bool mVisitorIsInsideComment;
    private int mDocTraversalDepth;
    private readonly StringBuilder mBuilder;
  }
    //ExEnd
*/


  test('UtcDateTime', () => {
    //ExStart:UtcDateTime
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:aw.Comment.dateTimeUtc
    //ExSummary:Shows how to get UTC date and time.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let date = new Date(2021, 9, 21, 10, 0, 0);

    let comment = new aw.Comment(doc, "John Doe", "J.D.", date);
    comment.setText("My comment.");


    builder.currentParagraph.appendChild(comment);

    doc.save(base.artifactsDir + "Comment.UtcDateTime.docx");
    doc = new aw.Document(base.artifactsDir + "Comment.UtcDateTime.docx");

    comment = doc.getComment(0, true);
    // DateTimeUtc return data without milliseconds.
    let expected = moment(date).add(date.getTimezoneOffset(), 'm').toDate();
    expect(comment.dateTimeUtc).toEqual(expected);
    //ExEnd:UtcDateTime
  });
});
