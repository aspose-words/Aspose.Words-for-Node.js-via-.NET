// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithComments", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });


    test('AddComments', () => {
        //ExStart:AddComments
        //GistId:f8f4978e43b554cf1c3f88982244c535
        //ExStart:CreateSimpleDocumentUsingDocumentBuilder
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.write("Some text is added.");
        //ExEnd:CreateSimpleDocumentUsingDocumentBuilder

        let comment = new aw.Comment(doc, "Awais Hafeez", "AH", new Date());
        comment.setText("Comment text.");

        builder.currentParagraph.appendChild(comment);

        doc.save(base.artifactsDir + "WorkingWithComments.AddComments.docx");
        //ExEnd:AddComments
    });

    test('AnchorComment', () => {
        //ExStart:AnchorComment
        //GistId:f8f4978e43b554cf1c3f88982244c535
        let doc = new aw.Document();

        let para1 = new aw.Paragraph(doc);
        let run1 = new aw.Run(doc, "Some ");
        let run2 = new aw.Run(doc, "text ");
        para1.appendChild(run1);
        para1.appendChild(run2);
        doc.firstSection.body.appendChild(para1);

        let para2 = new aw.Paragraph(doc);
        let run3 = new aw.Run(doc, "is ");
        let run4 = new aw.Run(doc, "added ");
        para2.appendChild(run3);
        para2.appendChild(run4);
        doc.firstSection.body.appendChild(para2);

        let comment = new aw.Comment(doc, "Awais Hafeez", "AH", new Date());
        comment.paragraphs.add(new aw.Paragraph(doc));
        comment.firstParagraph.runs.add(new aw.Run(doc, "Comment text."));

        let commentRangeStart = new aw.CommentRangeStart(doc, comment.id);
        let commentRangeEnd = new aw.CommentRangeEnd(doc, comment.id);

        run1.parentNode.insertAfter(commentRangeStart, run1);
        run3.parentNode.insertAfter(commentRangeEnd, run3);
        commentRangeEnd.parentNode.insertAfter(comment, commentRangeEnd);

        doc.save(base.artifactsDir + "WorkingWithComments.AnchorComment.doc");
        //ExEnd:AnchorComment
    });

    test('AddRemoveCommentReply', () => {
        //ExStart:AddRemoveCommentReply
        //GistId:f8f4978e43b554cf1c3f88982244c535
        let doc = new aw.Document(base.myDir + "Comments.docx");

        let comment = doc.getChild(aw.NodeType.Comment, 0, true).asComment();
        comment.removeReply(comment.replies.at(0));

        comment.addReply("John Doe", "JD", new Date(2017, 8, 25, 12, 15, 0), "New reply");

        doc.save(base.artifactsDir + "WorkingWithComments.AddRemoveCommentReply.docx");
        //ExEnd:AddRemoveCommentReply
    });

    test('ProcessComments', () => {
        //ExStart:ProcessComments
        //GistId:f8f4978e43b554cf1c3f88982244c535
        let doc = new aw.Document(base.myDir + "Comments.docx");

        // Extract the information about the comments of all the authors.
        for (let comment of extractComments(doc))
            process.stdout.write(comment);

        // Remove comments by the "pm" author.
        removeComments(doc, "pm");
        console.log("Comments from \"pm\" are removed!");

        // Extract the information about the comments of the "ks" author.
        for (let comment of extractComments(doc, "ks"))
            process.stdout.write(comment);

        // Read the comment's reply and resolve them.
        commentResolvedAndReplies(doc);

        // Remove all comments.
        removeComments(doc);
        console.log("All comments are removed!");

        doc.save(base.artifactsDir + "WorkingWithComments.ProcessComments.docx");
        //ExEnd:ProcessComments
    });

    //ExStart:ExtractComments
    //GistId:f8f4978e43b554cf1c3f88982244c535
    function extractComments(doc) {
        let collectedComments = [];
        let comments = doc.getChildNodes(aw.NodeType.Comment, true);

        for (let comment of comments) {
            comment = comment.asComment();
            collectedComments.push(comment.author + " " + comment.dateTime + " " +
                comment.toString(aw.SaveFormat.Text));
        }

        return collectedComments;
    }
    //ExEnd:ExtractComments

    //ExStart:ExtractCommentsByAuthor
    //GistId:f8f4978e43b554cf1c3f88982244c535
    function extractComments(doc, authorName) {
        let collectedComments = [];
        let comments = doc.getChildNodes(aw.NodeType.Comment, true);

        for (let comment of comments) {
            if (comment.author == authorName)
                comment = comment.asComment();
            collectedComments.push(comment.author + " " + comment.dateTime + " " +
                comment.toString(aw.SaveFormat.Text));
        }

        return collectedComments;
    }
    //ExEnd:ExtractCommentsByAuthor

    //ExStart:RemoveComments
    //GistId:f8f4978e43b554cf1c3f88982244c535
    function removeComments(doc) {
        let comments = doc.getChildNodes(aw.NodeType.Comment, true);
        comments.clear();
    }
    //ExEnd:RemoveComments

    //ExStart:RemoveCommentsByAuthor
    //GistId:f8f4978e43b554cf1c3f88982244c535
    function removeComments(doc, authorName) {
        let comments = doc.getChildNodes(aw.NodeType.Comment, true);

        // Look through all comments and remove those written by the authorName.
        for (let comment of comments) {
            comment = comment.asComment();
            if (comment.author == authorName)
                comment.remove();
        }
    }
    //ExEnd:RemoveCommentsByAuthor

    //ExStart:CommentResolvedAndReplies
    //GistId:f8f4978e43b554cf1c3f88982244c535
    function commentResolvedAndReplies(doc) {
        let comments = doc.getChildNodes(aw.NodeType.Comment, true);

        let parentComment = comments.at(0).asComment();
        for (let childComment of parentComment.replies) {
            childComment = childComment.asComment();
            // Get comment parent and status.
            console.log(childComment.ancestor.id);
            console.log(childComment.done);

            // And update comment Done mark.
            childComment.done = true;
        }
    }
    //ExEnd:CommentResolvedAndReplies

    test('RemoveRangeText', () => {
        //ExStart:RemoveRangeText
        //GistId:f8f4978e43b554cf1c3f88982244c535
        let doc = new aw.Document(base.myDir + "Comments.docx");

        let commentStart = doc.getChild(aw.NodeType.CommentRangeStart, 0, true);
        let currentNode = commentStart;

        let isRemoving = true;
        while (currentNode != null && isRemoving) {
            if (currentNode.nodeType == aw.NodeType.CommentRangeEnd)
                isRemoving = false;

            let nextNode = currentNode.nextPreOrder(doc);
            currentNode.remove();
            currentNode = nextNode;
        }

        doc.save(base.artifactsDir + "WorkingWithComments.RemoveRangeText.docx");
        //ExEnd:RemoveRangeText
    });
});