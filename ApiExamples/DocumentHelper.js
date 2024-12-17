// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const uniqueFilename = require('unique-filename');


    /// <summary>
    /// Create simple document without run in the paragraph
    /// </summary>
    function createDocumentWithoutDummyText() {
        const doc = new aw.Document();
        //Remove the previous changes of the document
        doc.removeAllChildren();
        //Set the document author
        doc.builtInDocumentProperties.author = "Test Author";
        //Create paragraph without run
        const builder = new aw.DocumentBuilder(doc);
        builder.writeln();
        return doc;
    }

    /// <summary>
    /// Create new document with text
    /// </summary>
    function createDocumentFillWithDummyText() {
        const doc = new aw.Document();
        //Remove the previous changes of the document
        doc.removeAllChildren();
        //Set the document author
        doc.builtInDocumentProperties.author = "Test Author";
        const builder = new aw.DocumentBuilder(doc);
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");
        //Insert new table with two rows and two cells
        insertTable(builder);
        builder.writeln("Hello World!");
        // Continued on page 2 of the document content
        builder.insertBreak(aw.BreakType.PageBreak);
        //Insert TOC entries
        insertToc(builder);
        return doc;
    }

    function findTextInFile(path, expression) {
        const text = fs.readFileSync(path).toString();
        const lines = text.split(/\r?\n/);

        for (let i = 0; i < lines.count; i++) {
            if (lines[i] == "")
                continue;

            let includes = lines[i].includes(expression);
            if (includes)
                console.log(lines[i]);
            expect(includes).toBe(true);
        }
    }

/*
    /// <summary>
    /// Create new document template for reporting engine
    /// </summary>
    internal static Document CreateSimpleDocument(string templateText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write(templateText);
        return doc;
    }
    /// <summary>
    /// Create new document with textbox shape and some query
    /// </summary>
    internal static Document CreateTemplateDocumentWithDrawObjects(string templateText, ShapeType shapeType)
    {
        Document doc = new Document();
        // Create textbox shape.
        Shape shape = new Shape(doc, shapeType);
        shape.Width = 431.5;
        shape.Height = 346.35;
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, templateText));
        // Insert paragraph into the textbox.
        shape.AppendChild(paragraph);
        // Insert textbox into the document.
        doc.FirstSection.Body.FirstParagraph.AppendChild(shape);
        return doc;
    }*/

    /// <summary>
    /// Compare word documents
    /// </summary>
    /// <param name="filePathDoc1">First document path</param>
    /// <param name="filePathDoc2">Second document path</param>
    /// <returns>Result of compare document</returns>
    function compareDocs(filePathDoc1, filePathDoc2) {
      let doc1 = new aw.Document(filePathDoc1);
      let doc2 = new aw. Document(filePathDoc2);
      return doc1.getText() == doc2.getText();
    }

    /// <summary>
    /// Insert run into the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="text">Custom text</param>
    /// <param name="paraIndex">Paragraph index</param>
    function insertNewRun(doc, text, paraIndex) {
        let para = getParagraph(doc, paraIndex);
        let run = new aw.Run(doc);
        run.text = text;
        para.appendChild(run);
        return run;
    }

    /*
    /// <summary>
    /// Insert text into the current document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    /// <param name="textStrings">Custom text</param>
    internal static void InsertBuilderText(DocumentBuilder builder, string[] textStrings)
    {
        foreach (string textString in textStrings)
        {
            builder.Writeln(textString);
        }
    }
*/        
    /// <summary>
    /// Insert new table in the document
    /// </summary>
    /// <param name="builder">Current document builder</param>
    function insertTable(builder)
    {
        //Start creating a new table
        let table = builder.startTable();
        //Insert Row 1 Cell 1
        builder.insertCell();
        builder.write("Date");
        //Set width to fit the table contents
        table.autoFit(aw.Tables.AutoFitBehavior.AutoFitToContents);
        //Insert Row 1 Cell 2
        builder.insertCell();
        builder.write(" ");
        builder.endRow();
        //Insert Row 2 Cell 1
        builder.insertCell();
        builder.write("Author");
        //Insert Row 2 Cell 2
        builder.insertCell();
        builder.write(" ");
        builder.endRow();
        builder.endTable();
        return table;
    }
    
    /// <summary>
    /// Insert TOC entries in the document
    /// </summary>
    /// <param name="builder">
    /// The builder.
    /// </param>
    function insertToc(builder)
    {
        // Creating TOC entries
        builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
        builder.writeln("Heading 1");
        builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading2;
        builder.writeln("Heading 1.1");
        builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading4;
        builder.writeln("Heading 1.1.1.1");
        builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading5;
        builder.writeln("Heading 1.1.1.1.1");
        builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading9;
        builder.writeln("Heading 1.1.1.1.1.1.1.1.1");
    }

    /// <summary>
    /// Get section text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="secIndex">Section number from collection</param>
    function getSectionText(doc, secIndex) {
        return doc.sections.at(secIndex).getText();
    }

    /// <summary>
    /// Get paragraph text of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    function getParagraphText(doc, paraIndex) {
        return doc.firstSection.body.paragraphs.at(paraIndex).getText();
    }

    /// <summary>
    /// Get paragraph of the current document
    /// </summary>
    /// <param name="doc">Current document</param>
    /// <param name="paraIndex">Paragraph number from collection</param>
    function getParagraph(doc, paraIndex) {
        return doc.firstSection.body.paragraphs.at(paraIndex);
    }

    /// <summary>
    /// Save the document to a stream, immediately re-open it and return the newly opened version
    /// </summary>
    /// <remarks>
    /// Used for testing how document features are preserved after saving/loading
    /// </remarks>
    /// <param name="doc">The document we wish to re-open</param>
    function saveOpen(doc) {
        const dstFile = uniqueFilename(base.artifactsDir, 'saveopen-temp') + ".docx";
        try {
            doc.save(dstFile, new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx));
            return new aw.Document(dstFile);
        }
        finally {
            if (fs.existsSync(dstFile)) {
                try {
                    fs.unlinkSync(dstFile)
                } catch(err) {
                    console.warn(`Can't unlink ${dstFile} - ${err}`);
                }
            }
        }
    }

module.exports = {
    createDocumentWithoutDummyText,
    createDocumentFillWithDummyText,
    findTextInFile,
    compareDocs,
    insertNewRun,
    getParagraphText,
    saveOpen,
    getParagraph,
    insertTable,
    insertToc,
    getSectionText
};
