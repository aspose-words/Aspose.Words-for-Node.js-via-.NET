// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithFields", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });


    test('FieldCode', () => {
        //ExStart:FieldCode
        //GistId:56db351e3569b23ecfe91a2ef9339fa7
        let doc = new aw.Document(base.myDir + "Hyperlinks.docx");

        for (let field of doc.range.fields) {
            let fieldCode = field.getFieldCode();
            let fieldResult = field.result;
        }
        //ExEnd:FieldCode
    });

    test('SpecifyLocaleAtFieldLevel', () => {
        //ExStart:SpecifyLocaleAtFieldLevel
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let builder = new aw.DocumentBuilder();

        let field = builder.insertField(aw.Fields.FieldType.FieldDate, true);
        field.localeId = 1049;

        builder.document.save(base.artifactsDir + "WorkingWithFields.SpecifyLocaleAtFieldLevel.docx");
        //ExEnd:SpecifyLocaleAtFieldLevel
    });

    test('ReplaceHyperlinks', () => {
        //ExStart:ReplaceHyperlinks
        //GistId:9b6efb87f331ae61c0100e106c9c1738
        let doc = new aw.Document(base.myDir + "Hyperlinks.docx");

        for (let field of doc.range.fields) {
            field = field.asFieldHyperlink()
            if (field.type == aw.Fields.FieldType.FieldHyperlink) {
                let hyperlink = field;

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink.subAddress != null)
                    continue;

                hyperlink.address = "http://www.aspose.com";
                hyperlink.result = "Aspose - The .NET & Java Component Publisher";
            }
        }

        doc.save(base.artifactsDir + "WorkingWithFields.ReplaceHyperlinks.docx");
        //ExEnd:ReplaceHyperlinks
    });

    test('RenameMergeFields', () => {
        //ExStart:RenameMergeFields
        //GistId:ce43c0268e53b9e7df2f581cafc2d748
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        for (let f of doc.range.fields) {
            if (f.type == aw.Fields.FieldType.FieldMergeField) {
                let mergeField = f.asFieldMergeField();
                mergeField.fieldName = mergeField.fieldName + "_Renamed";
                mergeField.update();
            }
        }

        doc.save(base.artifactsDir + "WorkingWithFields.RenameMergeFields.docx");
        //ExEnd:RenameMergeFields
    });

    test('RemoveField', () => {
        //ExStart:RemoveField
        //GistId:87f60ea5f7e177ac68f6daae9ff2e883
        let doc = new aw.Document(base.myDir + "Various fields.docx");

        let field = doc.range.fields.at(0);
        field.remove();
        //ExEnd:RemoveField
    });

    test('UnlinkFields', () => {
        //ExStart:UnlinkFields
        //GistId:5745cef9bae16cdf430ed2906034a61e
        let doc = new aw.Document(base.myDir + "Various fields.docx");
        doc.unlinkFields();
        //ExEnd:UnlinkFields
    });

    test('InsertToaFieldWithoutDocumentBuilder', () => {
        //ExStart:InsertToaFieldWithoutDocumentBuilder
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let para = new aw.Paragraph(doc);

        // We want to insert TA and TOA fields like this:
        // { TA  \c 1 \l "Value 0" }
        // { TOA  \c 1 }

        let fieldTA = para.appendField(aw.Fields.FieldType.FieldTOAEntry, false).asFieldTA();
        fieldTA.entryCategory = "1";
        fieldTA.longCitation = "Value 0";

        doc.firstSection.body.appendChild(para);

        para = new aw.Paragraph(doc);

        let fieldToa = para.appendField(aw.Fields.FieldType.FieldTOA, false).asFieldToa();
        fieldToa.entryCategory = "1";
        doc.firstSection.body.appendChild(para);

        fieldToa.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertToaFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertToaFieldWithoutDocumentBuilder
    });

    test('InsertNestedFields', () => {
        //ExStart:InsertNestedFields
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        for (let i = 0; i < 5; i++)
            builder.insertBreak(aw.BreakType.PageBreak);

        builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        let field = builder.insertField("IF ");
        builder.moveTo(field.separator);
        builder.insertField("PAGE");
        builder.write(" <> ");
        builder.insertField("NUMPAGES");
        builder.write(" \"See Next Page\" \"Last Page\" ");

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertNestedFields.docx");
        //ExEnd:InsertNestedFields
    });

    test('InsertMergeFieldUsingDom', () => {
        //ExStart:InsertMergeFieldUsingDom
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let para = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();
        builder.moveTo(para);

        // We want to insert a merge field like this:
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        let field = builder.insertField(aw.Fields.FieldType.FieldMergeField, false).asFieldMergeField();
        // { " MERGEFIELD Test1" }
        field.fieldName = "Test1";
        // { " MERGEFIELD Test1 \\b Test2" }
        field.textBefore = "Test2";
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field.textAfter = "Test3";
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field.isMapped = true;
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field.isVerticalFormatting = true;

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertMergeFieldUsingDom.docx");
        //ExEnd:InsertMergeFieldUsingDom
    });

    test('InsertAddressBlockFieldUsingDom', () => {
        //ExStart:InsertAddressBlockFieldUsingDom
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let para = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();
        builder.moveTo(para);

        // We want to insert a mail merge address block like this:
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        let field = builder.insertField(aw.Fields.FieldType.FieldAddressBlock, false).asFieldAddressBlock();
        // { ADDRESSBLOCK \\c 1" }
        field.includeCountryOrRegionName = "1";
        // { ADDRESSBLOCK \\c 1 \\d" }
        field.formatAddressOnCountryOrRegion = true;
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field.excludedCountryOrRegionName = "Test2";
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field.nameAndAddressFormat = "Test3";
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field.languageId = "Test 4";

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertAddressBlockFieldUsingDom.docx");
        //ExEnd:InsertAddressBlockFieldUsingDom
    });

    test('InsertFieldIncludeTextWithoutDocumentBuilder', () => {
        //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();

        let para = new aw.Paragraph(doc);

        // We want to insert an INCLUDETEXT field like this:
        // { INCLUDETEXT  "file path" }
        let fieldIncludeText = para.appendField(aw.Fields.FieldType.FieldIncludeText, false).asFieldIncludeText();
        fieldIncludeText.bookmarkName = "bookmark";
        fieldIncludeText.sourceFullName = base.myDir + "IncludeText.docx";

        doc.firstSection.body.appendChild(para);

        fieldIncludeText.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertIncludeFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
    });

    test('InsertFieldNone', () => {
        //ExStart:InsertFieldNone
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let field = builder.insertField(aw.Fields.FieldType.FieldNone, false);

        doc.save(base.artifactsDir + "WorkingWithFields.InsertFieldNone.docx");
        //ExEnd:InsertFieldNone
    });

    test('InsertField', () => {
        //ExStart:InsertField
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.insertField("MERGEFIELD MyFieldName \\* MERGEFORMAT");

        doc.save(base.artifactsDir + "WorkingWithFields.InsertField.docx");
        //ExEnd:InsertField
    });

    test('InsertFieldUsingFieldBuilder', () => {
        //ExStart:InsertFieldUsingFieldBuilder
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();

        // Prepare IF field with two nested MERGEFIELD fields: { IF "left expression" = "right expression" "Firstname: { MERGEFIELD firstname }" "Lastname: { MERGEFIELD lastname }"}
        let fieldBuilder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
            .addArgument("left expression")
            .addArgument("=")
            .addArgument("right expression")
            .addArgument(
                new aw.Fields.FieldArgumentBuilder()
                    .addText("Firstname: ")
                    .addField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldMergeField).addArgument("firstname")))
            .addArgument(
                new aw.Fields.FieldArgumentBuilder()
                    .addText("Lastname: ")
                    .addField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldMergeField).addArgument("lastname")));

        // Insert IF field in exact location
        let field = fieldBuilder.buildAndInsert(doc.firstSection.body.firstParagraph);
        field.update();

        doc.save(base.artifactsDir + "Field.InsertFieldUsingFieldBuilder.docx");
        //ExEnd:InsertFieldUsingFieldBuilder
    });

    test('InsertAuthorField', () => {
        //ExStart:InsertAuthorField
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();

        let para = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }
        let field = para.appendField(aw.Fields.FieldType.FieldAuthor, false).asFieldAuthor();
        field.authorName = "Test1";

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertAuthorField.docx");
        //ExEnd:InsertAuthorField
    });

    test('InsertAskFieldWithoutDocumentBuilder', () => {
        //ExStart:InsertAskFieldWithoutDocumentBuilder
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();

        let para = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();
        // We want to insert an Ask field like this:
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        let field = para.appendField(aw.Fields.FieldType.FieldAsk, false).asFieldAsk();
        // { ASK \"Test 1\" " }
        field.bookmarkName = "Test 1";
        // { ASK \"Test 1\" Test2 }
        field.promptText = "Test2";
        // { ASK \"Test 1\" Test2 \\d Test3 }
        field.defaultResponse = "Test3";
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        field.promptOnceOnMailMerge = true;

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertAskFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertAskFieldWithoutDocumentBuilder
    });

    test('InsertAdvanceFieldWithoutDocumentBuilder', () => {
        //ExStart:InsertAdvanceFieldWithoutDocumentBuilder
        //GistId:045f68a3af8a7ef327733a8b74034ec5
        let doc = new aw.Document();

        let para = doc.getChild(aw.NodeType.Paragraph, 0, true).asParagraph();
        // We want to insert an Advance field like this:
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        let field = para.appendField(aw.Fields.FieldType.FieldAdvance, false).asFieldAdvance();
        // { ADVANCE \\d 10 " }
        field.downOffset = "10";
        // { ADVANCE \\d 10 \\l 10 }
        field.leftOffset = "10";
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field.rightOffset = "-3.3";
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field.upOffset = "0";
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field.horizontalPosition = "100";
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field.verticalPosition = "100";

        field.update();

        doc.save(base.artifactsDir + "WorkingWithFields.InsertAdvanceFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
    });

    test('FieldDisplayResults', () => {
        //ExStart:FieldDisplayResults
        //GistId:ce43c0268e53b9e7df2f581cafc2d748
        //ExStart:UpdateDocFields
        //GistId:c75335b04abcee0bc8636813bd1b02e8
        let document = new aw.Document(base.myDir + "Various fields.docx");

        document.updateFields();
        //ExEnd:UpdateDocFields

        for (let field of document.range.fields)
            console.log(field.displayResult);
        //ExEnd:FieldDisplayResults
    });

    test('EvaluateIfCondition', () => {
        //ExStart:EvaluateIfCondition
        //GistId:b62cbbccff1a140de484012aafd71fa2
        let builder = new aw.DocumentBuilder();

        let field = builder.insertField("IF 1 = 1", null).asFieldIf();
        let actualResult = field.evaluateCondition();

        console.log(actualResult);
        //ExEnd:EvaluateIfCondition
    });

    test('UnlinkFieldsInParagraph', () => {
        //ExStart:UnlinkFieldsInParagraph
        //GistId:5745cef9bae16cdf430ed2906034a61e
        let doc = new aw.Document(base.myDir + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last
        // paragraph of the document.
        let fields = Array.from(doc.firstSection.body.lastParagraph.range.fields)
            .filter(f => f.type == aw.Fields.FieldType.FieldIf);

        for (let field of fields) {
            field.unlink();
        }

        doc.save(base.artifactsDir + "WorkingWithFields.UnlinkFieldsInParagraph.docx");
        //ExEnd:UnlinkFieldsInParagraph
    });

    test('UnlinkFieldsInDocument', () => {
        //ExStart:UnlinkFieldsInDocument
        //GistId:5745cef9bae16cdf430ed2906034a61e
        let doc = new aw.Document(base.myDir + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        let fields = Array.from(doc.range.fields)
            .filter(f => f.type == aw.Fields.FieldType.FieldIf);

        for (let field of fields) {
            field.unlink();
        }

        // Save the document with fields transformed to disk
        doc.save(base.artifactsDir + "WorkingWithFields.UnlinkFieldsInDocument.docx");
        //ExEnd:UnlinkFieldsInDocument
    });

    test('UnlinkFieldsInBody', () => {
        //ExStart:UnlinkFieldsInBody
        //GistId:5745cef9bae16cdf430ed2906034a61e
        let doc = new aw.Document(base.myDir + "Linked fields.docx");

        // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        let fields = Array.from(doc.firstSection.body.range.fields)
            .filter(f => f.type == aw.Fields.FieldType.FieldPage);

        for (let field of fields) {
            field.unlink();
        }

        doc.save(base.artifactsDir + "WorkingWithFields.UnlinkFieldsInBody.docx");
        //ExEnd:UnlinkFieldsInBody
    });

    //ExStart:ConvertFieldsToStaticText
    //GistId:5745cef9bae16cdf430ed2906034a61e
    /// <summary>
    /// Converts any fields of the specified type found in the descendants of the node into static text.
    /// </summary>
    /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
    /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
    function convertFieldsToStaticText(compositeNode, targetFieldType) {
        let fields = Array.from(compositeNode.range.fields)
            .filter(f => f.type == targetFieldType);

        for (let field of fields) {
            field.unlink();
        }
    }
    //ExEnd:ConvertFieldsToStaticText

});