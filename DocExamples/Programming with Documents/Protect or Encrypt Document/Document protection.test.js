// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("DocumentProtection", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });


    test('PasswordProtection', () => {
        //ExStart:PasswordProtection
        //GistId:d9e52f106d399d80f5df382419349f58
        let doc = new aw.Document();

        // Apply document protection.
        doc.protect(aw.ProtectionType.NoProtection, "password");

        doc.save(base.artifactsDir + "DocumentProtection.PasswordProtection.docx");
        //ExEnd:PasswordProtection
    });

    test('AllowOnlyFormFieldsProtect', () => {
        //ExStart:AllowOnlyFormFieldsProtect
        //GistId:d9e52f106d399d80f5df382419349f58
        // Insert two sections with some text.
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);
        builder.writeln("Text added to a document.");

        // A document protection only works when document protection is turned and only editing in form fields is allowed.
        doc.protect(aw.ProtectionType.AllowOnlyFormFields, "password");

        // Save the protected document.
        doc.save(base.artifactsDir + "DocumentProtection.AllowOnlyFormFieldsProtect.docx");
        //ExEnd:AllowOnlyFormFieldsProtect
    });

    test('RemoveDocumentProtection', () => {
        //ExStart:RemoveDocumentProtection
        //GistId:d9e52f106d399d80f5df382419349f58
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.writeln("Text added to a document.");

        // Documents can have protection removed either with no password, or with the correct password.
        doc.unprotect();
        doc.protect(aw.ProtectionType.ReadOnly, "newPassword");
        doc.unprotect("newPassword");

        doc.save(base.artifactsDir + "DocumentProtection.RemoveDocumentProtection.docx");
        //ExEnd:RemoveDocumentProtection
    });

    test('UnrestrictedEditableRegions', () => {
        //ExStart:UnrestrictedEditableRegions
        //GistId:d9e52f106d399d80f5df382419349f58
        // Upload a document and make it as read-only.
        let doc = new aw.Document(base.myDir + "Document.docx");
        let builder = new aw.DocumentBuilder(doc);

        doc.protect(aw.ProtectionType.ReadOnly, "MyPassword");

        builder.writeln("Hello world! Since we have set the document's protection level to read-only, " + "we cannot edit this paragraph without the password.");

        // Start an editable range.
        let edRangeStart = builder.startEditableRange();
        // An EditableRange object is created for the EditableRangeStart that we just made.
        let editableRange = edRangeStart.editableRange;

        // Put something inside the editable range.
        builder.writeln("Paragraph inside first editable range");

        // An editable range is well-formed if it has a start and an end.
        let edRangeEnd = builder.endEditableRange();

        builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.");

        doc.save(base.artifactsDir + "DocumentProtection.UnrestrictedEditableRegions.docx");
        //ExEnd:UnrestrictedEditableRegions
    });

    test('UnrestrictedSection', () => {
        //ExStart:UnrestrictedSection
        //GistId:d9e52f106d399d80f5df382419349f58
        // Insert two sections with some text.
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.writeln("Section 1. Unprotected.");
        builder.insertBreak(aw.BreakType.SectionBreakContinuous);
        builder.writeln("Section 2. Protected.");

        // Section protection only works when document protection is turned and only editing in form fields is allowed.
        doc.protect(aw.ProtectionType.AllowOnlyFormFields, "password");

        // By default, all sections are protected, but we can selectively turn protection off.
        doc.sections.at(0).protectedForForms = false;
        doc.save(base.artifactsDir + "DocumentProtection.UnrestrictedSection.docx");

        doc = new aw.Document(base.artifactsDir + "DocumentProtection.UnrestrictedSection.docx");
        expect(doc.sections.at(0).protectedForForms).toBe(false);
        expect(doc.sections.at(1).protectedForForms).toBe(true);
        //ExEnd:UnrestrictedSection
    });

    test('GetProtectionType', () => {
        //ExStart:GetProtectionType
        let doc = new aw.Document(base.myDir + "Document.docx");
        let protectionType = doc.protectionType;
        //ExEnd:GetProtectionType
    });

    test('ReadOnlyProtection', () => {
        //ExStart:ReadOnlyProtection
        //GistId:2a464f0279e5751f4ef94d7daf395e52
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.write("Open document as read-only");

        // Enter a password that's up to 15 characters long.
        doc.writeProtection.setPassword("MyPassword");

        // Make the document as read-only.
        doc.writeProtection.readOnlyRecommended = true;

        // Apply write protection as read-only.
        doc.protect(aw.ProtectionType.ReadOnly);
        doc.save(base.artifactsDir + "DocumentProtection.ReadOnlyProtection.docx");
        //ExEnd:ReadOnlyProtection
    });

    test('RemoveReadOnlyRestriction', () => {
        //ExStart:RemoveReadOnlyRestriction
        //GistId:2a464f0279e5751f4ef94d7daf395e52
        let doc = new aw.Document();

        // Enter a password that's up to 15 characters long.
        doc.writeProtection.setPassword("MyPassword");

        // Remove the read-only option.
        doc.writeProtection.readOnlyRecommended = false;

        // Apply write protection without any protection.
        doc.protect(aw.ProtectionType.NoProtection);
        doc.save(base.artifactsDir + "DocumentProtection.RemoveReadOnlyRestriction.docx");
        //ExEnd:RemoveReadOnlyRestriction
    });
});