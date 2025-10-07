// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithDigitalSignatures", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });


    test('AccessAndVerifySignature', () => {
        //ExStart:AccessAndVerifySignature
        let doc = new aw.Document(base.myDir + "Digitally signed.docx");

        for (let signature of doc.digitalSignatures) {
            console.log("*** Signature Found ***");
            console.log("Is valid: " + signature.isValid);
            // This property is available in MS Word documents only.
            console.log("Reason for signing: " + signature.comments);
            console.log("Time of signing: " + signature.signTime);
            console.log();
        }
        //ExEnd:AccessAndVerifySignature
    });

    test('RemoveSignatures', async () => {
        //ExStart:RemoveSignatures
        //GistId:246abc8bf535665565cc872be9b805ac
        // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
        // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
        // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
        aw.DigitalSignatures.DigitalSignatureUtil.removeAllSignatures(base.myDir + "Digitally signed.docx",
            base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

        // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
        let fs = require('fs');
        let streamIn = base.loadFileToBuffer(base.myDir + "Digitally signed.docx");
        let streamOut = fs.createWriteStream(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx");

        aw.DigitalSignatures.DigitalSignatureUtil.removeAllSignatures(streamIn, streamOut);
        await new Promise(resolve => streamOut.on("finish", resolve));

        // Verify that both our output documents have no digital signatures.
        expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx").count).toBe(0);
        expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx").count).toBe(0);
        //ExEnd:RemoveSignatures
    });

    test('SignatureValue', () => {
        //ExStart:SignatureValue
        //GistId:246abc8bf535665565cc872be9b805ac
        let doc = new aw.Document(base.myDir + "Digitally signed.docx");

        for (let digitalSignature of doc.digitalSignatures) {
            let signatureValue = Buffer.from(digitalSignature.signatureValue).toString('base64');
            expect(signatureValue).toBe("K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD" +
                "MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" +
                "+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=");
        }
        //ExEnd:SignatureValue
    });
});