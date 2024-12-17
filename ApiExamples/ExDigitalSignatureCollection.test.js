// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');


describe("ExDigitalSignatureCollection", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('GetEnumerator', () => {
    //ExStart
    //ExFor:aw.DigitalSignatures.DigitalSignatureCollection.getEnumerator
    //ExSummary:Shows how to print all the digital signatures of a signed document.
    let digitalSignatures =
      aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.myDir + "Digitally signed.docx");

    for (let ds of digitalSignatures) {
      console.log(ds.toString());
    }
    //ExEnd

    expect(digitalSignatures.count).toEqual(1);

    let signature = digitalSignatures.at(0);

    expect(signature.isValid).toEqual(true);
    expect(signature.signatureType).toEqual(aw.DigitalSignatures.DigitalSignatureType.XmlDsig);
    expect(signature.signTime.toISOString()).toEqual("2010-12-23T00:14:40.000Z");
    expect(signature.comments).toEqual("Test Sign");

    /* TODO signature.certificateHolder.certificate is not supported
    expect(signature.certificateHolder.certificate.issuerName.name).toEqual(signature.issuerName);
    expect(signature.certificateHolder.certificate.subjectName.name).toEqual(signature.subjectName);
    */

    expect(signature.issuerName).toEqual("CN=VeriSign Class 3 Code Signing 2009-2 CA, " +
                "OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
                "OU=VeriSign Trust Network, " +
                "O=\"VeriSign, Inc.\", " +
                "C=US");

    expect(signature.subjectName).toEqual("CN=Aspose Pty Ltd, " +
                "OU=Digital ID Class 3 - Microsoft Software Validation v2, " +
                "O=Aspose Pty Ltd, " +
                "L=Lane Cove, " +
                "S=New South Wales, " +
                "C=AU");
  });
});
