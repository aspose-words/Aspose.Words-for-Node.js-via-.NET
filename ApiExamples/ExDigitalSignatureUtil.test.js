// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const fs = require('fs');


describe("ExDigitalSignatureUtil", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Load', () => {
    //ExStart
    //ExFor:DigitalSignatureUtil
    //ExFor:DigitalSignatureUtil.loadSignatures(String)
    //ExFor:DigitalSignatureUtil.loadSignatures(Stream)
    //ExSummary:Shows how to load signatures from a digitally signed document.
    // There are two ways of loading a signed document's collection of digital signatures using the DigitalSignatureUtil class.
    // 1 -  Load from a document from a local file system filename:
    let digitalSignatures =
      aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.myDir + "Digitally signed.docx");

    // If this collection is nonempty, then we can verify that the document is digitally signed.
    expect(digitalSignatures.count).toEqual(1);

    // 2 -  Load from a document from a Buffer:
    let data = base.loadFileToBuffer(base.myDir + "Digitally signed.docx");
    digitalSignatures = aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(data);
    expect(digitalSignatures.count).toEqual(1);
    //ExEnd
  });


  test('Remove', async () => {
    //ExStart
    //ExFor:DigitalSignatureUtil
    //ExFor:DigitalSignatureUtil.loadSignatures(String)
    //ExFor:DigitalSignatureUtil.removeAllSignatures(Stream, Stream)
    //ExFor:DigitalSignatureUtil.removeAllSignatures(String, String)
    //ExSummary:Shows how to remove digital signatures from a digitally signed document.
    // There are two ways of using the DigitalSignatureUtil class to remove digital signatures
    // from a signed document by saving an unsigned copy of it somewhere else in the local file system.
    // 1 - Determine the locations of both the signed document and the unsigned copy by filename strings:
    aw.DigitalSignatures.DigitalSignatureUtil.removeAllSignatures(base.myDir + "Digitally signed.docx",
      base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx");

    // 2 - Determine the locations of both the signed document and the unsigned copy by file streams:
    let streamIn = base.loadFileToBuffer(base.myDir + "Digitally signed.docx");
    let streamOut = fs.createWriteStream(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx");

    aw.DigitalSignatures.DigitalSignatureUtil.removeAllSignatures(streamIn, streamOut);
    await new Promise(resolve => streamOut.on("finish", resolve));

    // Verify that both our output documents have no digital signatures.
    expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromString.docx").count).toEqual(0);
    expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "DigitalSignatureUtil.LoadAndRemove.FromStream.docx").count).toEqual(0);
    //ExEnd
  });


  test('RemoveSignatures', () => {
    aw.DigitalSignatures.DigitalSignatureUtil.removeAllSignatures(base.myDir + "Digitally signed.odt",
      base.artifactsDir + "DigitalSignatureUtil.RemoveSignatures.odt");

    expect(aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(base.artifactsDir + "DigitalSignatureUtil.RemoveSignatures.odt").count).toEqual(0);
  });

/*
    [Description("WORDSNET-16868")]
    [AotTests.IgnoreAot("CertificateHolder.Create is not AOT compatible.")]
  test('SignDocument', () => {
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:CertificateHolder.create(String, String)
    //ExFor:DigitalSignatureUtil.sign(Stream, Stream, CertificateHolder, SignOptions)
    //ExFor:DigitalSignatures.signOptions
    //ExFor:SignOptions.comments
    //ExFor:SignOptions.signTime
    //ExSummary:Shows how to digitally sign documents.
    // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    // Create a comment and date which will be applied with our new digital signature.
    let signOptions = new aw.DigitalSignatures.SignOptions
    {
      Comments = "My comment",
      SignTime = Date.now()
    };

    // Take an unsigned document from the local file system via a file stream,
    // then create a signed copy of it determined by the filename of the output file stream.
    using (Stream streamIn = new FileStream(base.myDir + "Document.docx", FileMode.open))
    {
      using (Stream streamOut = new FileStream(base.artifactsDir + "DigitalSignatureUtil.SignDocument.docx", FileMode.OpenOrCreate))
      {
        aw.DigitalSignatures.DigitalSignatureUtil.sign(streamIn, streamOut, certificateHolder, signOptions);
      }
    }
    //ExEnd

    using (Stream stream = new FileStream(base.artifactsDir + "DigitalSignatureUtil.SignDocument.docx", FileMode.open))
    {
      let digitalSignatures = aw.DigitalSignatures.DigitalSignatureUtil.loadSignatures(stream);
      expect(digitalSignatures.count).toEqual(1);

      let signature = digitalSignatures.at(0);

      expect(signature.isValid).toEqual(true);
      expect(signature.signatureType).toEqual(aw.DigitalSignatures.DigitalSignatureType.XmlDsig);
      expect(signature.signTime.toString()).toEqual(signOptions.signTime.toString());
      expect(signature.comments).toEqual("My comment");
    }
  });


    [Description("WORDSNET-16868")]
    [AotTests.IgnoreAot("CertificateHolder.Create is not AOT compatible.")]
  test('DecryptionPassword', () => {
    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignOptions.decryptionPassword
    //ExFor:LoadOptions.password
    //ExSummary:Shows how to sign encrypted document file.
    // Create an X.509 certificate from a PKCS#12 store, which should contain a private key.
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    // Create a comment, date, and decryption password which will be applied with our new digital signature.
    let signOptions = new aw.DigitalSignatures.SignOptions
    {
      Comments = "Comment",
      SignTime = Date.now(),
      DecryptionPassword = "docPassword"
    };

    // Set a local system filename for the unsigned input document, and an output filename for its new digitally signed copy.
    string inputFileName = base.myDir + "Encrypted.docx";
    string outputFileName = base.artifactsDir + "DigitalSignatureUtil.decryptionPassword.docx";

    aw.DigitalSignatures.DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd

    // Open encrypted document from a file.
    let loadOptions = new aw.Loading.LoadOptions("docPassword");
    expect(loadOptions.password).toEqual(signOptions.decryptionPassword);

    // Check that encrypted document was successfully signed.
    let signedDoc = new aw.Document(outputFileName, loadOptions);
    let signatures = signedDoc.digitalSignatures;

    expect(signatures.count).toEqual(1);
    expect(signatures.isValid).toEqual(true);
  });


    [Description("WORDSNET-13036, WORDSNET-16868")]
    [AotTests.IgnoreAot("CertificateHolder.Create is not AOT compatible.")]
  test('SignDocumentObfuscationBug', () => {
    let ch = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    let doc = new aw.Document(base.myDir + "Structured document tags.docx");
    string outputFileName = base.artifactsDir + "DigitalSignatureUtil.SignDocumentObfuscationBug.doc";

    let signOptions = new aw.DigitalSignatures.SignOptions { Comments = "Comment", SignTime = Date.now() };

    aw.DigitalSignatures.DigitalSignatureUtil.sign(doc.originalFileName, outputFileName, ch, signOptions);
  });


    [Description("WORDSNET-16868")]
    [AotTests.IgnoreAot("CertificateHolder.Create is not AOT compatible.")]
  test('IncorrectDecryptionPassword', () => {
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");

    let doc = new aw.Document(base.myDir + "Encrypted.docx", new aw.Loading.LoadOptions("docPassword"));
    string outputFileName = base.artifactsDir + "DigitalSignatureUtil.IncorrectDecryptionPassword.docx";

    let signOptions = new aw.DigitalSignatures.SignOptions
    {
      Comments = "Comment",
      SignTime = Date.now(),
      DecryptionPassword = "docPassword1"
    };

    Assert.Throws<IncorrectPasswordException>(
      () => aw.DigitalSignatures.DigitalSignatureUtil.sign(doc.originalFileName, outputFileName, certificateHolder, signOptions),
      "The document password is incorrect.");
  });
*/

  test('NoArgumentsForSing', () => {
    let signOptions = new aw.DigitalSignatures.SignOptions();
    signOptions.comments = '';
    signOptions.signTime = Date.now();
    signOptions.decryptionPassword = '';

    expect(() => aw.DigitalSignatures.DigitalSignatureUtil.sign('', '', null, signOptions))
      .toThrow("The argument cannot be null or empty string. (Parameter 'srcFileName')");
  });


  test('NoCertificateForSign', () => {
    let doc = new aw.Document(base.myDir + "Digitally signed.docx");
    let outputFileName = base.artifactsDir + "DigitalSignatureUtil.NoCertificateForSign.docx";

    let signOptions = new aw.DigitalSignatures.SignOptions();
    signOptions.comments = "Comment";
    signOptions.signTime = Date.now();
    signOptions.decryptionPassword = "docPassword";

    expect(() => aw.DigitalSignatures.DigitalSignatureUtil.sign(doc.originalFileName, outputFileName, null, signOptions))
      .toThrow("Value cannot be null.");
  });


  test('XmlDsig', () => {
    //ExStart:XmlDsig
    //GistId:e06aa7a168b57907a5598e823a22bf0a
    //ExFor:SignOptions.xmlDsigLevel
    //ExFor:XmlDsigLevel
    //ExSummary:Shows how to sign document based on XML-DSig standard.
    let certificateHolder = aw.DigitalSignatures.CertificateHolder.create(base.myDir + "morzal.pfx", "aw");
    let signOptions = new aw.DigitalSignatures.SignOptions();
    signOptions.xmlDsigLevel = aw.DigitalSignatures.XmlDsigLevel.XAdEsEpes;

    let inputFileName = base.myDir + "Document.docx";
    let outputFileName = base.artifactsDir + "DigitalSignatureUtil.xmlDsig.docx";
    aw.DigitalSignatures.DigitalSignatureUtil.sign(inputFileName, outputFileName, certificateHolder, signOptions);
    //ExEnd:XmlDsig
  });

});
