// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const fs = require('fs');


describe("ExPlainTextDocument", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });
 
  test('Load', () => {
    //ExStart
    //ExFor:PlainTextDocument
    //ExFor:PlainTextDocument.#ctor(String)
    //ExFor:aw.PlainTextDocument.text
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext.
    let doc = new aw.Document(); 
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    doc.save(base.artifactsDir + "PlainTextDocument.load.docx");

    let plaintext = new aw.PlainTextDocument(base.artifactsDir + "PlainTextDocument.load.docx");

    expect(plaintext.text.trim()).toEqual("Hello world!");
    //ExEnd
  });


  test.skip('LoadFromStream - ', () => {
    //ExStart
    //ExFor:PlainTextDocument.#ctor(Stream)
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext using stream.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    doc.save(base.artifactsDir + "PlainTextDocument.LoadFromStream.docx");

    var buffer = base.loadFileToBuffer(base.artifactsDir + "PlainTextDocument.LoadFromStream.docx");
    let plaintext = new aw.PlainTextDocument(buffer);
    expect(plaintext.text.trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('LoadEncrypted', () => {
    //ExStart
    //ExFor:PlainTextDocument.#ctor(String, LoadOptions)
    //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.password = "MyPassword";

    doc.save(base.artifactsDir + "PlainTextDocument.LoadEncrypted.docx", saveOptions);

    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.password = "MyPassword";

    let plaintext = new aw.PlainTextDocument(base.artifactsDir + "PlainTextDocument.LoadEncrypted.docx", loadOptions);

    expect(plaintext.text.trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('LoadEncryptedUsingStream', () => {
    //ExStart
    //ExFor:PlainTextDocument.#ctor(Stream, LoadOptions)
    //ExSummary:Shows how to load the contents of an encrypted Microsoft Word document in plaintext using stream.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    let saveOptions = new aw.Saving.OoxmlSaveOptions();
    saveOptions.password = "MyPassword";

    doc.save(base.artifactsDir + "PlainTextDocument.LoadFromStreamWithOptions.docx", saveOptions);

    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.password = "MyPassword";

    var buffer = base.loadFileToBuffer(base.artifactsDir + "PlainTextDocument.LoadFromStreamWithOptions.docx");
    let plaintext = new aw.PlainTextDocument(buffer, loadOptions);
    expect(plaintext.text.trim()).toEqual("Hello world!");
    //ExEnd
  });


  test('BuiltInProperties', () => {
    //ExStart
    //ExFor:aw.PlainTextDocument.builtInDocumentProperties
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's built-in properties.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    doc.builtInDocumentProperties.author = "John Doe";

    doc.save(base.artifactsDir + "PlainTextDocument.BuiltInProperties.docx");

    let plaintext = new aw.PlainTextDocument(base.artifactsDir + "PlainTextDocument.BuiltInProperties.docx");

    expect(plaintext.text.trim()).toEqual("Hello world!");
    expect(plaintext.builtInDocumentProperties.author).toEqual("John Doe");
    //ExEnd
  });


  test('CustomDocumentProperties', () => {
    //ExStart
    //ExFor:aw.PlainTextDocument.customDocumentProperties
    //ExSummary:Shows how to load the contents of a Microsoft Word document in plaintext and then access the original document's custom properties.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    doc.customDocumentProperties.add("Location of writing", "123 Main St, London, UK");

    doc.save(base.artifactsDir + "PlainTextDocument.customDocumentProperties.docx");

    let plaintext = new aw.PlainTextDocument(base.artifactsDir + "PlainTextDocument.customDocumentProperties.docx");

    expect(plaintext.text.trim()).toEqual("Hello world!");
    expect(plaintext.customDocumentProperties.at("Location of writing").toString()).toEqual("123 Main St, London, UK");
    //ExEnd
  });
});
