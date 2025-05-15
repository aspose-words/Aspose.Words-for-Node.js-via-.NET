// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;

describe("HelloWorld", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('SimpleHelloWorld', () => {
    //ExStart:HelloWorld
    //GistId:43e1e4f8d1f0c53662b750993b354108
    let docA = new aw.Document();            
    let builder = new aw.DocumentBuilder(docA);

    // Insert text to the document start.
    builder.moveToDocumentStart();
    builder.write("First Hello World paragraph");

    let docB = new aw.Document(base.myDir + "Document.docx");
    // Add document B to the and of document A, preserving document B formatting.
    docA.appendDocument(docB, aw.ImportFormatMode.KeepSourceFormatting);
            
    docA.save(base.artifactsDir + "HelloWorld.SimpleHelloWorld.pdf");
    //ExEnd:HelloWorld
  });
 });

