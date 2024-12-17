// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExComHelper", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.skip('ComHelper - TODO: The document appears to be corrupted and cannot be loaded on comHelper.open(stream).', () => {
    //ExStart
    //ExFor:ComHelper
    //ExFor:ComHelper.#ctor
    //ExFor:aw.ComHelper.open(Stream)
    //ExFor:aw.ComHelper.open(String)
    //ExSummary:Shows how to open documents using the ComHelper class.
    // The ComHelper class allows us to load documents from within COM clients.
    let comHelper = new aw.ComHelper();

    // 1 -  Using a local system filename:
    let doc = comHelper.open(base.myDir + "Document.docx");

    expect(doc.getText().trim()).toEqual("Hello World!\r\rHello Word!\r\r\rHello World!");

    // 2 -  From a stream:
    let stream = base.loadFileToBuffer(base.myDir + "Document.docx");
    doc = comHelper.open(stream);
    expect(doc.getText().trim()).toEqual("Hello World!\r\rHello Word!\r\r\rHello World!");
    //ExEnd
  });
});
