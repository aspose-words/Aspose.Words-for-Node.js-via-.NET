// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithRtfLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('RecognizeUtf8Text', () => {
    //ExStart:RecognizeUtf8Text
    let loadOptions = new aw.Loading.RtfLoadOptions();
    loadOptions.recognizeUtf8Text = true;

    let doc = new aw.Document(base.myDir + "UTF-8 characters.rtf", loadOptions);
    doc.save(base.artifactsDir + "WorkingWithRtfLoadOptions.RecognizeUtf8Text.rtf");
    //ExEnd:RecognizeUtf8Text
  });

});