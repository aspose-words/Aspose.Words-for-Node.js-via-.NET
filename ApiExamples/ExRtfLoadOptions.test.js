// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExRtfLoadOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.each([false, true])('RecognizeUtf8Text(%o)', (recognizeUtf8Text) => {
    //ExStart
    //ExFor:RtfLoadOptions
    //ExFor:RtfLoadOptions.#ctor
    //ExFor:aw.Loading.RtfLoadOptions.recognizeUtf8Text
    //ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
    // Create an "RtfLoadOptions" object to modify how we load an RTF document.
    let loadOptions = new aw.Loading.RtfLoadOptions();

    // Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
    // and loads every character in the document.
    // Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
    loadOptions.recognizeUtf8Text = recognizeUtf8Text;

    let doc = new aw.Document(base.myDir + "UTF-8 characters.rtf", loadOptions);

    expect(doc.firstSection.body.getText().trim()).toEqual(recognizeUtf8Text
      ? "“John Doe´s list of currency symbols”™\r€, ¢, £, ¥, ¤"
      : "â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\râ‚¬, Â¢, Â£, Â¥, Â¤");
    //ExEnd
  });


});
