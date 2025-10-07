// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithRanges", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('RangesDeleteText', () => {
      //ExStart:RangesDeleteText
      //GistId:5abf4b66965fca92533f9a266a06c7ed
      let doc = new aw.Document(base.myDir + "Document.docx");
      doc.sections.at(0).range.delete();
      //ExEnd:RangesDeleteText
  });

  test('RangesGetText', () => {
      //ExStart:RangesGetText
      //GistId:5abf4b66965fca92533f9a266a06c7ed
      let doc = new aw.Document(base.myDir + "Document.docx");
      let text = doc.range.text;
      //ExEnd:RangesGetText
  });
});