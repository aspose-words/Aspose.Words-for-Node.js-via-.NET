// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithHyphenation", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('HyphenateWords', () => {
    //ExStart:HyphenateWords
    //GistId:b8ff496530bd59b589a3beb20281c4de
    let doc = new aw.Document(base.myDir + "German text.docx");
    aw.Hyphenation.registerDictionary("en-US", base.myDir + "hyph_en_US.dic");
    aw.Hyphenation.registerDictionary("de-CH", base.myDir + "hyph_de_CH.dic");
    doc.save(base.artifactsDir + "WorkingWithHyphenation.HyphenateWords.pdf");
    //ExEnd:HyphenateWords
  });

  test('LoadHyphenationDictionary', () => {
    //ExStart:LoadHyphenationDictionary
    //GistId:b8ff496530bd59b589a3beb20281c4de
    let doc = new aw.Document(base.myDir + "German text.docx");

    let stream = base.loadFileToBuffer(base.myDir + "hyph_de_CH.dic");
    aw.Hyphenation.registerDictionary("de-CH", stream);
    doc.save(base.artifactsDir + "WorkingWithHyphenation.LoadHyphenationDictionary.pdf");
    //ExEnd:LoadHyphenationDictionary
  });

});