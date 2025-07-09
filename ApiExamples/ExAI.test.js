// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;


describe("ExAI", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('AiSummarize', () => {
    //ExStart:AiSummarize
    //GistId:757cf7d3534a39730cf3290d418681ab
    //ExFor:GoogleAiModel
    //ExFor:OpenAiModel
    //ExFor:OpenAiModel.setOrganization(String)
    //ExFor:OpenAiModel.setProject(String)
    //ExFor:IAiModelText
    //ExFor:IAiModelText.summarize(Document, SummarizeOptions)
    //ExFor:IAiModelText.summarize(Document[], SummarizeOptions)
    //ExFor:SummarizeOptions
    //ExFor:SummarizeOptions.#ctor
    //ExFor:SummarizeOptions.summaryLength
    //ExFor:SummaryLength
    //ExFor:AiModel
    //ExFor:AiModel.create(AiModelType)
    //ExFor:AiModel.setApiKey(String)
    //ExFor:AiModelType
    //ExSummary:Shows how to summarize text using OpenAI and Google models.
    let firstDoc = new aw.Document(base.myDir + "Big document.docx");
    let secondDoc = new aw.Document(base.myDir + "Document.docx");

    const apiKey = process.env.API_KEY;
    if (!apiKey) {
      console.warn("API_KEY environment variable is not set.");
      return;
    }

    // Use OpenAI or Google generative language models.
    let model = aw.AI.AiModel.createGpt4OMini();
    model.setApiKey(apiKey);
    model.setOrganization("Organization");
    model.setProject("Project");

    let options = new aw.AI.SummarizeOptions();

    options.summaryLength = aw.AI.SummaryLength.Short;
    let oneDocumentSummary = model.summarize(firstDoc, options);
    oneDocumentSummary.save(base.artifactsDir + "AI.AiSummarize.one.docx");

    options.summaryLength = aw.AI.SummaryLength.Long;
    let multiDocumentSummary = model.summarize([firstDoc, secondDoc], options);
    multiDocumentSummary.save(base.artifactsDir + "AI.AiSummarize.multi.docx");
    //ExEnd:AiSummarize
  });


  test('AiTranslate', () => {
    //ExStart:AiTranslate
    //GistId:757cf7d3534a39730cf3290d418681ab
    //ExFor:IAiModelText.translate(Document, AI.language)
    //ExFor:AI.language
    //ExSummary:Shows how to translate text using Google models.
    let doc = new aw.Document(base.myDir + "Document.docx");

    const apiKey = process.env.API_KEY;
    if (!apiKey) {
      console.warn("API_KEY environment variable is not set.");
      return;
    }

    // Use Google generative language models.
    let model = aw.AI.AiModel.createGemini15Flash();
    model.setApiKey(apiKey);

    let translatedDoc = model.translate(doc, aw.AI.Language.Arabic);
    translatedDoc.save(base.artifactsDir + "AI.AiTranslate.docx");
    //ExEnd:AiTranslate
  });


  test('AiGrammar', () => {
    //ExStart:AiGrammar
    //GistId:757cf7d3534a39730cf3290d418681ab
    //ExFor:IAiModelText.checkGrammar(Document, CheckGrammarOptions)
    //ExFor:CheckGrammarOptions
    //ExSummary:Shows how to check the grammar of a document.
    let doc = new aw.Document(base.myDir + "Big document.docx");

    const apiKey = process.env.API_KEY;
    if (!apiKey) {
      console.warn("API_KEY environment variable is not set.");
      return;
    }

    // Use OpenAI generative language models.
    let model = aw.AI.AiModel.createGpt4OMini();
    model.setApiKey(apiKey);

    let grammarOptions = new aw.AI.CheckGrammarOptions();
    grammarOptions.improveStylistics = true;

    let proofedDoc = model.checkGrammar(doc, grammarOptions);
    proofedDoc.save(base.artifactsDir + "AI.AiGrammar.docx");
    //ExEnd:AiGrammar
  });
});
