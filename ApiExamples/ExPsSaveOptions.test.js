// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExPsSaveOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.each([false, true])('UseBookFoldPrintingSettings(%o)', (renderTextAsBookFold) => {
    //ExStart
    //ExFor:PsSaveOptions
    //ExFor:PsSaveOptions.saveFormat
    //ExFor:PsSaveOptions.useBookFoldPrintingSettings
    //ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");

    // Create a "PsSaveOptions" object that we can pass to the document's "Save" method
    // to modify how that method converts the document to PostScript.
    // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
    // in the output Postscript document in a way that helps us make a booklet out of it.
    // Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
    let saveOptions = new aw.Saving.PsSaveOptions()
    saveOptions.saveFormat = aw.SaveFormat.Ps;
    saveOptions.useBookFoldPrintingSettings = renderTextAsBookFold;

    // If we are rendering the document as a booklet, we must set the "MultiplePages"
    // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
    for (let s of doc.sections)
    {
      let section = s.asSection();
      section.pageSetup.multiplePages = aw.Settings.MultiplePagesType.BookFoldPrinting;
    }

    // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
    // and the contents will line up in a way that creates a booklet.
    doc.save(base.artifactsDir + "PsSaveOptions.useBookFoldPrintingSettings.ps", saveOptions);
    //ExEnd
  });


});
