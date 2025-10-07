// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithHtmlLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('PreferredControlType', () => {
    //ExStart:LoadHtmlElementsWithPreferredControlType
    let html = `
        <html>
            <select name='ComboBox' size='1'>
                <option value='val1'>item1</option>
                <option value='val2'></option>
            </select>
        </html>`;

    let loadOptions = new aw.Loading.HtmlLoadOptions();
    loadOptions.preferredControlType = aw.Loading.HtmlControlType.StructuredDocumentTag;

    let doc = new aw.Document(Buffer.from(html, 'utf8'), loadOptions);
    doc.save(base.artifactsDir + "WorkingWithHtmlLoadOptions.PreferredControlType.docx", aw.SaveFormat.Docx);
    //ExEnd:LoadHtmlElementsWithPreferredControlType

  });

});