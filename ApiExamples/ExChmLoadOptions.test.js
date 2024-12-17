// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExChmLoadOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('OriginalFileName', () => {
    //ExStart
    //ExFor:aw.Loading.ChmLoadOptions.originalFileName
    //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
    // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
    // so file links don't work after saving it to HTML.
    // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
    let loadOptions = new aw.Loading.ChmLoadOptions();
    loadOptions.originalFileName = "amhelp.chm";

    let doc = new aw.Document(base.myDir + "Document with ms-its links.chm", loadOptions);
            
    doc.save(base.artifactsDir + "ExChmLoadOptions.originalFileName.html");
    //ExEnd
  });

});
