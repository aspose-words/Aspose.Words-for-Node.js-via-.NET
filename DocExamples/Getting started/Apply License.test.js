// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;

describe("ApplyLicense", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('ApplyLicenseFromFile', () => {
    //ExStart:ApplyLicenseFromFile
    //GistId:c65baf943b2663664d4cc2492cd33464
    let license = new aw.License();

    // This line attempts to set a license from location relative to the executable.
    // You can also use the additional overload to load a license from a Buffer, this is useful,
    // for instance, when the license is stored as an embedded resource.
    try {
      license.setLicense("Aspose.Words.NodeJs.NET.lic");
                
      console.log("License set successfully.");
    } catch(err) {
      // We do not ship any license with this example,
      // visit the Aspose site to obtain either a temporary or permanent license. 
      console.error(`\nThere was an error setting the license:  ${err}`);
    }
    //ExEnd:ApplyLicenseFromFile
  });


  test('ApplyLicenseFromStream', () => {
    //ExStart:ApplyLicenseFromStream
    //GistId:c65baf943b2663664d4cc2492cd33464
    let license = new aw.License();

    try {
      license.setLicense(base.loadFileToBuffer("Aspose.Words.NodeJs.NET.lic"));
                
      console.log("License set successfully.");
    } catch(err){
      // We do not ship any license with this example,
      // visit the Aspose site to obtain either a temporary or permanent license. 
      console.error(`\nThere was an error setting the license:  ${err}`);
    }
    //ExEnd:ApplyLicenseFromStream
  });


  test.skip('ApplyMeteredLicense - TODO: not implemented yet', () => {
    //ExStart:ApplyMeteredLicense
    //GistId:c65baf943b2663664d4cc2492cd33464
    try
    {
      let metered = new aw.Metered();
      metered.setMeteredKey("*****", "*****");

      let doc = new aw.Document(base.myDir + "Document.docx");

      console.log(doc.pageCount);
    } catch(err) {
      console.log(`\nThere was an error setting the license:  ${err}`);
    }
    //ExEnd:ApplyMeteredLicense
  });
});