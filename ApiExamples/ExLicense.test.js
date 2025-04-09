// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const fs = require('fs');
const path = require('path');


describe("ExLicense", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('LicenseFromFileNoPath', () => {
    //ExStart
    //ExFor:License
    //ExFor:License.#ctor
    //ExFor:aw.License.setLicense(String)
    //ExSummary:Shows how initialize a license for Aspose.words using a license file in the local file system.
    // Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
    const licFilename = "Aspose.Words.NodeJs.NET.lic";
    var licenseFileName =  path.join(base.licenseDir, licFilename);

    let license = new aw.License();
    try {
      license.setLicense(licenseFileName);

      // Create a copy of our license file in the binaries folder of our application.
      var licenseCopyFileName = path.join(base.codeBaseDir, licFilename);
      fs.copyFileSync(licenseFileName, licenseCopyFileName);

      // If we pass a file's name without a path,
      // the SetLicense will search several local file system locations for this file.
      // One of those locations will be the "bin" folder, which contains a copy of our license file.
      license.setLicense(licFilename);
      //ExEnd

      fs.unlinkSync(licenseCopyFileName);

      expect(() => license.setLicense(licFilename)).toThrow(`Cannot find license '${licFilename}'.`);
    }
    finally {
      license.setLicense("");
    }
  });


  test('LicenseFromStream', () => {
    //ExStart
    //ExFor:aw.License.setLicense(Stream)
    //ExSummary:Shows how to initialize a license for Aspose.words from a stream.
    // Set the license for our Aspose.words product by passing a stream for a valid license file in our local file system.
    let data = base.loadFileToBuffer(path.join(base.licenseDir, "Aspose.Words.NodeJs.NET.lic"));
    let license = new aw.License();
    license.setLicense(data);
    //ExEnd
  });
});
