// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExBuildVersion", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('PrintBuildVersionInfo', () => {
    //ExStart
    //ExFor:BuildVersionInfo
    //ExFor:BuildVersionInfo.product
    //ExFor:BuildVersionInfo.version
    //ExSummary:Shows how to display information about your installed version of Aspose.words.
    console.log(`I am currently using ${aw.BuildVersionInfo.product}, version number ${aw.BuildVersionInfo.version}!`);
    //ExEnd

    expect(aw.BuildVersionInfo.product).toEqual("Aspose.Words for Node.js via .NET");
    const r = new RegExp("[0-9]{2}.[0-9]{1,2}");
    expect(r.test(aw.BuildVersionInfo.version)).toEqual(true);
  });
});
