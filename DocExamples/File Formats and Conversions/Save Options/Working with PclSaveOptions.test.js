// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("WorkingWithPclSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('RasterizeTransformedElements', () => {
    //ExStart:RasterizeTransformedElements
    //GistId:757cf7d3534a39730cf3290d418681ab
    let doc = new aw.Document(base.myDir + "Rendering.docx");

    let saveOptions = new aw.Saving.PclSaveOptions();
    saveOptions.saveFormat = aw.SaveFormat.Pcl;
    saveOptions.rasterizeTransformedElements = false;

    doc.save(base.artifactsDir + "WorkingWithPclSaveOptions.rasterizeTransformedElements.pcl", saveOptions);
    //ExEnd:RasterizeTransformedElements
  });

});
