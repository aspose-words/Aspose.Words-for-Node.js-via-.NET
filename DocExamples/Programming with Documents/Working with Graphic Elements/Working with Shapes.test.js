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


describe("WorkingWithShapes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('GetActualShapeBoundsPoints', () => {
    //ExStart:GetActualShapeBoundsPoints
    //GistId:3a90c8783e87c53371d103d9350f1d31
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let shape = builder.insertImage(base.imagesDir + "Transparent background logo.png");
    shape.aspectRatioLocked = false;

    console.log("\nGets the actual bounds of the shape in points: ");
    console.log(shape.getShapeRenderer().boundsInPoints2);
    //ExEnd:GetActualShapeBoundsPoints
  });

});
