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


describe("WorkingWithRevisions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  
  test('ShowRevisionsInBalloons', () => {
    //ExStart:ShowRevisionsInBalloons
    //GistId:829442fe4196eb8eb1ec945902f2f8ae
    let doc = new aw.Document(base.myDir + "Revisions.docx");

    doc.layoutOptions.revisionOptions.showInBalloons = aw.Layout.ShowInBalloons.FormatAndDelete;
    doc.layoutOptions.revisionOptions.measurementUnit = aw.MeasurementUnits.Inches;
    // Renders revision bars on the right side of a page.
    doc.layoutOptions.revisionOptions.revisionBarsPosition = aw.Drawing.HorizontalAlignment.Right;

    doc.save(base.artifactsDir + "WorkingWithRevisions.ShowRevisionsInBalloons.pdf");
    //ExEnd:ShowRevisionsInBalloons
  });

});
