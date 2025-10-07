// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;

describe("WorkingWithOfficeMath", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('MathEquations', () => {
    //ExStart:MathEquations
    //GistId:7a6483a866277e7ecf45b19756b1da06
    let doc = new aw.Document(base.myDir + "Office math.docx");

    let officeMath = doc.getChild(aw.NodeType.OfficeMath, 0, true);
    // OfficeMath display type represents whether an equation is displayed inline with the text or displayed on its line.
    officeMath.displayType = aw.Math.OfficeMathDisplayType.Display;
    officeMath.justification = aw.Math.OfficeMathJustification.Left;

    doc.save(base.artifactsDir + "WorkingWithOfficeMath.MathEquations.docx");
    //ExEnd:MathEquations
  });
});
