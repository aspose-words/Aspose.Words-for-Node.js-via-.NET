// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithTextboxes", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CreateLink', () => {
    //ExStart:CreateLink
    //GistId:e78f2e5545401312af45ab0be0f09bb2
    let doc = new aw.Document();

    let shape1 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    let shape2 = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);

    let textBox1 = shape1.textBox;
    let textBox2 = shape2.textBox;

    if (textBox1.isValidLinkTarget(textBox2))
      textBox1.next = textBox2;
    //ExEnd:CreateLink
  });

  test('CheckSequence', () => {
    //ExStart:CheckSequence
    //GistId:e78f2e5545401312af45ab0be0f09bb2
    let doc = new aw.Document();

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    let textBox = shape.textBox;

    if (textBox.next != null && textBox.previous == null)
      console.log("The head of the sequence");

    if (textBox.next != null && textBox.previous != null)
      console.log("The Middle of the sequence.");

    if (textBox.next == null && textBox.previous != null)
      console.log("The Tail of the sequence.");
    //ExEnd:CheckSequence
  });

  test('BreakLink', () => {
    //ExStart:BreakLink
    //GistId:e78f2e5545401312af45ab0be0f09bb2
    let doc = new aw.Document();

    let shape = new aw.Drawing.Shape(doc, aw.Drawing.ShapeType.TextBox);
    let textBox = shape.textBox;

    // Break a forward link.
    textBox.breakForwardLink();

    // Break a forward link by setting a null.
    textBox.next = null;

    // Break a link, which leads to this textbox.
    if (textBox.previous != null)
      textBox.previous.breakForwardLink();
    //ExEnd:BreakLink
  });

});