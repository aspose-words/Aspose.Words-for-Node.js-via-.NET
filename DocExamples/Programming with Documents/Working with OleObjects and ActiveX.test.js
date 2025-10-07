// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithOleObjectsAndActiveX", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('InsertOleObject', () => {
    //ExStart:InsertOleObject
    //GistId:82ca803e5833cb807b7e1c5111066cb0
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

    doc.save(base.artifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
    //ExEnd:InsertOleObject
  });

  test('InsertOleObjectAsIcon', () => {
    //ExStart:InsertOleObjectAsIcon
    //GistId:82ca803e5833cb807b7e1c5111066cb0
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertOleObjectAsIcon(base.myDir + "Presentation.pptx", false, base.imagesDir + "Logo icon.ico",
        "My embedded file");

    doc.save(base.artifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
    //ExEnd:InsertOleObjectAsIcon
  });

  test('ReadActiveXControlProperties', () => {
    let doc = new aw.Document(base.myDir + "ActiveX controls.docx");

    let properties = "";
    for (let shape of doc.getChildNodes(aw.NodeType.Shape, true)) {
      shape = shape.asShape();
      if (shape.oleFormat === null) break;

        let oleControl = shape.oleFormat.oleControl;
        if (oleControl.isForms2OleControl) {
            let checkBox = oleControl;
            properties = properties + "\nCaption: " + checkBox.caption;
            properties = properties + "\nValue: " + checkBox.value;
            properties = properties + "\nEnabled: " + checkBox.enabled;
            properties = properties + "\nType: " + checkBox.type;
            if (checkBox.childNodes != null) {
                properties = properties + "\nChildNodes: " + checkBox.childNodes;
            }

            properties += "\n";
        }
    }

    properties = properties + "\nTotal ActiveX Controls found: " + doc.getChildNodes(aw.NodeType.Shape, true).count;
    console.log("\n" + properties);
  });

  test('InsertOnlineVideo', () => {
    //ExStart:InsertOnlineVideo
    //GistId:82ca803e5833cb807b7e1c5111066cb0
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let url = "https://youtu.be/t_1LYZ102RA";
    let width = 360;
    let height = 270;

    let shape = builder.insertOnlineVideo(url, width, height);

    doc.save(base.artifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideo.docx");
    //ExEnd:InsertOnlineVideo
  });

});