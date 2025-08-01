﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');
const MemoryStream = require('memorystream');


describe("WorkingWithMarkdownSaveOptions", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('MarkdownTableContentAlignment', () => {
    //ExStart:MarkdownTableContentAlignment
    //GistId:757cf7d3534a39730cf3290d418681ab
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertCell();
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Right;
    builder.write("Cell1");
    builder.insertCell();
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
    builder.write("Cell2");

    // Makes all paragraphs inside the table to be aligned.
    let saveOptions = new aw.Saving.MarkdownSaveOptions();
    saveOptions.tableContentAlignment = aw.Saving.TableContentAlignment.Left;
    doc.save(base.artifactsDir + "WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

    saveOptions.tableContentAlignment = aw.Saving.TableContentAlignment.Right;
    doc.save(base.artifactsDir + "WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

    saveOptions.tableContentAlignment = aw.Saving.TableContentAlignment.Center;
    doc.save(base.artifactsDir + "WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

    // The alignment in this case will be taken from the first paragraph in corresponding table column.
    saveOptions.tableContentAlignment = aw.Saving.TableContentAlignment.Auto;
    doc.save(base.artifactsDir + "WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
    //ExEnd:MarkdownTableContentAlignment
  });


  test('ImagesFolder', () => {
    //ExStart:ImagesFolder
    //GistId:a2fee7fa3d8e5704ce24f041be9a4821
    let doc = new aw.Document(base.myDir + "Image bullet points.docx");

    let saveOptions = new aw.Saving.MarkdownSaveOptions();
    saveOptions.imagesFolder = base.artifactsDir + "Images";

    let stream = new MemoryStream()
    doc.save(stream, saveOptions);
    //ExEnd:ImagesFolder
  });

});