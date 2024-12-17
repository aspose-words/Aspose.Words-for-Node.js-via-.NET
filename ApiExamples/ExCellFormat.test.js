// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExCellFormat", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('VerticalMerge', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.endRow
    //ExFor:CellMerge
    //ExFor:aw.Tables.CellFormat.verticalMerge
    //ExSummary:Shows how to merge table cells vertically.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a cell into the first column of the first row.
    // This cell will be the first in a range of vertically merged cells.
    builder.insertCell();
    builder.cellFormat.verticalMerge = aw.Tables.CellMerge.First;
    builder.write("Text in merged cells.");

    // Insert a cell into the second column of the first row, then end the row.
    // Also, configure the builder to disable vertical merging in created cells.
    builder.insertCell();
    builder.cellFormat.verticalMerge = aw.Tables.CellMerge.None;
    builder.write("Text in unmerged cell.");
    builder.endRow();

    // Insert a cell into the first column of the second row. 
    // Instead of adding text contents, we will merge this cell with the first cell that we added directly above.
    builder.insertCell();
    builder.cellFormat.verticalMerge = aw.Tables.CellMerge.Previous;

    // Insert another independent cell in the second column of the second row.
    builder.insertCell();
    builder.cellFormat.verticalMerge = aw.Tables.CellMerge.None;
    builder.write("Text in unmerged cell.");
    builder.endRow();
    builder.endTable();

    doc.save(base.artifactsDir + "CellFormat.verticalMerge.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "CellFormat.verticalMerge.docx");
    let table = doc.firstSection.body.tables.at(0);

    expect(table.rows.at(0).cells.at(0).cellFormat.verticalMerge).toEqual(aw.Tables.CellMerge.First);
    expect(table.rows.at(1).cells.at(0).cellFormat.verticalMerge).toEqual(aw.Tables.CellMerge.Previous);
    expect(table.rows.at(0).cells.at(0).getText().trim()).toEqual("Text in merged cells.\u0007");
    expect(table.rows.at(1).cells.at(0).getText()).not.toEqual(table.rows.at(0).cells.at(0).getText());
  });

  test('HorizontalMerge', () => {
    //ExStart
    //ExFor:CellMerge
    //ExFor:aw.Tables.CellFormat.horizontalMerge
    //ExSummary:Shows how to merge table cells horizontally.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a cell into the first column of the first row.
    // This cell will be the first in a range of horizontally merged cells.
    builder.insertCell();
    builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.First;
    builder.write("Text in merged cells.");

    // Insert a cell into the second column of the first row. Instead of adding text contents,
    // we will merge this cell with the first cell that we added directly to the left.
    builder.insertCell();
    builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.Previous;
    builder.endRow();

    // Insert two more unmerged cells to the second row.
    builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.None;
    builder.insertCell();
    builder.write("Text in unmerged cell.");
    builder.insertCell();
    builder.write("Text in unmerged cell.");
    builder.endRow();
    builder.endTable();

    doc.save(base.artifactsDir + "CellFormat.horizontalMerge.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "CellFormat.horizontalMerge.docx");
    let table = doc.firstSection.body.tables.at(0);

    expect(table.rows.at(0).cells.count).toEqual(1);
    expect(table.rows.at(0).cells.at(0).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.None);
    expect(table.rows.at(0).cells.at(0).getText().trim()).toEqual("Text in merged cells.\u0007");
  });

  test('Padding', () => {
    //ExStart
    //ExFor:aw.Tables.CellFormat.setPaddings
    //ExSummary:Shows how to pad the contents of a cell with whitespace.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set a padding distance (in points) between the border and the text contents
    // of each table cell we create with the document builder. 
    builder.cellFormat.setPaddings(5, 10, 40, 50);

    // Create a table with one cell whose contents will have whitespace padding.
    builder.startTable();
    builder.insertCell();
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

    doc.save(base.artifactsDir + "CellFormat.Padding.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "CellFormat.Padding.docx");

    let table = doc.firstSection.body.tables.at(0);
    let cell = table.rows.at(0).cells.at(0);

    expect(cell.cellFormat.leftPadding).toEqual(5);
    expect(cell.cellFormat.topPadding).toEqual(10);
    expect(cell.cellFormat.rightPadding).toEqual(40);
    expect(cell.cellFormat.bottomPadding).toEqual(50);
  });
});
