// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');
const DocumentHelper = require('./DocumentHelper');


describe("ExTable", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CreateTable', () => {
    //ExStart
    //ExFor:Table
    //ExFor:Row
    //ExFor:Cell
    //ExFor:Table.#ctor(DocumentBase)
    //ExSummary:Shows how to create a table.
    let doc = new aw.Document();
    let table = new aw.Tables.Table(doc);
    doc.firstSection.body.appendChild(table);

    // Tables contain rows, which contain cells, which may have paragraphs
    // with typical elements such as runs, shapes, and even other tables.
    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one row, cell, and paragraph.
    let firstRow = new aw.Tables.Row(doc);
    table.appendChild(firstRow);

    let firstCell = new aw.Tables.Cell(doc);
    firstRow.appendChild(firstCell);

    let paragraph = new aw.Paragraph(doc);
    firstCell.appendChild(paragraph);

    // Add text to the first cell in the first row of the table.
    let run = new aw.Run(doc, "Hello world!");
    paragraph.appendChild(run);

    doc.save(base.artifactsDir + "Table.CreateTable.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.CreateTable.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.rows.count).toEqual(1);
    expect(table.firstRow.cells.count).toEqual(1);
    expect(table.getText().trim()).toEqual("Hello world!\u0007\u0007");
  });


  test('Padding', () => {
    //ExStart
    //ExFor:Table.leftPadding
    //ExFor:Table.rightPadding
    //ExFor:Table.topPadding
    //ExFor:Table.bottomPadding
    //ExSummary:Shows how to configure content padding in a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Row 1, cell 1.");
    builder.insertCell();
    builder.write("Row 1, cell 2.");
    builder.endTable();

    // For every cell in the table, set the distance between its contents and each of its borders. 
    // This table will maintain the minimum padding distance by wrapping text.
    table.leftPadding = 30;
    table.rightPadding = 60;
    table.topPadding = 10;
    table.bottomPadding = 90;
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(250);

    doc.save(base.artifactsDir + "DocumentBuilder.SetRowFormatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "DocumentBuilder.SetRowFormatting.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.leftPadding).toEqual(30.0);
    expect(table.rightPadding).toEqual(60.0);
    expect(table.topPadding).toEqual(10.0);
    expect(table.bottomPadding).toEqual(90.0);
  });


  test('RowCellFormat', () => {
    //ExStart
    //ExFor:Row.rowFormat
    //ExFor:RowFormat
    //ExFor:Cell.cellFormat
    //ExFor:CellFormat
    //ExFor:CellFormat.shading
    //ExSummary:Shows how to modify the format of rows and cells in a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("City");
    builder.insertCell();
    builder.write("Country");
    builder.endRow();
    builder.insertCell();
    builder.write("London");
    builder.insertCell();
    builder.write("U.K.");
    builder.endTable();

    // Use the first row's "RowFormat" property to modify the formatting
    // of the contents of all cells in this row.
    let rowFormat = table.firstRow.rowFormat;
    rowFormat.height = 25;
    rowFormat.borders.at(aw.BorderType.Bottom).color = "#FF0000";

    // Use the "CellFormat" property of the first cell in the last row to modify the formatting of that cell's contents.
    let cellFormat = table.lastRow.firstCell.cellFormat;
    cellFormat.width = 100;
    cellFormat.shading.backgroundPatternColor = "#FFA500";

    doc.save(base.artifactsDir + "Table.RowCellFormat.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.RowCellFormat.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.getText().trim()).toEqual("City\u0007Country\u0007\u0007London\u0007U.K.\u0007\u0007");

    rowFormat = table.firstRow.rowFormat;

    expect(rowFormat.height).toEqual(25.0);
    expect(rowFormat.borders.at(aw.BorderType.Bottom).color).toEqual("#FF0000");

    cellFormat = table.lastRow.firstCell.cellFormat;

    expect(cellFormat.width).toEqual(110.8);
    expect(cellFormat.shading.backgroundPatternColor).toEqual("#FFA500");
  });


  test('DisplayContentOfTables', () => {
    //ExStart
    //ExFor:Cell
    //ExFor:CellCollection
    //ExFor:CellCollection.item(Int32)
    //ExFor:CellCollection.toArray
    //ExFor:Row
    //ExFor:Row.cells
    //ExFor:RowCollection
    //ExFor:RowCollection.item(Int32)
    //ExFor:RowCollection.toArray
    //ExFor:Table
    //ExFor:Table.rows
    //ExFor:TableCollection.item(Int32)
    //ExFor:TableCollection.toArray
    //ExSummary:Shows how to iterate through all tables in the document and print the contents of each cell.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let tables = doc.firstSection.body.tables;

    expect(tables.toArray().length).toEqual(2);

    for (let i = 0; i < tables.count; i++)
    {
      console.log(`Start of Table ${i}`);

      let rows = tables.at(i).rows;

      for (let j = 0; j < rows.count; j++)
      {
        console.log(`\tStart of Row ${j}`);

        let cells = rows.at(j).cells;

        for (let k = 0; k < cells.count; k++)
        {
          let cellText = cells.at(k).toString(aw.SaveFormat.Text).trim();
          console.log(`\t\tContents of Cell:${k} = \"${cellText}\"`);
        }

        console.log(`\tEnd of Row ${j}`);
      }

      console.log(`End of Table ${i}\n`);
    }
    //ExEnd
  });


  //ExStart
  //ExFor:Node.GetAncestor(NodeType)
  //ExFor:Node.GetAncestor(Type)
  //ExFor:Table.NodeType
  //ExFor:Cell.Tables
  //ExFor:TableCollection
  //ExFor:NodeCollection.Count
  //ExSummary:Shows how to find out if a tables are nested.
  test('CalculateDepthOfNestedTables', () => {
    let doc = new aw.Document(base.myDir + "Nested tables.docx");
    let tableNodes = doc.getChildNodes(aw.NodeType.Table, true);
    expect(tableNodes.count).toEqual(5);

    for (let i = 0; i < tableNodes.count; i++)
    {
      let table = tableNodes.at(i).asTable();

      // Find out if any cells in the table have other tables as children.
      let count = getChildTableCount(table);
      console.log("Table #{0} has {1} tables directly within its cells", i, count);

      // Find out if the table is nested inside another table, and, if so, at what depth.
      let tableDepth = getNestedDepthOfTable(table);

      if (tableDepth > 0)
        console.log("Table #{0} is nested inside another table at depth of {1}", i,
          tableDepth);
      else
        console.log("Table #{0} is a non nested table (is not a child of another table)", i);
    }
  });


  /// <summary>
  /// Calculates what level a table is nested inside other tables.
  /// </summary>
  /// <returns>
  /// An integer indicating the nesting depth of the table (number of parent table nodes).
  /// </returns>
  function getNestedDepthOfTable(table) {
    let depth = 0;
    let parent = table.getAncestor(aw.NodeType.Table);

    while (parent != null)
    {
      depth++;
      parent = parent.getAncestor(aw.NodeType.Table);
    }

    return depth;
  }

  
  /// <summary>
  /// Determines if a table contains any immediate child table within its cells.
  /// Do not recursively traverse through those tables to check for further tables.
  /// </summary>
  /// <returns>
  /// Returns true if at least one child cell contains a table.
  /// Returns false if no cells in the table contain a table.
  /// </returns>
  function getChildTableCount(table) {
    let childTableCount = 0;

    for (let row of table.rows.toArray())
    {
      for (let cell of row.cells.toArray())
      {
        if (cell.tables.count > 0)
          childTableCount++;
      }
    }

    return childTableCount;
  }
  //ExEnd

  
  test('EnsureTableMinimum', () => {
    //ExStart
    //ExFor:Table.ensureMinimum
    //ExSummary:Shows how to ensure that a table node contains the nodes we need to add content.
    let doc = new aw.Document();
    let table = new aw.Tables.Table(doc);
    doc.firstSection.body.appendChild(table);

    // Tables contain rows, which contain cells, which may contain paragraphs
    // with typical elements such as runs, shapes, and even other tables.
    // Our new table has none of these nodes, and we cannot add contents to it until it does.
    expect(table.getChildNodes(aw.NodeType.Any, true).count).toEqual(0);

    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one row and one cell with an empty paragraph.
    table.ensureMinimum();
    table.firstRow.firstCell.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));
    //ExEnd

    expect(table.getChildNodes(aw.NodeType.Any, true).count).toEqual(4);
  });


  test('EnsureRowMinimum', () => {
    //ExStart
    //ExFor:Row.ensureMinimum
    //ExSummary:Shows how to ensure a row node contains the nodes we need to begin adding content to it.
    let doc = new aw.Document();
    let table = new aw.Tables.Table(doc);
    doc.firstSection.body.appendChild(table);
    let row = new aw.Tables.Row(doc);
    table.appendChild(row);

    // Rows contain cells, containing paragraphs with typical elements such as runs, shapes, and even other tables.
    // Our new row has none of these nodes, and we cannot add contents to it until it does.
    expect(row.getChildNodes(aw.NodeType.Any, true).count).toEqual(0);

    // Calling the "EnsureMinimum" method on a table will ensure that
    // the table has at least one cell with an empty paragraph.
    row.ensureMinimum();
    row.firstCell.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));
    //ExEnd

    expect(row.getChildNodes(aw.NodeType.Any, true).count).toEqual(3);
  });


  test('EnsureCellMinimum', () => {
    //ExStart
    //ExFor:Cell.ensureMinimum
    //ExSummary:Shows how to ensure a cell node contains the nodes we need to begin adding content to it.
    let doc = new aw.Document();
    let table = new aw.Tables.Table(doc);
    doc.firstSection.body.appendChild(table);
    let row = new aw.Tables.Row(doc);
    table.appendChild(row);
    let cell = new aw.Tables.Cell(doc);
    row.appendChild(cell);

    // Cells may contain paragraphs with typical elements such as runs, shapes, and even other tables.
    // Our new cell does not have any paragraphs, and we cannot add contents such as run and shape nodes to it until it does.
    expect(cell.getChildNodes(aw.NodeType.Any, true).count).toEqual(0);

    // Calling the "EnsureMinimum" method on a cell will ensure that
    // the cell has at least one empty paragraph, which we can then add contents to.
    cell.ensureMinimum();
    cell.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));
    //ExEnd

    expect(cell.getChildNodes(aw.NodeType.Any, true).count).toEqual(2);
  });


  test('SetOutlineBorders', () => {
    //ExStart
    //ExFor:Table.alignment
    //ExFor:TableAlignment
    //ExFor:Table.clearBorders
    //ExFor:Table.clearShading
    //ExFor:Table.setBorder
    //ExFor:TextureIndex
    //ExFor:Table.setShading
    //ExSummary:Shows how to apply an outline border to a table.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Align the table to the center of the page.
    table.alignment = aw.Tables.TableAlignment.Center;

    // Clear any existing borders and shading from the table.
    table.clearBorders();
    table.clearShading();

    // Add green borders to the outline of the table.
    table.setBorder(aw.BorderType.Left, aw.LineStyle.Single, 1.5, "#008000", true);
    table.setBorder(aw.BorderType.Right, aw.LineStyle.Single, 1.5, "#008000", true);
    table.setBorder(aw.BorderType.Top, aw.LineStyle.Single, 1.5, "#008000", true);
    table.setBorder(aw.BorderType.Bottom, aw.LineStyle.Single, 1.5, "#008000", true);

    // Fill the cells with a light green solid color.
    table.setShading(aw.TextureIndex.TextureSolid, "#90EE90", "#000000");

    doc.save(base.artifactsDir + "Table.SetOutlineBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.SetOutlineBorders.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.alignment).toEqual(aw.Tables.TableAlignment.Center);

    let borders = table.firstRow.rowFormat.borders;

    expect(borders.top.color).toEqual("#008000");
    expect(borders.left.color).toEqual("#008000");
    expect(borders.right.color).toEqual("#008000");
    expect(borders.bottom.color).toEqual("#008000");
    expect(borders.horizontal.color).not.toEqual("#008000");
    expect(borders.vertical.color).not.toEqual("#008000");
    expect(table.firstRow.firstCell.cellFormat.shading.foregroundPatternColor).toEqual("#90EE90");
  });


  test('SetBorders', () => {
    //ExStart
    //ExFor:Table.setBorders
    //ExSummary:Shows how to format of all of a table's borders at once.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Clear all existing borders from the table.
    table.clearBorders();

    // Set a single green line to serve as every outer and inner border of this table.
    table.setBorders(aw.LineStyle.Single, 1.5, "#008000");

    doc.save(base.artifactsDir + "Table.setBorders.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.setBorders.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.firstRow.rowFormat.borders.top.color).toEqual("#008000");
    expect(table.firstRow.rowFormat.borders.left.color).toEqual("#008000");
    expect(table.firstRow.rowFormat.borders.right.color).toEqual("#008000");
    expect(table.firstRow.rowFormat.borders.bottom.color).toEqual("#008000");
    expect(table.firstRow.rowFormat.borders.horizontal.color).toEqual("#008000");
    expect(table.firstRow.rowFormat.borders.vertical.color).toEqual("#008000");
  });


  test('RowFormat', () => {
    //ExStart
    //ExFor:RowFormat
    //ExFor:Row.rowFormat
    //ExSummary:Shows how to modify formatting of a table row.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Use the first row's "RowFormat" property to set formatting that modifies that entire row's appearance.
    let firstRow = table.firstRow;
    firstRow.rowFormat.borders.lineStyle = aw.LineStyle.None;
    firstRow.rowFormat.heightRule = aw.HeightRule.Auto;
    firstRow.rowFormat.allowBreakAcrossPages = true;

    doc.save(base.artifactsDir + "Table.rowFormat.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.rowFormat.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.firstRow.rowFormat.borders.lineStyle).toEqual(aw.LineStyle.None);
    expect(table.firstRow.rowFormat.heightRule).toEqual(aw.HeightRule.Auto);
    expect(table.firstRow.rowFormat.allowBreakAcrossPages).toEqual(true);
  });


  test('CellFormat', () => {
    //ExStart
    //ExFor:CellFormat
    //ExFor:Cell.cellFormat
    //ExSummary:Shows how to modify formatting of a table cell.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.firstSection.body.tables.at(0);
    let firstCell = table.firstRow.firstCell;

    // Use a cell's "CellFormat" property to set formatting that modifies the appearance of that cell.
    firstCell.cellFormat.width = 30;
    firstCell.cellFormat.orientation = aw.TextOrientation.Downward;
    firstCell.cellFormat.shading.foregroundPatternColor = "#90EE90";

    doc.save(base.artifactsDir + "Table.cellFormat.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.cellFormat.docx");

    table = doc.firstSection.body.tables.at(0);
    expect(table.firstRow.firstCell.cellFormat.width).toEqual(30);
    expect(table.firstRow.firstCell.cellFormat.orientation).toEqual(aw.TextOrientation.Downward);
    expect(table.firstRow.firstCell.cellFormat.shading.foregroundPatternColor).toEqual("#90EE90");
  });


  test('DistanceBetweenTableAndText', () => {
    //ExStart
    //ExFor:Table.distanceBottom
    //ExFor:Table.distanceLeft
    //ExFor:Table.distanceRight
    //ExFor:Table.distanceTop
    //ExSummary:Shows how to set distance between table boundaries and text.
    let doc = new aw.Document(base.myDir + "Table wrapped by text.docx");

    let table = doc.firstSection.body.tables.at(0);
    expect(table.distanceTop).toEqual(25.9);
    expect(table.distanceBottom).toEqual(25.9);
    expect(table.distanceLeft).toEqual(17.3);
    expect(table.distanceRight).toEqual(17.3);

    // Set distance between table and surrounding text.
    table.distanceLeft = 24;
    table.distanceRight = 24;
    table.distanceTop = 3;
    table.distanceBottom = 3;

    doc.save(base.artifactsDir + "Table.DistanceBetweenTableAndText.docx");
    //ExEnd
  });


  test('Borders', () => {
    //ExStart
    //ExFor:Table.clearBorders
    //ExSummary:Shows how to remove all borders from a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Hello world!");
    builder.endTable();

    // Modify the color and thickness of the top border.
    let topBorder = table.firstRow.rowFormat.borders.at(aw.BorderType.Top);
    table.setBorder(aw.BorderType.Top, aw.LineStyle.Double, 1.5, "#FF0000", true);

    expect(topBorder.lineWidth).toEqual(1.5);
    expect(topBorder.color).toEqual("#FF0000");
    expect(topBorder.lineStyle).toEqual(aw.LineStyle.Double);

    // Clear the borders of all cells in the table, and then save the document.
    table.clearBorders();
    expect(topBorder.color).not.toEqual("");
    doc.save(base.artifactsDir + "Table.clearBorders.docx");

    // Verify the values of the table's properties after re-opening the document.
    doc = new aw.Document(base.artifactsDir + "Table.clearBorders.docx");
    table = doc.firstSection.body.tables.at(0);
    topBorder = table.firstRow.rowFormat.borders.at(aw.BorderType.Top);

    expect(topBorder.lineWidth).toEqual(0.0);
    expect(topBorder.color).toEqual(base.emptyColor);
    expect(topBorder.lineStyle).toEqual(aw.LineStyle.None);
    //ExEnd
  });


  test('ReplaceCellText', () => {
    //ExStart
    //ExFor:Range.replace(String, String, FindReplaceOptions)
    //ExSummary:Shows how to replace all instances of String of text in a table and cell.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Carrots");
    builder.insertCell();
    builder.write("50");
    builder.endRow();
    builder.insertCell();
    builder.write("Potatoes");
    builder.insertCell();
    builder.write("50");
    builder.endTable();

    let options = new aw.Replacing.FindReplaceOptions();
    options.matchCase = true;
    options.findWholeWordsOnly = true;

    // Perform a find-and-replace operation on an entire table.
    table.range.replace("Carrots", "Eggs", options);

    // Perform a find-and-replace operation on the last cell of the last row of the table.
    table.lastRow.lastCell.range.replace("50", "20", options);

    expect(table.getText().trim()).toEqual("Eggs\u000750\u0007\u0007" +
                            "Potatoes\u000720\u0007\u0007");
    //ExEnd
  });


  test.skip.each([true,
    false])('RemoveParagraphTextAndMark - TODO: Regex not supported yet', (isSmartParagraphBreakReplacement) => {
    //ExStart
    //ExFor:FindReplaceOptions.smartParagraphBreakReplacement
    //ExSummary:Shows how to remove paragraph from a table cell with a nested table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create table with paragraph and inner table in first cell.
    builder.startTable();
    builder.insertCell();
    builder.write("TEXT1");
    builder.startTable();
    builder.insertCell();
    builder.endTable();
    builder.endTable();
    builder.writeln();

    let options = new aw.Replacing.FindReplaceOptions();
    // When the following option is set to 'true', Aspose.words will remove paragraph's text
    // completely with its paragraph mark. Otherwise, Aspose.words will mimic Word and remove
    // only paragraph's text and leaves the paragraph mark intact (when a table follows the text).
    options.smartParagraphBreakReplacement = isSmartParagraphBreakReplacement;
    doc.range.replace(new Regex("TEXT1&p"), "", options);

    doc.save(base.artifactsDir + "Table.RemoveParagraphTextAndMark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.RemoveParagraphTextAndMark.docx");

    expect(doc.firstSection.body.tables.at(0).rows.at(0).cells[0].paragraphs.count).toEqual(isSmartParagraphBreakReplacement ? 1 : 2);
  });


  test('PrintTableRange', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.firstSection.body.tables.at(0);

    // The range text will include control characters such as "\a" for a cell.
    // You can call ToString on the desired node to retrieve the plain text content.

    // Print the plain text range of the table to the screen.
    console.log("Contents of the table: ");
    console.log(table.range.text);

    // Print the contents of the second row to the screen.
    console.log("\nContents of the row: ");
    console.log(table.rows.at(1).range.text);

    // Print the contents of the last cell in the table to the screen.
    console.log("\nContents of the cell: ");
    console.log(table.lastRow.lastCell.range.text);

    expect(table.rows.at(1).range.text).toEqual("\u0007Column 1\u0007Column 2\u0007Column 3\u0007Column 4\u0007\u0007");
    expect(table.lastRow.lastCell.range.text).toEqual("Cell 12 contents\u0007");
  });


  test('CloneTable', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.firstSection.body.tables.at(0);

    let tableClone = table.clone(true).asTable();

    // Insert the cloned table into the document after the original.
    table.parentNode.insertAfter(tableClone, table);

    // Insert an empty paragraph between the two tables.
    table.parentNode.insertAfter(new aw.Paragraph(doc), table);

    doc.save(base.artifactsDir + "Table.CloneTable.doc");

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(3);
    expect(tableClone.range.text).toEqual(table.range.text);

    for (var cellNode of tableClone.getChildNodes(aw.NodeType.Cell, true))
      cellNode.asCell().removeAllChildren();

    expect(tableClone.toString(aw.SaveFormat.Text).trim()).toEqual('');
  });


  test.each([false,
    true])('AllowBreakAcrossPages', (allowBreakAcrossPages) => {
    //ExStart
    //ExFor:RowFormat.allowBreakAcrossPages
    //ExSummary:Shows how to disable rows breaking across pages for every row in a table.
    let doc = new aw.Document(base.myDir + "Table spanning two pages.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Set the "AllowBreakAcrossPages" property to "false" to keep the row
    // in one piece if a table spans two pages, which break up along that row.
    // If the row is too big to fit in one page, Microsoft Word will push it down to the next page.
    // Set the "AllowBreakAcrossPages" property to "true" to allow the row to break up across two pages.
    for (let row of table.rows.toArray())
      row.rowFormat.allowBreakAcrossPages = allowBreakAcrossPages;

    doc.save(base.artifactsDir + "Table.allowBreakAcrossPages.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.allowBreakAcrossPages.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.rows.toArray().filter(r => r.rowFormat.allowBreakAcrossPages == allowBreakAcrossPages).length).toEqual(3);
  });


  test.each([false,
    true])('AllowAutoFitOnTable', (allowAutoFit) => {
    //ExStart
    //ExFor:Table.allowAutoFit
    //ExSummary:Shows how to enable/disable automatic table cell resizing.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPoints(100);
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
          "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    builder.insertCell();
    builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.auto;
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
          "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder.endRow();
    builder.endTable();

    // Set the "AllowAutoFit" property to "false" to get the table to maintain the dimensions
    // of all its rows and cells, and truncate contents if they get too large to fit.
    // Set the "AllowAutoFit" property to "true" to allow the table to change its cells' width and height
    // to accommodate their contents.
    table.allowAutoFit = allowAutoFit;

    doc.save(base.artifactsDir + "Table.AllowAutoFitOnTable.html");
    //ExEnd

    if (allowAutoFit)
    {
      TestUtil.fileContainsString(
        "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
        base.artifactsDir + "Table.AllowAutoFitOnTable.html");
      TestUtil.fileContainsString(
        "<td style=\"border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
        base.artifactsDir + "Table.AllowAutoFitOnTable.html");
    }
    else
    {
      TestUtil.fileContainsString(
        "<td style=\"width:89.2pt; border-right-style:solid; border-right-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-right:0.5pt single\">",
        base.artifactsDir + "Table.AllowAutoFitOnTable.html");
      TestUtil.fileContainsString(
        "<td style=\"width:7.2pt; border-left-style:solid; border-left-width:0.75pt; padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-left:0.5pt single\">",
        base.artifactsDir + "Table.AllowAutoFitOnTable.html");
    }
  });


  test('KeepTableTogether', () => {
    //ExStart
    //ExFor:ParagraphFormat.keepWithNext
    //ExFor:Row.isLastRow
    //ExFor:Paragraph.isEndOfCell
    //ExFor:Paragraph.isInCell
    //ExFor:Cell.parentRow
    //ExFor:Cell.paragraphs
    //ExSummary:Shows how to set a table to stay together on the same page.
    let doc = new aw.Document(base.myDir + "Table spanning two pages.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Enabling KeepWithNext for every paragraph in the table except for the
    // last ones in the last row will prevent the table from splitting across multiple pages.
    for (var cellNode of table.getChildNodes(aw.NodeType.Cell, true)) {
      var cell = cellNode.asCell();
      for (let para of cell.paragraphs.toArray())
      {
        expect(para.isInCell).toEqual(true);

        if (!(cell.parentRow.isLastRow && para.isEndOfCell))
          para.paragraphFormat.keepWithNext = true;
      }
    }

    doc.save(base.artifactsDir + "Table.KeepTableTogether.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.KeepTableTogether.docx");
    table = doc.firstSection.body.tables.at(0);

    for (var paraNode of table.getChildNodes(aw.NodeType.Paragraph, true)) {
      var para = paraNode.asParagraph();
      if (para.isEndOfCell && para.parentNode.asCell().parentRow.isLastRow)
        expect(para.paragraphFormat.keepWithNext).toEqual(false);
      else
        expect(para.paragraphFormat.keepWithNext).toEqual(true);
    }
  });


  test('GetIndexOfTableElements', () => {
    //ExStart
    //ExFor:NodeCollection.indexOf(Node)
    //ExSummary:Shows how to get the index of a node in a collection.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.firstSection.body.tables.at(0);
    let allTables = doc.getChildNodes(aw.NodeType.Table, true);

    expect(allTables.indexOf(table)).toEqual(0);

    let row = table.rows.at(2);

    expect(table.indexOf(row)).toEqual(2);

    let cell = row.lastCell;

    expect(row.indexOf(cell)).toEqual(4);
    //ExEnd
  });


  test('GetPreferredWidthTypeAndValue', () => {
    //ExStart
    //ExFor:PreferredWidthType
    //ExFor:PreferredWidth.type
    //ExFor:PreferredWidth.value
    //ExSummary:Shows how to verify the preferred width type and value of a table cell.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.firstSection.body.tables.at(0);
    let firstCell = table.firstRow.firstCell;

    expect(firstCell.cellFormat.preferredWidth.type).toEqual(aw.Tables.PreferredWidthType.Percent);
    expect(firstCell.cellFormat.preferredWidth.value).toEqual(11.16);
    //ExEnd
  });


  test.each([false,
    true])('AllowCellSpacing', (allowCellSpacing) => {
    //ExStart
    //ExFor:Table.allowCellSpacing
    //ExFor:Table.cellSpacing
    //ExSummary:Shows how to enable spacing between individual cells in a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Animal");
    builder.insertCell();
    builder.write("Class");
    builder.endRow();
    builder.insertCell();
    builder.write("Dog");
    builder.insertCell();
    builder.write("Mammal");
    builder.endTable();

    table.cellSpacing = 3;

    // Set the "AllowCellSpacing" property to "true" to enable spacing between cells
    // with a magnitude equal to the value of the "CellSpacing" property, in points.
    // Set the "AllowCellSpacing" property to "false" to disable cell spacing
    // and ignore the value of the "CellSpacing" property.
    table.allowCellSpacing = allowCellSpacing;

    doc.save(base.artifactsDir + "Table.allowCellSpacing.html");

    // Adjusting the "CellSpacing" property will automatically enable cell spacing.
    table.cellSpacing = 5;

    expect(table.allowCellSpacing).toEqual(true);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.allowCellSpacing.html");
    table = doc.getTable(0, true);

    expect(table.allowCellSpacing).toEqual(allowCellSpacing);

    if (allowCellSpacing)
      expect(table.cellSpacing).toEqual(3.0);
    else
      expect(table.cellSpacing).toEqual(0.0);

    TestUtil.fileContainsString(
      allowCellSpacing
        ? "<td style=\"border-style:solid; border-width:0.75pt; padding-right:5.4pt; padding-left:5.4pt; vertical-align:top; -aw-border:0.5pt single\">"
        : "<td style=\"border-right-style:solid; border-right-width:0.75pt; border-bottom-style:solid; border-bottom-width:0.75pt; " +
        "padding-right:5.03pt; padding-left:5.03pt; vertical-align:top; -aw-border-bottom:0.5pt single; -aw-border-right:0.5pt single\">",
      base.artifactsDir + "Table.allowCellSpacing.html");
  });


  //ExStart
  //ExFor:Table
  //ExFor:Row
  //ExFor:Cell
  //ExFor:Table.#ctor(DocumentBase)
  //ExFor:Table.Title
  //ExFor:Table.Description
  //ExFor:Row.#ctor(DocumentBase)
  //ExFor:Cell.#ctor(DocumentBase)
  //ExFor:Cell.FirstParagraph
  //ExSummary:Shows how to build a nested table without using a document builder.
  test('CreateNestedTable', () => {
    let doc = new aw.Document();

    // Create the outer table with three rows and four columns, and then add it to the document.
    let outerTable = createTable(doc, 3, 4, "Outer Table");
    doc.firstSection.body.appendChild(outerTable);

    // Create another table with two rows and two columns and then insert it into the first table's first cell.
    let innerTable = createTable(doc, 2, 2, "Inner Table");
    outerTable.firstRow.firstCell.appendChild(innerTable);

    doc.save(base.artifactsDir + "Table.CreateNestedTable.docx");
    testCreateNestedTable(new aw.Document(base.artifactsDir + "Table.CreateNestedTable.docx")); //ExSkip
  });


  /// <summary>
  /// Creates a new table in the document with the given dimensions and text in each cell.
  /// </summary>
  function createTable(doc, rowCount, cellCount, cellText)
  {
    let table = new aw.Tables.Table(doc);

    for (let rowId = 1; rowId <= rowCount; rowId++)
    {
      let row = new aw.Tables.Row(doc);
      table.appendChild(row);

      for (let cellId = 1; cellId <= cellCount; cellId++)
      {
        let cell = new aw.Tables.Cell(doc);
        cell.appendChild(new aw.Paragraph(doc));
        cell.firstParagraph.appendChild(new aw.Run(doc, cellText));

        row.appendChild(cell);
      }
    }

    // You can use the "Title" and "Description" properties to add a title and description respectively to your table.
    // The table must have at least one row before we can use these properties.
    // These properties are meaningful for ISO / IEC 29500 compliant .docx documents (see the OoxmlCompliance class).
    // If we save the document to pre-ISO/IEC 29500 formats, Microsoft Word ignores these properties.
    table.title = "Aspose table title";
    table.description = "Aspose table description";

    return table;
  }
  //ExEnd

  function testCreateNestedTable(doc)
  {
    let outerTable = doc.firstSection.body.tables.at(0);
    let innerTable = doc.getTable(1, true);

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(2);
    expect(outerTable.firstRow.firstCell.tables.count).toEqual(1);
    expect(outerTable.getChildNodes(aw.NodeType.Cell, true).count).toEqual(16);
    expect(innerTable.getChildNodes(aw.NodeType.Cell, true).count).toEqual(4);
    expect(innerTable.title).toEqual("Aspose table title");
    expect(innerTable.description).toEqual("Aspose table description");
  }

  //ExStart
  //ExFor:CellFormat.HorizontalMerge
  //ExFor:CellFormat.VerticalMerge
  //ExFor:CellMerge
  //ExSummary:Prints the horizontal and vertical merge type of a cell.
  test('CheckCellsMerged', () => {
    let doc = new aw.Document(base.myDir + "Table with merged cells.docx");
    let table = doc.firstSection.body.tables.at(0);

    for (let row of table.rows.toArray())
      for (let cell of row.cells.toArray())
        console.log(printCellMergeType(cell));
    expect(printCellMergeType(table.firstRow.firstCell)).toEqual("The cell at R1, C1 is vertically merged");
  });


  function printCellMergeType(cell)
  {
    let isHorizontallyMerged = cell.cellFormat.horizontalMerge != aw.Tables.CellMerge.None;
    let isVerticallyMerged = cell.cellFormat.verticalMerge != aw.Tables.CellMerge.None;
    let cellLocation =
      `R${cell.parentRow.parentTable.indexOf(cell.parentRow) + 1}, C${cell.parentRow.indexOf(cell) + 1}`;

    if (isHorizontallyMerged && isVerticallyMerged)
      return `The cell at ${cellLocation} is both horizontally and vertically merged`;
    if (isHorizontallyMerged)
      return `The cell at ${cellLocation} is horizontally merged.`;

    return isVerticallyMerged ? `The cell at ${cellLocation} is vertically merged` : `The cell at ${cellLocation} is not merged`;
  }
  //ExEnd


  test('MergeCellRange', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let table = doc.firstSection.body.tables.at(0);

    // We want to merge the range of cells found in between these two cells.
    let cellStartRange = table.rows.at(2).cells.at(2);
    let cellEndRange = table.rows.at(3).cells.at(3);

    // Merge all the cells between the two specified cells into one.
    mergeCells(cellStartRange, cellEndRange);

    doc.save(base.artifactsDir + "Table.MergeCellRange.doc");

    let mergedCellsCount = 0;
    for (var node of table.getChildNodes(aw.NodeType.Cell, true))
    {
      let cell = node.asCell();
      if (cell.cellFormat.horizontalMerge != aw.Tables.CellMerge.None ||
        cell.cellFormat.verticalMerge != aw.Tables.CellMerge.None)
        mergedCellsCount++;
    }

    expect(mergedCellsCount).toEqual(4);
    expect(table.rows.at(2).cells.at(2).cellFormat.horizontalMerge == aw.Tables.CellMerge.First).toEqual(true);
    expect(table.rows.at(2).cells.at(2).cellFormat.verticalMerge == aw.Tables.CellMerge.First).toEqual(true);
    expect(table.rows.at(3).cells.at(3).cellFormat.horizontalMerge == aw.Tables.CellMerge.Previous).toEqual(true);
    expect(table.rows.at(3).cells.at(3).cellFormat.verticalMerge == aw.Tables.CellMerge.Previous).toEqual(true);
  });


  /// <summary>
  /// Merges the range of cells found between the two specified cells both horizontally and vertically.
  /// Can span over multiple rows.
  /// </summary>
  function mergeCells(startCell, endCell)
  {
    let parentTable = startCell.parentRow.parentTable;

    // Find the row and cell indices for the start and end cells.
    let startCellPos = new aw.JSPoint(startCell.parentRow.indexOf(startCell),
      parentTable.indexOf(startCell.parentRow));
    let endCellPos = new aw.JSPoint(endCell.parentRow.indexOf(endCell), parentTable.indexOf(endCell.parentRow));

    // Create a range of cells to be merged based on these indices.
    // Inverse each index if the end cell is before the start cell.
    let mergeRange = new aw.JSRectangle(
      Math.min(startCellPos.X, endCellPos.X),
      Math.min(startCellPos.Y, endCellPos.Y),
      Math.abs(endCellPos.X - startCellPos.X) + 1,
      Math.abs(endCellPos.Y - startCellPos.Y) + 1);

    for (let row of parentTable.rows.toArray())
    {
      for (let cell of row.cells.toArray())
      {
        let currentPos = new aw.JSPoint(row.indexOf(cell), parentTable.indexOf(row));

          // Check if the current cell is inside our merge range, then merge it.
        if (mergeRange.contains(currentPos))
        {
          cell.cellFormat.horizontalMerge =
            currentPos.X == mergeRange.X ? aw.Tables.CellMerge.First : aw.Tables.CellMerge.Previous;
          cell.cellFormat.verticalMerge =
            currentPos.Y == mergeRange.Y ? aw.Tables.CellMerge.First : aw.Tables.CellMerge.Previous;
        }
      }
    }
  }

  
  test('CombineTables', () => {
    //ExStart
    //ExFor:Cell.cellFormat
    //ExFor:CellFormat.borders
    //ExFor:Table.rows
    //ExFor:Table.firstRow
    //ExFor:CellFormat.clearFormatting
    //ExFor:CompositeNode.hasChildNodes
    //ExSummary:Shows how to combine the rows from two tables into one.
    let doc = new aw.Document(base.myDir + "Tables.docx");

    // Below are two ways of getting a table from a document.
    // 1 -  From the "Tables" collection of a Body node:
    let firstTable = doc.firstSection.body.tables.at(0);

    // 2 -  Using the "GetChild" method:
    let secondTable = doc.getTable(1, true);

    // Append all rows from the current table to the next.
    while (secondTable.hasChildNodes)
      firstTable.rows.add(secondTable.firstRow);

    // Remove the empty table container.
    secondTable.remove();

    doc.save(base.artifactsDir + "Table.CombineTables.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.CombineTables.docx");

    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(1);
    expect(doc.firstSection.body.tables.at(0).rows.count).toEqual(9);
    expect(doc.firstSection.body.tables.at(0).getChildNodes(aw.NodeType.Cell, true).count).toEqual(42);
  });


  test.skip('SplitTable - TODO: Failed.', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");

    let firstTable = doc.firstSection.body.tables.at(0);

    // We will split the table at the third row (inclusive).
    let row = firstTable.rows.at(2);

    // Create a new container for the split table.
    let table = firstTable.clone(false).asTable();

    // Insert the container after the original.
    firstTable.parentNode.insertAfter(table, firstTable);

    // Add a buffer paragraph to ensure the tables stay apart.
    firstTable.parentNode.insertAfter(new aw.Paragraph(doc), firstTable);

    var currentRow;
    do
    {
      currentRow = firstTable.lastRow;
      table.prependChild(currentRow);
    } while (currentRow != row);

    doc = DocumentHelper.saveOpen(doc);

    expect(table.firstRow).toEqual(row);
    expect(firstTable.rows.count).toEqual(2);
    expect(table.rows.count).toEqual(3);
    expect(doc.getChildNodes(aw.NodeType.Table, true).count).toEqual(3);
  });


  test('WrapText', () => {
    //ExStart
    //ExFor:Table.textWrapping
    //ExFor:TextWrapping
    //ExSummary:Shows how to work with table text wrapping.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Cell 1");
    builder.insertCell();
    builder.write("Cell 2");
    builder.endTable();
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(300);

    builder.font.size = 16;
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

    // Set the "TextWrapping" property to "TextWrapping.Around" to get the table to wrap text around it,
    // and push it down into the paragraph below by setting the position.
    table.textWrapping = aw.Tables.TextWrapping.Around;
    table.absoluteHorizontalDistance = 100;
    table.absoluteVerticalDistance = 20;

    doc.save(base.artifactsDir + "Table.wrapText.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.wrapText.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.textWrapping).toEqual(aw.Tables.TextWrapping.Around);
    expect(table.absoluteHorizontalDistance).toEqual(100.0);
    expect(table.absoluteVerticalDistance).toEqual(20.0);
  });


  test('GetFloatingTableProperties', () => {
    //ExStart
    //ExFor:Table.horizontalAnchor
    //ExFor:Table.verticalAnchor
    //ExFor:Table.allowOverlap
    //ExFor:ShapeBase.allowOverlap
    //ExSummary:Shows how to work with floating tables properties.
    let doc = new aw.Document(base.myDir + "Table wrapped by text.docx");

    let table = doc.firstSection.body.tables.at(0);

    if (table.textWrapping == aw.Tables.TextWrapping.Around)
    {
      expect(table.horizontalAnchor).toEqual(aw.Drawing.RelativeHorizontalPosition.Margin);
      expect(table.verticalAnchor).toEqual(aw.Drawing.RelativeVerticalPosition.Paragraph);
      expect(table.allowOverlap).toEqual(false);

      // Only Margin, Page, Column available in RelativeHorizontalPosition for HorizontalAnchor setter.
      // The ArgumentException will be thrown for any other values.
      table.horizontalAnchor = aw.Drawing.RelativeHorizontalPosition.Column;

      // Only Margin, Page, Paragraph available in RelativeVerticalPosition for VerticalAnchor setter.
      // The ArgumentException will be thrown for any other values.
      table.verticalAnchor = aw.Drawing.RelativeVerticalPosition.Page;
    }
    //ExEnd
  });


  test('ChangeFloatingTableProperties', () => {
    //ExStart
    //ExFor:Table.relativeHorizontalAlignment
    //ExFor:Table.relativeVerticalAlignment
    //ExFor:Table.absoluteHorizontalDistance
    //ExFor:Table.absoluteVerticalDistance
    //ExSummary:Shows how set the location of floating tables.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Table 1, cell 1");
    builder.endTable();
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(300);

    // Set the table's location to a place on the page, such as, in this case, the bottom right corner.
    table.relativeVerticalAlignment = aw.Drawing.VerticalAlignment.Bottom;
    table.relativeHorizontalAlignment = aw.Drawing.HorizontalAlignment.Right;

    table = builder.startTable();
    builder.insertCell();
    builder.write("Table 2, cell 1");
    builder.endTable();
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(300);

    // We can also set a horizontal and vertical offset in points from the paragraph's location where we inserted the table. 
    table.absoluteVerticalDistance = 50;
    table.absoluteHorizontalDistance = 100;

    doc.save(base.artifactsDir + "Table.ChangeFloatingTableProperties.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.ChangeFloatingTableProperties.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.relativeVerticalAlignment).toEqual(aw.Drawing.VerticalAlignment.Bottom);
    expect(table.relativeHorizontalAlignment).toEqual(aw.Drawing.HorizontalAlignment.Right);

    table = doc.getTable(1, true);

    expect(table.absoluteVerticalDistance).toEqual(50.0);
    expect(table.absoluteHorizontalDistance).toEqual(100.0);
  });


  test('TableStyleCreation', () => {
    //ExStart
    //ExFor:Table.bidi
    //ExFor:Table.cellSpacing
    //ExFor:Table.style
    //ExFor:Table.styleName
    //ExFor:TableStyle
    //ExFor:TableStyle.allowBreakAcrossPages
    //ExFor:TableStyle.bidi
    //ExFor:TableStyle.cellSpacing
    //ExFor:TableStyle.bottomPadding
    //ExFor:TableStyle.leftPadding
    //ExFor:TableStyle.rightPadding
    //ExFor:TableStyle.topPadding
    //ExFor:TableStyle.shading
    //ExFor:TableStyle.borders
    //ExFor:TableStyle.verticalAlignment
    //ExSummary:Shows how to create custom style settings for the table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
 
    let table = builder.startTable();
    builder.insertCell();
    builder.write("Name");
    builder.insertCell();
    builder.write("مرحبًا");
    builder.endRow();
    builder.insertCell();
    builder.insertCell();
    builder.endTable();
 
    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
    tableStyle.allowBreakAcrossPages = true;
    tableStyle.bidi = true;
    tableStyle.cellSpacing = 5;
    tableStyle.bottomPadding = 20;
    tableStyle.leftPadding = 5;
    tableStyle.rightPadding = 10;
    tableStyle.topPadding = 20;
    tableStyle.shading.backgroundPatternColor = "#FAEBD7";
    tableStyle.borders.color = "#0000FF";
    tableStyle.borders.lineStyle = aw.LineStyle.DotDash;
    tableStyle.verticalAlignment = aw.Tables.CellVerticalAlignment.Center;

    table.style = tableStyle;

    // Setting the style properties of a table may affect the properties of the table itself.
    expect(table.bidi).toEqual(true);
    expect(table.cellSpacing).toEqual(5.0);
    expect(table.styleName).toEqual("MyTableStyle1");

    doc.save(base.artifactsDir + "Table.TableStyleCreation.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.TableStyleCreation.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.bidi).toEqual(true);
    expect(table.cellSpacing).toEqual(5.0);
    expect(table.styleName).toEqual("MyTableStyle1");
    expect(tableStyle.bottomPadding).toEqual(20.0);
    expect(tableStyle.leftPadding).toEqual(5.0);
    expect(tableStyle.rightPadding).toEqual(10.0);
    expect(tableStyle.topPadding).toEqual(20.0);
    expect([...table.firstRow.rowFormat.borders].filter(b => b.color == "#0000FF").length).toEqual(6);
    expect(tableStyle.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);

    tableStyle = doc.styles.at("MyTableStyle1").asTableStyle();

    expect(tableStyle.allowBreakAcrossPages).toEqual(true);
    expect(tableStyle.bidi).toEqual(true);
    expect(tableStyle.cellSpacing).toEqual(5.0);
    expect(tableStyle.bottomPadding).toEqual(20.0);
    expect(tableStyle.leftPadding).toEqual(5.0);
    expect(tableStyle.rightPadding).toEqual(10.0);
    expect(tableStyle.topPadding).toEqual(20.0);
    expect(tableStyle.shading.backgroundPatternColor).toEqual("#FAEBD7");
    expect(tableStyle.borders.color).toEqual("#0000FF");
    expect(tableStyle.borders.lineStyle).toEqual(aw.LineStyle.DotDash);
    expect(tableStyle.verticalAlignment).toEqual(aw.Tables.CellVerticalAlignment.Center);
  });


  test('SetTableAlignment', () => {
    //ExStart
    //ExFor:TableStyle.alignment
    //ExFor:TableStyle.leftIndent
    //ExSummary:Shows how to set the position of a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways of aligning a table horizontally.
    // 1 -  Use the "Alignment" property to align it to a location on the page, such as the center:
    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
    tableStyle.alignment = aw.Tables.TableAlignment.Center;
    tableStyle.borders.color = "#0000FF";
    tableStyle.borders.lineStyle = aw.LineStyle.Single;

    // Insert a table and apply the style we created to it.
    let table = builder.startTable();
    builder.insertCell();
    builder.write("Aligned to the center of the page");
    builder.endTable();
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(300);
            
    table.style = tableStyle;

    // 2 -  Use the "LeftIndent" to specify an indent from the left margin of the page:
    tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle2").asTableStyle();
    tableStyle.leftIndent = 55;
    tableStyle.borders.color = "#008000";
    tableStyle.borders.lineStyle = aw.LineStyle.Single;

    table = builder.startTable();
    builder.insertCell();
    builder.write("Aligned according to left indent");
    builder.endTable();
    table.preferredWidth = aw.Tables.PreferredWidth.fromPoints(300);

    table.style = tableStyle;

    doc.save(base.artifactsDir + "Table.SetTableAlignment.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.SetTableAlignment.docx");

    tableStyle = doc.styles.at("MyTableStyle1").asTableStyle();

    expect(tableStyle.alignment).toEqual(aw.Tables.TableAlignment.Center);
    expect(doc.firstSection.body.tables.at(0).style).toEqual(tableStyle);

    tableStyle = doc.styles.at("MyTableStyle2").asTableStyle();

    expect(tableStyle.leftIndent).toEqual(55.0);
    expect((doc.getTable( 1, true)).style).toEqual(tableStyle);
  });


  test.skip('ConditionalStyles - TODO: WORDSNODEJS-84.', () => {
    //ExStart
    //ExFor:ConditionalStyle
    //ExFor:ConditionalStyle.shading
    //ExFor:ConditionalStyle.borders
    //ExFor:ConditionalStyle.paragraphFormat
    //ExFor:ConditionalStyle.bottomPadding
    //ExFor:ConditionalStyle.leftPadding
    //ExFor:ConditionalStyle.rightPadding
    //ExFor:ConditionalStyle.topPadding
    //ExFor:ConditionalStyle.font
    //ExFor:ConditionalStyle.type
    //ExFor:ConditionalStyleCollection.getEnumerator
    //ExFor:ConditionalStyleCollection.firstRow
    //ExFor:ConditionalStyleCollection.lastRow
    //ExFor:ConditionalStyleCollection.lastColumn
    //ExFor:ConditionalStyleCollection.count
    //ExFor:ConditionalStyleCollection
    //ExFor:ConditionalStyleCollection.bottomLeftCell
    //ExFor:ConditionalStyleCollection.bottomRightCell
    //ExFor:ConditionalStyleCollection.evenColumnBanding
    //ExFor:ConditionalStyleCollection.evenRowBanding
    //ExFor:ConditionalStyleCollection.firstColumn
    //ExFor:ConditionalStyleCollection.item(ConditionalStyleType)
    //ExFor:ConditionalStyleCollection.item(Int32)
    //ExFor:ConditionalStyleCollection.oddColumnBanding
    //ExFor:ConditionalStyleCollection.oddRowBanding
    //ExFor:ConditionalStyleCollection.topLeftCell
    //ExFor:ConditionalStyleCollection.topRightCell
    //ExFor:ConditionalStyleType
    //ExFor:TableStyle.conditionalStyles
    //ExSummary:Shows how to work with certain area styles of a table.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("Cell 1");
    builder.insertCell();
    builder.write("Cell 2");
    builder.endRow();
    builder.insertCell();
    builder.write("Cell 3");
    builder.insertCell();
    builder.write("Cell 4");
    builder.endTable();

    // Create a custom table style.
    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();

    // Conditional styles are formatting changes that affect only some of the table's cells
    // based on a predicate, such as the cells being in the last row.
    // Below are three ways of accessing a table style's conditional styles from the "ConditionalStyles" collection.
    // 1 -  By style type:
    tableStyle.conditionalStyles.at(aw.ConditionalStyleType.FirstRow).shading.backgroundPatternColor = "#F0F8FF";

    // 2 -  By index:
    tableStyle.conditionalStyles.at(0).borders.color = "#000000";
    tableStyle.conditionalStyles.at(0).borders.lineStyle = aw.LineStyle.DotDash;
    expect(tableStyle.conditionalStyles.at(0).type).toEqual(aw.ConditionalStyleType.FirstRow);

    // 3 -  As a property:
    tableStyle.conditionalStyles.firstRow.paragraphFormat.alignment = aw.ParagraphAlignment.Center;

    // Apply padding and text formatting to conditional styles.
    tableStyle.conditionalStyles.lastRow.bottomPadding = 10;
    tableStyle.conditionalStyles.lastRow.leftPadding = 10;
    tableStyle.conditionalStyles.lastRow.rightPadding = 10;
    tableStyle.conditionalStyles.lastRow.topPadding = 10;
    tableStyle.conditionalStyles.lastColumn.font.bold = true;

    // List all possible style conditions.
    for (var currentStyle of tableStyle.conditionalStyles)
    {
        if (currentStyle != null) console.log(currentStyle.type);
    }

    // Apply the custom style, which contains all conditional styles, to the table.
    table.style = tableStyle;

    // Our style applies some conditional styles by default.
    expect(table.styleOptions).toEqual(aw.Tables.TableStyleOptions.FirstRow | aw.Tables.TableStyleOptions.FirstColumn | aw.Tables.TableStyleOptions.RowBands);

    // We will need to enable all other styles ourselves via the "StyleOptions" property.
    table.styleOptions = table.styleOptions | aw.Tables.TableStyleOptions.LastRow | aw.Tables.TableStyleOptions.LastColumn;

    doc.save(base.artifactsDir + "Table.conditionalStyles.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.conditionalStyles.docx");
    table = doc.firstSection.body.tables.at(0);

    expect(table.styleOptions).toEqual(aw.Tables.TableStyleOptions.Default | aw.Tables.TableStyleOptions.LastRow | aw.Tables.TableStyleOptions.LastColumn);
    let conditionalStyles = doc.styles.at("MyTableStyle1").asTableStyle().conditionalStyles;

    expect(conditionalStyles.at(0).type).toEqual(aw.ConditionalStyleType.FirstRow);
    expect(conditionalStyles.at(0).shading.backgroundPatternColor).toEqual("#F0F8FF");
    expect(conditionalStyles.at(0).borders.color).toEqual("#000000");
    expect(conditionalStyles.at(0).borders.lineStyle).toEqual(aw.LineStyle.DotDash);
    expect(conditionalStyles.at(0).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Center);

    expect(conditionalStyles.at(2).type).toEqual(aw.ConditionalStyleType.LastRow);
    expect(conditionalStyles.at(2).bottomPadding).toEqual(10.0);
    expect(conditionalStyles.at(2).leftPadding).toEqual(10.0);
    expect(conditionalStyles.at(2).rightPadding).toEqual(10.0);
    expect(conditionalStyles.at(2).topPadding).toEqual(10.0);

    expect(conditionalStyles.at(3).type).toEqual(aw.ConditionalStyleType.LastColumn);
    expect(conditionalStyles.at(3).font.bold).toEqual(true);
  });


  test('ClearTableStyleFormatting', () => {
    //ExStart
    //ExFor:ConditionalStyle.clearFormatting
    //ExFor:ConditionalStyleCollection.clearFormatting
    //ExSummary:Shows how to reset conditional table styles.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let table = builder.startTable();
    builder.insertCell();
    builder.write("First row");
    builder.endRow();
    builder.insertCell();
    builder.write("Last row");
    builder.endTable();

    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
    table.style = tableStyle;

    // Set the table style to color the borders of the first row of the table in red.
    tableStyle.conditionalStyles.firstRow.borders.color = "#FF0000";

    // Set the table style to color the borders of the last row of the table in blue.
    tableStyle.conditionalStyles.lastRow.borders.color = "#0000FF";

    // Below are two ways of using the "ClearFormatting" method to clear the conditional styles.
    // 1 -  Clear the conditional styles for a specific part of a table:
    tableStyle.conditionalStyles.at(0).clearFormatting();

    expect(tableStyle.conditionalStyles.firstRow.borders.color).toEqual(base.emptyColor);

    // 2 -  Clear the conditional styles for the entire table:
    tableStyle.conditionalStyles.clearFormatting();

    expect([...tableStyle.conditionalStyles].every(s => s.borders.color == "")).toEqual(true);
    //ExEnd
  });


  test('AlternatingRowStyles', () => {
    //ExStart
    //ExFor:TableStyle.columnStripe
    //ExFor:TableStyle.rowStripe
    //ExSummary:Shows how to create conditional table styles that alternate between rows.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // We can configure a conditional style of a table to apply a different color to the row/column,
    // based on whether the row/column is even or odd, creating an alternating color pattern.
    // We can also apply a number n to the row/column banding,
    // meaning that the color alternates after every n rows/columns instead of one.
    // Create a table where single columns and rows will band the columns will banded in threes.
    let table = builder.startTable();
    for (let i = 0; i < 15; i++)
    {
      for (let j = 0; j < 4; j++)
      {
        builder.insertCell();
        builder.writeln(`${(j % 2 == 0 ? "Even" : "Odd")} column.`);
        builder.write(`Row banding ${(i % 3 == 0 ? "start" : "continuation")}.`);
      }
      builder.endRow();
    }
    builder.endTable();

    // Apply a line style to all the borders of the table.
    let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
    tableStyle.borders.color = "#000000";
    tableStyle.borders.lineStyle = aw.LineStyle.Double;

    // Set the two colors, which will alternate over every 3 rows.
    tableStyle.rowStripe = 3;
    tableStyle.conditionalStyles.at(aw.ConditionalStyleType.OddRowBanding).shading.backgroundPatternColor = "#ADD8E6";
    tableStyle.conditionalStyles.at(aw.ConditionalStyleType.EvenRowBanding).shading.backgroundPatternColor = "#E0FFFF";

    // Set a color to apply to every even column, which will override any custom row coloring.
    tableStyle.columnStripe = 1;
    tableStyle.conditionalStyles.at(aw.ConditionalStyleType.EvenColumnBanding).shading.backgroundPatternColor = "#FFA07A";

    table.style = tableStyle;

    // The "StyleOptions" property enables row banding by default.
    expect(table.styleOptions).toEqual(aw.Tables.TableStyleOptions.FirstRow | aw.Tables.TableStyleOptions.FirstColumn | aw.Tables.TableStyleOptions.RowBands);

    // Use the "StyleOptions" property also to enable column banding.
    table.styleOptions = table.styleOptions | aw.Tables.TableStyleOptions.ColumnBands;

    doc.save(base.artifactsDir + "Table.AlternatingRowStyles.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Table.AlternatingRowStyles.docx");
    table = doc.firstSection.body.tables.at(0);
    tableStyle = doc.styles.at("MyTableStyle1").asTableStyle();

    expect(table.style).toEqual(tableStyle);
    expect(table.styleOptions).toEqual(table.styleOptions | aw.Tables.TableStyleOptions.ColumnBands);

    expect(tableStyle.borders.color).toEqual("#000000");
    expect(tableStyle.borders.lineStyle).toEqual(aw.LineStyle.Double);
    expect(tableStyle.rowStripe).toEqual(3);
    expect(tableStyle.conditionalStyles.at(aw.ConditionalStyleType.OddRowBanding).shading.backgroundPatternColor).toEqual("#ADD8E6");
    expect(tableStyle.conditionalStyles.at(aw.ConditionalStyleType.EvenRowBanding).shading.backgroundPatternColor).toEqual("#E0FFFF");
    expect(tableStyle.columnStripe).toEqual(1);
    expect(tableStyle.conditionalStyles.at(aw.ConditionalStyleType.EvenColumnBanding).shading.backgroundPatternColor).toEqual("#FFA07A");
  });


  test('ConvertToHorizontallyMergedCells', () => {
    //ExStart
    //ExFor:Table.convertToHorizontallyMergedCells
    //ExSummary:Shows how to convert cells horizontally merged by width to cells merged by CellFormat.horizontalMerge.
    let doc = new aw.Document(base.myDir + "Table with merged cells.docx");

    // Microsoft Word does not write merge flags anymore, defining merged cells by width instead.
    // Aspose.words by default define only 5 cells in a row, and none of them have the horizontal merge flag,
    // even though there were 7 cells in the row before the horizontal merging took place.
    let table = doc.firstSection.body.tables.at(0);
    let row = table.rows.at(0);

    expect(row.cells.count).toEqual(5);
    expect(row.cells.toArray().every(c => c.cellFormat.horizontalMerge == aw.Tables.CellMerge.None)).toEqual(true);

    // Use the "ConvertToHorizontallyMergedCells" method to convert cells horizontally merged
    // by its width to the cell horizontally merged by flags.
    // Now, we have 7 cells, and some of them have horizontal merge values.
    table.convertToHorizontallyMergedCells();
    row = table.rows.at(0);

    expect(row.cells.count).toEqual(7);

    expect(row.cells.at(0).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.None);
    expect(row.cells.at(1).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.First);
    expect(row.cells.at(2).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.Previous);
    expect(row.cells.at(3).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.None);
    expect(row.cells.at(4).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.First);
    expect(row.cells.at(5).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.Previous);
    expect(row.cells.at(6).cellFormat.horizontalMerge).toEqual(aw.Tables.CellMerge.None);
    //ExEnd
  });


  test('GetTextFromCells', () => {
    //ExStart
    //ExFor:Row.nextRow
    //ExFor:Row.previousRow
    //ExFor:Cell.nextCell
    //ExFor:Cell.previousCell
    //ExSummary:Shows how to enumerate through all table cells.
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.firstSection.body.tables.at(0);

    // Enumerate through all cells of the table.
    for (let row = table.firstRow; row != null; row = row.nextRow)
    {
      for (let cell = row.firstCell; cell != null; cell = cell.nextCell)
      {
        console.log(cell.getText());
      }
    }
    //ExEnd
  });

});
