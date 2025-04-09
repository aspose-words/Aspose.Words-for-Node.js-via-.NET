// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');

/// <summary>
/// Represents a facade object for a column of a table in a Microsoft Word document.
/// </summary>
class Column {
  #table;
  #columnIndex;

  constructor(table, columnIndex) {
    this.#table = table;
    this.#columnIndex = columnIndex;
  }

  /// <summary>
  /// Returns a new column facade from the table and supplied zero-based index.
  /// </summary>
  static fromIndex(table, columnIndex) {
    return new Column(table, columnIndex);
  }

  /// <summary>
  /// Returns the cells which make up the column.
  /// </summary>
  get cells() {
    return this.getColumnCells();
  }

  /// <summary>
  /// Returns the index of the given cell in the column.
  /// </summary>
  indexOf(cell) {
    return this.getColumnCells().findIndex(c => c.referenceEquals(cell));
  }

  /// <summary>
  /// Inserts a new column before this column into the table.
  /// </summary>
  insertColumnBefore() {
    let columnCells = this.cells;
    if (columnCells.length == 0)
      throw new Error("Column must not be empty");

    // Create a clone of this column
    for (let cell of columnCells)
      cell.parentRow.insertBefore(cell.clone(false), cell);
    let newColumn = new Column(columnCells.at(0).parentRow.parentTable, this.#columnIndex);
    // We want to make sure that the cells are all valid to work with (have at least one paragraph).
    for (let cell of newColumn.cells)
      cell.ensureMinimum();
    // Increment the index of this column represents since there is a new column before it.
    this.#columnIndex++;
    return newColumn;
  }

  /// <summary>
  /// Removes the column from the table.
  /// </summary>
  remove() {
    for (let cell of this.cells)
      cell.remove();
  }

  /// <summary>
  /// Returns the text of the column. 
  /// </summary>
  toTxt() {
    let text = '';
    for (let cell of this.cells)
      text += cell.toString(aw.SaveFormat.Text);
    return text;
  }

  /// <summary>
  /// Provides an up-to-date collection of cells which make up the column represented by this facade.
  /// </summary>
  getColumnCells() {
    let columnCells = [];
    for (var row of this.#table.rows.toArray()) {
      let cell = row.cells.at(this.#columnIndex);
      if (cell != null)
        columnCells.push(cell);
    }
    return columnCells;
  }
}


describe("ExTableColumn", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('RemoveColumnFromTable', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.getTable(1, true);

    let column = Column.fromIndex(table, 2);
    column.remove();

    doc.save(base.artifactsDir + "TableColumn.RemoveColumn.doc");

    expect(table.getChildNodes(aw.NodeType.Cell, true).count).toEqual(16);
    expect(table.rows.at(2).cells.at(2).toString(aw.SaveFormat.Text).trim()).toEqual("Cell 7 contents");
    expect(table.lastRow.cells.at(2).toString(aw.SaveFormat.Text).trim()).toEqual("Cell 11 contents");
  });


  test('Insert', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.getTable(1, true);

    let column = Column.fromIndex(table, 1);

    // Create a new column to the left of this column.
    // This is the same as using the "Insert Column Before" command in Microsoft Word.
    let newColumn = column.insertColumnBefore();

    // Add some text to each cell in the column.
    for (let cell of newColumn.cells)
      cell.firstParagraph.appendChild(new aw.Run(doc, "Column Text " + newColumn.indexOf(cell)));

    doc.save(base.artifactsDir + "TableColumn.insert.doc");

    expect(table.getChildNodes(aw.NodeType.Cell, true).count).toEqual(24);
    expect(table.firstRow.cells.at(1).toString(aw.SaveFormat.Text).trim()).toEqual("Column Text 0");
    expect(table.lastRow.cells.at(1).toString(aw.SaveFormat.Text).trim()).toEqual("Column Text 3");
  });


  test('TableColumnToTxt', () => {
    let doc = new aw.Document(base.myDir + "Tables.docx");
    let table = doc.getTable(1, true);

    let column = Column.fromIndex(table, 0);
    console.log(column.toTxt());

    expect(column.toTxt()).toEqual("\rRow 1\rRow 2\rRow 3\r");
  });
});
