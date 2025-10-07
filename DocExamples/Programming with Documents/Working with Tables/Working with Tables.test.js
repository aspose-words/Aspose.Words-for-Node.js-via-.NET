// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithTables", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('RemoveColumn', () => {
        //ExStart:RemoveColumn
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 1, true).asTable();

        let column = Column.fromIndex(table, 2);
        column.remove();
        //ExEnd:RemoveColumn
    });

    test('InsertBlankColumn', () => {
        //ExStart:InsertBlankColumn
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        //ExStart:GetPlainText
        let column = Column.fromIndex(table, 0);
        // Print the plain text of the column to the screen.
        console.log(column.toTxt());
        //ExEnd:GetPlainText

        // Create a new column to the left of this column.
        // This is the same as using the "Insert Column Before" command in Microsoft Word.
        let newColumn = column.insertColumnBefore();

        for (let cell of newColumn.cells)
            cell.firstParagraph.appendChild(new aw.Run(doc, "Column Text " + newColumn.indexOf(cell)));
        //ExEnd:InsertBlankColumn
    });

    //ExStart:ColumnClass
    //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
    /// <summary>
    /// Represents a facade object for a column of a table in a Microsoft Word document.
    /// </summary>
    class Column {
        constructor(table, columnIndex) {
            if (!table) throw new Error("table");
            this.mTable = table;
            this.mColumnIndex = columnIndex;
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
            return this.getColumnCells().indexOf(cell);
        }

        /// <summary>
        /// Inserts a brand new column before this column into the table.
        /// </summary>
        insertColumnBefore() {
            let columnCells = this.cells;

            if (columnCells.length === 0)
                throw new Error("Column must not be empty");

            // Create a clone of this column.
            for (let cell of columnCells)
                cell.parentRow.insertBefore(cell.clone(false), cell);

            // This is the new column.
            let column = new Column(columnCells[0].parentRow.parentTable, this.mColumnIndex);

            // We want to make sure that the cells are all valid to work with (have at least one paragraph).
            for (let cell of column.cells)
                cell.ensureMinimum();

            // Increase the index which this column represents since there is now one extra column in front.
            this.mColumnIndex++;

            return column;
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
            let builder = '';

            for (let cell of this.cells)
                builder += cell.toString(aw.SaveFormat.Text);

            return builder;
        }

        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        getColumnCells() {
            let columnCells = [];

            for (let row of this.mTable.rows) {
                let cell = row.asRow().cells.at(this.mColumnIndex);
                if (cell != null)
                    columnCells.push(cell);
            }

            return columnCells;
        }
    }
    //ExEnd:ColumnClass

    test('AutoFitTableToContents', () => {
        //ExStart:AutoFitTableToContents
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        table.autoFit(aw.Tables.AutoFitBehavior.AutoFitToContents);

        doc.save(base.artifactsDir + "WorkingWithTables.AutoFitTableToContents.docx");
        //ExEnd:AutoFitTableToContents
    });

    test('AutoFitTableToFixedColumnWidths', () => {
        //ExStart:AutoFitTableToFixedColumnWidths
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        // Disable autofitting on this table.
        table.autoFit(aw.Tables.AutoFitBehavior.FixedColumnWidths);

        doc.save(base.artifactsDir + "WorkingWithTables.AutoFitTableToFixedColumnWidths.docx");
        //ExEnd:AutoFitTableToFixedColumnWidths
    });

    test('AutoFitTableToPageWidth', () => {
        //ExStart:AutoFitTableToPageWidth
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        // Autofit the first table to the page width.
        table.autoFit(aw.Tables.AutoFitBehavior.AutoFitToWindow);

        doc.save(base.artifactsDir + "WorkingWithTables.AutoFitTableToWindow.docx");
        //ExEnd:AutoFitTableToPageWidth
    });

    test('BuildTableFromDataTable', () => {
        //ExStart:BuildTableFromDataTable
        //GistId:9bd44c1142bfae4a4e088b1ff8ccb6ab
        let doc = new aw.Document();
        // We can position where we want the table to be inserted and specify any extra formatting to the table.
        let builder = new aw.DocumentBuilder(doc);

        // We want to rotate the page landscape as we expect a wide table.
        doc.firstSection.pageSetup.orientation = aw.Orientation.Landscape;

        // Simulate DataTable - in Node.js this would be an array of objects
        let dataTable = [
            {Name: "John Doe", Age: 30, City: "New York"},
            {Name: "Jane Smith", Age: 25, City: "London"},
            {Name: "Bob Johnson", Age: 35, City: "Paris"}
        ];

        // Build a table in the document from the data contained in the DataTable.
        let table = importTableFromDataTable(builder, dataTable, true);

        // We can apply a table style as a very quick way to apply formatting to the entire table.
        table.styleIdentifier = aw.StyleIdentifier.MediumList2Accent1;
        table.styleOptions = aw.Tables.TableStyleOptions.FirstRow | aw.Tables.TableStyleOptions.RowBands | aw.Tables.TableStyleOptions.LastColumn;

        // For our table, we want to remove the heading for the image column.
        table.firstRow.lastCell.removeAllChildren();

        doc.save(base.artifactsDir + "WorkingWithTables.BuildTableFromDataTable.docx");
        //ExEnd:BuildTableFromDataTable
    });

    //ExStart:ImportTableFromDataTable
    //GistId:9bd44c1142bfae4a4e088b1ff8ccb6ab
    /// <summary>
    /// Imports the content from the specified DataTable into a new Aspose.Words Table object.
    /// The table is inserted at the document builder's current position and using the current builder's formatting if any is defined.
    /// </summary>
    function importTableFromDataTable(builder, dataTable, importColumnHeadings) {
        let table = builder.startTable();

        // Check if the columns' names from the data source are to be included in a header row.
        if (importColumnHeadings) {
            // Store the original values of these properties before changing them.
            let boldValue = builder.font.bold;
            let paragraphAlignmentValue = builder.paragraphFormat.alignment;

            // Format the heading row with the appropriate properties.
            builder.font.bold = true;
            builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;

            // Create a new row and insert the name of each column into the first row of the table.
            let columns = Object.keys(dataTable[0]);
            for (let column of columns) {
                builder.insertCell();
                builder.writeln(column);
            }

            builder.endRow();

            // Restore the original formatting.
            builder.font.bold = boldValue;
            builder.paragraphFormat.alignment = paragraphAlignmentValue;
        }

        for (let dataRow of dataTable) {
            for (let key of Object.keys(dataRow)) {
                // Insert a new cell for each object.
                builder.insertCell();

                let item = dataRow[key];
                if (item instanceof Date) {
                    // Define a custom format for dates and times.
                    builder.write(item.toLocaleDateString('en-US', {year: 'numeric', month: 'long', day: 'numeric'}));
                } else {
                    // By default any other item will be inserted as text.
                    builder.write(item.toString());
                }
            }

            // After we insert all the data from the current record, we can end the table row.
            builder.endRow();
        }

        // We have finished inserting all the data from the DataTable, we can end the table.
        builder.endTable();

        return table;
    }
    //ExEnd:ImportTableFromDataTable

    test('CloneCompleteTable', () => {
        //ExStart:CloneCompleteTable
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // Clone the table and insert it into the document after the original.
        let tableClone = table.clone(true);
        table.parentNode.insertAfter(tableClone, table);

        // Insert an empty paragraph between the two tables,
        // or else they will be combined into one upon saving this has to do with document validation.
        table.parentNode.insertAfter(new aw.Paragraph(doc), table);

        doc.save(base.artifactsDir + "WorkingWithTables.CloneCompleteTable.docx");
        //ExEnd:CloneCompleteTable
    });

    test('CloneLastRow', () => {
        //ExStart:CloneLastRow
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        let clonedRow = table.lastRow.clone(true).asRow();
        // Remove all content from the cloned row's cells. This makes the row ready for new content to be inserted into.
        for (let cell of clonedRow.cells) {
            cell = cell.asCell();
            cell.removeAllChildren();
        }

        table.appendChild(clonedRow);

        doc.save(base.artifactsDir + "WorkingWithTables.CloneLastRow.docx");
        //ExEnd:CloneLastRow
    });

    test('FindingIndex', () => {
        let doc = new aw.Document(base.myDir + "Tables.docx");

        //ExStart:RetrieveTableIndex
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        let allTables = doc.getChildNodes(aw.NodeType.Table, true);
        let tableIndex = allTables.indexOf(table);
        //ExEnd:RetrieveTableIndex
        console.log("\nTable index is " + tableIndex);

        //ExStart:RetrieveRowIndex
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let rowIndex = table.indexOf(table.lastRow);
        //ExEnd:RetrieveRowIndex
        console.log("\nRow index is " + rowIndex);

        let row = table.lastRow;
        //ExStart:RetrieveCellIndex
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let cellIndex = row.indexOf(row.cells.at(4));
        //ExEnd:RetrieveCellIndex
        console.log("\nCell index is " + cellIndex);
    });

    test('InsertTableDirectly', () => {
        //ExStart:InsertTableDirectly
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document();

        // We start by creating the table object. Note that we must pass the document object
        // to the constructor of each node. This is because every node we create must belong
        // to some document.
        let table = new aw.Tables.Table(doc);
        doc.firstSection.body.appendChild(table);

        // Here we could call EnsureMinimum to create the rows and cells for us. This method is used
        // to ensure that the specified node is valid. In this case, a valid table should have at least one Row and one cell.

        // Instead, we will handle creating the row and table ourselves.
        // This would be the best way to do this if we were creating a table inside an algorithm.
        let row = new aw.Tables.Row(doc);
        row.rowFormat.allowBreakAcrossPages = true;
        table.appendChild(row);

        let cell = new aw.Tables.Cell(doc);
        cell.cellFormat.shading.backgroundPatternColor = "#ADD8E6";
        cell.cellFormat.width = 80;
        cell.appendChild(new aw.Paragraph(doc));
        cell.firstParagraph.appendChild(new aw.Run(doc, "Row 1, Cell 1 Text"));

        row.appendChild(cell);

        // We would then repeat the process for the other cells and rows in the table.
        // We can also speed things up by cloning existing cells and rows.
        row.appendChild(cell.clone(false));
        row.lastCell.appendChild(new aw.Paragraph(doc));
        row.lastCell.firstParagraph.appendChild(new aw.Run(doc, "Row 1, Cell 2 Text"));

        // We can now apply any auto fit settings.
        table.autoFit(aw.Tables.AutoFitBehavior.FixedColumnWidths);

        doc.save(base.artifactsDir + "WorkingWithTables.InsertTableDirectly.docx");
        //ExEnd:InsertTableDirectly
    });

    test('InsertTableFromHtml', () => {
        //ExStart:InsertTableFromHtml
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        // Note that AutoFitSettings does not apply to tables inserted from HTML.
        builder.insertHtml("<table>" +
            "<tr>" +
            "<td>Row 1, Cell 1</td>" +
            "<td>Row 1, Cell 2</td>" +
            "</tr>" +
            "<tr>" +
            "<td>Row 2, Cell 2</td>" +
            "<td>Row 2, Cell 2</td>" +
            "</tr>" +
            "</table>");

        doc.save(base.artifactsDir + "WorkingWithTables.InsertTableFromHtml.docx");
        //ExEnd:InsertTableFromHtml
    });

    test('CreateSimpleTable', () => {
        //ExStart:CreateSimpleTable
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        // Start building the table.
        builder.startTable();
        builder.insertCell();
        builder.write("Row 1, Cell 1 Content.");

        // Build the second cell.
        builder.insertCell();
        builder.write("Row 1, Cell 2 Content.");

        // Call the following method to end the row and start a new row.
        builder.endRow();

        // Build the first cell of the second row.
        builder.insertCell();
        builder.write("Row 2, Cell 1 Content");

        // Build the second cell.
        builder.insertCell();
        builder.write("Row 2, Cell 2 Content.");
        builder.endRow();

        // Signal that we have finished building the table.
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTables.CreateSimpleTable.docx");
        //ExEnd:CreateSimpleTable
    });

    test('FormattedTable', () => {
        //ExStart:FormattedTable
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();

        // Table wide formatting must be applied after at least one row is present in the table.
        table.leftIndent = 20.0;

        // Set height and define the height rule for the header row.
        builder.rowFormat.height = 40.0;
        builder.rowFormat.heightRule = aw.HeightRule.AtLeast;

        builder.cellFormat.shading.backgroundPatternColor = "#C6D9F1";
        builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
        builder.font.size = 16;
        builder.font.name = "Arial";
        builder.font.bold = true;

        builder.cellFormat.width = 100.0;
        builder.write("Header Row,\n Cell 1");

        // We don't need to specify this cell's width because it's inherited from the previous cell.
        builder.insertCell();
        builder.write("Header Row,\n Cell 2");

        builder.insertCell();
        builder.cellFormat.width = 200.0;
        builder.write("Header Row,\n Cell 3");
        builder.endRow();

        builder.cellFormat.shading.backgroundPatternColor = "#FFFFFF";
        builder.cellFormat.width = 100.0;
        builder.cellFormat.verticalAlignment = aw.Tables.CellVerticalAlignment.Center;

        // Reset height and define a different height rule for table body.
        builder.rowFormat.height = 30.0;
        builder.rowFormat.heightRule = aw.HeightRule.Auto;
        builder.insertCell();

        // Reset font formatting.
        builder.font.size = 12;
        builder.font.bold = false;

        builder.write("Row 1, Cell 1 Content");
        builder.insertCell();
        builder.write("Row 1, Cell 2 Content");

        builder.insertCell();
        builder.cellFormat.width = 200.0;
        builder.write("Row 1, Cell 3 Content");
        builder.endRow();

        builder.insertCell();
        builder.cellFormat.width = 100.0;
        builder.write("Row 2, Cell 1 Content");

        builder.insertCell();
        builder.write("Row 2, Cell 2 Content");

        builder.insertCell();
        builder.cellFormat.width = 200.0;
        builder.write("Row 2, Cell 3 Content.");
        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTables.FormattedTable.docx");
        //ExEnd:FormattedTable
    });

    test('NestedTable', () => {
        //ExStart:NestedTable
        //GistId:ba24a0bcb1eecc75eb8db4c8e7f5616c
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let cell = builder.insertCell();
        builder.writeln("Outer Table Cell 1");

        builder.insertCell();
        builder.writeln("Outer Table Cell 2");

        // This call is important to create a nested table within the first table.
        // Without this call, the cells inserted below will be appended to the outer table.
        builder.endTable();

        // Move to the first cell of the outer table.
        builder.moveTo(cell.firstParagraph);

        // Build the inner table.
        builder.insertCell();
        builder.writeln("Inner Table Cell 1");
        builder.insertCell();
        builder.writeln("Inner Table Cell 2");
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTables.NestedTable.docx");
        //ExEnd:NestedTable
    });

    test('CombineRows', () => {
        //ExStart:CombineRows
        //GistId:9cc6f2ce785d8c91aa932c98aeed304d
        let doc = new aw.Document(base.myDir + "Tables.docx");

        // The rows from the second table will be appended to the end of the first table.
        let firstTable = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        let secondTable = doc.getChild(aw.NodeType.Table, 1, true).asTable();

        // Append all rows from the current table to the next tables
        // with different cell count and widths can be joined into one table.
        while (secondTable.hasChildNodes)
            firstTable.rows.add(secondTable.firstRow);

        secondTable.remove();

        doc.save(base.artifactsDir + "WorkingWithTables.CombineRows.docx");
        //ExEnd:CombineRows
    });

    test('SplitTable', () => {
        //ExStart:SplitTable
        //GistId:b8cd11852d8ab0968ecdda0e2baeda15
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let firstTable = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // We will split the table at the third row (inclusive).
        let row = firstTable.rows.at(2);

        // Create a new container for the split table.
        let table = firstTable.clone(false).asTable();

        // Insert the container after the original.
        firstTable.parentNode.insertAfter(table, firstTable);

        // Add a buffer paragraph to ensure the tables stay apart.
        firstTable.parentNode.insertAfter(new aw.Paragraph(doc), firstTable);

        let currentRow;
        do {
            currentRow = firstTable.lastRow;
            table.prependChild(currentRow);
        } while (base.compareNodes(currentRow, row));

        doc.save(base.artifactsDir + "WorkingWithTables.SplitTable.docx");
        //ExEnd:SplitTable
    });

    test('RowFormatDisableBreakAcrossPages', () => {
        //ExStart:RowFormatDisableBreakAcrossPages
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let doc = new aw.Document(base.myDir + "Table spanning two pages.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // Disable breaking across pages for all rows in the table.
        for (let row of table.rows) {
            row = row.asRow();
            row.rowFormat.allowBreakAcrossPages = false;
        }

        doc.save(base.artifactsDir + "WorkingWithTables.RowFormatDisableBreakAcrossPages.docx");
        //ExEnd:RowFormatDisableBreakAcrossPages
    });

    test('KeepTableTogether', () => {
        //ExStart:KeepTableTogether
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let doc = new aw.Document(base.myDir + "Table spanning two pages.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // We need to enable KeepWithNext for every paragraph in the table to keep it from breaking across a page,
        // except for the last paragraphs in the last row of the table.
        for (let cell of table.getChildNodes(aw.NodeType.Cell, true)) {
            cell = cell.asCell();
            cell.ensureMinimum();

            for (let para of cell.paragraphs) {
                para = para.asParagraph();
                if (!(cell.parentRow.isLastRow && para.isEndOfCell))
                    para.paragraphFormat.keepWithNext = true;
            }
        }

        doc.save(base.artifactsDir + "WorkingWithTables.KeepTableTogether.docx");
        //ExEnd:KeepTableTogether
    });

    test('CheckCellsMerged', () => {
        //ExStart:CheckCellsMerged
        //GistId:a2e5839d12017f76e67d145b434558bc
        let doc = new aw.Document(base.myDir + "Table with merged cells.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        for (let row of table.rows) {
            row = row.asRow();
            for (let cell of row.cells) {
                cell = cell.asCell();
                console.log(printCellMergeType(cell));
            }
        }
        //ExEnd:CheckCellsMerged
    });

    //ExStart:PrintCellMergeType
    function printCellMergeType(cell) {
        let isHorizontallyMerged = cell.cellFormat.horizontalMerge !== aw.Tables.CellMerge.None;
        let isVerticallyMerged = cell.cellFormat.verticalMerge !== aw.Tables.CellMerge.None;

        let cellLocation =
            `R${cell.parentRow.parentTable.indexOf(cell.parentRow) + 1}, C${cell.parentRow.indexOf(cell) + 1}`;

        if (isHorizontallyMerged && isVerticallyMerged)
            return `The cell at ${cellLocation} is both horizontally and vertically merged`;

        if (isHorizontallyMerged)
            return `The cell at ${cellLocation} is horizontally merged.`;

        if (isVerticallyMerged)
            return `The cell at ${cellLocation} is vertically merged`;

        return `The cell at ${cellLocation} is not merged`;
    }
    //ExEnd:PrintCellMergeType

    test('VerticalMerge', () => {
        //ExStart:VerticalMerge
        //GistId:a2e5839d12017f76e67d145b434558bc
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.insertCell();
        builder.cellFormat.verticalMerge = aw.Tables.CellMerge.First;
        builder.write("Text in merged cells.");

        builder.insertCell();
        builder.cellFormat.verticalMerge = aw.Tables.CellMerge.None;
        builder.write("Text in one cell");
        builder.endRow();

        builder.insertCell();
        // This cell is vertically merged to the cell above and should be empty.
        builder.cellFormat.verticalMerge = aw.Tables.CellMerge.Previous;

        builder.insertCell();
        builder.cellFormat.verticalMerge = aw.Tables.CellMerge.None;
        builder.write("Text in another cell");
        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTables.VerticalMerge.docx");
        //ExEnd:VerticalMerge
    });

    test('HorizontalMerge', () => {
        //ExStart:HorizontalMerge
        //GistId:a2e5839d12017f76e67d145b434558bc
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.insertCell();
        builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.First;
        builder.write("Text in merged cells.");

        builder.insertCell();
        // This cell is merged to the previous and should be empty.
        builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.Previous;
        builder.endRow();

        builder.insertCell();
        builder.cellFormat.horizontalMerge = aw.Tables.CellMerge.None;
        builder.write("Text in one cell.");

        builder.insertCell();
        builder.write("Text in another cell.");
        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTables.HorizontalMerge.docx");
        //ExEnd:HorizontalMerge
    });

    test('MergeCellRange', () => {
        //ExStart:MergeCellRange
        //GistId:a2e5839d12017f76e67d145b434558bc
        let doc = new aw.Document(base.myDir + "Table with merged cells.docx");

        let table = doc.firstSection.body.tables.at(0);

        // We want to merge the range of cells found inbetween these two cells.
        let cellStartRange = table.rows.at(0).cells.at(0);
        let cellEndRange = table.rows.at(1).cells.at(1);

        // Merge all the cells between the two specified cells into one.
        mergeCells(cellStartRange, cellEndRange);

        doc.save(base.artifactsDir + "WorkingWithTables.MergeCellRange.docx");
        //ExEnd:MergeCellRange
    });

    //ExStart:MergeCells
    //GistId:a2e5839d12017f76e67d145b434558bc
    function mergeCells(startCell, endCell) {
        let parentTable = startCell.parentRow.parentTable;

        // Find the row and cell indices for the start and end cell.
        let startCellPos = {
            x: startCell.parentRow.indexOf(startCell),
            y: parentTable.indexOf(startCell.parentRow)
        };
        let endCellPos = {
            x: endCell.parentRow.indexOf(endCell),
            y: parentTable.indexOf(endCell.parentRow)
        };

        // Create a range of cells to be merged based on these indices.
        // Inverse each index if the end cell is before the start cell.
        let mergeRange = {
            x: Math.min(startCellPos.x, endCellPos.x),
            y: Math.min(startCellPos.y, endCellPos.y),
            width: Math.abs(endCellPos.x - startCellPos.x) + 1,
            height: Math.abs(endCellPos.y - startCellPos.y) + 1
        };

        for (let row of parentTable.rows) {
            row = row.asRow();
            for (let cell of row.cells) {
                cell = cell.asCell();
                let currentPos = {
                    x: row.indexOf(cell),
                    y: parentTable.indexOf(row)
                };

                // Check if the current cell is inside our merge range, then merge it.
                if (currentPos.x >= mergeRange.x && currentPos.x < mergeRange.x + mergeRange.width &&
                    currentPos.y >= mergeRange.y && currentPos.y < mergeRange.y + mergeRange.height) {
                    cell.cellFormat.horizontalMerge = currentPos.x === mergeRange.x ? aw.Tables.CellMerge.First : aw.Tables.CellMerge.Previous;
                    cell.cellFormat.verticalMerge = currentPos.y === mergeRange.y ? aw.Tables.CellMerge.First : aw.Tables.CellMerge.Previous;
                }
            }
        }
    }
    //ExEnd:MergeCells

    test('ConvertToHorizontallyMergedCells', () => {
        //ExStart:ConvertToHorizontallyMergedCells
        //GistId:a2e5839d12017f76e67d145b434558bc
        let doc = new aw.Document(base.myDir + "Table with merged cells.docx");

        let table = doc.firstSection.body.tables.at(0);
        // Now merged cells have appropriate merge flags.
        table.convertToHorizontallyMergedCells();
        //ExEnd:ConvertToHorizontallyMergedCells
    });

    test('RepeatRowsOnSubsequentPages', () => {
        //ExStart:RepeatRowsOnSubsequentPages
        //GistId:0b4aa2dc6bae9b78989a4a7283d7c8da
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.startTable();
        builder.rowFormat.headingFormat = true;
        builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;
        builder.cellFormat.width = 100;
        builder.insertCell();
        builder.writeln("Heading row 1");
        builder.endRow();
        builder.insertCell();
        builder.writeln("Heading row 2");
        builder.endRow();

        builder.cellFormat.width = 50;
        builder.paragraphFormat.clearFormatting();

        for (let i = 0; i < 50; i++) {
            builder.insertCell();
            builder.rowFormat.headingFormat = false;
            builder.write("Column 1 Text");
            builder.insertCell();
            builder.write("Column 2 Text");
            builder.endRow();
        }

        doc.save(base.artifactsDir + "WorkingWithTables.RepeatRowsOnSubsequentPages.docx");
        //ExEnd:RepeatRowsOnSubsequentPages
    });

    test('AutoFitPageWidth', () => {
        //ExStart:AutoFitPageWidth
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        // Insert a table with a width that takes up half the page width.
        let table = builder.startTable();

        builder.insertCell();
        table.preferredWidth = aw.Tables.PreferredWidth.fromPercent(50);
        builder.writeln("Cell #1");

        builder.insertCell();
        builder.writeln("Cell #2");

        builder.insertCell();
        builder.writeln("Cell #3");

        doc.save(base.artifactsDir + "WorkingWithTables.AutoFitPageWidth.docx");
        //ExEnd:AutoFitPageWidth
    });

    test('PreferredWidthSettings', () => {
        //ExStart:PreferredWidthSettings
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        // Insert a table row made up of three cells which have different preferred widths.
        // Insert an absolute sized cell.
        builder.insertCell();
        builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPoints(40);
        builder.cellFormat.shading.backgroundPatternColor = "#FFFFE0";
        builder.writeln("Cell at 40 points width");

        // Insert a relative (percent) sized cell.
        builder.insertCell();
        builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.fromPercent(20);
        builder.cellFormat.shading.backgroundPatternColor = "#ADD8E6";
        builder.writeln("Cell at 20% width");

        // Insert a auto sized cell.
        builder.insertCell();
        builder.cellFormat.preferredWidth = aw.Tables.PreferredWidth.auto;
        builder.cellFormat.shading.backgroundPatternColor = "#90EE90";
        builder.writeln(
            "Cell automatically sized. The size of this cell is calculated from the table preferred width.");
        builder.writeln("In this case the cell will fill up the rest of the available space.");

        doc.save(base.artifactsDir + "WorkingWithTables.PreferredWidthSettings.docx");
        //ExEnd:PreferredWidthSettings
    });

    test('RetrievePreferredWidthType', () => {
        //ExStart:RetrievePreferredWidthType
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        //ExStart:AllowAutoFit
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        table.allowAutoFit = true;
        //ExEnd:AllowAutoFit

        let firstCell = table.firstRow.firstCell;
        let type = firstCell.cellFormat.preferredWidth.type;
        let value = firstCell.cellFormat.preferredWidth.value;
        //ExEnd:RetrievePreferredWidthType
    });

    test('GetTablePosition', () => {
        //ExStart:GetTablePosition
        //GistId:b55c18ec2f5fe3f033f8d24c508b6a0f
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        if (table.textWrapping === aw.Tables.TextWrapping.Around) {
            console.log(table.relativeHorizontalAlignment);
            console.log(table.relativeVerticalAlignment);
        } else {
            console.log(table.alignment);
        }
        //ExEnd:GetTablePosition
    });

    test('GetFloatingTablePosition', () => {
        //ExStart:GetFloatingTablePosition
        //GistId:b55c18ec2f5fe3f033f8d24c508b6a0f
        let doc = new aw.Document(base.myDir + "Table wrapped by text.docx");

        for (let table of doc.firstSection.body.tables) {
            // If the table is floating type, then print its positioning properties.
            if (table.textWrapping === aw.Tables.TextWrapping.Around) {
                console.log(table.horizontalAnchor);
                console.log(table.verticalAnchor);
                console.log(table.absoluteHorizontalDistance);
                console.log(table.absoluteVerticalDistance);
                console.log(table.allowOverlap);
                console.log(table.absoluteHorizontalDistance);
                console.log(table.relativeVerticalAlignment);
                console.log("..............................");
            }
        }
        //ExEnd:GetFloatingTablePosition
    });

    test('FloatingTablePosition', () => {
        //ExStart:FloatingTablePosition
        //GistId:b55c18ec2f5fe3f033f8d24c508b6a0f
        let doc = new aw.Document(base.myDir + "Table wrapped by text.docx");

        let table = doc.firstSection.body.tables.at(0);
        table.absoluteHorizontalDistance = 10;
        table.relativeVerticalAlignment = aw.Drawing.VerticalAlignment.Center;

        doc.save(base.artifactsDir + "WorkingWithTables.FloatingTablePosition.docx");
        //ExEnd:FloatingTablePosition
    });

    test('RelativeHorizontalOrVerticalPosition', () => {
        //ExStart:RelativeHorizontalOrVerticalPosition
        let doc = new aw.Document(base.myDir + "Table wrapped by text.docx");

        let table = doc.firstSection.body.tables.at(0);
        table.horizontalAnchor = aw.Drawing.RelativeHorizontalPosition.Column;
        table.verticalAnchor = aw.Drawing.RelativeVerticalPosition.Page;

        doc.save(base.artifactsDir + "WorkingWithTables.RelativeHorizontalOrVerticalPosition.docx");
        //ExEnd:RelativeHorizontalOrVerticalPosition
    });
});