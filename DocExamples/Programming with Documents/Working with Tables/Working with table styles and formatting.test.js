// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithTableStylesAndFormatting", () => {
    beforeAll(() => {
        base.oneTimeSetup();
    });

    afterAll(() => {
        base.oneTimeTearDown();
    });

    test('DistanceBetweenTableSurroundingText', () => {
        //ExStart:DistanceBetweenTableSurroundingText
        //GistId:b55c18ec2f5fe3f033f8d24c508b6a0f
        let doc = new aw.Document(base.myDir + "Tables.docx");

        console.log("\nGet distance between table left, right, bottom, top and the surrounding text.");
        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        console.log(table.distanceTop);
        console.log(table.distanceBottom);
        console.log(table.distanceRight);
        console.log(table.distanceLeft);
        //ExEnd:DistanceBetweenTableSurroundingText
    });

    test('ApplyOutlineBorder', () => {
        //ExStart:ApplyOutlineBorder
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        //ExStart:InlineTablePosition
        //GistId:b55c18ec2f5fe3f033f8d24c508b6a0f
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        // Align the table to the center of the page.
        table.alignment = aw.Tables.TableAlignment.Center;
        //ExEnd:InlineTablePosition
        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around the table but not inside.
        table.setBorder(aw.BorderType.Left, aw.LineStyle.Single, 1.5, "#008000", true);
        table.setBorder(aw.BorderType.Right, aw.LineStyle.Single, 1.5, "#008000", true);
        table.setBorder(aw.BorderType.Top, aw.LineStyle.Single, 1.5, "#008000", true);
        table.setBorder(aw.BorderType.Bottom, aw.LineStyle.Single, 1.5, "#008000", true);

        // Fill the cells with a light green solid color.
        table.setShading(aw.TextureIndex.TextureSolid, "#90EE90", "#000000");

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.ApplyOutlineBorder.docx");
        //ExEnd:ApplyOutlineBorder
    });

    test('BuildTableWithBorders', () => {
        //ExStart:BuildTableWithBorders
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around and inside the table.
        table.setBorders(aw.LineStyle.Single, 1.5, "#008000");

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.BuildTableWithBorders.docx");
        //ExEnd:BuildTableWithBorders
    });

    test('ModifyRowFormatting', () => {
        //ExStart:ModifyRowFormatting
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        // Retrieve the first row in the table.
        let firstRow = table.firstRow;
        firstRow.rowFormat.borders.lineStyle = aw.LineStyle.None;
        firstRow.rowFormat.heightRule = aw.HeightRule.Auto;
        firstRow.rowFormat.allowBreakAcrossPages = true;
        //ExEnd:ModifyRowFormatting
    });

    test('ApplyRowFormatting', () => {
        //ExStart:ApplyRowFormatting
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();

        let rowFormat = builder.rowFormat;
        rowFormat.height = 100;
        rowFormat.heightRule = aw.HeightRule.Exactly;

        // These formatting properties are set on the table and are applied to all rows in the table.
        table.leftPadding = 30;
        table.rightPadding = 30;
        table.topPadding = 30;
        table.bottomPadding = 30;

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.ApplyRowFormatting.docx");
        //ExEnd:ApplyRowFormatting
    });

    test('CellPadding', () => {
        //ExStart:CellPadding
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        // Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
        builder.cellFormat.setPaddings(30, 50, 30, 50);
        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.CellPadding.docx");
        //ExEnd:CellPadding
    });

    test('ModifyCellFormatting', () => {
        //ExStart:ModifyCellFormatting
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");
        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();

        let firstCell = table.firstRow.firstCell;
        firstCell.cellFormat.width = 30;
        firstCell.cellFormat.orientation = aw.TextOrientation.Downward;
        firstCell.cellFormat.shading.foregroundPatternColor = "#90EE90";
        //ExEnd:ModifyCellFormatting
    });

    test('FormatTableAndCellWithDifferentBorders', () => {
        //ExStart:FormatTableAndCellWithDifferentBorders
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();

        // Set the borders for the entire table.
        table.setBorders(aw.LineStyle.Single, 2.0, "#000000");

        // Set the cell shading for this cell.
        builder.cellFormat.shading.backgroundPatternColor = "#FF0000";
        builder.writeln("Cell #1");

        builder.insertCell();

        // Specify a different cell shading for the second cell.
        builder.cellFormat.shading.backgroundPatternColor = "#008000";
        builder.writeln("Cell #2");

        builder.endRow();

        // Clear the cell formatting from previous operations.
        builder.cellFormat.clearFormatting();

        builder.insertCell();

        // Create larger borders for the first cell of this row. This will be different
        // compared to the borders set for the table.
        builder.cellFormat.borders.left.lineWidth = 4.0;
        builder.cellFormat.borders.right.lineWidth = 4.0;
        builder.cellFormat.borders.top.lineWidth = 4.0;
        builder.cellFormat.borders.bottom.lineWidth = 4.0;
        builder.writeln("Cell #3");

        builder.insertCell();
        builder.cellFormat.clearFormatting();
        builder.writeln("Cell #4");

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.FormatTableAndCellWithDifferentBorders.docx");
        //ExEnd:FormatTableAndCellWithDifferentBorders
    });

    test('TableTitleAndDescription', () => {
        //ExStart:TableTitleAndDescription
        //GistId:1693b4ac01f19ec81c9618649b62acb8
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        table.title = "Test title";
        table.description = "Test description";

        let options = new aw.Saving.OoxmlSaveOptions();
        options.compliance = aw.Saving.OoxmlCompliance.Iso29500_2008_Strict;

        doc.compatibilityOptions.optimizeFor(aw.Settings.MsWordVersion.Word2016);

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.TableTitleAndDescription.docx", options);
        //ExEnd:TableTitleAndDescription
    });

    test('AllowCellSpacing', () => {
        //ExStart:AllowCellSpacing
        //GistId:6d14807d3df5bb7a531673f3b67ed3f7
        let doc = new aw.Document(base.myDir + "Tables.docx");

        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        table.allowCellSpacing = true;
        table.cellSpacing = 2;

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.AllowCellSpacing.docx");
        //ExEnd:AllowCellSpacing
    });

    test('BuildTableWithStyle', () => {
        //ExStart:BuildTableWithStyle
        //GistId:a79ed2d7052cbfbbbc1215708bb4ac4b
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();

        // We must insert at least one row first before setting any table formatting.
        builder.insertCell();

        // Set the table style used based on the unique style identifier.
        table.styleIdentifier = aw.StyleIdentifier.MediumShading1Accent1;

        // Apply which features should be formatted by the style.
        table.styleOptions =
            aw.Tables.TableStyleOptions.FirstColumn | aw.Tables.TableStyleOptions.RowBands | aw.Tables.TableStyleOptions.FirstRow;
        table.autoFit(aw.Tables.AutoFitBehavior.AutoFitToContents);

        builder.writeln("Item");
        builder.cellFormat.rightPadding = 40;
        builder.insertCell();
        builder.writeln("Quantity (kg)");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Apples");
        builder.insertCell();
        builder.writeln("20");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Bananas");
        builder.insertCell();
        builder.writeln("40");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Carrots");
        builder.insertCell();
        builder.writeln("50");
        builder.endRow();

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.BuildTableWithStyle.docx");
        //ExEnd:BuildTableWithStyle
    });

    test('ExpandFormattingOnCellsAndRowFromStyle', () => {
        //ExStart:ExpandFormattingOnCellsAndRowFromStyle
        //GistId:a79ed2d7052cbfbbbc1215708bb4ac4b
        let doc = new aw.Document(base.myDir + "Tables.docx");

        // Get the first cell of the first table in the document.
        let table = doc.getChild(aw.NodeType.Table, 0, true).asTable();
        let firstCell = table.firstRow.firstCell;

        // First print the color of the cell shading.
        // This should be empty as the current shading is stored in the table style.
        let cellShadingBefore = firstCell.cellFormat.shading.backgroundPatternColor;
        console.log("Cell shading before style expansion: " + cellShadingBefore);

        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles.
        // A blue background pattern color should have been applied from the table style.
        let cellShadingAfter = firstCell.cellFormat.shading.backgroundPatternColor;
        console.log("Cell shading after style expansion: " + cellShadingAfter);
        //ExEnd:ExpandFormattingOnCellsAndRowFromStyle
    });

    test('CreateTableStyle', () => {
        //ExStart:CreateTableStyle
        //GistId:a79ed2d7052cbfbbbc1215708bb4ac4b
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();
        builder.write("Name");
        builder.insertCell();
        builder.write("Value");
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();

        let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
        tableStyle.borders.lineStyle = aw.LineStyle.Double;
        tableStyle.borders.lineWidth = 1;
        tableStyle.leftPadding = 18;
        tableStyle.rightPadding = 18;
        tableStyle.topPadding = 12;
        tableStyle.bottomPadding = 12;

        table.style = tableStyle;

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.CreateTableStyle.docx");
        //ExEnd:CreateTableStyle
    });

    test('DefineConditionalFormatting', () => {
        //ExStart:DefineConditionalFormatting
        //GistId:a79ed2d7052cbfbbbc1215708bb4ac4b
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();
        builder.write("Name");
        builder.insertCell();
        builder.write("Value");
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();

        let tableStyle = doc.styles.add(aw.StyleType.Table, "MyTableStyle1").asTableStyle();
        tableStyle.conditionalStyles.at(aw.ConditionalStyleType.FirstRow).shading.backgroundPatternColor = "#ADFF2F";
        tableStyle.conditionalStyles.at(aw.ConditionalStyleType.FirstRow).shading.texture = aw.TextureIndex.TextureNone;

        table.style = tableStyle;

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.DefineConditionalFormatting.docx");
        //ExEnd:DefineConditionalFormatting
    });

    test('SetTableCellFormatting', () => {
        //ExStart:SetTableCellFormatting
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        let cellFormat = builder.cellFormat;
        cellFormat.width = 250;
        cellFormat.leftPadding = 30;
        cellFormat.rightPadding = 30;
        cellFormat.topPadding = 30;
        cellFormat.bottomPadding = 30;

        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.SetTableCellFormatting.docx");
        //ExEnd:SetTableCellFormatting
    });

    test('SetTableRowFormatting', () => {
        //ExStart:SetTableRowFormatting
        let doc = new aw.Document();
        let builder = new aw.DocumentBuilder(doc);

        let table = builder.startTable();
        builder.insertCell();

        let rowFormat = builder.rowFormat;
        rowFormat.height = 100;
        rowFormat.heightRule = aw.HeightRule.Exactly;

        // These formatting properties are set on the table and are applied to all rows in the table.
        table.leftPadding = 30;
        table.rightPadding = 30;
        table.topPadding = 30;
        table.bottomPadding = 30;

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();

        doc.save(base.artifactsDir + "WorkingWithTableStylesAndFormatting.SetTableRowFormatting.docx");
        //ExEnd:SetTableRowFormatting
    });
});