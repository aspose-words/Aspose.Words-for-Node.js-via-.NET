// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');
const MemoryStream = require('memorystream');
const moment = require('moment');

describe("ExField", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('GetFieldFromDocument', () => {
    //ExStart
    //ExFor:FieldType
    //ExFor:FieldChar
    //ExFor:FieldChar.fieldType
    //ExFor:FieldChar.isDirty
    //ExFor:FieldChar.isLocked
    //ExFor:FieldChar.getField
    //ExFor:Field.isLocked
    //ExSummary:Shows how to work with a FieldStart node.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldDate, true).asFieldDate();
    field.format.dateTimeFormat = "dddd, MMMM dd, yyyy";
    field.update();

    let fieldStart = field.start;

    expect(fieldStart.fieldType).toEqual(aw.Fields.FieldType.FieldDate);
    expect(fieldStart.isDirty).toEqual(false);
    expect(fieldStart.isLocked).toEqual(false);

    // Retrieve the facade object which represents the field in the document.
    field = fieldStart.getField().asFieldDate();

    expect(field.isLocked).toEqual(false);
    expect(field.getFieldCode()).toEqual(" DATE  \\@ \"dddd, MMMM dd, yyyy\"");

    // Update the field to show the current date.
    field.update();
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, " DATE  \\@ \"dddd, MMMM dd, yyyy\"", moment(new Date()).format("dddd, MMMM DD, yyyy"), doc.range.fields.at(0));
  });


  test('GetFieldData', () => {
    //ExStart
    //ExFor:FieldStart.fieldData
    //ExSummary:Shows how to get data associated with the field.
    let doc = new aw.Document(base.myDir + "Field sample - Field with data.docx");

    let field = doc.range.fields.at(2);
    console.log(new Buffer.from(field.start.fieldData).toString('utf-8'));
    //ExEnd
  });


  test('GetFieldCode', () => {
    //ExStart
    //ExFor:Field.getFieldCode
    //ExFor:Field.getFieldCode(bool)
    //ExSummary:Shows how to get a field's field code.
    // Open a document which contains a MERGEFIELD inside an IF field.
    let doc = new aw.Document(base.myDir + "Nested fields.docx");
    let fieldIf = doc.range.fields.at(0).asFieldIf();

    // There are two ways of getting a field's field code:
    // 1 -  Omit its inner fields:
    expect(fieldIf.getFieldCode(false)).toEqual(" IF  > 0 \" (surplus of ) \" \"\" ");

    // 2 -  Include its inner fields:
    expect(fieldIf.getFieldCode(true)).toEqual(` IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" `);

    // By default, the GetFieldCode method displays inner fields.
    expect(fieldIf.getFieldCode(true)).toEqual(fieldIf.getFieldCode());
    //ExEnd
  });


  test('DisplayResult', () => {
    //ExStart
    //ExFor:Field.displayResult
    //ExSummary:Shows how to get the real text that a field displays in the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("This document was written by ");
    let fieldAuthor = builder.insertField(aw.Fields.FieldType.FieldAuthor, true).asFieldAuthor();
    fieldAuthor.authorName = "John Doe";

    // We can use the DisplayResult property to verify what exact text
    // a field would display in its place in the document.
    expect(fieldAuthor.displayResult).toEqual('');

    // Fields do not maintain accurate result values in real-time. 
    // To make sure our fields display accurate results at any given time,
    // such as right before a save operation, we need to update them manually.
    fieldAuthor.update();

    expect(fieldAuthor.displayResult).toEqual("John Doe");

    doc.save(base.artifactsDir + "Field.displayResult.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.displayResult.docx");

    expect(doc.range.fields.at(0).displayResult).toEqual("John Doe");
  });


  test('CreateWithFieldBuilder', () => {
    //ExStart
    //ExFor:FieldBuilder.#ctor(FieldType)
    //ExFor:FieldBuilder.buildAndInsert(Inline)
    //ExSummary:Shows how to create and insert a field using a field builder.
    let doc = new aw.Document();

    // A convenient way of adding text content to a document is with a document builder.
    let builder = new aw.DocumentBuilder(doc);
    builder.write(" Hello world! This text is one Run, which is an inline node.");

    // Fields have their builder, which we can use to construct a field code piece by piece.
    // In this case, we will construct a BARCODE field representing a US postal code,
    // and then insert it in front of a Run.
    let fieldBuilder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldBarcode);
    fieldBuilder.addArgument("90210");
    fieldBuilder.addSwitch("\\f", "A");
    fieldBuilder.addSwitch("\\u");

    fieldBuilder.buildAndInsert(doc.firstSection.body.firstParagraph.runs.at(0));

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.CreateWithFieldBuilder.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.CreateWithFieldBuilder.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldBarcode, " BARCODE 90210 \\f A \\u ", '', doc.range.fields.at(0));

    expect(doc.range.fields.at(0).end).toEqual(doc.firstSection.body.firstParagraph.runs.at(11).previousSibling);
    expect(doc.getText().trim()).toEqual(`${aw.ControlChar.fieldStartChar} BARCODE 90210 \\f A \\u ${aw.ControlChar.fieldEndChar} Hello world! This text is one Run, which is an inline node.`);
  });


  test('RevNum', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.revisionNumber
    //ExFor:FieldRevNum
    //ExSummary:Shows how to work with REVNUM fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Current revision #");

    // Insert a REVNUM field, which displays the document's current revision number property.
    let field = builder.insertField(aw.Fields.FieldType.FieldRevisionNum, true).asFieldRevNum();

    expect(field.getFieldCode()).toEqual(" REVNUM ");
    expect(field.result).toEqual("1");
    expect(doc.builtInDocumentProperties.revisionNumber).toEqual(1);

    // This property counts how many times a document has been saved in Microsoft Word,
    // and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
    // via Properties -> Details. We can update this property manually.
    doc.builtInDocumentProperties.revisionNumber++;
    expect(field.result).toEqual("1");
    field.update();

    expect(field.result).toEqual("2");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    expect(doc.builtInDocumentProperties.revisionNumber).toEqual(2);

    TestUtil.verifyField(aw.Fields.FieldType.FieldRevisionNum, " REVNUM ", "2", doc.range.fields.at(0));
  });


  test('InsertFieldNone', () => {
    //ExStart
    //ExFor:FieldUnknown
    //ExSummary:Shows how to work with 'FieldNone' field in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a field that does not denote an objective field type in its field code.
    let field = builder.insertField(" NOTAREALFIELD //a");

    // The "FieldNone" field type is reserved for fields such as these.
    expect(field.type).toEqual(aw.Fields.FieldType.FieldNone);

    // We can also still work with these fields and assign them as instances of the FieldUnknown class.
    let fieldUnknown = field.asFieldUnknown();
    expect(fieldUnknown.getFieldCode()).toEqual(" NOTAREALFIELD //a");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    TestUtil.verifyField(aw.Fields.FieldType.FieldNone, " NOTAREALFIELD //a", "Error! Bookmark not defined.", doc.range.fields.at(0));
  });


  test('InsertTcField', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a TC field at the current document builder position.
    builder.insertField("TC \"Entry Text\" \\f t");
  });


  /*//Commented
  test('InsertTcFieldsAtText', () => {
    let doc = new aw.Document();

    let options = new aw.Replacing.FindReplaceOptions();
    options.replacingCallback = new InsertTcFieldHandler("Chapter 1", "\\l 1");

    // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    doc.range.replace(new Regex("The Beginning"), "", options);
  });


  private class InsertTcFieldHandler : IReplacingCallback
  {
      // Store the text and switches to be used for the TC fields.
    private readonly string mFieldText;
    private readonly string mFieldSwitches;

      /// <summary>
      /// The display text and switches to use for each TC field. Display name can be an empty String or null.
      /// </summary>
    public InsertTcFieldHandler(string text, string switches)
    {
      mFieldText = text;
      mFieldSwitches = switches;
    }

    ReplaceAction aw.Replacing.IReplacingCallback.replacing(ReplacingArgs args)
    {
      let builder = new aw.DocumentBuilder((Document)args.matchNode.document);
      builder.moveTo(args.matchNode);

        // If the user-specified text is used in the field as display text, use that, otherwise
        // use the match String as the display text.
      string insertText = !string.IsNullOrEmpty(mFieldText) ? mFieldText : args.match.value;

        // Insert the TC field before this node using the specified String
        // as the display text and user-defined switches.
      builder.insertField(`TC \"${insertText}\" ${mFieldSwitches}`);

      return aw.Replacing.ReplaceAction.Skip;
    }
  }
  //EndCommented*/

  test.skip('FieldLocale: unsupported CultureInfo', () => {
    //ExStart
    //ExFor:Field.localeId
    //ExSummary:Shows how to insert a field and work with its locale.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a DATE field, and then print the date it will display.
    // Your thread's current culture determines the formatting of the date.
    let field = builder.insertField("DATE");
    //console.log(`Today's date, as displayed in the \"${CultureInfo.CurrentCulture.EnglishName}\" culture: ${field.result}`);

    expect(field.localeId).toEqual(1033);
    expect(doc.fieldOptions.fieldUpdateCultureSource).toEqual(aw.Fields.FieldUpdateCultureSource.CurrentThread);

    // Changing the culture of our thread will impact the result of the DATE field.
    // Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
    // This way allows us to avoid changing the thread's culture to get this effect.
    doc.fieldOptions.fieldUpdateCultureSource = aw.Fields.FieldUpdateCultureSource.FieldCode;
    let de = new CultureInfo("de-DE");
    field.localeId = de.LCID;
    field.update();

    console.log(`Today's date, as displayed according to the \"${CultureInfo.GetCultureInfo(field.localeId).EnglishName}\" culture: ${field.result}`);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    field = doc.range.fields.at(0); 

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, "DATE", Date.now().ToString(de.dateTimeFormat.ShortDatePattern), field);
    expect(field.localeId).toEqual(new CultureInfo("de-DE").LCID);
  });


  test.skip.each([true, false])('UpdateDirtyFields(%o) - TODO: WORDSNODEJS-118', (updateDirtyFields) => {
    //ExStart
    //ExFor:Field.isDirty
    //ExFor:LoadOptions.updateDirtyFields
    //ExSummary:Shows how to use special property for updating field result.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Give the document's built-in "Author" property value, and then display it with a field.
    doc.builtInDocumentProperties.author = "John Doe";
    let field = builder.insertField(aw.Fields.FieldType.FieldAuthor, true).asFieldAuthor();

    expect(field.isDirty).toEqual(false);
    expect(field.result).toEqual("John Doe");

    // Update the property. The field still displays the old value.
    doc.builtInDocumentProperties.author = "John & Jane Doe";

    expect(field.result).toEqual("John Doe");

    // Since the field's value is out of date, we can mark it as "dirty".
    // This value will stay out of date until we update the field manually with the Field.update() method.
    field.isDirty = true;

    let docStream = new MemoryStream();
    // If we save without calling an update method,
    // the field will keep displaying the out of date value in the output document.
    doc.save(docStream, aw.SaveFormat.Docx);

    // The LoadOptions object has an option to update all fields
    // marked as "dirty" when loading the document.
    let options = new aw.Loading.LoadOptions();
    options.updateDirtyFields = updateDirtyFields;
    doc = new aw.Document(docStream, options);

    expect(doc.builtInDocumentProperties.author).toEqual("John & Jane Doe");

    field = doc.range.fields.at(0).asFieldAuthor();

    // Updating dirty fields like this automatically set their "IsDirty" flag to false.
    if (updateDirtyFields)
    {
      expect(field.result).toEqual("John & Jane Doe");
      expect(field.isDirty).toEqual(false);
    }
    else
    {
      expect(field.result).toEqual("John Doe");
      expect(field.isDirty).toEqual(true);
    }
    //ExEnd
  });


  test('InsertFieldWithFieldBuilderException', () => {
    let doc = new aw.Document();

    let run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

    let argumentBuilder = new aw.Fields.FieldArgumentBuilder();
    argumentBuilder.addField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldMergeField));
    argumentBuilder.addNode(run);
    argumentBuilder.addText("Text argument builder");

    let fieldBuilder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIncludeText);

    expect(() => fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField").addArgument(10).addArgument(20.0).buildAndInsert(run))
      .toThrow("Cannot add a node before/after itself.");
  });


/*#if !WORDS_AOT
  test('BarCodeWord2Pdf', () => {
    let doc = new aw.Document(base.myDir + "Field sample - BARCODE.docx");

    doc.fieldOptions.barcodeGenerator = new CustomBarcodeGenerator();

    doc.save(base.artifactsDir + "Field.BarCodeWord2Pdf.pdf");

    using (BarCodeReader barCodeReader = BarCodeReaderPdf(base.artifactsDir + "Field.BarCodeWord2Pdf.pdf"))
    {
      expect(barCodeReader.FoundBarCodes.at(0).CodeTypeName).toEqual("QR");
    }
  });


  private BarCodeReader BarCodeReaderPdf(string filename)
  {
      // Set license for Aspose.BarCode.
    Aspose.BarCode.License licenceBarCode = new Aspose.BarCode.License();
    licenceBarCode.setLicense(base.licenseDir + "Aspose.Total.NET.lic");

    Aspose.pdf.Facades.PdfExtractor pdfExtractor = new Aspose.pdf.Facades.PdfExtractor();
    pdfExtractor.BindPdf(filename);

      // Set page range for image extraction.
    pdfExtractor.startPage = 1;
    pdfExtractor.endPage = 1;

    pdfExtractor.ExtractImage();

    let imageStream = new MemoryStream();
    pdfExtractor.GetNextImage(imageStream);
    imageStream.position = 0;

      // Recognize the barcode from the image stream above.
    let barcodeReader = new BarCodeReader(imageStream, DecodeType.QR);

    foreach (BarCodeResult result in barcodeReader.ReadBarCodes())
      console.log("Codetext found: " + result.CodeText + ", Symbology: " + result.CodeTypeName);

    return barcodeReader;
  }
#endif*/

  /*#if WORDS_AOT
    [Ignore("OLEDB is not supported in AOT")]
  #endif
  test('FieldDatabase', () => {
    //ExStart
    //ExFor:FieldDatabase
    //ExFor:FieldDatabase.connection
    //ExFor:FieldDatabase.fileName
    //ExFor:FieldDatabase.firstRecord
    //ExFor:FieldDatabase.formatAttributes
    //ExFor:FieldDatabase.insertHeadings
    //ExFor:FieldDatabase.insertOnceOnMailMerge
    //ExFor:FieldDatabase.lastRecord
    //ExFor:FieldDatabase.query
    //ExFor:FieldDatabase.tableFormat
    //ExFor:FieldDatabaseDataTable
    //ExFor:IFieldDatabaseProvider
    //ExFor:IFieldDatabaseProvider.getQueryResult(String,String,String,FieldDatabase)
    //ExFor:FieldOptions.fieldDatabaseProvider
    //ExSummary:Shows how to extract data from a database and insert it as a field into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // This DATABASE field will run a query on a database, and display the result in a table.
    let field = (FieldDatabase)builder.insertField(aw.Fields.FieldType.FieldDatabase, true);
    field.fileName = base.databaseDir + "Northwind.accdb";
    field.connection = "Provider=Microsoft.ACE.OLEDB.12.0";
    field.query = "SELECT * FROM [Products]";

    expect(field.getFieldCode()).toEqual(` DATABASE  \\d {base.databaseDir.replace(`\\", "\\\\") + "Northwind.accdb"} \\c Provider=Microsoft.ACE.OLEDB.12.0 \\s \"SELECT * FROM [Products]\"");

    // Insert another DATABASE field with a more complex query that sorts all products in descending order by gross sales.
    field = (FieldDatabase)builder.insertField(aw.Fields.FieldType.FieldDatabase, true);
    field.fileName = base.databaseDir + "Northwind.accdb";
    field.connection = "Provider=Microsoft.ACE.OLEDB.12.0";
    field.query =
      "SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
      "FROM([Products] " +
      "LEFT JOIN.at(Order Details) ON.at(Products).[ProductID] = [Order Details].[ProductID]) " +
      "GROUP BY.at(Products).productName " +
      "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC";

    // These properties have the same function as LIMIT and TOP clauses.
    // Configure them to display only rows 1 to 10 of the query result in the field's table.
    field.firstRecord = "1";
    field.lastRecord = "10";

    // This property is the index of the format we want to use for our table. The list of table formats is in the "Table AutoFormat..." menu
    // that shows up when we create a DATABASE field in Microsoft Word. Index #10 corresponds to the "Colorful 3" format.
    field.tableFormat = "10";

    // The FormatAttribute property is a string representation of an integer which stores multiple flags.
    // We can patrially apply the format which the TableFormat property points to by setting different flags in this property.
    // The number we use is the sum of a combination of values corresponding to different aspects of the table style.
    // 63 represents 1 (borders) + 2 (shading) + 4 (font) + 8 (color) + 16 (autofit) + 32 (heading rows).
    field.formatAttributes = "63";
    field.insertHeadings = true;
    field.insertOnceOnMailMerge = true;

    doc.fieldOptions.fieldDatabaseProvider = new OleDbFieldDatabaseProvider();
    doc.updateFields();

    doc.save(base.artifactsDir + "Field.DATABASE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.DATABASE.docx");

    expect(doc.range.fields.count).toEqual(2);

    let table = doc.firstSection.body.tables.at(0);

    expect(table.rows.count).toEqual(77);
    expect(table.rows.at(0).cells.count).toEqual(10);

    field = (FieldDatabase)doc.range.fields.at(0);

    expect(field.getFieldCode()).toEqual(` DATABASE  \\d {base.databaseDir.replace(`\\", "\\\\") + "Northwind.accdb"} \\c Provider=Microsoft.ACE.OLEDB.12.0 \\s \"SELECT * FROM [Products]\"");

    TestUtil.TableMatchesQueryResult(table, base.databaseDir + "Northwind.accdb", field.query);

    table = (Table)doc.getChild(aw.NodeType.Table, 1, true);
    field = (FieldDatabase)doc.range.fields.at(1);

    expect(table.rows.count).toEqual(11);
    expect(table.rows.at(0).cells.count).toEqual(2);
    expect(table.rows.at(0).cells.at(0).getText()).toEqual("ProductName\u0007");
    expect(table.rows.at(0).cells.at(1).getText()).toEqual("GrossSales\u0007");

    expect(field.getFieldCode()).toEqual(` DATABASE  \\d {base.databaseDir.replace(`\\", "\\\\") + "Northwind.accdb"} \\c Provider=Microsoft.ACE.OLEDB.12.0 " +
                            `\\s \"SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales ` +
                            "FROM([Products] " +
                            "LEFT JOIN.at(Order Details) ON.at(Products).[ProductID] = [Order Details].[ProductID]) " +
                            "GROUP BY[Products].productName " +
                            "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o");

    table.rows.at(0).remove();

    TestUtil.TableMatchesQueryResult(table, base.databaseDir + "Northwind.accdb", field.query.insert(7, " TOP 10 "));
  });


  public class OleDbFieldDatabaseProvider : IFieldDatabaseProvider
  {
    FieldDatabaseDataTable aw.Fields.IFieldDatabaseProvider.getQueryResult(string fileName, string connection, string query, FieldDatabase field)
    {
      let connectionStringBuilder = new OleDbConnectionStringBuilder(connection);
      connectionStringBuilder.dataSource = fileName;

      {
        let oleDbDataAdapter = new OleDbDataAdapter(query, oleDbConnection);
        let dataTable = new DataTable();
        oleDbDataAdapter.fill(dataTable);

        return aw.Fields.FieldDatabaseDataTable.createFrom(dataTable);
      }
    }
  }*/

  test.skip.each([false, true])('PreserveIncludePicture(%o) - TODO: WORDSNODEJS-118', (preserveIncludePictureField) => {
    //ExStart
    //ExFor:Field.update(bool)
    //ExFor:LoadOptions.preserveIncludePictureField
    //ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let includePicture = builder.insertField(aw.Fields.FieldType.FieldIncludePicture, true).asFieldIncludePicture();
    includePicture.sourceFullName = base.imageDir + "Transparent background logo.png";
    includePicture.update(true);

    let docStream = new MemoryStream();
    doc.save(docStream, new aw.Saving.OoxmlSaveOptions(aw.SaveFormat.Docx));

    // We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
    // into image shapes when loading a document that contains them.
    let loadOptions = new aw.Loading.LoadOptions();
    loadOptions.preserveIncludePictureField = preserveIncludePictureField;

    doc = new aw.Document(docStream, loadOptions);

    if (preserveIncludePictureField)
    {
      expect(doc.range.fields.some(f => f.type == aw.Fields.FieldType.FieldIncludePicture)).toBeTruthy();

      doc.updateFields();
      doc.save(base.artifactsDir + "Field.PreserveIncludePicture.docx");
    }
    else
    {
      expect(doc.range.fields.some(f => f.type == aw.Fields.FieldType.FieldIncludePicture)).toBeFalsy();
    }
    //ExEnd
  });


  test('FieldFormat', () => {
    //ExStart
    //ExFor:Field.format
    //ExFor:Field.update
    //ExFor:FieldFormat
    //ExFor:FieldFormat.dateTimeFormat
    //ExFor:FieldFormat.numericFormat
    //ExFor:FieldFormat.generalFormats
    //ExFor:GeneralFormat
    //ExFor:GeneralFormatCollection
    //ExFor:GeneralFormatCollection.add(GeneralFormat)
    //ExFor:GeneralFormatCollection.count
    //ExFor:GeneralFormatCollection.item(Int32)
    //ExFor:GeneralFormatCollection.remove(GeneralFormat)
    //ExFor:GeneralFormatCollection.removeAt(Int32)
    //ExFor:GeneralFormatCollection.getEnumerator
    //ExSummary:Shows how to format field results.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a document builder to insert a field that displays a result with no format applied.
    let field = builder.insertField("= 2 + 3");

    expect(field.getFieldCode()).toEqual("= 2 + 3");
    expect(field.result).toEqual("5");

    // We can apply a format to a field's result using the field's properties.
    // Below are three types of formats that we can apply to a field's result.
    // 1 -  Numeric format:
    let format = field.format;
    format.numericFormat = "$###.00";
    field.update();

    expect(field.getFieldCode()).toEqual("= 2 + 3 \\# $###.00");
    expect(field.result).toEqual("$  5.00");

    // 2 -  Date/time format:
    field = builder.insertField("DATE");
    format = field.format;
    format.dateTimeFormat = "dddd, MMMM dd, yyyy";
    field.update();

    expect(field.getFieldCode()).toEqual("DATE \\@ \"dddd, MMMM dd, yyyy\"");
    console.log(`Today's date, in ${format.dateTimeFormat} format:\n\t${field.result}`);

    // 3 -  General format:
    field = builder.insertField("= 25 + 33");
    format = field.format;
    format.generalFormats.add(aw.Fields.GeneralFormat.LowercaseRoman);
    format.generalFormats.add(aw.Fields.GeneralFormat.Upper);
    field.update();

    let index = 0;
    for (let generalFormat of format.generalFormats)
      console.log(`General format index ${index++}: ${generalFormat}`);

    expect(field.getFieldCode()).toEqual("= 25 + 33 \\* roman \\* Upper");
    expect(field.result).toEqual("LVIII");
    expect(format.generalFormats.count).toEqual(2);
    expect(format.generalFormats.at(0)).toEqual(aw.Fields.GeneralFormat.LowercaseRoman);

    // We can remove our formats to revert the field's result to its original form.
    format.generalFormats.remove(aw.Fields.GeneralFormat.LowercaseRoman);
    format.generalFormats.removeAt(0);
    expect(format.generalFormats.count).toEqual(0);
    field.update();

    expect(field.getFieldCode()).toEqual("= 25 + 33  ");
    expect(field.result).toEqual("58");
    expect(format.generalFormats.count).toEqual(0);
    //ExEnd
  });


  test('Unlink', () => {
    //ExStart
    //ExFor:Document.unlinkFields
    //ExSummary:Shows how to unlink all fields in the document.
    let doc = new aw.Document(base.myDir + "Linked fields.docx");

    doc.unlinkFields();
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    let paraWithFields = DocumentHelper.getParagraphText(doc, 0);

    expect(paraWithFields).toEqual("Fields.Docx   Элементы указателя не найдены.     1.\r");
  });


  test('UnlinkAllFieldsInRange', () => {
    //ExStart
    //ExFor:Range.unlinkFields
    //ExSummary:Shows how to unlink all fields in a range.
    let doc = new aw.Document(base.myDir + "Linked fields.docx");

    let newSection = doc.sections.at(0).clone();
    doc.sections.add(newSection);

    doc.sections.at(1).range.unlinkFields();
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    let secWithFields = DocumentHelper.getSectionText(doc, 1);
    expect(secWithFields.includes(
                "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4.")).toBe(true);
  });


  test('UnlinkSingleField', () => {
    //ExStart
    //ExFor:Field.unlink
    //ExSummary:Shows how to unlink a field.
    let doc = new aw.Document(base.myDir + "Linked fields.docx");
    doc.range.fields.at(1).unlink();
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    let paraWithFields = DocumentHelper.getParagraphText(doc, 0);

    expect(paraWithFields.trim().endsWith(
                "FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015")).toBeTruthy();
  });


  test('UpdateTocPageNumbers', () => {
    let doc = new aw.Document(base.myDir + "Field sample - TOC.docx");

    let startNode = DocumentHelper.getParagraph(doc, 2);
    let endNode = null;

    let paragraphCollection = doc.getChildNodes(aw.NodeType.Paragraph, true).toArray().map(node => node.asParagraph());

    for (let para of paragraphCollection)
    {
      for (let run of para.runs.toArray().map(node => node.asRun()))
      {
        if (run.text.includes(aw.ControlChar.pageBreak))
        {
          endNode = run;
          break;
        }
      }
    }

    if (startNode != null && endNode != null)
    {
      removeSequence(startNode, endNode);

      startNode.remove();
      endNode.remove();
    }

    let fStart = doc.getChildNodes(aw.NodeType.FieldStart, true).toArray().map(node => node.asFieldStart());

    for (let field of fStart)
    {
      let fType = field.fieldType;
      if (fType == aw.Fields.FieldType.FieldTOC)
      {
        let para = field.getAncestor(aw.NodeType.Paragraph).asParagraph();
        para.range.updateFields();
        break;
      }
    }

    doc.save(base.artifactsDir + "Field.UpdateTocPageNumbers.docx");
  });


  function removeSequence(start, end) {
    let curNode = start.nextPreOrder(start.document);
    while (curNode != null && !curNode.referenceEquals(end))
    {
      let nextNode = curNode.nextPreOrder(start.document);

      if (curNode.isComposite)
      {
        let curComposite = curNode.asCompositeNode();
        if (!curComposite.getChildNodes(aw.NodeType.Any, true).toArray().includes(end) &&
          !curComposite.getChildNodes(aw.NodeType.Any, true).toArray().includes(start))
        {
          nextNode = curNode.nextSibling;
          curNode.remove();
        }
      }
      else
      {
        curNode.remove();
      }

      curNode = nextNode;
    }
  }

  /*  //ExStart
    //ExFor:FieldAsk
    //ExFor:FieldAsk.BookmarkName
    //ExFor:FieldAsk.DefaultResponse
    //ExFor:FieldAsk.PromptOnceOnMailMerge
    //ExFor:FieldAsk.PromptText
    //ExFor:FieldOptions.UserPromptRespondent
    //ExFor:IFieldUserPromptRespondent
    //ExFor:IFieldUserPromptRespondent.Respond(String,String)
    //ExSummary:Shows how to create an ASK field, and set its properties.
  test('FieldAsk', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Place a field where the response to our ASK field will be placed.
    let fieldRef = (FieldRef)builder.insertField(aw.Fields.FieldType.FieldRef, true);
    fieldRef.bookmarkName = "MyAskField";
    builder.writeln();

    expect(fieldRef.getFieldCode()).toEqual(" REF  MyAskField");

    // Insert the ASK field and edit its properties to reference our REF field by bookmark name.
    let fieldAsk = (FieldAsk)builder.insertField(aw.Fields.FieldType.FieldAsk, true);
    fieldAsk.bookmarkName = "MyAskField";
    fieldAsk.promptText = "Please provide a response for this ASK field";
    fieldAsk.defaultResponse = "Response from within the field.";
    fieldAsk.promptOnceOnMailMerge = true;
    builder.writeln();

    Assert.AreEqual(
      " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
      fieldAsk.getFieldCode());

    // ASK fields apply the default response to their respective REF fields during a mail merge.
    let table = new DataTable("My Table");
    table.columns.add("Column 1");
    table.rows.add("Row 1");
    table.rows.add("Row 2");

    let fieldMergeField = (FieldMergeField)builder.insertField(aw.Fields.FieldType.FieldMergeField, true);
    fieldMergeField.fieldName = "Column 1";

    // We can modify or override the default response in our ASK fields with a custom prompt responder,
    // which will occur during a mail merge.
    doc.fieldOptions.userPromptRespondent = new MyPromptRespondent();
    doc.mailMerge.execute(table);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.ASK.docx");
    TestFieldAsk(table, doc); //ExSkip
  });


    /// <summary>
    /// Prepends text to the default response of an ASK field during a mail merge.
    /// </summary>
  private class MyPromptRespondent : IFieldUserPromptRespondent
  {
    public string Respond(string promptText, string defaultResponse)
    {
      return "Response from MyPromptRespondent. " + defaultResponse;
    }
  }
    //ExEnd

  private void TestFieldAsk(DataTable dataTable, Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);

    let fieldRef = (FieldRef)doc.range.fields.first(f => f.type == aw.Fields.FieldType.FieldRef);
    TestUtil.VerifyField(aw.Fields.FieldType.FieldRef, 
      " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", fieldRef);

    let fieldAsk = (FieldAsk)doc.range.fields.first(f => f.type == aw.Fields.FieldType.FieldAsk);
    TestUtil.VerifyField(aw.Fields.FieldType.FieldAsk, 
      " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o", 
      "Response from MyPromptRespondent. Response from within the field.", fieldAsk);

    expect(fieldAsk.bookmarkName).toEqual("MyAskField");
    expect(fieldAsk.promptText).toEqual("Please provide a response for this ASK field");
    expect(fieldAsk.defaultResponse).toEqual("Response from within the field.");
    expect(fieldAsk.promptOnceOnMailMerge).toEqual(true);

    TestUtil.MailMergeMatchesDataTable(dataTable, doc, true);
  }*/

  test('FieldAdvance', () => {
    //ExStart
    //ExFor:FieldAdvance
    //ExFor:FieldAdvance.downOffset
    //ExFor:FieldAdvance.horizontalPosition
    //ExFor:FieldAdvance.leftOffset
    //ExFor:FieldAdvance.rightOffset
    //ExFor:FieldAdvance.upOffset
    //ExFor:FieldAdvance.verticalPosition
    //ExSummary:Shows how to insert an ADVANCE field, and edit its properties. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("This text is in its normal place.");

    // Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
    // The effects of an ADVANCE field continue to be applied until the paragraph ends,
    // or another ADVANCE field updates the offset/coordinate values.
    // 1 -  Specify a directional offset:
    let field = builder.insertField(aw.Fields.FieldType.FieldAdvance, true).asFieldAdvance();
    expect(field.type).toEqual(aw.Fields.FieldType.FieldAdvance);
    expect(field.getFieldCode()).toEqual(" ADVANCE ");
    field.rightOffset = "5";
    field.upOffset = "5";

    expect(field.getFieldCode()).toEqual(" ADVANCE  \\r 5 \\u 5");

    builder.write("This text will be moved up and to the right.");

    field = builder.insertField(aw.Fields.FieldType.FieldAdvance, true).asFieldAdvance();
    field.downOffset = "5";
    field.leftOffset = "100";

    expect(field.getFieldCode()).toEqual(" ADVANCE  \\d 5 \\l 100");

    builder.writeln("This text is moved down and to the left, overlapping the previous text.");

    // 2 -  Move text to a position specified by coordinates:
    field = builder.insertField(aw.Fields.FieldType.FieldAdvance, true).asFieldAdvance();
    field.horizontalPosition = "-100";
    field.verticalPosition = "200";

    expect(field.getFieldCode()).toEqual(" ADVANCE  \\x -100 \\y 200");

    builder.write("This text is in a custom position.");

    doc.save(base.artifactsDir + "Field.ADVANCE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.ADVANCE.docx");

    field = doc.range.fields.at(0).asFieldAdvance();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAdvance, " ADVANCE  \\r 5 \\u 5", '', field);
    expect(field.rightOffset).toEqual("5");
    expect(field.upOffset).toEqual("5");

    field = doc.range.fields.at(1).asFieldAdvance();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAdvance, " ADVANCE  \\d 5 \\l 100", '', field);
    expect(field.downOffset).toEqual("5");
    expect(field.leftOffset).toEqual("100");

    field = doc.range.fields.at(2).asFieldAdvance();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAdvance, " ADVANCE  \\x -100 \\y 200", '', field);
    expect(field.horizontalPosition).toEqual("-100");
    expect(field.verticalPosition).toEqual("200");
  });


  test('FieldAddressBlock', () => {
    //ExStart
    //ExFor:FieldAddressBlock.excludedCountryOrRegionName
    //ExFor:FieldAddressBlock.formatAddressOnCountryOrRegion
    //ExFor:FieldAddressBlock.includeCountryOrRegionName
    //ExFor:FieldAddressBlock.languageId
    //ExFor:FieldAddressBlock.nameAndAddressFormat
    //ExSummary:Shows how to insert an ADDRESSBLOCK field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldAddressBlock, true).asFieldAddressBlock();

    expect(field.getFieldCode()).toEqual(" ADDRESSBLOCK ");

    // Setting this to "2" will include all countries and regions,
    // unless it is the one specified in the ExcludedCountryOrRegionName property.
    field.includeCountryOrRegionName = "2";
    field.formatAddressOnCountryOrRegion = true;
    field.excludedCountryOrRegionName = "United States";
    field.nameAndAddressFormat = "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>";

    // By default, this property will contain the language ID of the first character of the document.
    // We can set a different culture for the field to format the result with like this.
    // new CultureInfo("en-US").LCID.toString() == "1033"
    field.languageId = "1033";

    expect(field.getFieldCode()).toEqual(" ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    field = doc.range.fields.at(0).asFieldAddressBlock();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAddressBlock, 
      " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033", 
      "«AddressBlock»", field);
    expect(field.includeCountryOrRegionName).toEqual("2");
    expect(field.formatAddressOnCountryOrRegion).toEqual(true);
    expect(field.excludedCountryOrRegionName).toEqual("United States");
    expect(field.nameAndAddressFormat).toEqual("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");
    expect(field.languageId).toEqual("1033");
  });


  /*  //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a collection of fields.
  test('FieldCollection', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder.insertField(" TIME ");
    builder.insertField(" REVNUM ");
    builder.insertField(" AUTHOR  \"John Doe\" ");
    builder.insertField(" SUBJECT \"My Subject\" ");
    builder.insertField(" QUOTE \"Hello world!\" ");
    doc.updateFields();

    let fields = doc.range.fields;

    expect(fields.count).toEqual(6);

    // Iterate over the field collection, and print contents and type
    // of every field using a custom visitor implementation.
    let fieldVisitor = new FieldVisitor();

    using (IEnumerator<Field> fieldEnumerator = fields.getEnumerator())
    {
      while (fieldEnumerator.moveNext())
      {
        if (fieldEnumerator.current != null)
        {
          fieldEnumerator.current.start.accept(fieldVisitor);
          fieldEnumerator.current.separator?.Accept(fieldVisitor);
          fieldEnumerator.current.end.accept(fieldVisitor);
        }
        else
        {
          console.log("There are no fields in the document.");
        }
      }
    }

    console.log(fieldVisitor.getText());
    TestFieldCollection(fieldVisitor.getText()); //ExSkip
  });


    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
  public class FieldVisitor : DocumentVisitor
  {
    public FieldVisitor()
    {
      mBuilder = new StringBuilder();
    }

      /// <summary>
      /// Gets the plain text of the document that was accumulated by the visitor.
      /// </summary>
    public string GetText()
    {
      return mBuilder.toString();
    }

      /// <summary>
      /// Called when a FieldStart node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldStart(FieldStart fieldStart)
    {
      mBuilder.AppendLine("Found field: " + fieldStart.fieldType);
      mBuilder.AppendLine("\tField code: " + fieldStart.getField().GetFieldCode());
      mBuilder.AppendLine("\tDisplayed as: " + fieldStart.getField().Result);

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a FieldSeparator node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
    {
      mBuilder.AppendLine("\tFound separator: " + fieldSeparator.getText());

      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Called when a FieldEnd node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
    {
      mBuilder.AppendLine("End of field: " + fieldEnd.fieldType);

      return aw.VisitorAction.Continue;
    }

    private readonly StringBuilder mBuilder;
  }
    //ExEnd

  private void TestFieldCollection(string fieldVisitorText)
  {
    expect(fieldVisitorText.contains("Found field: FieldDate")).toEqual(true);
    expect(fieldVisitorText.contains("Found field: FieldTime")).toEqual(true);
    expect(fieldVisitorText.contains("Found field: FieldRevisionNum")).toEqual(true);
    expect(fieldVisitorText.contains("Found field: FieldAuthor")).toEqual(true);
    expect(fieldVisitorText.contains("Found field: FieldSubject")).toEqual(true);
    expect(fieldVisitorText.contains("Found field: FieldQuote")).toEqual(true);
  }*/

  test('RemoveFields', () => {
    //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.count
    //ExFor:FieldCollection.clear
    //ExFor:FieldCollection.item(Int32)
    //ExFor:FieldCollection.remove(Field)
    //ExFor:FieldCollection.removeAt(Int32)
    //ExFor:Field.remove
    //ExSummary:Shows how to remove fields from a field collection.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder.insertField(" TIME ");
    builder.insertField(" REVNUM ");
    builder.insertField(" AUTHOR  \"John Doe\" ");
    builder.insertField(" SUBJECT \"My Subject\" ");
    builder.insertField(" QUOTE \"Hello world!\" ");
    doc.updateFields();

    let fields = doc.range.fields;

    expect(fields.count).toEqual(6);

    // Below are four ways of removing fields from a field collection.
    // 1 -  Get a field to remove itself:
    fields.at(0).remove();
    expect(fields.count).toEqual(5);

    // 2 -  Get the collection to remove a field that we pass to its removal method:
    let lastField = fields.at(3);
    fields.remove(lastField);
    expect(fields.count).toEqual(4);

    // 3 -  Remove a field from a collection at an index:
    fields.removeAt(2);
    expect(fields.count).toEqual(3);

    // 4 -  Remove all the fields from the collection at once:
    fields.clear();
    expect(fields.count).toEqual(0);
    //ExEnd
  });


  test('FieldCompare', () => {
    //ExStart
    //ExFor:FieldCompare
    //ExFor:FieldCompare.comparisonOperator
    //ExFor:FieldCompare.leftExpression
    //ExFor:FieldCompare.rightExpression
    //ExSummary:Shows how to compare expressions using a COMPARE field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldCompare, true).asFieldCompare();
    field.leftExpression = "3";
    field.comparisonOperator = "<";
    field.rightExpression = "2";
    field.update();

    // The COMPARE field displays a "0" or a "1", depending on its statement's truth.
    // The result of this statement is false so that this field will display a "0".
    expect(field.getFieldCode()).toEqual(" COMPARE  3 < 2");
    expect(field.result).toEqual("0");

    builder.writeln();

    field = builder.insertField(aw.Fields.FieldType.FieldCompare, true).asFieldCompare();
    field.leftExpression = "5";
    field.comparisonOperator = "=";
    field.rightExpression = "2 + 3";
    field.update();

    // This field displays a "1" since the statement is true.
    expect(field.getFieldCode()).toEqual(" COMPARE  5 = \"2 + 3\"");
    expect(field.result).toEqual("1");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.COMPARE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.COMPARE.docx");

    field = doc.range.fields.at(0).asFieldCompare();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCompare, " COMPARE  3 < 2", "0", field);
    expect(field.leftExpression).toEqual("3");
    expect(field.comparisonOperator).toEqual("<");
    expect(field.rightExpression).toEqual("2");

    field = doc.range.fields.at(1).asFieldCompare();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCompare, " COMPARE  5 = \"2 + 3\"", "1", field);
    expect(field.leftExpression).toEqual("5");
    expect(field.comparisonOperator).toEqual("=");
    expect(field.rightExpression).toEqual("\"2 + 3\"");
  });


  test('FieldIf', () => {
    //ExStart
    //ExFor:FieldIf
    //ExFor:FieldIf.comparisonOperator
    //ExFor:FieldIf.evaluateCondition
    //ExFor:FieldIf.falseText
    //ExFor:FieldIf.leftExpression
    //ExFor:FieldIf.rightExpression
    //ExFor:FieldIf.trueText
    //ExFor:FieldIfComparisonResult
    //ExSummary:Shows how to insert an IF field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Statement 1: ");
    let field = builder.insertField(aw.Fields.FieldType.FieldIf, true).asFieldIf();
    field.leftExpression = "0";
    field.comparisonOperator = "=";
    field.rightExpression = "1";

    // The IF field will display a string from either its "TrueText" property,
    // or its "FalseText" property, depending on the truth of the statement that we have constructed.
    field.trueText = "True";
    field.falseText = "False";
    field.update();

    // In this case, "0 = 1" is incorrect, so the displayed result will be "False".
    expect(field.getFieldCode()).toEqual(" IF  0 = 1 True False");
    expect(field.evaluateCondition()).toEqual(aw.Fields.FieldIfComparisonResult.False);
    expect(field.result).toEqual("False");

    builder.write("\nStatement 2: ");
    field = builder.insertField(aw.Fields.FieldType.FieldIf, true).asFieldIf();
    field.leftExpression = "5";
    field.comparisonOperator = "=";
    field.rightExpression = "2 + 3";
    field.trueText = "True";
    field.falseText = "False";
    field.update();

    // This time the statement is correct, so the displayed result will be "True".
    expect(field.getFieldCode()).toEqual(" IF  5 = \"2 + 3\" True False");
    expect(field.evaluateCondition()).toEqual(aw.Fields.FieldIfComparisonResult.True);
    expect(field.result).toEqual("True");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.IF.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.IF.docx");
    field = doc.range.fields.at(0).asFieldIf();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIf, " IF  0 = 1 True False", "False", field);
    expect(field.leftExpression).toEqual("0");
    expect(field.comparisonOperator).toEqual("=");
    expect(field.rightExpression).toEqual("1");
    expect(field.trueText).toEqual("True");
    expect(field.falseText).toEqual("False");

    field = doc.range.fields.at(1).asFieldIf();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIf, " IF  5 = \"2 + 3\" True False", "True", field);
    expect(field.leftExpression).toEqual("5");
    expect(field.comparisonOperator).toEqual("=");
    expect(field.rightExpression).toEqual("\"2 + 3\"");
    expect(field.trueText).toEqual("True");
    expect(field.falseText).toEqual("False");
  });


  test('FieldAutoNum', () => {
    //ExStart
    //ExFor:FieldAutoNum
    //ExFor:FieldAutoNum.separatorCharacter
    //ExSummary:Shows how to number paragraphs using autonum fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Each AUTONUM field displays the current value of a running count of AUTONUM fields,
    // allowing us to automatically number items like a numbered list.
    // This field will display a number "1.".
    let field = builder.insertField(aw.Fields.FieldType.FieldAutoNum, true).asFieldAutoNum();
    builder.writeln("\tParagraph 1.");

    expect(field.getFieldCode()).toEqual(" AUTONUM ");

    field = builder.insertField(aw.Fields.FieldType.FieldAutoNum, true).asFieldAutoNum();
    builder.writeln("\tParagraph 2.");

    // The separator character, which appears in the field result immediately after the number,is a full stop by default.
    // If we leave this property null, our second AUTONUM field will display "2." in the document.
    expect(field.separatorCharacter).toBe(null);

    // We can set this property to apply the first character of its string as the new separator character.
    // In this case, our AUTONUM field will now display "2:".
    field.separatorCharacter = ":";

    expect(field.getFieldCode()).toEqual(" AUTONUM  \\s :");

    doc.save(base.artifactsDir + "Field.AUTONUM.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.AUTONUM.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldAutoNum, " AUTONUM ", '', doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldAutoNum, " AUTONUM  \\s :", '', doc.range.fields.at(1));
  });


  //ExStart
  //ExFor:FieldAutoNumLgl
  //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
  //ExFor:FieldAutoNumLgl.SeparatorCharacter
  //ExSummary:Shows how to organize a document using AUTONUMLGL fields.
  test('FieldAutoNumLgl', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    const fillerText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
      "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

    // AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
    // These fields maintain a separate count for each heading level,
    // and each field also displays the AUTONUMLGL field counts for all heading levels below its own. 
    // Changing the count for any heading level resets the counts for all levels above that level to 1.
    // This allows us to organize our document in the form of an outline list.
    // This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
    insertNumberedClause(builder, "\tHeading 1", fillerText, aw.StyleIdentifier.Heading1);

    // This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
    insertNumberedClause(builder, "\tHeading 2", fillerText, aw.StyleIdentifier.Heading1);

    // This is the first AUTONUMLGL field at a heading level of 2,
    // and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
    insertNumberedClause(builder, "\tHeading 3", fillerText, aw.StyleIdentifier.Heading2);

    // This is the first AUTONUMLGL field at a heading level of 3. 
    // Working in the same way as the field above, it will display "2.1.1.".
    insertNumberedClause(builder, "\tHeading 4", fillerText, aw.StyleIdentifier.Heading3);

    // This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
    insertNumberedClause(builder, "\tHeading 5", fillerText, aw.StyleIdentifier.Heading2);

    // Incrementing the AUTONUMLGL count for a heading level below this one
    // has reset the count for this level so that this field will display "2.2.1.".
    insertNumberedClause(builder, "\tHeading 6", fillerText, aw.StyleIdentifier.Heading3);

    for (let field of Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldAutoNumLegal).map(node => node.asFieldAutoNumLgl()))
    {
      // The separator character, which appears in the field result immediately after the number,
      // is a full stop by default. If we leave this property null,
      // our last AUTONUMLGL field will display "2.2.1." in the document.
      expect(field.separatorCharacter).toBeNull();

      // Setting a custom separator character and removing the trailing period
      // will change that field's appearance from "2.2.1." to "2:2:1".
      // We will apply this to all the fields that we have created.
      field.separatorCharacter = ":";
      field.removeTrailingPeriod = true;
      expect(field.getFieldCode()).toEqual(" AUTONUMLGL  \\s : \\e");
    }

    doc.save(base.artifactsDir + "Field.AUTONUMLGL.docx");
    testFieldAutoNumLgl(doc); //ExSkip
  });


  /// <summary>
  /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
  /// </summary>
  function insertNumberedClause(builder, heading, contents, headingStyle) {
    builder.insertField(aw.Fields.FieldType.FieldAutoNumLegal, true);
    builder.currentParagraph.paragraphFormat.styleIdentifier = headingStyle;
    builder.writeln(heading);

    // This text will belong to the auto num legal field above it.
    // It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
    builder.currentParagraph.paragraphFormat.styleIdentifier = aw.StyleIdentifier.BodyText;
    builder.writeln(contents);
  }
  //ExEnd

  function testFieldAutoNumLgl(doc) {
    doc = DocumentHelper.saveOpen(doc);

    for (let field of Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldAutoNumLegal).map(node => node.asFieldAutoNumLgl()))
    {
      TestUtil.verifyField(aw.Fields.FieldType.FieldAutoNumLegal, " AUTONUMLGL  \\s : \\e", '', field);

      expect(field.separatorCharacter).toEqual(":");
      expect(field.removeTrailingPeriod).toEqual(true);
    }
  }

  test('FieldAutoNumOut', () => {
    //ExStart
    //ExFor:FieldAutoNumOut
    //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
    // Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
    // which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
    // This allows us to automatically number items like a numbered list.
    // LISTNUM fields are a newer alternative to AUTONUMOUT fields.
    // This field will display "1.".
    builder.insertField(aw.Fields.FieldType.FieldAutoNumOutline, true);
    builder.writeln("\tParagraph 1.");

    // This field will display "2.".
    builder.insertField(aw.Fields.FieldType.FieldAutoNumOutline, true);
    builder.writeln("\tParagraph 2.");

    for (let field of Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldAutoNumOutline))
      expect(field.getFieldCode()).toEqual(" AUTONUMOUT ");

    doc.save(base.artifactsDir + "Field.AUTONUMOUT.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.AUTONUMOUT.docx");

    for (let field of doc.range.fields)
      TestUtil.verifyField(aw.Fields.FieldType.FieldAutoNumOutline, " AUTONUMOUT ", '', field);
  });


  test('FieldAutoText', () => {
    //ExStart
    //ExFor:FieldAutoText
    //ExFor:FieldAutoText.entryName
    //ExFor:FieldOptions.builtInTemplatesPaths
    //ExFor:FieldGlossary
    //ExFor:FieldGlossary.entryName
    //ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields. 
    let doc = new aw.Document();

    // Create a glossary document and add an AutoText building block to it.
    doc.glossaryDocument = new aw.BuildingBlocks.GlossaryDocument();
    let buildingBlock = new aw.BuildingBlocks.BuildingBlock(doc.glossaryDocument);
    buildingBlock.name = "MyBlock";
    buildingBlock.gallery = aw.BuildingBlocks.BuildingBlockGallery.AutoText;
    buildingBlock.category = "General";
    buildingBlock.description = "MyBlock description";
    buildingBlock.behavior = aw.BuildingBlocks.BuildingBlockBehavior.Paragraph;
    doc.glossaryDocument.appendChild(buildingBlock);

    // Create a source and add it as text to our building block.
    let buildingBlockSource = new aw.Document();
    let buildingBlockSourceBuilder = new aw.DocumentBuilder(buildingBlockSource);
    buildingBlockSourceBuilder.writeln("Hello World!");

    let buildingBlockContent = doc.glossaryDocument.importNode(buildingBlockSource.firstSection, true);
    buildingBlock.appendChild(buildingBlockContent);

    // Set a file which contains parts that our document, or its attached template may not contain.
    doc.fieldOptions.builtInTemplatesPaths = new Array(base.myDir + "Busniess brochure.dotx");

    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways to use fields to display the contents of our building block.
    // 1 -  Using an AUTOTEXT field:
    let fieldAutoText = builder.insertField(aw.Fields.FieldType.FieldAutoText, true).asFieldAutoText();
    fieldAutoText.entryName = "MyBlock";

    expect(fieldAutoText.getFieldCode()).toEqual(" AUTOTEXT  MyBlock");

    // 2 -  Using a GLOSSARY field:
    let fieldGlossary = builder.insertField(aw.Fields.FieldType.FieldGlossary, true).asFieldGlossary();
    fieldGlossary.entryName = "MyBlock";

    expect(fieldGlossary.getFieldCode()).toEqual(" GLOSSARY  MyBlock");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.AUTOTEXT.GLOSSARY.dotx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.AUTOTEXT.GLOSSARY.dotx");

    expect(doc.fieldOptions.builtInTemplatesPaths.length).toEqual(0);

    fieldAutoText = doc.range.fields.at(0).asFieldAutoText();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAutoText, " AUTOTEXT  MyBlock", "Hello World!\r", fieldAutoText);
    expect(fieldAutoText.entryName).toEqual("MyBlock");

    fieldGlossary = doc.range.fields.at(1).asFieldGlossary();

    TestUtil.verifyField(aw.Fields.FieldType.FieldGlossary, " GLOSSARY  MyBlock", "Hello World!\r", fieldGlossary);
    expect(fieldGlossary.entryName).toEqual("MyBlock");
  });


  //ExStart
  //ExFor:FieldAutoTextList
  //ExFor:FieldAutoTextList.EntryName
  //ExFor:FieldAutoTextList.ListStyle
  //ExFor:FieldAutoTextList.ScreenTip
  //ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
  test('FieldAutoTextList', () => {
    let doc = new aw.Document();

    // Create a glossary document and populate it with auto text entries.
    doc.glossaryDocument = new aw.BuildingBlocks.GlossaryDocument();
    appendAutoTextEntry(doc.glossaryDocument, "AutoText 1", "Contents of AutoText 1");
    appendAutoTextEntry(doc.glossaryDocument, "AutoText 2", "Contents of AutoText 2");
    appendAutoTextEntry(doc.glossaryDocument, "AutoText 3", "Contents of AutoText 3");

    let builder = new aw.DocumentBuilder(doc);

    // Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
    // Set the text to prompt the user to right-click this field to select an AutoText building block,
    // whose contents the field will display.
    let field = builder.insertField(aw.Fields.FieldType.FieldAutoTextList, true).asFieldAutoTextList();
    field.entryName = "Right click here to select an AutoText block";
    field.listStyle = "Heading 1";
    field.screenTip = "Hover tip text for AutoTextList goes here";

    expect(field.getFieldCode()).toEqual(" AUTOTEXTLIST  \"Right click here to select an AutoText block\" " +
                            "\\s \"Heading 1\" " +
                            "\\t \"Hover tip text for AutoTextList goes here\"");

    doc.save(base.artifactsDir + "Field.AUTOTEXTLIST.dotx");
    testFieldAutoTextList(doc); //ExSkip
  });


  /// <summary>
  /// Create an AutoText-type building block and add it to a glossary document.
  /// </summary>
  function appendAutoTextEntry(glossaryDoc, name, contents) {
    let buildingBlock = new aw.BuildingBlocks.BuildingBlock(glossaryDoc);
    buildingBlock.name = name;
    buildingBlock.gallery = aw.BuildingBlocks.BuildingBlockGallery.AutoText;
    buildingBlock.category = "General";
    buildingBlock.behavior = aw.BuildingBlocks.BuildingBlockBehavior.Paragraph;

    let section = new aw.Section(glossaryDoc);
    section.appendChild(new aw.Body(glossaryDoc));
    section.body.appendParagraph(contents);
    buildingBlock.appendChild(section);

    glossaryDoc.appendChild(buildingBlock);
  }
    //ExEnd

  function testFieldAutoTextList(doc) {
    doc = DocumentHelper.saveOpen(doc);

    expect(doc.glossaryDocument.count).toEqual(3);
    expect(doc.glossaryDocument.buildingBlocks.at(0).name).toEqual("AutoText 1");
    expect(doc.glossaryDocument.buildingBlocks.at(0).getText().trim()).toEqual("Contents of AutoText 1");
    expect(doc.glossaryDocument.buildingBlocks.at(1).name).toEqual("AutoText 2");
    expect(doc.glossaryDocument.buildingBlocks.at(1).getText().trim()).toEqual("Contents of AutoText 2");
    expect(doc.glossaryDocument.buildingBlocks.at(2).name).toEqual("AutoText 3");
    expect(doc.glossaryDocument.buildingBlocks.at(2).getText().trim()).toEqual("Contents of AutoText 3");

    let field = doc.range.fields.at(0).asFieldAutoTextList();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAutoTextList,
      " AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
      '', field);
    expect(field.entryName).toEqual("Right click here to select an AutoText block");
    expect(field.listStyle).toEqual("Heading 1");
    expect(field.screenTip).toEqual("Hover tip text for AutoTextList goes here");
  }

  test.skip('FieldGreetingLine: DataTable', () => {
    //ExStart
    //ExFor:FieldGreetingLine
    //ExFor:FieldGreetingLine.alternateText
    //ExFor:FieldGreetingLine.getFieldNames
    //ExFor:FieldGreetingLine.languageId
    //ExFor:FieldGreetingLine.nameFormat
    //ExSummary:Shows how to insert a GREETINGLINE field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a generic greeting using a GREETINGLINE field, and some text after it.
    let field = builder.insertField(aw.Fields.FieldType.FieldGreetingLine, true).asFieldGreetingLine();
    builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

    // A GREETINGLINE field accepts values from a data source during a mail merge, like a MERGEFIELD.
    // It can also format how the source's data is written in its place once the mail merge is complete.
    // The field names collection corresponds to the columns from the data source
    // that the field will take values from.
    expect(field.getFieldNames().Length).toEqual(0);

    // To populate that array, we need to specify a format for our greeting line.
    field.nameFormat = "<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ";

    // Now, our field will accept values from these two columns in the data source.
    expect(field.getFieldNames()[0]).toEqual("Courtesy Title");
    expect(field.getFieldNames()[1]).toEqual("Last Name");
    expect(field.getFieldNames().Length).toEqual(2);

    // This string will cover any cases where the data table data is invalid
    // by substituting the malformed name with a string.
    field.alternateText = "Sir or Madam";

    // Set a locale to format the result.
    field.languageId = new CultureInfo("en-US").LCID.toString();

    expect(field.getFieldCode()).toEqual(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033");

    // Create a data table with columns whose names match elements
    // from the field's field names collection, and then carry out the mail merge.
    let table = new DataTable("Employees");
    table.columns.add("Courtesy Title");
    table.columns.add("First Name");
    table.columns.add("Last Name");
    table.rows.add("Mr.", "John", "Doe");
    table.rows.add("Mrs.", "Jane", "Cardholder");

    // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text.
    table.rows.add("", "No", "Name");

    doc.mailMerge.execute(table);

    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getText().trim()).toEqual("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                            "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                            "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!");
    //ExEnd
  });


  test('FieldListNum', () => {
    //ExStart
    //ExFor:FieldListNum
    //ExFor:FieldListNum.hasListName
    //ExFor:FieldListNum.listLevel
    //ExFor:FieldListNum.listName
    //ExFor:FieldListNum.startingNumber
    //ExSummary:Shows how to number paragraphs with LISTNUM fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // LISTNUM fields display a number that increments at each LISTNUM field.
    // These fields also have a variety of options that allow us to use them to emulate numbered lists.
    let field = builder.insertField(aw.Fields.FieldType.FieldListNum, true).asFieldListNum();

    // Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
    // This field will display "0)".
    field.startingNumber = "0";
    builder.writeln("Paragraph 1");

    expect(field.getFieldCode()).toEqual(" LISTNUM  \\s 0");

    // LISTNUM fields maintain separate counts for each list level. 
    // Inserting a LISTNUM field in the same paragraph as another LISTNUM field
    // increases the list level instead of the count.
    // The next field will continue the count we started above and display a value of "1" at list level 1.
    builder.insertField(aw.Fields.FieldType.FieldListNum, true);

    // This field will start a count at list level 2. It will display a value of "1".
    builder.insertField(aw.Fields.FieldType.FieldListNum, true);

    // This field will start a count at list level 3. It will display a value of "1".
    // Different list levels have different formatting,
    // so these fields combined will display a value of "1)a)i)".
    builder.insertField(aw.Fields.FieldType.FieldListNum, true);
    builder.writeln("Paragraph 2");

    // The next LISTNUM field that we insert will continue the count at the list level
    // that the previous LISTNUM field was on.
    // We can use the "ListLevel" property to jump to a different list level.
    // If this LISTNUM field stayed on list level 3, it would display "ii)",
    // but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
    field = builder.insertField(aw.Fields.FieldType.FieldListNum, true).asFieldListNum();
    field.listLevel = "2";
    builder.writeln("Paragraph 3");

    expect(field.getFieldCode()).toEqual(" LISTNUM  \\l 2");

    // We can set the ListName property to get the field to emulate a different AUTONUM field type.
    // "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
    // and "LegalDefault" emulates AUTONUMLGL fields.
    // The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
    field = builder.insertField(aw.Fields.FieldType.FieldListNum, true).asFieldListNum();
    field.startingNumber = "1";
    field.listName = "OutlineDefault";
    builder.writeln("Paragraph 4");

    expect(field.hasListName).toEqual(true);
    expect(field.getFieldCode()).toEqual(" LISTNUM  OutlineDefault \\s 1");

    // The ListName does not carry over from the previous field, so we will need to set it for each new field.
    // This field continues the count with the different list name and displays "II.".
    field = builder.insertField(aw.Fields.FieldType.FieldListNum, true).asFieldListNum();
    field.listName = "OutlineDefault";
    builder.writeln("Paragraph 5");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.LISTNUM.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.LISTNUM.docx");

    expect(doc.range.fields.count).toEqual(7);

    field = doc.range.fields.at(0).asFieldListNum();

    TestUtil.verifyField(aw.Fields.FieldType.FieldListNum, " LISTNUM  \\s 0", '', field);
    expect(field.startingNumber).toEqual("0");
    expect(field.listLevel).toBe(null);
    expect(field.hasListName).toEqual(false);
    expect(field.listName).toBe(null);

    for (let i = 1; i < 4; i++)
    {
      field = doc.range.fields.at(i).asFieldListNum();

      TestUtil.verifyField(aw.Fields.FieldType.FieldListNum, " LISTNUM ", '', field);
      expect(field.startingNumber).toBe(null);
      expect(field.listLevel).toBe(null);
      expect(field.hasListName).toEqual(false);
      expect(field.listName).toBe(null);
    }

    field = doc.range.fields.at(4).asFieldListNum();

    TestUtil.verifyField(aw.Fields.FieldType.FieldListNum, " LISTNUM  \\l 2", '', field);
    expect(field.startingNumber).toBe(null);
    expect(field.listLevel).toEqual("2");
    expect(field.hasListName).toEqual(false);
    expect(field.listName).toBe(null);

    field = doc.range.fields.at(5).asFieldListNum();

    TestUtil.verifyField(aw.Fields.FieldType.FieldListNum, " LISTNUM  OutlineDefault \\s 1", '', field);
    expect(field.startingNumber).toEqual("1");
    expect(field.listLevel).toBe(null);
    expect(field.hasListName).toEqual(true);
    expect(field.listName).toEqual("OutlineDefault");
  });


  test.skip('MergeField: DataTable', () => {
    //ExStart
    //ExFor:FieldMergeField
    //ExFor:FieldMergeField.fieldName
    //ExFor:FieldMergeField.fieldNameNoPrefix
    //ExFor:FieldMergeField.isMapped
    //ExFor:FieldMergeField.isVerticalFormatting
    //ExFor:FieldMergeField.textAfter
    //ExFor:FieldMergeField.textBefore
    //ExFor:FieldMergeField.type
    //ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a data table to be used as a mail merge data source.
    let table = new DataTable("Employees");
    table.columns.add("Courtesy Title");
    table.columns.add("First Name");
    table.columns.add("Last Name");
    table.rows.add("Mr.", "John", "Doe");
    table.rows.add("Mrs.", "Jane", "Cardholder");

    // Insert a MERGEFIELD with a FieldName property set to the name of a column in the data source.
    let fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Courtesy Title";
    fieldMergeField.isMapped = true;
    fieldMergeField.isVerticalFormatting = false;

    // We can apply text before and after the value that this field accepts when the merge takes place.
    fieldMergeField.textBefore = "Dear ";
    fieldMergeField.textAfter = " ";

    expect(fieldMergeField.getFieldCode()).toEqual(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"");
    expect(fieldMergeField.type).toEqual(aw.Fields.FieldType.FieldMergeField);

    // Insert another MERGEFIELD for a different column in the data source.
    fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Last Name";
    fieldMergeField.textAfter = ":";

    doc.updateFields();
    doc.mailMerge.execute(table);

    expect(doc.getText().trim()).toEqual("Dear Mr. Doe:\u000cDear Mrs. Cardholder:");
    //ExEnd

    expect(doc.range.fields.count).toEqual(0);
  });


  //ExStart
  //ExFor:FieldToc
  //ExFor:FieldToc.BookmarkName
  //ExFor:FieldToc.CustomStyles
  //ExFor:FieldToc.EntrySeparator
  //ExFor:FieldToc.HeadingLevelRange
  //ExFor:FieldToc.HideInWebLayout
  //ExFor:FieldToc.InsertHyperlinks
  //ExFor:FieldToc.PageNumberOmittingLevelRange
  //ExFor:FieldToc.PreserveLineBreaks
  //ExFor:FieldToc.PreserveTabs
  //ExFor:FieldToc.UpdatePageNumbers
  //ExFor:FieldToc.UseParagraphOutlineLevel
  //ExFor:FieldOptions.CustomTocStyleSeparator
  //ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
  test('FieldToc', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("MyBookmark");

    // Insert a TOC field, which will compile all headings into a table of contents.
    // For each heading, this field will create a line with the text in that heading style to the left,
    // and the page the heading appears on to the right.
    let field = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();

    // Use the BookmarkName property to only list headings
    // that appear within the bounds of a bookmark with the "MyBookmark" name.
    field.bookmarkName = "MyBookmark";

    // Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
    // We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
    field.customStyles = "Quote; 6; Intense Quote; 7";

    // By default, Styles/TOC levels are separated in the CustomStyles property by a comma,
    // but we can set a custom delimiter in this property.
    doc.fieldOptions.customTocStyleSeparator = ";";

    // Configure the field to exclude any headings that have TOC levels outside of this range.
    field.headingLevelRange = "1-3";

    // The TOC will not display the page numbers of headings whose TOC levels are within this range.
    field.pageNumberOmittingLevelRange = "2-5";

    // Set a custom string that will separate every heading from its page number. 
    field.entrySeparator = "-";
    field.insertHyperlinks = true;
    field.hideInWebLayout = false;
    field.preserveLineBreaks = true;
    field.preserveTabs = true;
    field.useParagraphOutlineLevel = false;

    insertNewPageWithHeading(builder, "First entry", "Heading 1");
    builder.writeln("Paragraph text.");
    insertNewPageWithHeading(builder, "Second entry", "Heading 1");
    insertNewPageWithHeading(builder, "Third entry", "Quote");
    insertNewPageWithHeading(builder, "Fourth entry", "Intense Quote");

    // These two headings will have the page numbers omitted because they are within the "2-5" range.
    insertNewPageWithHeading(builder, "Fifth entry", "Heading 2");
    insertNewPageWithHeading(builder, "Sixth entry", "Heading 3");

    // This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
    insertNewPageWithHeading(builder, "Seventh entry", "Heading 4");

    builder.endBookmark("MyBookmark");
    builder.writeln("Paragraph text.");

    // This entry does not appear because it is outside the bookmark specified by the TOC.
    insertNewPageWithHeading(builder, "Eighth entry", "Heading 1");

    expect(field.getFieldCode()).toEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w");

    field.updatePageNumbers();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TOC.docx");
    testFieldToc(doc); //ExSkip
  });


  /// <summary>
  /// Start a new page and insert a paragraph of a specified style.
  /// </summary>
  function insertNewPageWithHeading(builder, captionText, styleName) {
    builder.insertBreak(aw.BreakType.PageBreak);
    let originalStyle = builder.paragraphFormat.styleName;
    builder.paragraphFormat.style = builder.document.styles.at(styleName);
    builder.writeln(captionText);
    builder.paragraphFormat.style = builder.document.styles.at(originalStyle);
  }
  //ExEnd

  function testFieldToc(doc) {
    doc = DocumentHelper.saveOpen(doc);
    let field = doc.range.fields.at(0).asFieldToc();

    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.customStyles).toEqual("Quote; 6; Intense Quote; 7");
    expect(field.entrySeparator).toEqual("-");
    expect(field.headingLevelRange).toEqual("1-3");
    expect(field.pageNumberOmittingLevelRange).toEqual("2-5");
    expect(field.hideInWebLayout).toEqual(false);
    expect(field.insertHyperlinks).toEqual(true);
    expect(field.preserveLineBreaks).toEqual(true);
    expect(field.preserveTabs).toEqual(true);
    expect(field.updatePageNumbers()).toEqual(true);
    expect(field.useParagraphOutlineLevel).toEqual(false);
    expect(field.getFieldCode()).toEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w");
    expect(field.result).toEqual("\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" +
                            "\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r");
  }

  //ExStart
    //ExFor:FieldToc.EntryIdentifier
    //ExFor:FieldToc.EntryLevelRange
    //ExFor:FieldTC
    //ExFor:FieldTC.OmitPageNumber
    //ExFor:FieldTC.Text
    //ExFor:FieldTC.TypeIdentifier
    //ExFor:FieldTC.EntryLevel
    //ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
  test('FieldTocEntryIdentifier', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a TOC field, which will compile all TC fields into a table of contents.
    let fieldToc = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();

    // Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
    fieldToc.entryIdentifier = "A";
    fieldToc.entryLevelRange = "1-3";

    expect(fieldToc.getFieldCode()).toEqual(" TOC  \\f A \\l 1-3");

    // These two entries will appear in the table.
    builder.insertBreak(aw.BreakType.PageBreak);
    insertTocEntry(builder, "TC field 1", "A", "1");
    insertTocEntry(builder, "TC field 2", "A", "2");

    expect(doc.range.fields.at(1).getFieldCode()).toEqual(" TC  \"TC field 1\" \\n \\f A \\l 1");

    // This entry will be omitted from the table because it has a different type from "A".
    insertTocEntry(builder, "TC field 3", "B", "1");

    // This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
    insertTocEntry(builder, "TC field 4", "A", "5");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TC.docx");
    testFieldTocEntryIdentifier(doc); //ExSkip
  });


  /// <summary>
  /// Use a document builder to insert a TC field.
  /// </summary>
  function insertTocEntry(builder, text, typeIdentifier, entryLevel) {
    let fieldTc = builder.insertField(aw.Fields.FieldType.FieldTOCEntry, true).asFieldTC();
    fieldTc.omitPageNumber = true;
    fieldTc.text = text;
    fieldTc.typeIdentifier = typeIdentifier;
    fieldTc.entryLevel = entryLevel;
  }
  //ExEnd

  function testFieldTocEntryIdentifier(doc) {
    doc = DocumentHelper.saveOpen(doc);
    let fieldToc = doc.range.fields.at(0).asFieldToc();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOC, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", fieldToc);
    expect(fieldToc.entryIdentifier).toEqual("A");
    expect(fieldToc.entryLevelRange).toEqual("1-3");

    let fieldTc = doc.range.fields.at(1).asFieldTC();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOCEntry, " TC  \"TC field 1\" \\n \\f A \\l 1", '', fieldTc);
    expect(fieldTc.omitPageNumber).toEqual(true);
    expect(fieldTc.text).toEqual("TC field 1");
    expect(fieldTc.typeIdentifier).toEqual("A");
    expect(fieldTc.entryLevel).toEqual("1");

    fieldTc = doc.range.fields.at(2).asFieldTC();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOCEntry, " TC  \"TC field 2\" \\n \\f A \\l 2", '', fieldTc);
    expect(fieldTc.omitPageNumber).toEqual(true);
    expect(fieldTc.text).toEqual("TC field 2");
    expect(fieldTc.typeIdentifier).toEqual("A");
    expect(fieldTc.entryLevel).toEqual("2");

    fieldTc = doc.range.fields.at(3).asFieldTC();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOCEntry, " TC  \"TC field 3\" \\n \\f B \\l 1", '', fieldTc);
    expect(fieldTc.omitPageNumber).toEqual(true);
    expect(fieldTc.text).toEqual("TC field 3");
    expect(fieldTc.typeIdentifier).toEqual("B");
    expect(fieldTc.entryLevel).toEqual("1");

    fieldTc = doc.range.fields.at(4).asFieldTC();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOCEntry, " TC  \"TC field 4\" \\n \\f A \\l 5", '', fieldTc);
    expect(fieldTc.omitPageNumber).toEqual(true);
    expect(fieldTc.text).toEqual("TC field 4");
    expect(fieldTc.typeIdentifier).toEqual("A");
    expect(fieldTc.entryLevel).toEqual("5");
  }

  test('TocSeqPrefix', () => {
    //ExStart
    //ExFor:FieldToc
    //ExFor:FieldToc.tableOfFiguresLabel
    //ExFor:FieldToc.prefixedSequenceIdentifier
    //ExFor:FieldToc.sequenceSeparator
    //ExFor:FieldSeq
    //ExFor:FieldSeq.sequenceIdentifier
    //ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
    // Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
    let fieldToc = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();

    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Use the "TableOfFiguresLabel" property to name a main sequence for the TOC.
    // Now, this TOC will only create entries out of SEQ fields with their "SequenceIdentifier" set to "MySequence".
    fieldToc.tableOfFiguresLabel = "MySequence";

    // We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" property.
    // SEQ fields from this prefix sequence will not create TOC entries. 
    // Every TOC entry created from a main sequence SEQ field will now also display the count that
    // the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
    fieldToc.prefixedSequenceIdentifier = "PrefixSequence";

    // Each TOC entry will display the prefix sequence count immediately to the left
    // of the page number that the main sequence SEQ field appears on.
    // We can specify a custom separator that will appear between these two numbers.
    fieldToc.sequenceSeparator = ">";

    expect(fieldToc.getFieldCode()).toEqual(" TOC  \\c MySequence \\s PrefixSequence \\d >");

    builder.insertBreak(aw.BreakType.PageBreak);

    // There are two ways of using SEQ fields to populate this TOC.
    // 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
    // This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
    // Since this field does not belong to the main sequence identified
    // by the "TableOfFiguresLabel" property of the TOC, it will not appear as an entry.
    let fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "PrefixSequence";
    builder.insertParagraph();

    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  PrefixSequence");

    // 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
    // This SEQ field will create an entry in the TOC.
    // The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
    // This entry will also display the count that the prefix sequence is currently at,
    // separated from the page number by the value in the TOC's SeqenceSeparator property.
    // The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
    // and the separator is ">", so entry will display "1>2".
    builder.write("First TOC entry, MySequence #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";

    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  MySequence");

    // Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
    // The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
    // so the TOC entry will display "2>3" at its page count.
    builder.insertBreak(aw.BreakType.PageBreak);
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "PrefixSequence";
    builder.insertParagraph();
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    builder.write("Second TOC entry, MySequence #");
    fieldSeq.sequenceIdentifier = "MySequence";

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TOC.SEQ.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.TOC.SEQ.docx");

    expect(doc.range.fields.count).toEqual(9);

    fieldToc = doc.range.fields.at(0).asFieldToc();
    console.log(fieldToc.displayResult);
    TestUtil.verifyField(aw.Fields.FieldType.FieldTOC, " TOC  \\c MySequence \\s PrefixSequence \\d >",
      "First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2" +
      "Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r", 
      fieldToc);
    expect(fieldToc.tableOfFiguresLabel).toEqual("MySequence");
    expect(fieldToc.prefixedSequenceIdentifier).toEqual("PrefixSequence");
    expect(fieldToc.sequenceSeparator).toEqual(">");

    fieldSeq = doc.range.fields.at(1).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");

    // Byproduct field created by Aspose.words
    let fieldPageRef = doc.range.fields.at(2).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");
    expect(fieldPageRef.bookmarkName).toEqual("_Toc256000000");

    fieldSeq = doc.range.fields.at(3).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ PrefixSequence _Toc256000001 \\* ARABIC ", "2", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");

    fieldPageRef = doc.range.fields.at(4).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF _Toc256000001 \\h ", "3", fieldPageRef);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");
    expect(fieldPageRef.bookmarkName).toEqual("_Toc256000001");

    fieldSeq = doc.range.fields.at(5).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  PrefixSequence", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");

    fieldSeq = doc.range.fields.at(6).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(7).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  PrefixSequence", "2", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("PrefixSequence");

    fieldSeq = doc.range.fields.at(8).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "2", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");
  });


  test('TocSeqNumbering', () => {
    //ExStart
    //ExFor:FieldSeq
    //ExFor:FieldSeq.insertNextNumber
    //ExFor:FieldSeq.resetHeadingLevel
    //ExFor:FieldSeq.resetNumber
    //ExFor:FieldSeq.sequenceIdentifier
    //ExSummary:Shows create numbering using SEQ fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field that will display the current count value of "MySequence",
    // after using the "ResetNumber" property to set it to 100.
    builder.write("#");
    let fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    fieldSeq.resetNumber = "100";
    fieldSeq.update();

    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  MySequence \\r 100");
    expect(fieldSeq.result).toEqual("100");

    // Display the next number in this sequence with another SEQ field.
    builder.write(", #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    fieldSeq.update();

    expect(fieldSeq.result).toEqual("101");

    // Insert a level 1 heading.
    builder.insertBreak(aw.BreakType.ParagraphBreak);
    builder.paragraphFormat.style = doc.styles.at("Heading 1");
    builder.writeln("This level 1 heading will reset MySequence to 1");
    builder.paragraphFormat.style = doc.styles.at("Normal");

    // Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
    builder.write("\n#");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    fieldSeq.resetHeadingLevel = "1";
    fieldSeq.update();

    // The above heading is a level 1 heading, so the count for this sequence is reset to 1.
    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  MySequence \\s 1");
    expect(fieldSeq.result).toEqual("1");

    // Move to the next number of this sequence.
    builder.write(", #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    fieldSeq.insertNextNumber = true;
    fieldSeq.update();

    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  MySequence \\n");
    expect(fieldSeq.result).toEqual("2");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.SEQ.ResetNumbering.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SEQ.ResetNumbering.docx");

    expect(doc.range.fields.count).toEqual(4);

    fieldSeq = doc.range.fields.at(0).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence \\r 100", "100", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(1).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "101", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(2).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence \\s 1", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(3).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence \\n", "2", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");
  });


  test('TocSeqBookmark', () => {
    //ExStart
    //ExFor:FieldSeq
    //ExFor:FieldSeq.bookmarkName
    //ExSummary:Shows how to combine table of contents and sequence fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
    // Each entry contains the paragraph that contains the SEQ field,
    // and the number of the page that the field appears on.
    let fieldToc = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();

    // Configure this TOC field to have a SequenceIdentifier property with a value of "MySequence".
    fieldToc.tableOfFiguresLabel = "MySequence";

    // Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
    // named "TOCBookmark".
    fieldToc.bookmarkName = "TOCBookmark";
    builder.insertBreak(aw.BreakType.PageBreak);

    expect(fieldToc.getFieldCode()).toEqual(" TOC  \\c MySequence \\b TOCBookmark");

    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field that has a sequence identifier that matches the TOC's
    // TableOfFiguresLabel property. This field will not create an entry in the TOC since it is outside
    // the bookmark's bounds designated by "BookmarkName".
    builder.write("MySequence #");
    let fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    builder.writeln(", will not show up in the TOC because it is outside of the bookmark.");

    builder.startBookmark("TOCBookmark");

    // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bookmark's bounds.
    // The paragraph that contains this field will show up in the TOC as an entry.
    builder.write("MySequence #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    builder.writeln(", will show up in the TOC next to the entry for the above caption.");

    // This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" property,
    // and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
    builder.write("MySequence #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "OtherSequence";
    builder.writeln(", will not show up in the TOC because it's from a different sequence identifier.");

    // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bounds of the bookmark.
    // This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
    // The SEQ field itself will not display the contents of that bookmark.
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    fieldSeq.bookmarkName = "SEQBookmark";
    expect(fieldSeq.getFieldCode()).toEqual(" SEQ  MySequence SEQBookmark");

    // Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.startBookmark("SEQBookmark");
    builder.write("MySequence #");
    fieldSeq = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    fieldSeq.sequenceIdentifier = "MySequence";
    builder.writeln(", text from inside SEQBookmark.");
    builder.endBookmark("SEQBookmark");

    builder.endBookmark("TOCBookmark");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.SEQ.bookmark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SEQ.bookmark.docx");

    expect(doc.range.fields.count).toEqual(8);

    fieldToc = doc.range.fields.at(0).asFieldToc();
    let pageRefIds = fieldToc.result.split(' ').filter(s => s.startsWith("_Toc"));

    expect(fieldToc.type).toEqual(aw.Fields.FieldType.FieldTOC);
    expect(fieldToc.tableOfFiguresLabel).toEqual("MySequence");
    TestUtil.verifyField(aw.Fields.FieldType.FieldTOC, " TOC  \\c MySequence \\b TOCBookmark",
      `MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF ${pageRefIds[0]} \\h \u00142\u0015\r` +
      `3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF ${pageRefIds[1]} \\h \u00142\u0015\r`, fieldToc);

    let fieldPageRef = doc.range.fields.at(1).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, ` PAGEREF ${pageRefIds[0]} \\h `, "2", fieldPageRef);
    expect(fieldPageRef.bookmarkName).toEqual(pageRefIds.at(0));

    fieldPageRef = doc.range.fields.at(2).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, ` PAGEREF ${pageRefIds[1]} \\h `, "2", fieldPageRef);
    expect(fieldPageRef.bookmarkName).toEqual(pageRefIds.at(1));

    fieldSeq = doc.range.fields.at(3).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(4).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "2", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");

    fieldSeq = doc.range.fields.at(5).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  OtherSequence", "1", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("OtherSequence");

    fieldSeq = doc.range.fields.at(6).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence SEQBookmark", "3", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");
    expect(fieldSeq.bookmarkName).toEqual("SEQBookmark");

    fieldSeq = doc.range.fields.at(7).asFieldSeq();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSequence, " SEQ  MySequence", "3", fieldSeq);
    expect(fieldSeq.sequenceIdentifier).toEqual("MySequence");
  });


  test('FieldCitation', () => {
    //ExStart
    //ExFor:FieldCitation
    //ExFor:FieldCitation.anotherSourceTag
    //ExFor:FieldCitation.formatLanguageId
    //ExFor:FieldCitation.pageNumber
    //ExFor:FieldCitation.prefix
    //ExFor:FieldCitation.sourceTag
    //ExFor:FieldCitation.suffix
    //ExFor:FieldCitation.suppressAuthor
    //ExFor:FieldCitation.suppressTitle
    //ExFor:FieldCitation.suppressYear
    //ExFor:FieldCitation.volumeNumber
    //ExFor:FieldBibliography
    //ExFor:FieldBibliography.formatLanguageId
    //ExFor:FieldBibliography.filterLanguageId
    //ExFor:FieldBibliography.sourceTag
    //ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
    // Open a document containing bibliographical sources that we can find in
    // Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
    let doc = new aw.Document(base.myDir + "Bibliography.docx");
    expect(doc.range.fields.count).toEqual(2);

    let builder = new aw.DocumentBuilder(doc);
    builder.write("Text to be cited with one source.");

    // Create a citation with just the page number and the author of the referenced book.
    let fieldCitation = builder.insertField(aw.Fields.FieldType.FieldCitation, true).asFieldCitation();

    // We refer to sources using their tag names.
    fieldCitation.sourceTag = "Book1";
    fieldCitation.pageNumber = "85";
    fieldCitation.suppressAuthor = false;
    fieldCitation.suppressTitle = true;
    fieldCitation.suppressYear = true;

    expect(fieldCitation.getFieldCode()).toEqual(" CITATION  Book1 \\p 85 \\t \\y");

    // Create a more detailed citation which cites two sources.
    builder.insertParagraph();
    builder.write("Text to be cited with two sources.");
    fieldCitation = builder.insertField(aw.Fields.FieldType.FieldCitation, true).asFieldCitation();
    fieldCitation.sourceTag = "Book1";
    fieldCitation.anotherSourceTag = "Book2";
    fieldCitation.formatLanguageId = "en-US";
    fieldCitation.pageNumber = "19";
    fieldCitation.prefix = "Prefix ";
    fieldCitation.suffix = " Suffix";
    fieldCitation.suppressAuthor = false;
    fieldCitation.suppressTitle = false;
    fieldCitation.suppressYear = false;
    fieldCitation.volumeNumber = "VII";

    expect(fieldCitation.getFieldCode()).toEqual(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII");

    // We can use a BIBLIOGRAPHY field to display all the sources within the document.
    builder.insertBreak(aw.BreakType.PageBreak);
    let fieldBibliography = builder.insertField(aw.Fields.FieldType.FieldBibliography, true).asFieldBibliography();
    fieldBibliography.formatLanguageId = "5129";
    fieldBibliography.filterLanguageId = "5129";
    fieldBibliography.sourceTag = "Book2";

    expect(fieldBibliography.getFieldCode()).toEqual(" BIBLIOGRAPHY  \\l 5129 \\f 5129 \\m Book2");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.CITATION.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.CITATION.docx");

    expect(doc.range.fields.count).toEqual(5);

    fieldCitation = doc.range.fields.at(0).asFieldCitation();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCitation, " CITATION  Book1 \\p 85 \\t \\y", "(Doe, p. 85)", fieldCitation);
    expect(fieldCitation.sourceTag).toEqual("Book1");
    expect(fieldCitation.pageNumber).toEqual("85");
    expect(fieldCitation.suppressAuthor).toEqual(false);
    expect(fieldCitation.suppressTitle).toEqual(true);
    expect(fieldCitation.suppressYear).toEqual(true);

    fieldCitation = doc.range.fields.at(1).asFieldCitation();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCitation, 
      " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", 
      "(Doe, 2018; Prefix Cardholder, 2018, VII:19 Suffix)", fieldCitation);
    expect(fieldCitation.sourceTag).toEqual("Book1");
    expect(fieldCitation.anotherSourceTag).toEqual("Book2");
    expect(fieldCitation.formatLanguageId).toEqual("en-US");
    expect(fieldCitation.prefix).toEqual("Prefix ");
    expect(fieldCitation.suffix).toEqual(" Suffix");
    expect(fieldCitation.pageNumber).toEqual("19");
    expect(fieldCitation.suppressAuthor).toEqual(false);
    expect(fieldCitation.suppressTitle).toEqual(false);
    expect(fieldCitation.suppressYear).toEqual(false);
    expect(fieldCitation.volumeNumber).toEqual("VII");

    fieldBibliography = doc.range.fields.at(2).asFieldBibliography();

    
    TestUtil.verifyField(aw.Fields.FieldType.FieldBibliography, " BIBLIOGRAPHY  \\l 5129 \\f 5129 \\m Book2",
      "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\r", fieldBibliography);
      
    expect(fieldBibliography.formatLanguageId).toEqual("5129");
    expect(fieldBibliography.filterLanguageId).toEqual("5129");

    fieldCitation = doc.range.fields.at(3).asFieldCitation();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCitation, " CITATION Book1 \\l 1033 ", " (Doe, 2018)", fieldCitation);
    expect(fieldCitation.sourceTag).toEqual("Book1");
    expect(fieldCitation.formatLanguageId).toEqual("1033");

    fieldBibliography = doc.range.fields.at(4).asFieldBibliography();

    TestUtil.verifyField(aw.Fields.FieldType.FieldBibliography, " BIBLIOGRAPHY ", 
      "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);

  });

/*
  //ExStart
  //ExFor:Bibliography.BibliographyStyle
  //ExFor:IBibliographyStylesProvider
  //ExFor:IBibliographyStylesProvider.GetStyle(String)
  //ExFor:FieldOptions.BibliographyStylesProvider
  //ExSummary:Shows how to override built-in styles or provide custom one.
  test('ChangeBibliographyStyles', () => {
    var oldCulture = Thread.currentThread.CurrentCulture; //ExSkip
    Thread.currentThread.CurrentCulture = new CultureInfo("en-nz", false); //ExSkip

    let doc = new aw.Document(base.myDir + "Bibliography.docx");

    // If the document already has a style you can change it with the following code:
    // doc.bibliography.bibliographyStyle = "Bibliography custom style.xsl";

    doc.fieldOptions.bibliographyStylesProvider = new BibliographyStylesProvider();
    doc.updateFields();

    doc.save(base.artifactsDir + "Field.ChangeBibliographyStyles.docx");

    Thread.currentThread.CurrentCulture = oldCulture; //ExSkip
  });


  public class BibliographyStylesProvider : IBibliographyStylesProvider
  {
    Stream aw.Fields.IBibliographyStylesProvider.getStyle(string styleFileName)
    {
      return File.OpenRead(base.myDir + "Bibliography custom style.xsl");
    }
  }
    //ExEnd
*/    

  test('FieldData', () => {
    //ExStart
    //ExFor:FieldData
    //ExSummary:Shows how to insert a DATA field into a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldData, true).asFieldData();
    expect(field.getFieldCode()).toEqual(" DATA ");
    //ExEnd

    TestUtil.verifyField(aw.Fields.FieldType.FieldData, " DATA ", '', DocumentHelper.saveOpen(doc).range.fields.at(0));
  });


  test('FieldInclude', () => {
    //ExStart
    //ExFor:FieldInclude
    //ExFor:FieldInclude.bookmarkName
    //ExFor:FieldInclude.lockFields
    //ExFor:FieldInclude.sourceFullName
    //ExFor:FieldInclude.textConverter
    //ExSummary:Shows how to create an INCLUDE field, and set its properties.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // We can use an INCLUDE field to import a portion of another document in the local file system.
    // The bookmark from the other document that we reference with this field contains this imported portion.
    let field = builder.insertField(aw.Fields.FieldType.FieldInclude, true).asFieldInclude();
    field.sourceFullName = base.myDir + "Bookmarks.docx";
    field.bookmarkName = "MyBookmark1";
    field.lockFields = false;
    field.textConverter = "Microsoft Word";

    expect(new RegExp(" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").test(field.getFieldCode())).toBeTruthy();

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INCLUDE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INCLUDE.docx");
    field = doc.range.fields.at(0).asFieldInclude();

    expect(field.type).toEqual(aw.Fields.FieldType.FieldInclude);
    expect(field.result).toEqual("First bookmark.");
    expect(new RegExp(" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").test(field.getFieldCode())).toBeTruthy();

    expect(field.sourceFullName).toEqual(base.myDir + "Bookmarks.docx");
    expect(field.bookmarkName).toEqual("MyBookmark1");
    expect(field.lockFields).toEqual(false);
    expect(field.textConverter).toEqual("Microsoft Word");
  });


  test('FieldIncludePicture', () => {
    //ExStart
    //ExFor:FieldIncludePicture
    //ExFor:FieldIncludePicture.graphicFilter
    //ExFor:FieldIncludePicture.isLinked
    //ExFor:FieldIncludePicture.resizeHorizontally
    //ExFor:FieldIncludePicture.resizeVertically
    //ExFor:FieldIncludePicture.sourceFullName
    //ExFor:FieldImport
    //ExFor:FieldImport.graphicFilter
    //ExFor:FieldImport.isLinked
    //ExFor:FieldImport.sourceFullName
    //ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two similar field types that we can use to display images linked from the local file system.
    // 1 -  The INCLUDEPICTURE field:
    let fieldIncludePicture = builder.insertField(aw.Fields.FieldType.FieldIncludePicture, true).asFieldIncludePicture();
    fieldIncludePicture.sourceFullName = base.imageDir + "Transparent background logo.png";

    expect(new RegExp(" INCLUDEPICTURE  .*").test(fieldIncludePicture.getFieldCode())).toBeTruthy();

    // Apply the PNG32.FLT filter.
    fieldIncludePicture.graphicFilter = "PNG32";
    fieldIncludePicture.isLinked = true;
    fieldIncludePicture.resizeHorizontally = true;
    fieldIncludePicture.resizeVertically = true;

    // 2 -  The IMPORT field:
    let fieldImport = builder.insertField(aw.Fields.FieldType.FieldImport, true).asFieldImport();
    fieldImport.sourceFullName = base.imageDir + "Transparent background logo.png";
    fieldImport.graphicFilter = "PNG32";
    fieldImport.isLinked = true;

    expect(new RegExp(" IMPORT  .* \\\\c PNG32 \\\\d").test(fieldImport.getFieldCode())).toBeTruthy();

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.IMPORT.INCLUDEPICTURE.docx");
    //ExEnd

    expect(fieldIncludePicture.sourceFullName).toEqual(base.imageDir + "Transparent background logo.png");
    expect(fieldIncludePicture.graphicFilter).toEqual("PNG32");
    expect(fieldIncludePicture.isLinked).toBeTruthy();
    expect(fieldIncludePicture.resizeHorizontally).toBeTruthy();
    expect(fieldIncludePicture.resizeVertically).toBeTruthy();

    expect(fieldImport.sourceFullName).toEqual(base.imageDir + "Transparent background logo.png");
    expect(fieldImport.graphicFilter).toEqual("PNG32");
    expect(fieldImport.isLinked).toBeTruthy();

    doc = new aw.Document(base.artifactsDir + "Field.IMPORT.INCLUDEPICTURE.docx");

    // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(2);

    let image = doc.getShape(0, true);

    expect(image.isImage).toBeTruthy();
    expect(image.imageData.imageBytes).toBeNull();
    expect(image.imageData.sourceFullName.replace("%20", " ")).toEqual(base.imageDir + "Transparent background logo.png");

    image = doc.getShape(1, true);

    expect(image.isImage).toBeTruthy();
    expect(image.imageData.imageBytes).toBeNull();
    expect(image.imageData.sourceFullName.replace("%20", " ")).toEqual(base.imageDir + "Transparent background logo.png");
  });


  /*//ExStart
    //ExFor:FieldIncludeText
    //ExFor:FieldIncludeText.BookmarkName
    //ExFor:FieldIncludeText.Encoding
    //ExFor:FieldIncludeText.LockFields
    //ExFor:FieldIncludeText.MimeType
    //ExFor:FieldIncludeText.NamespaceMappings
    //ExFor:FieldIncludeText.SourceFullName
    //ExFor:FieldIncludeText.TextConverter
    //ExFor:FieldIncludeText.XPath
    //ExFor:FieldIncludeText.XslTransformation
    //ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
  test('FieldIncludeText', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways to use INCLUDETEXT fields to display the contents of an XML file in the local file system.
    // 1 -  Perform an XSL transformation on an XML document:
    let fieldIncludeText = CreateFieldIncludeText(builder, base.myDir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
    fieldIncludeText.xslTransformation = base.myDir + "CD collection XSL transformation.xsl";

    builder.writeln();

    // 2 -  Use an XPath to take specific elements from an XML document:
    fieldIncludeText = CreateFieldIncludeText(builder, base.myDir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
    fieldIncludeText.namespaceMappings = "xmlns:n='myNamespace'";
    fieldIncludeText.xPath = "/catalog/cd/title";

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INCLUDETEXT.docx");
    TestFieldIncludeText(new aw.Document(base.artifactsDir + "Field.INCLUDETEXT.docx")); //ExSkip
  });


    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field with custom properties.
    /// </summary>
  public FieldIncludeText CreateFieldIncludeText(DocumentBuilder builder, string sourceFullName, bool lockFields, string mimeType, string textConverter, string encoding)
  {
    let fieldIncludeText = (FieldIncludeText)builder.insertField(aw.Fields.FieldType.FieldIncludeText, true);
    fieldIncludeText.sourceFullName = sourceFullName;
    fieldIncludeText.lockFields = lockFields;
    fieldIncludeText.mimeType = mimeType;
    fieldIncludeText.textConverter = textConverter;
    fieldIncludeText.encoding = encoding;

    return fieldIncludeText;
  }
    //ExEnd

  function TestFieldIncludeText(doc) {
    doc = DocumentHelper.saveOpen(doc);

    let fieldIncludeText = doc.range.fields.at(0).asFieldIncludeText();
    expect(fieldIncludeText.sourceFullName).toEqual(base.myDir + "CD collection data.xml");
    expect(fieldIncludeText.xslTransformation).toEqual(base.myDir + "CD collection XSL transformation.xsl");
    expect(fieldIncludeText.lockFields).toEqual(false);
    expect(fieldIncludeText.mimeType).toEqual("text/xml");
    expect(fieldIncludeText.textConverter).toEqual("XML");
    expect(fieldIncludeText.encoding).toEqual("ISO-8859-1");
    expect(fieldIncludeText.getFieldCode()).toEqual(" INCLUDETEXT  \"" + base.myDir.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" + 
                            base.myDir.replace("\\", "\\\\") + "CD collection XSL transformation.xsl\"");
    expect(fieldIncludeText.result.startsWith("My CD Collection")).toEqual(true);

    let cdCollectionData = new XmlDocument();
    cdCollectionData.loadXml(File.ReadAllText(base.myDir + "CD collection data.xml"));
    XmlNode catalogData = cdCollectionData.childNodes.at(0);

    let cdCollectionXslTransformation = new XmlDocument();
    cdCollectionXslTransformation.loadXml(File.ReadAllText(base.myDir + "CD collection XSL transformation.xsl"));

    let table = doc.firstSection.body.tables.at(0);

    let manager = new XmlNamespaceManager(cdCollectionXslTransformation.nameTable);
    manager.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");

    for (let i = 0; i < table.rows.count; i++)
      for (let j = 0; j < table.rows.at(i).count; j++)
      {
        if (i == 0)
        {
            // When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
          for (let k = 0; k < table.rows.count - 1; k++)
            expect(table.rows.at(i).cells.at(j).getText().Replace(aw.ControlChar.cell, '').ToLower()).toEqual(catalogData.childNodes.at(k).childNodes[j].name);

            // Also, make sure that the whole first row has the same color as the XSL transform.
          expect(manager)[0].Attributes.GetNamedItem("bgcolor").Value, ColorTranslator.ToHtml(table.rows.at(i).cells.at(j).cellFormat.shading.backgroundPatternColor).ToLower()).toEqual(cdCollectionXslTransformation.selectNodes("//xsl:stylesheet/xsl:template/html/body/table/tr");
        }
        else
        {
            // When on all other rows of the input document's table, ensure that cell contents match XML element Values.
          expect(table.rows.at(i).cells.at(j).getText().Replace(aw.ControlChar.cell, '')).toEqual(catalogData.childNodes.at(i - 1).childNodes[j].firstChild.value);
          expect(table.rows.at(i).cells.at(j).cellFormat.shading.backgroundPatternColor).toEqual(base.emptyColor);
        }

        Assert.AreEqual(
          double.parse(cdCollectionXslTransformation.selectNodes("//xsl:stylesheet/xsl:template/html/body/table", manager)[0].Attributes.GetNamedItem("border").Value) * 0.75, 
          table.firstRow.rowFormat.borders.bottom.lineWidth);
      }

    fieldIncludeText = (FieldIncludeText)doc.range.fields.at(1);
    expect(fieldIncludeText.sourceFullName).toEqual(base.myDir + "CD collection data.xml");
    expect(fieldIncludeText.xslTransformation).toBe(null);
    expect(fieldIncludeText.lockFields).toEqual(false);
    expect(fieldIncludeText.mimeType).toEqual("text/xml");
    expect(fieldIncludeText.textConverter).toEqual("XML");
    expect(fieldIncludeText.encoding).toEqual("ISO-8859-1");
    expect(fieldIncludeText.getFieldCode()).toEqual(" INCLUDETEXT  \"" + base.myDir.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title");

    string expectedFieldResult = "";
    for (let i = 0; i < catalogData.childNodes.count; i++)
    {
      expectedFieldResult += catalogData.childNodes.at(i).childNodes.at(0).childNodes[0].value;
    }

    expect(fieldIncludeText.result).toEqual(expectedFieldResult);
  }*/

  test('FieldHyperlink', () => {
    //ExStart
    //ExFor:FieldHyperlink
    //ExFor:FieldHyperlink.address
    //ExFor:FieldHyperlink.isImageMap
    //ExFor:FieldHyperlink.openInNewWindow
    //ExFor:FieldHyperlink.screenTip
    //ExFor:FieldHyperlink.subAddress
    //ExFor:FieldHyperlink.target
    //ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldHyperlink, true).asFieldHyperlink();

    // When we click this HYPERLINK field in Microsoft Word,
    // it will open the linked document and then place the cursor at the specified bookmark.
    field.address = base.myDir + "Bookmarks.docx";
    field.subAddress = "MyBookmark3";
    field.screenTip = "Open " + field.address + " on bookmark " + field.subAddress + " in a new window";

    builder.writeln();

    // When we click this HYPERLINK field in Microsoft Word,
    // it will open the linked document, and automatically scroll down to the specified iframe.
    field = builder.insertField(aw.Fields.FieldType.FieldHyperlink, true).asFieldHyperlink();
    field.address = base.myDir + "Iframes.html";
    field.screenTip = "Open " + field.address;
    field.target = "iframe_3";
    field.openInNewWindow = true;
    field.isImageMap = false;

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.HYPERLINK.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.HYPERLINK.docx");
    field = doc.range.fields.at(0).asFieldHyperlink();

    let myDir = base.myDir.replaceAll("\\", "\\\\");
    TestUtil.verifyField(aw.Fields.FieldType.FieldHyperlink, 
      " HYPERLINK \"" + myDir + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + myDir + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
      base.myDir + "Bookmarks.docx - MyBookmark3", field);
    expect(field.address).toEqual(base.myDir + "Bookmarks.docx");
    expect(field.subAddress).toEqual("MyBookmark3");
    expect(field.screenTip).toEqual("Open " + field.address + " on bookmark " + field.subAddress + " in a new window");

    field = doc.range.fields.at(1).asFieldHyperlink();

    TestUtil.verifyField(aw.Fields.FieldType.FieldHyperlink, " HYPERLINK \"" + myDir.replace(" ", "%20") + "Iframes.html\" \\o \"Open " + myDir + "Iframes.html\" \\t \"iframe_3\" ",
      base.myDir + "Iframes.html", field);

    expect(field.address).toEqual(base.myDir.replaceAll(" ","%20") + "Iframes.html");
    expect(field.screenTip).toEqual("Open " + base.myDir + "Iframes.html");
    expect(field.target).toEqual("iframe_3");
    expect(field.openInNewWindow).toEqual(false);
    expect(field.isImageMap).toEqual(false);
  });


  /*  //ExStart
    //ExFor:MergeFieldImageDimension
    //ExFor:MergeFieldImageDimension.#ctor(Double)
    //ExFor:MergeFieldImageDimension.#ctor(Double,MergeFieldImageDimensionUnit)
    //ExFor:MergeFieldImageDimension.Unit
    //ExFor:MergeFieldImageDimension.Value
    //ExFor:MergeFieldImageDimensionUnit
    //ExFor:ImageFieldMergingArgs
    //ExFor:ImageFieldMergingArgs.ImageFileName
    //ExFor:ImageFieldMergingArgs.ImageWidth
    //ExFor:ImageFieldMergingArgs.ImageHeight
    //ExFor:ImageFieldMergingArgs.Shape
    //ExSummary:Shows how to set the dimensions of images as MERGEFIELDS accepts them during a mail merge.
  test('MergeFieldImageDimension', () => {
    let doc = new aw.Document();

    // Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
    // a column in the data source containing local system filenames of images we wish to use in the mail merge.
    let builder = new aw.DocumentBuilder(doc);
    let field = (FieldMergeField)builder.insertField("MERGEFIELD Image:ImageColumn");

    // The data source should have such a column named "ImageColumn".
    expect(field.fieldName).toEqual("Image:ImageColumn");

    // Create a suitable data source.
    let dataTable = new DataTable("Images");
    dataTable.columns.add(new DataColumn("ImageColumn"));
    dataTable.rows.add(base.imageDir + "Logo.jpg");
    dataTable.rows.add(base.imageDir + "Transparent background logo.png");
    dataTable.rows.add(base.imageDir + "Enhanced Windows MetaFile.emf");

    // Configure a callback to modify the sizes of images at merge time, then execute the mail merge.
    doc.mailMerge.fieldMergingCallback = new MergedImageResizer(200, 200, aw.Fields.MergeFieldImageDimensionUnit.Point);
    doc.mailMerge.execute(dataTable);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.MERGEFIELD.ImageDimension.docx");
    TestMergeFieldImageDimension(doc); //ExSkip
  });


    /// <summary>
    /// Sets the size of all mail merged images to one defined width and height.
    /// </summary>
  private class MergedImageResizer : IFieldMergingCallback
  {
    public MergedImageResizer(double imageWidth, double imageHeight, MergeFieldImageDimensionUnit unit)
    {
      mImageWidth = imageWidth;
      mImageHeight = imageHeight;
      mUnit = unit;
    }

    public void FieldMerging(FieldMergingArgs e)
    {
      throw new NotImplementedException();
    }

    public void ImageFieldMerging(ImageFieldMergingArgs args)
    {
      args.imageFileName = args.fieldValue.toString();
      args.imageWidth = new aw.Fields.MergeFieldImageDimension(mImageWidth, mUnit);
      args.imageHeight = new aw.Fields.MergeFieldImageDimension(mImageHeight, mUnit);

      expect(args.imageWidth.value).toEqual(mImageWidth);
      expect(args.imageWidth.unit).toEqual(mUnit);
      expect(args.imageHeight.value).toEqual(mImageHeight);
      expect(args.imageHeight.unit).toEqual(mUnit);
      expect(args.shape).toBe(null);
    }

    private readonly double mImageWidth;
    private readonly double mImageHeight;
    private readonly MergeFieldImageDimensionUnit mUnit;
  }
    //ExEnd

  private void TestMergeFieldImageDimension(Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);

    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(3);

    let shape = (Shape)doc.getShape(0, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.width).toEqual(200.0);
    expect(shape.height).toEqual(200.0);

    shape = (Shape)doc.getShape(1, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Png, shape);
    expect(shape.width).toEqual(200.0);
    expect(shape.height).toEqual(200.0);

    shape = (Shape)doc.getShape(2, true);

    TestUtil.VerifyImageInShape(534, 534, aw.Drawing.ImageType.Emf, shape);
    expect(shape.width).toEqual(200.0);
    expect(shape.height).toEqual(200.0);
  }*/

  /*  //ExStart
    //ExFor:ImageFieldMergingArgs.Image
    //ExSummary:Shows how to use a callback to customize image merging logic.
  test('MergeFieldImages', () => {
    let doc = new aw.Document();

    // Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
    // a column in the data source which contains local system filenames of images we wish to use in the mail merge.
    let builder = new aw.DocumentBuilder(doc);
    let field = (FieldMergeField)builder.insertField("MERGEFIELD Image:ImageColumn");

    // In this case, the field expects the data source to have such a column named "ImageColumn".
    expect(field.fieldName).toEqual("Image:ImageColumn");

    // Filenames can be lengthy, and if we can find a way to avoid storing them in the data source,
    // we may considerably reduce its size.
    // Create a data source that refers to images using short names.
    let dataTable = new DataTable("Images");
    dataTable.columns.add(new DataColumn("ImageColumn"));
    dataTable.rows.add("Dark logo");
    dataTable.rows.add("Transparent logo");

    // Assign a merging callback that contains all logic that processes those names,
    // and then execute the mail merge. 
    doc.mailMerge.fieldMergingCallback = new ImageFilenameCallback();
    doc.mailMerge.execute(dataTable);

    doc.save(base.artifactsDir + "Field.MERGEFIELD.Images.docx");
    TestMergeFieldImages(new aw.Document(base.artifactsDir + "Field.MERGEFIELD.Images.docx")); //ExSkip
  });


    /// <summary>
    /// Contains a dictionary that maps names of images to local system filenames that contain these images.
    /// If a mail merge data source uses one of the dictionary's names to refer to an image,
    /// this callback will pass the respective filename to the merge destination.
    /// </summary>
  private class ImageFilenameCallback : IFieldMergingCallback
  {
    public ImageFilenameCallback()
    {
      mImageFilenames = new Dictionary<string, string>();
      mImageFilenames.add("Dark logo", base.imageDir + "Logo.jpg");
      mImageFilenames.add("Transparent logo", base.imageDir + "Transparent background logo.png");
    }

    void aw.MailMerging.IFieldMergingCallback.fieldMerging(FieldMergingArgs args)
    {
      throw new NotImplementedException();
    }

    void aw.MailMerging.IFieldMergingCallback.imageFieldMerging(ImageFieldMergingArgs args)
    {
      if (mImageFilenames.containsKey(args.fieldValue.toString()))
      {
#if NET461_OR_GREATER || JAVA
        args.image = Image.FromFile(mImageFilenames.at(args.fieldValue.toString()));
#elif NET5_0_OR_GREATER
        args.image = SKBitmap.Decode(mImageFilenames.at(args.fieldValue.toString()));
        args.imageFileName = mImageFilenames.at(args.fieldValue.toString());
#endif
      }

      expect(args.image).not.toBe(null);
    }

    private readonly Dictionary<string, string> mImageFilenames;
  }
    //ExEnd

  private void TestMergeFieldImages(Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);

    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getChildNodes(aw.NodeType.Shape, true).Count).toEqual(2);

    let shape = (Shape)doc.getShape(0, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Jpeg, shape);
    expect(shape.width).toEqual(300.0);
    expect(shape.height).toEqual(300.0);

    shape = (Shape)doc.getShape(1, true);

    TestUtil.VerifyImageInShape(400, 400, aw.Drawing.ImageType.Png, shape);
    expect(shape.width, 1).toEqual(300.0);
    expect(shape.height, 1).toEqual(300.0);
  }*/

  test('FieldIndexFilter', () => {
    //ExStart
    //ExFor:FieldIndex
    //ExFor:FieldIndex.bookmarkName
    //ExFor:FieldIndex.entryType
    //ExFor:FieldXE
    //ExFor:FieldXE.entryType
    //ExFor:FieldXE.text
    //ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side
    // and the page containing the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // Configure the INDEX field only to display XE fields that are within the bounds
    // of a bookmark named "MainBookmark", and whose "EntryType" properties have a value of "A".
    // For both INDEX and XE fields, the "EntryType" property only uses the first character of its string value.
    index.bookmarkName = "MainBookmark";
    index.entryType = "A";

    expect(index.getFieldCode()).toEqual(" INDEX  \\b MainBookmark \\f A");

    // On a new page, start the bookmark with a name that matches the value
    // of the INDEX field's "BookmarkName" property.
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.startBookmark("MainBookmark");

    // The INDEX field will pick up this entry because it is inside the bookmark,
    // and its entry type also matches the INDEX field's entry type.
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Index entry 1";
    indexEntry.entryType = "A";

    expect(indexEntry.getFieldCode()).toEqual(" XE  \"Index entry 1\" \\f A");

    // Insert an XE field that will not appear in the INDEX because the entry types do not match.
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Index entry 2";
    indexEntry.entryType = "B";

    // End the bookmark and insert an XE field afterwards.
    // It is of the same type as the INDEX field, but will not appear
    // since it is outside the bookmark's boundaries.
    builder.endBookmark("MainBookmark");
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Index entry 3";
    indexEntry.entryType = "A";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.Filtering.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.Filtering.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index);
    expect(index.bookmarkName).toEqual("MainBookmark");
    expect(index.entryType).toEqual("A");

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"Index entry 1\" \\f A", '', indexEntry);
    expect(indexEntry.text).toEqual("Index entry 1");
    expect(indexEntry.entryType).toEqual("A");

    indexEntry = doc.range.fields.at(2).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"Index entry 2\" \\f B", '', indexEntry);
    expect(indexEntry.text).toEqual("Index entry 2");
    expect(indexEntry.entryType).toEqual("B");

    indexEntry = doc.range.fields.at(3).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"Index entry 3\" \\f A", '', indexEntry);
    expect(indexEntry.text).toEqual("Index entry 3");
    expect(indexEntry.entryType).toEqual("A");
  });


  test('FieldIndexFormatting', () => {
    //ExStart
    //ExFor:FieldIndex
    //ExFor:FieldIndex.heading
    //ExFor:FieldIndex.numberOfColumns
    //ExFor:FieldIndex.languageId
    //ExFor:FieldIndex.letterRange
    //ExFor:FieldXE
    //ExFor:FieldXE.isBold
    //ExFor:FieldXE.isItalic
    //ExFor:FieldXE.text
    //ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();
    index.languageId = "1033";

    // Setting this property's value to "A" will group all the entries by their first letter,
    // and place that letter in uppercase above each group.
    index.heading = "A";

    // Set the table created by the INDEX field to span over 2 columns.
    index.numberOfColumns = "2";

    // Set any entries with starting letters outside the "a-c" character range to be omitted.
    index.letterRange = "a-c";

    expect(index.getFieldCode()).toEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c");

    // These next two XE fields will show up under the "A" heading,
    // with their respective text stylings also applied to their page numbers.
    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Apple";
    indexEntry.isItalic = true;

    expect(indexEntry.getFieldCode()).toEqual(" XE  Apple \\i");

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Apricot";
    indexEntry.isBold = true;

    expect(indexEntry.getFieldCode()).toEqual(" XE  Apricot \\b");

    // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Banana";

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Cherry";

    // INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Avocado";

    // This entry will not appear because it starts with the letter "D",
    // which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Durian";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.Formatting.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.Formatting.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    expect(index.languageId).toEqual("1033");
    expect(index.heading).toEqual("A");
    expect(index.numberOfColumns).toEqual("2");
    expect(index.letterRange).toEqual("a-c");
    expect(index.getFieldCode()).toEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c");
    expect(index.result).toEqual("\fA\r" +
                            "Apple, 2\r" +
                            "Apricot, 3\r" +
                            "Avocado, 6\r" +
                            "B\r" +
                            "Banana, 4\r" +
                            "C\r" +
                            "Cherry, 5\r\f");

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Apple \\i", '', indexEntry);
    expect(indexEntry.text).toEqual("Apple");
    expect(indexEntry.isBold).toEqual(false);
    expect(indexEntry.isItalic).toEqual(true);

    indexEntry = doc.range.fields.at(2).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Apricot \\b", '', indexEntry);
    expect(indexEntry.text).toEqual("Apricot");
    expect(indexEntry.isBold).toEqual(true);
    expect(indexEntry.isItalic).toEqual(false);

    indexEntry = doc.range.fields.at(3).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Banana", '', indexEntry);
    expect(indexEntry.text).toEqual("Banana");
    expect(indexEntry.isBold).toEqual(false);
    expect(indexEntry.isItalic).toEqual(false);

    indexEntry = doc.range.fields.at(4).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Cherry", '', indexEntry);
    expect(indexEntry.text).toEqual("Cherry");
    expect(indexEntry.isBold).toEqual(false);
    expect(indexEntry.isItalic).toEqual(false);

    indexEntry = doc.range.fields.at(5).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Avocado", '', indexEntry);
    expect(indexEntry.text).toEqual("Avocado");
    expect(indexEntry.isBold).toEqual(false);
    expect(indexEntry.isItalic).toEqual(false);

    indexEntry = doc.range.fields.at(6).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Durian", '', indexEntry);
    expect(indexEntry.text).toEqual("Durian");
    expect(indexEntry.isBold).toEqual(false);
    expect(indexEntry.isItalic).toEqual(false);
  });


  test('FieldIndexSequence', () => {
    //ExStart
    //ExFor:FieldIndex.hasSequenceName
    //ExFor:FieldIndex.sequenceName
    //ExFor:FieldIndex.sequenceSeparator
    //ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // If the XE fields have the same value in their "Text" property,
    // the INDEX field will group them into one entry.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // In the SequenceName property, name a SEQ field sequence. Each entry of this INDEX field will now also display
    // the number that the sequence count is on at the XE field location that created this entry.
    index.sequenceName = "MySequence";

    // Set text that will around the sequence and page numbers to explain their meaning to the user.
    // An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
    // PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
    index.pageNumberSeparator = "\tMySequence at ";
    index.sequenceSeparator = " on page ";
    expect(index.hasSequenceName).toEqual(true);

    expect(index.getFieldCode()).toEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"");

    // SEQ fields display a count that increments at each SEQ field.
    // These fields also maintain separate counts for each unique named sequence
    // identified by the SEQ field's "SequenceIdentifier" property.
    // Insert a SEQ field which moves the "MySequence" sequence to 1.
    // This field no different from normal document text. It will not appear on an INDEX field's table of contents.
    builder.insertBreak(aw.BreakType.PageBreak);
    let sequenceField = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    sequenceField.sequenceIdentifier = "MySequence";

    expect(sequenceField.getFieldCode()).toEqual(" SEQ  MySequence");

    // Insert an XE field which will create an entry in the INDEX field.
    // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
    // this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Cat";

    expect(indexEntry.getFieldCode()).toEqual(" XE  Cat");

    // Insert a page break and use SEQ fields to advance "MySequence" to 3.
    builder.insertBreak(aw.BreakType.PageBreak);
    sequenceField = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    sequenceField.sequenceIdentifier = "MySequence";
    sequenceField = builder.insertField(aw.Fields.FieldType.FieldSequence, true).asFieldSeq();
    sequenceField.sequenceIdentifier = "MySequence";

    // Insert an XE field with the same Text property as the one above.
    // The INDEX entry will group XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
    // The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Cat";

    // Insert an XE field with a new and unique Text property value.
    // This will add a new entry, with MySequence at 3 on page 4.
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Dog";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.Sequence.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.Sequence.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    expect(index.sequenceName).toEqual("MySequence");
    expect(index.pageNumberSeparator).toEqual("\tMySequence at ");
    expect(index.sequenceSeparator).toEqual(" on page ");
    expect(index.hasSequenceName).toEqual(true);
    expect(index.getFieldCode()).toEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"");
    expect(index.result).toEqual("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
                            "Dog\tMySequence at 3 on page 4\r");

    expect(Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldSequence).length).toEqual(3);
  });


  test('FieldIndexPageNumberSeparator', () => {
    //ExStart
    //ExFor:FieldIndex.hasPageNumberSeparator
    //ExFor:FieldIndex.pageNumberSeparator
    //ExFor:FieldIndex.pageNumberListSeparator
    //ExSummary:Shows how to edit the page number separator in an INDEX field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will group XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // If our INDEX field has an entry for a group of XE fields,
    // this entry will display the number of each page that contains an XE field that belongs to this group.
    // We can set custom separators to customize the appearance of these page numbers.
    index.pageNumberSeparator = ", on page(s) ";
    index.pageNumberListSeparator = " & ";

    expect(index.getFieldCode()).toEqual(" INDEX  \\e \", on page(s) \" \\l \" & \"");
    expect(index.hasPageNumberSeparator).toEqual(true);

    // After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "First entry";

    expect(indexEntry.getFieldCode()).toEqual(" XE  \"First entry\"");

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "First entry";

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "First entry";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.PageNumberList.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.PageNumberList.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index);
    expect(index.pageNumberSeparator).toEqual(", on page(s) ");
    expect(index.pageNumberListSeparator).toEqual(" & ");
    expect(index.hasPageNumberSeparator).toEqual(true);
  });


  test('FieldIndexPageRangeBookmark', () => {
    //ExStart
    //ExFor:FieldIndex.pageRangeSeparator
    //ExFor:FieldXE.pageRangeBookmarkName
    //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // For INDEX entries that display page ranges, we can specify a separator string
    // which will appear between the number of the first page, and the number of the last.
    index.pageNumberSeparator = ", on page(s) ";
    index.pageRangeSeparator = " to ";

    expect(index.getFieldCode()).toEqual(" INDEX  \\e \", on page(s) \" \\g \" to \"");

    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "My entry";

    // If an XE field names a bookmark using the PageRangeBookmarkName property,
    // its INDEX entry will show the range of pages that the bookmark spans
    // instead of the number of the page that contains the XE field.
    indexEntry.pageRangeBookmarkName = "MyBookmark";

    expect(indexEntry.getFieldCode()).toEqual(" XE  \"My entry\" \\r MyBookmark");
    expect(indexEntry.pageRangeBookmarkName).toEqual("MyBookmark");

    // Insert a bookmark that starts on page 3 and ends on page 5.
    // The INDEX entry for the XE field that references this bookmark will display this page range.
    // In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.startBookmark("MyBookmark");
    builder.write("Start of MyBookmark");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.write("End of MyBookmark");
    builder.endBookmark("MyBookmark");

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.PageRangeBookmark.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.PageRangeBookmark.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index);
    expect(index.pageNumberSeparator).toEqual(", on page(s) ");
    expect(index.pageRangeSeparator).toEqual(" to ");

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"My entry\" \\r MyBookmark", '', indexEntry);
    expect(indexEntry.text).toEqual("My entry");
    expect(indexEntry.pageRangeBookmarkName).toEqual("MyBookmark");
  });


  test('FieldIndexCrossReferenceSeparator', () => {
    //ExStart
    //ExFor:FieldIndex.crossReferenceSeparator
    //ExFor:FieldXE.pageNumberReplacement
    //ExSummary:Shows how to define cross references in an INDEX field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // We can configure an XE field to get its INDEX entry to display a string instead of a page number.
    // First, for entries that substitute a page number with a string,
    // specify a custom separator between the XE field's Text property value and the string.
    index.crossReferenceSeparator = ", see: ";

    expect(index.getFieldCode()).toEqual(" INDEX  \\k \", see: \"");

    // Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
    // and does not invoke the CrossReferenceSeparator value.
    // The entry for this XE field will display "Apple, 2".
    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Apple";

    expect(indexEntry.getFieldCode()).toEqual(" XE  Apple");

    // Insert another XE field on page 3 and set a value for the PageNumberReplacement property.
    // This value will show up instead of the number of the page that this field is on,
    // and the INDEX field's CrossReferenceSeparator value will appear in front of it.
    // The entry for this XE field will display "Banana, see: Tropical fruit".
    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Banana";
    indexEntry.pageNumberReplacement = "Tropical fruit";

    expect(indexEntry.getFieldCode()).toEqual(" XE  Banana \\t \"Tropical fruit\"");

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.crossReferenceSeparator.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.crossReferenceSeparator.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\k \", see: \"",
      "Apple, 2\r" +
      "Banana, see: Tropical fruit\r", index);
    expect(index.crossReferenceSeparator).toEqual(", see: ");

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Apple", '', indexEntry);
    expect(indexEntry.text).toEqual("Apple");
    expect(indexEntry.pageNumberReplacement).toBe(null);

    indexEntry = doc.range.fields.at(2).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  Banana \\t \"Tropical fruit\"", '', indexEntry);
    expect(indexEntry.text).toEqual("Banana");
    expect(indexEntry.pageNumberReplacement).toEqual("Tropical fruit");
  });


  test.each([true, false])('FieldIndexSubheading(%o)', (runSubentriesOnTheSameLine) => {
    //ExStart
    //ExFor:FieldIndex.runSubentriesOnSameLine
    //ExSummary:Shows how to work with subentries in an INDEX field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();
    index.pageNumberSeparator = ", see page ";
    index.heading = "A";

    // XE fields that have a Text property whose value becomes the heading of the INDEX entry.
    // If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
    // the first segment is heading, and the second segment will become the subheading.
    // The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
    // headings, the INDEX field will further subgroup them by the values of these headings.
    // There can be multiple subgrouping layers, depending on how many times
    // the Text properties of XE fields get segmented like this.
    // By default, an INDEX field entry group will create a new line for every subheading within this group. 
    // We can set the RunSubentriesOnSameLine flag to true to keep the heading,
    // and every subheading for the group on one line instead, which will make the INDEX field more compact.
    index.runSubentriesOnSameLine = runSubentriesOnTheSameLine;

    if (runSubentriesOnTheSameLine)
      expect(index.getFieldCode()).toEqual(" INDEX  \\e \", see page \" \\h A \\r");
    else
      expect(index.getFieldCode()).toEqual(" INDEX  \\e \", see page \" \\h A");

    // Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
    // which the INDEX field will use to group them.
    // If RunSubentriesOnSameLine is false, then the INDEX table will create three lines:
    // one line for the grouping heading "Heading 1", and one more line for each subheading.
    // If RunSubentriesOnSameLine is true, then the INDEX table will create a one-line
    // entry that encompasses the heading and every subheading.
    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Heading 1:Subheading 1";

    expect(indexEntry.getFieldCode()).toEqual(" XE  \"Heading 1:Subheading 1\"");

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "Heading 1:Subheading 2";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + `Field.INDEX.XE.Subheading.docx`);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + `Field.INDEX.XE.Subheading.docx`);
    index = doc.range.fields.at(0).asFieldIndex();

    if (runSubentriesOnTheSameLine)
    {
      TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\e \", see page \" \\h A \\r",
        "H\r" +
        "Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
      expect(index.runSubentriesOnSameLine).toEqual(true);
    }
    else
    {
      TestUtil.verifyField(aw.Fields.FieldType.FieldIndex, " INDEX  \\e \", see page \" \\h A",
        "H\r" +
        "Heading 1\r" +
        "Subheading 1, see page 2\r" +
        "Subheading 2, see page 3\r", index);
      expect(index.runSubentriesOnSameLine).toEqual(false);
    }

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"Heading 1:Subheading 1\"", '', indexEntry);
    expect(indexEntry.text).toEqual("Heading 1:Subheading 1");

    indexEntry = doc.range.fields.at(2).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  \"Heading 1:Subheading 2\"", '', indexEntry);
    expect(indexEntry.text).toEqual("Heading 1:Subheading 2");
  });


  test.skip.each([true, false])('FieldIndexYomi(%o): WORDSNET-24595', (sortEntriesUsingYomi) => {
    //ExStart
    //ExFor:FieldIndex.useYomi
    //ExFor:FieldXE.yomi
    //ExSummary:Shows how to sort INDEX field entries phonetically.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create an INDEX field which will display an entry for each XE field found in the document.
    // Each entry will display the XE field's Text property value on the left side,
    // and the number of the page that contains the XE field on the right.
    // The INDEX entry will collect all XE fields with matching values in the "Text" property
    // into one entry as opposed to making an entry for each XE field.
    let index = builder.insertField(aw.Fields.FieldType.FieldIndex, true).asFieldIndex();

    // The INDEX table automatically sorts its entries by the values of their Text properties in alphabetic order.
    // Set the INDEX table to sort entries phonetically using Hiragana instead.
    index.useYomi = sortEntriesUsingYomi;

    if (sortEntriesUsingYomi)
      expect(index.getFieldCode()).toEqual(" INDEX  \\y");
    else
      expect(index.getFieldCode()).toEqual(" INDEX ");

    // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents.
    // The "Text" property may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
    // while the "Yomi" version of the word will spell exactly how it is pronounced using Hiragana.
    // If we set our INDEX field to use Yomi, it will sort these entries
    // by the value of their Yomi properties, instead of their Text values.
    builder.insertBreak(aw.BreakType.PageBreak);
    let indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "愛子";
    indexEntry.yomi = "あ";

    expect(indexEntry.getFieldCode()).toEqual(" XE  愛子 \\y あ");

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "明美";
    indexEntry.yomi = "あ";

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "恵美";
    indexEntry.yomi = "え";

    builder.insertBreak(aw.BreakType.PageBreak);
    indexEntry = builder.insertField(aw.Fields.FieldType.FieldIndexEntry, true).asFieldXE();
    indexEntry.text = "愛美";
    indexEntry.yomi = "え";

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.INDEX.XE.yomi.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INDEX.XE.yomi.docx");
    index = doc.range.fields.at(0).asFieldIndex();

    if (sortEntriesUsingYomi)
    {
      expect(index.useYomi).toEqual(true);
      expect(index.getFieldCode()).toEqual(" INDEX  \\y");
      expect(index.result).toEqual("愛子, 2\r" +
                                "明美, 3\r" +
                                "恵美, 4\r" +
                                "愛美, 5\r");
    }
    else
    {
      expect(index.useYomi).toEqual(false);
      expect(index.getFieldCode()).toEqual(" INDEX ");
      expect(index.result).toEqual("恵美, 4\r" +
                                "愛子, 2\r" +
                                "愛美, 5\r" +
                                "明美, 3\r");
    }

    indexEntry = doc.range.fields.at(1).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  愛子 \\y あ", '', indexEntry);
    expect(indexEntry.text).toEqual("愛子");
    expect(indexEntry.yomi).toEqual("あ");

    indexEntry = doc.range.fields.at(2).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  明美 \\y あ", '', indexEntry);
    expect(indexEntry.text).toEqual("明美");
    expect(indexEntry.yomi).toEqual("あ");

    indexEntry = doc.range.fields.at(3).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  恵美 \\y え", '', indexEntry);
    expect(indexEntry.text).toEqual("恵美");
    expect(indexEntry.yomi).toEqual("え");

    indexEntry = doc.range.fields.at(4).asFieldXE();

    TestUtil.verifyField(aw.Fields.FieldType.FieldIndexEntry, " XE  愛美 \\y え", '', indexEntry);
    expect(indexEntry.text).toEqual("愛美");
    expect(indexEntry.yomi).toEqual("え");
  });


  test('FieldBarcode', () => {
    //ExStart
    //ExFor:FieldBarcode
    //ExFor:FieldBarcode.facingIdentificationMark
    //ExFor:FieldBarcode.isBookmark
    //ExFor:FieldBarcode.isUSPostalAddress
    //ExFor:FieldBarcode.postalAddress
    //ExSummary:Shows how to use the BARCODE field to display U.S. ZIP codes in the form of a barcode. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln();

    // Below are two ways of using BARCODE fields to display custom values as barcodes.
    // 1 -  Store the value that the barcode will display in the PostalAddress property:
    let field = builder.insertField(aw.Fields.FieldType.FieldBarcode, true).asFieldBarcode();

    // This value needs to be a valid ZIP code.
    field.postalAddress = "96801";
    field.isUSPostalAddress = true;
    field.facingIdentificationMark = "C";

    expect(field.getFieldCode()).toEqual(" BARCODE  96801 \\u \\f C");

    builder.insertBreak(aw.BreakType.LineBreak);

    // 2 -  Reference a bookmark that stores the value that this barcode will display:
    field = builder.insertField(aw.Fields.FieldType.FieldBarcode, true).asFieldBarcode();
    field.postalAddress = "BarcodeBookmark";
    field.isBookmark = true;

    expect(field.getFieldCode()).toEqual(" BARCODE  BarcodeBookmark \\b");

    // The bookmark that the BARCODE field references in its PostalAddress property
    // need to contain nothing besides the valid ZIP code.
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.startBookmark("BarcodeBookmark");
    builder.writeln("968877");
    builder.endBookmark("BarcodeBookmark");

    doc.save(base.artifactsDir + "Field.BARCODE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.BARCODE.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);

    field = doc.range.fields.at(0).asFieldBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldBarcode, " BARCODE  96801 \\u \\f C", '', field);
    expect(field.facingIdentificationMark).toEqual("C");
    expect(field.postalAddress).toEqual("96801");
    expect(field.isUSPostalAddress).toEqual(true);

    field = doc.range.fields.at(1).asFieldBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldBarcode, " BARCODE  BarcodeBookmark \\b", '', field);
    expect(field.postalAddress).toEqual("BarcodeBookmark");
    expect(field.isBookmark).toEqual(true);
  });


  test('FieldDisplayBarcode', () => {
    //ExStart
    //ExFor:FieldDisplayBarcode
    //ExFor:FieldDisplayBarcode.addStartStopChar
    //ExFor:FieldDisplayBarcode.backgroundColor
    //ExFor:FieldDisplayBarcode.barcodeType
    //ExFor:FieldDisplayBarcode.barcodeValue
    //ExFor:FieldDisplayBarcode.caseCodeStyle
    //ExFor:FieldDisplayBarcode.displayText
    //ExFor:FieldDisplayBarcode.errorCorrectionLevel
    //ExFor:FieldDisplayBarcode.fixCheckDigit
    //ExFor:FieldDisplayBarcode.foregroundColor
    //ExFor:FieldDisplayBarcode.posCodeStyle
    //ExFor:FieldDisplayBarcode.scalingFactor
    //ExFor:FieldDisplayBarcode.symbolHeight
    //ExFor:FieldDisplayBarcode.symbolRotation
    //ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldDisplayBarcode, true).asFieldDisplayBarcode();

    // Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
    // 1 -  QR code with custom colors:
    field.barcodeType = "QR";
    field.barcodeValue = "ABC123";
    field.backgroundColor = "0xF8BD69";
    field.foregroundColor = "0xB5413B";
    field.errorCorrectionLevel = "3";
    field.scalingFactor = "250";
    field.symbolHeight = "1000";
    field.symbolRotation = "0";

    expect(field.getFieldCode()).toEqual(" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0");
    builder.writeln();

    // 2 -  EAN13 barcode, with the digits displayed below the bars:
    field = builder.insertField(aw.Fields.FieldType.FieldDisplayBarcode, true).asFieldDisplayBarcode();
    field.barcodeType = "EAN13";
    field.barcodeValue = "501234567890";
    field.displayText = true;
    field.posCodeStyle = "CASE";
    field.fixCheckDigit = true;

    expect(field.getFieldCode()).toEqual(" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x");
    builder.writeln();

    // 3 -  CODE39 barcode:
    field = builder.insertField(aw.Fields.FieldType.FieldDisplayBarcode, true).asFieldDisplayBarcode();
    field.barcodeType = "CODE39";
    field.barcodeValue = "12345ABCDE";
    field.addStartStopChar = true;

    expect(field.getFieldCode()).toEqual(" DISPLAYBARCODE  12345ABCDE CODE39 \\d");
    builder.writeln();

    // 4 -  ITF4 barcode, with a specified case code:
    field = builder.insertField(aw.Fields.FieldType.FieldDisplayBarcode, true).asFieldDisplayBarcode();
    field.barcodeType = "ITF14";
    field.barcodeValue = "09312345678907";
    field.caseCodeStyle = "STD";

    expect(field.getFieldCode()).toEqual(" DISPLAYBARCODE  09312345678907 ITF14 \\c STD");

    doc.save(base.artifactsDir + "Field.DISPLAYBARCODE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.DISPLAYBARCODE.docx");

    expect(doc.getChildNodes(aw.NodeType.Shape, true).count).toEqual(0);

    field = doc.range.fields.at(0).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", '', field);
    expect(field.barcodeType).toEqual("QR");
    expect(field.barcodeValue).toEqual("ABC123");
    expect(field.backgroundColor).toEqual("0xF8BD69");
    expect(field.foregroundColor).toEqual("0xB5413B");
    expect(field.errorCorrectionLevel).toEqual("3");
    expect(field.scalingFactor).toEqual("250");
    expect(field.symbolHeight).toEqual("1000");
    expect(field.symbolRotation).toEqual("0");

    field = doc.range.fields.at(1).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", '', field);
    expect(field.barcodeType).toEqual("EAN13");
    expect(field.barcodeValue).toEqual("501234567890");
    expect(field.displayText).toEqual(true);
    expect(field.posCodeStyle).toEqual("CASE");
    expect(field.fixCheckDigit).toEqual(true);

    field = doc.range.fields.at(2).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", '', field);
    expect(field.barcodeType).toEqual("CODE39");
    expect(field.barcodeValue).toEqual("12345ABCDE");
    expect(field.addStartStopChar).toEqual(true);

    field = doc.range.fields.at(3).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", '', field);
    expect(field.barcodeType).toEqual("ITF14");
    expect(field.barcodeValue).toEqual("09312345678907");
    expect(field.caseCodeStyle).toEqual("STD");
  });


  test.skip('FieldMergeBarcode_QR: DataTable', () => {
    //ExStart
    //ExFor:FieldDisplayBarcode
    //ExFor:FieldMergeBarcode
    //ExFor:FieldMergeBarcode.backgroundColor
    //ExFor:FieldMergeBarcode.barcodeType
    //ExFor:FieldMergeBarcode.barcodeValue
    //ExFor:FieldMergeBarcode.errorCorrectionLevel
    //ExFor:FieldMergeBarcode.foregroundColor
    //ExFor:FieldMergeBarcode.scalingFactor
    //ExFor:FieldMergeBarcode.symbolHeight
    //ExFor:FieldMergeBarcode.symbolRotation
    //ExSummary:Shows how to perform a mail merge on QR barcodes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    // This field will convert all values in a merge data source's "MyQRCode" column into QR codes.
    let field = builder.insertField(aw.Fields.FieldType.FieldMergeBarcode, true).asFieldMergeBarcode();
    field.barcodeType = "QR";
    field.barcodeValue = "MyQRCode";

    // Apply custom colors and scaling.
    field.backgroundColor = "0xF8BD69";
    field.foregroundColor = "0xB5413B";
    field.errorCorrectionLevel = "3";
    field.scalingFactor = "250";
    field.symbolHeight = "1000";
    field.symbolRotation = "0";

    expect(field.type).toEqual(aw.Fields.FieldType.FieldMergeBarcode);
    expect(field.getFieldCode()).toEqual(" MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0");
    builder.writeln();

    // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    // which will display a QR code with the value from the merged row.
    let table = new DataTable("Barcodes");
    table.columns.add("MyQRCode");
    table.rows.add(["ABC123"]);
    table.rows.add(["DEF456"]);

    doc.mailMerge.execute(table);

    expect(doc.range.fields.at(0).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(0).getFieldCode()).toEqual("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B");
    expect(doc.range.fields.at(1).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(1).getFieldCode()).toEqual("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B");

    doc.save(base.artifactsDir + "Field.MERGEBARCODE.QR.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MERGEBARCODE.QR.docx");

    expect(Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldMergeBarcode).length).toEqual(0);

    let barcode = doc.range.fields.at(0).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, 
      "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", '', barcode);
    expect(barcode.barcodeValue).toEqual("ABC123");
    expect(barcode.barcodeType).toEqual("QR");

    barcode = doc.range.fields.at(1).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, 
      "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", '', barcode);
    expect(barcode.barcodeValue).toEqual("DEF456");
    expect(barcode.barcodeType).toEqual("QR");
  });


  test.skip('FieldMergeBarcode_EAN13: DataTable', () => {
    //ExStart
    //ExFor:FieldMergeBarcode
    //ExFor:FieldMergeBarcode.barcodeType
    //ExFor:FieldMergeBarcode.barcodeValue
    //ExFor:FieldMergeBarcode.displayText
    //ExFor:FieldMergeBarcode.fixCheckDigit
    //ExFor:FieldMergeBarcode.posCodeStyle
    //ExSummary:Shows how to perform a mail merge on EAN13 barcodes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    // This field will convert all values in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes.
    let field = builder.insertField(aw.Fields.FieldType.FieldMergeBarcode, true).asFieldMergeBarcode();
    field.barcodeType = "EAN13";
    field.barcodeValue = "MyEAN13Barcode";

    // Display the numeric value of the barcode underneath the bars.
    field.displayText = true;
    field.posCodeStyle = "CASE";
    field.fixCheckDigit = true;

    expect(field.type).toEqual(aw.Fields.FieldType.FieldMergeBarcode);
    expect(field.getFieldCode()).toEqual(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x");
    builder.writeln();

    // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    // which will display an EAN13 barcode with the value from the merged row.
    let table = new DataTable("Barcodes");
    table.columns.add("MyEAN13Barcode");
    table.rows.add(["501234567890"]);
    table.rows.add(["123456789012"]);

    doc.mailMerge.execute(table);

    expect(doc.range.fields.at(0).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(0).getFieldCode()).toEqual("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x");
    expect(doc.range.fields.at(1).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(1).getFieldCode()).toEqual("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x");

    doc.save(base.artifactsDir + "Field.MERGEBARCODE.EAN13.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MERGEBARCODE.EAN13.docx");

    expect(Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldMergeBarcode).length).toEqual(0);

    let barcode = doc.range.fields.at(0).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", '', barcode);
    expect(barcode.barcodeValue).toEqual("501234567890");
    expect(barcode.barcodeType).toEqual("EAN13");

    barcode = doc.range.fields.at(1).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", '', barcode);
    expect(barcode.barcodeValue).toEqual("123456789012");
    expect(barcode.barcodeType).toEqual("EAN13");
  });


  test.skip('FieldMergeBarcode_CODE39: DataTable', () => {
    //ExStart
    //ExFor:FieldMergeBarcode
    //ExFor:FieldMergeBarcode.addStartStopChar
    //ExFor:FieldMergeBarcode.barcodeType
    //ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    // This field will convert all values in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes.
    let field = builder.insertField(aw.Fields.FieldType.FieldMergeBarcode, true).asFieldMergeBarcode();
    field.barcodeType = "CODE39";
    field.barcodeValue = "MyCODE39Barcode";

    // Edit its appearance to display start/stop characters.
    field.addStartStopChar = true;

    expect(field.type).toEqual(aw.Fields.FieldType.FieldMergeBarcode);
    expect(field.getFieldCode()).toEqual(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d");
    builder.writeln();

    // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    // which will display a CODE39 barcode with the value from the merged row.
    let table = new DataTable("Barcodes");
    table.columns.add("MyCODE39Barcode");
    table.rows.add(["12345ABCDE"]);
    table.rows.add(["67890FGHIJ"]);

    doc.mailMerge.execute(table);

    expect(doc.range.fields.at(0).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(0).getFieldCode()).toEqual("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d");
    expect(doc.range.fields.at(1).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(1).getFieldCode()).toEqual("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d");

    doc.save(base.artifactsDir + "Field.MERGEBARCODE.CODE39.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MERGEBARCODE.CODE39.docx");

    expect(Array.from(doc.range.fields).filter(f => f.type == aw.Fields.FieldType.FieldMergeBarcode).length).toEqual(0);

    let barcode = doc.range.fields.at(0).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", '', barcode);
    expect(barcode.barcodeValue).toEqual("12345ABCDE");
    expect(barcode.barcodeType).toEqual("CODE39");

    barcode = doc.range.fields.at(1).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", '', barcode);
    expect(barcode.barcodeValue).toEqual("67890FGHIJ");
    expect(barcode.barcodeType).toEqual("CODE39");
  });


  test.skip('FieldMergeBarcode_ITF14: DataTable', () => {
    //ExStart
    //ExFor:FieldMergeBarcode
    //ExFor:FieldMergeBarcode.barcodeType
    //ExFor:FieldMergeBarcode.caseCodeStyle
    //ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    // This field will convert all values in a merge data source's "MyITF14Barcode" column into ITF14 barcodes.
    let field = builder.insertField(aw.Fields.FieldType.FieldMergeBarcode, true).asFieldMergeBarcode();
    field.barcodeType = "ITF14";
    field.barcodeValue = "MyITF14Barcode";
    field.caseCodeStyle = "STD";

    expect(field.type).toEqual(aw.Fields.FieldType.FieldMergeBarcode);
    expect(field.getFieldCode()).toEqual(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD");

    // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    // which will display an ITF14 barcode with the value from the merged row.
    let table = new DataTable("Barcodes");
    table.columns.add("MyITF14Barcode");
    table.rows.add(["09312345678907"]);
    table.rows.add(["1234567891234"]);

    doc.mailMerge.execute(table);

    expect(doc.range.fields.at(0).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(0).getFieldCode()).toEqual("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD");
    expect(doc.range.fields.at(1).type).toEqual(aw.Fields.FieldType.FieldDisplayBarcode);
    expect(doc.range.fields.at(1).getFieldCode()).toEqual("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD");

    doc.save(base.artifactsDir + "Field.MERGEBARCODE.ITF14.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MERGEBARCODE.ITF14.docx");

    expect(doc.range.fields.count(f => f.type == aw.Fields.FieldType.FieldMergeBarcode)).toEqual(0);

    let barcode = doc.range.fields.at(0).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", '', barcode);
    expect(barcode.barcodeValue).toEqual("09312345678907");
    expect(barcode.barcodeType).toEqual("ITF14");

    barcode = doc.range.fields.at(1).asFieldDisplayBarcode();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDisplayBarcode, "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD", '', barcode);
    expect(barcode.barcodeValue).toEqual("1234567891234");
    expect(barcode.barcodeType).toEqual("ITF14");
  });


  const InsertLinkedObjectAs = {
    // LinkedObjectAsText
    Text: 1,
    Unicode: 2,
    Html: 3,
    Rtf: 4,
    // LinkedObjectAsImage
    Picture: 5,
    Bitmap: 6
  }


  //ExStart
  //ExFor:FieldLink
  //ExFor:FieldLink.AutoUpdate
  //ExFor:FieldLink.FormatUpdateType
  //ExFor:FieldLink.InsertAsBitmap
  //ExFor:FieldLink.InsertAsHtml
  //ExFor:FieldLink.InsertAsPicture
  //ExFor:FieldLink.InsertAsRtf
  //ExFor:FieldLink.InsertAsText
  //ExFor:FieldLink.InsertAsUnicode
  //ExFor:FieldLink.IsLinked
  //ExFor:FieldLink.ProgId
  //ExFor:FieldLink.SourceFullName
  //ExFor:FieldLink.SourceItem
  //ExFor:FieldDde
  //ExFor:FieldDde.AutoUpdate
  //ExFor:FieldDde.InsertAsBitmap
  //ExFor:FieldDde.InsertAsHtml
  //ExFor:FieldDde.InsertAsPicture
  //ExFor:FieldDde.InsertAsRtf
  //ExFor:FieldDde.InsertAsText
  //ExFor:FieldDde.InsertAsUnicode
  //ExFor:FieldDde.IsLinked
  //ExFor:FieldDde.ProgId
  //ExFor:FieldDde.SourceFullName
  //ExFor:FieldDde.SourceItem
  //ExFor:FieldDdeAuto
  //ExFor:FieldDdeAuto.InsertAsBitmap
  //ExFor:FieldDdeAuto.InsertAsHtml
  //ExFor:FieldDdeAuto.InsertAsPicture
  //ExFor:FieldDdeAuto.InsertAsRtf
  //ExFor:FieldDdeAuto.InsertAsText
  //ExFor:FieldDdeAuto.InsertAsUnicode
  //ExFor:FieldDdeAuto.IsLinked
  //ExFor:FieldDdeAuto.ProgId
  //ExFor:FieldDdeAuto.SourceFullName
  //ExFor:FieldDdeAuto.SourceItem
  //ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
  test.each([InsertLinkedObjectAs.Text,
    InsertLinkedObjectAs.Unicode,
    InsertLinkedObjectAs.Html,
    InsertLinkedObjectAs.Rtf])('FieldLinkedObjectsAsText(%o)', (insertLinkedObjectAs) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are three types of fields we can use to display contents from a linked document in the form of text.
    // 1 -  A LINK field:
    builder.writeln("FieldLink:\n");
    insertFieldLink(builder, insertLinkedObjectAs, "Word.document.8", base.myDir + "Document.docx", null, true);

    // 2 -  A DDE field:
    builder.writeln("FieldDde:\n");
    insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", base.myDir + "Spreadsheet.xlsx",
      "Sheet1!R1C1", true, true);

    // 3 -  A DDEAUTO field:
    builder.writeln("FieldDdeAuto:\n");
    insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", base.myDir + "Spreadsheet.xlsx",
      "Sheet1!R1C1", true);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.LINK.DDE.DDEAUTO.docx");
  });


  test.each([InsertLinkedObjectAs.Picture, InsertLinkedObjectAs.Bitmap])('FieldLinkedObjectsAsImage(%o)', (insertLinkedObjectAs) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are three types of fields we can use to display contents from a linked document in the form of an image.
    // 1 -  A LINK field:
    builder.writeln("FieldLink:\n");
    insertFieldLink(builder, insertLinkedObjectAs, "Excel.Sheet", base.myDir + "MySpreadsheet.xlsx", "Sheet1!R2C2", true);

    // 2 -  A DDE field:
    builder.writeln("FieldDde:\n");
    insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", base.myDir + "Spreadsheet.xlsx", "Sheet1!R1C1", true, true);

    // 3 -  A DDEAUTO field:
    builder.writeln("FieldDdeAuto:\n");
    insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", base.myDir + "Spreadsheet.xlsx", "Sheet1!R1C1", true);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.LINK.DDE.DDEAUTO.AsImage.docx");
  });


  /// <summary>
  /// Use a document builder to insert a LINK field and set its properties according to parameters.
  /// </summary>
  function insertFieldLink(builder, insertLinkedObjectAs, progId, sourceFullName, sourceItem, shouldAutoUpdate) {
    let field = builder.insertField(aw.Fields.FieldType.FieldLink, true).asFieldLink();

    switch (insertLinkedObjectAs)
    {
      case InsertLinkedObjectAs.Text:
        field.insertAsText = true;
        break;
      case InsertLinkedObjectAs.Unicode:
        field.insertAsUnicode = true;
        break;
      case InsertLinkedObjectAs.Html:
        field.insertAsHtml = true;
        break;
      case InsertLinkedObjectAs.Rtf:
        field.insertAsRtf = true;
        break;
      case InsertLinkedObjectAs.Picture:
        field.insertAsPicture = true;
        break;
      case InsertLinkedObjectAs.Bitmap:
        field.insertAsBitmap = true;
        break;
    }

    field.autoUpdate = shouldAutoUpdate;
    field.progId = progId;
    field.sourceFullName = sourceFullName;
    field.sourceItem = sourceItem;

    builder.writeln("\n");
  }

  /// <summary>
  /// Use a document builder to insert a DDE field, and set its properties according to parameters.
  /// </summary>
  function insertFieldDde(builder, insertLinkedObjectAs, progId, sourceFullName, sourceItem, isLinked, shouldAutoUpdate) {
    let field = builder.insertField(aw.Fields.FieldType.FieldDDE, true).asFieldDde();

    switch (insertLinkedObjectAs)
    {
      case InsertLinkedObjectAs.Text:
        field.insertAsText = true;
        break;
      case InsertLinkedObjectAs.Unicode:
        field.insertAsUnicode = true;
        break;
      case InsertLinkedObjectAs.Html:
        field.insertAsHtml = true;
        break;
      case InsertLinkedObjectAs.Rtf:
        field.insertAsRtf = true;
        break;
      case InsertLinkedObjectAs.Picture:
        field.insertAsPicture = true;
        break;
      case InsertLinkedObjectAs.Bitmap:
        field.insertAsBitmap = true;
        break;
    }

    field.autoUpdate = shouldAutoUpdate;
    field.progId = progId;
    field.sourceFullName = sourceFullName;
    field.sourceItem = sourceItem;
    field.isLinked = isLinked;

    builder.writeln("\n");
  }

  /// <summary>
  /// Use a document builder to insert a DDEAUTO, field and set its properties according to parameters.
  /// </summary>
  function insertFieldDdeAuto(builder, insertLinkedObjectAs, progId, sourceFullName, sourceItem, isLinked) {
    let field = builder.insertField(aw.Fields.FieldType.FieldDDEAuto, true).asFieldDdeAuto();

    switch (insertLinkedObjectAs)
    {
      case InsertLinkedObjectAs.Text:
        field.insertAsText = true;
        break;
      case InsertLinkedObjectAs.Unicode:
        field.insertAsUnicode = true;
        break;
      case InsertLinkedObjectAs.Html:
        field.insertAsHtml = true;
        break;
      case InsertLinkedObjectAs.Rtf:
        field.insertAsRtf = true;
        break;
      case InsertLinkedObjectAs.Picture:
        field.insertAsPicture = true;
        break;
      case InsertLinkedObjectAs.Bitmap:
        field.insertAsBitmap = true;
        break;
    }

    field.progId = progId;
    field.sourceFullName = sourceFullName;
    field.sourceItem = sourceItem;
    field.isLinked = isLinked;
  }
  //ExEnd

  test('FieldUserAddress', () => {
    //ExStart
    //ExFor:FieldUserAddress
    //ExFor:FieldUserAddress.userAddress
    //ExSummary:Shows how to use the USERADDRESS field.
    let doc = new aw.Document();

    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    let userInformation = new aw.Fields.UserInformation();
    userInformation.address = "123 Main Street";
    doc.fieldOptions.currentUser = userInformation;

    // Create a USERADDRESS field to display the current user's address,
    // taken from the UserInformation object we created above.
    let builder = new aw.DocumentBuilder(doc);
    let fieldUserAddress = builder.insertField(aw.Fields.FieldType.FieldUserAddress, true).asFieldUserAddress();
    expect(fieldUserAddress.result).toEqual(userInformation.address);

    expect(fieldUserAddress.getFieldCode()).toEqual(" USERADDRESS ");
    expect(fieldUserAddress.result).toEqual("123 Main Street");

    // We can set this property to get our field to override the value currently stored in the UserInformation object.
    fieldUserAddress.userAddress = "456 North Road";
    fieldUserAddress.update();

    expect(fieldUserAddress.getFieldCode()).toEqual(" USERADDRESS  \"456 North Road\"");
    expect(fieldUserAddress.result).toEqual("456 North Road");

    // This does not affect the value in the UserInformation object.
    expect(doc.fieldOptions.currentUser.address).toEqual("123 Main Street");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.USERADDRESS.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.USERADDRESS.docx");

    fieldUserAddress = doc.range.fields.at(0).asFieldUserAddress();

    TestUtil.verifyField(aw.Fields.FieldType.FieldUserAddress, " USERADDRESS  \"456 North Road\"", "456 North Road", fieldUserAddress);
    expect(fieldUserAddress.userAddress).toEqual("456 North Road");
  });


  test('FieldUserInitials', () => {
    //ExStart
    //ExFor:FieldUserInitials
    //ExFor:FieldUserInitials.userInitials
    //ExSummary:Shows how to use the USERINITIALS field.
    let doc = new aw.Document();

    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    let userInformation = new aw.Fields.UserInformation();
    userInformation.initials = "J. D.";
    doc.fieldOptions.currentUser = userInformation;

    // Create a USERINITIALS field to display the current user's initials,
    // taken from the UserInformation object we created above.
    let builder = new aw.DocumentBuilder(doc);
    let fieldUserInitials = builder.insertField(aw.Fields.FieldType.FieldUserInitials, true).asFieldUserInitials();
    expect(fieldUserInitials.result).toEqual(userInformation.initials);

    expect(fieldUserInitials.getFieldCode()).toEqual(" USERINITIALS ");
    expect(fieldUserInitials.result).toEqual("J. D.");

    // We can set this property to get our field to override the value currently stored in the UserInformation object. 
    fieldUserInitials.userInitials = "J. C.";
    fieldUserInitials.update();

    expect(fieldUserInitials.getFieldCode()).toEqual(" USERINITIALS  \"J. C.\"");
    expect(fieldUserInitials.result).toEqual("J. C.");

    // This does not affect the value in the UserInformation object.
    expect(doc.fieldOptions.currentUser.initials).toEqual("J. D.");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.USERINITIALS.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.USERINITIALS.docx");

    fieldUserInitials = doc.range.fields.at(0).asFieldUserInitials();

    TestUtil.verifyField(aw.Fields.FieldType.FieldUserInitials, " USERINITIALS  \"J. C.\"", "J. C.", fieldUserInitials);
    expect(fieldUserInitials.userInitials).toEqual("J. C.");
  });


  test('FieldUserName', () => {
    //ExStart
    //ExFor:FieldUserName
    //ExFor:FieldUserName.userName
    //ExSummary:Shows how to use the USERNAME field.
    let doc = new aw.Document();

    // Create a UserInformation object and set it as the source of user information for any fields that we create.
    let userInformation = new aw.Fields.UserInformation();
    userInformation.name = "John Doe";
    doc.fieldOptions.currentUser = userInformation;

    let builder = new aw.DocumentBuilder(doc);

    // Create a USERNAME field to display the current user's name,
    // taken from the UserInformation object we created above.
    let fieldUserName = builder.insertField(aw.Fields.FieldType.FieldUserName, true).asFieldUserName();
    expect(fieldUserName.result).toEqual(userInformation.name);

    expect(fieldUserName.getFieldCode()).toEqual(" USERNAME ");
    expect(fieldUserName.result).toEqual("John Doe");

    // We can set this property to get our field to override the value currently stored in the UserInformation object. 
    fieldUserName.userName = "Jane Doe";
    fieldUserName.update();

    expect(fieldUserName.getFieldCode()).toEqual(" USERNAME  \"Jane Doe\"");
    expect(fieldUserName.result).toEqual("Jane Doe");

    // This does not affect the value in the UserInformation object.
    expect(doc.fieldOptions.currentUser.name).toEqual("John Doe");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.USERNAME.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.USERNAME.docx");

    fieldUserName = doc.range.fields.at(0).asFieldUserName();

    TestUtil.verifyField(aw.Fields.FieldType.FieldUserName, " USERNAME  \"Jane Doe\"", "Jane Doe", fieldUserName);
    expect(fieldUserName.userName).toEqual("Jane Doe");
  });


  test('FieldStyleRefParagraphNumbers', () => {
    //ExStart
    //ExFor:FieldStyleRef
    //ExFor:FieldStyleRef.insertParagraphNumber
    //ExFor:FieldStyleRef.insertParagraphNumberInFullContext
    //ExFor:FieldStyleRef.insertParagraphNumberInRelativeContext
    //ExFor:FieldStyleRef.insertRelativePosition
    //ExFor:FieldStyleRef.searchFromBottom
    //ExFor:FieldStyleRef.styleName
    //ExFor:FieldStyleRef.suppressNonDelimiters
    //ExSummary:Shows how to use STYLEREF fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a list based using a Microsoft Word list template.
    let list = doc.lists.add(aw.Lists.ListTemplate.NumberDefault);

    // This generated list will display "1.a )".
    // Space before the bracket is a non-delimiter character, which we can suppress. 
    list.listLevels.at(0).numberFormat = "\u0000.";
    list.listLevels.at(1).numberFormat = "\u0001 )";

    // Add text and apply paragraph styles that STYLEREF fields will reference.
    builder.listFormat.list = list;
    builder.listFormat.listIndent();
    builder.paragraphFormat.style = doc.styles.at("List Paragraph");
    builder.writeln("Item 1");
    builder.paragraphFormat.style = doc.styles.at("Quote");
    builder.writeln("Item 2");
    builder.paragraphFormat.style = doc.styles.at("List Paragraph");
    builder.writeln("Item 3");
    builder.listFormat.removeNumbers();
    builder.paragraphFormat.style = doc.styles.at("Normal");

    // Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    let field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "List Paragraph";

    // Place a STYLEREF field in the footer, and have it display the last text.
    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "List Paragraph";
    field.searchFromBottom = true;

    builder.moveToDocumentEnd();

    // We can also use STYLEREF fields to reference the list numbers of lists.
    builder.write("\nParagraph number: ");
    field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "Quote";
    field.insertParagraphNumber = true;

    builder.write("\nParagraph number, relative context: ");
    field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "Quote";
    field.insertParagraphNumberInRelativeContext = true;

    builder.write("\nParagraph number, full context: ");
    field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "Quote";
    field.insertParagraphNumberInFullContext = true;

    builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ");
    field = builder.insertField(aw.Fields.FieldType.FieldStyleRef, true).asFieldStyleRef();
    field.styleName = "Quote";
    field.insertParagraphNumberInFullContext = true;
    field.suppressNonDelimiters = true;

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.STYLEREF.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.STYLEREF.docx");

    field = doc.range.fields.at(0).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  \"List Paragraph\"", "Item 1", field);
    expect(field.styleName).toEqual("List Paragraph");

    field = doc.range.fields.at(1).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field);
    expect(field.styleName).toEqual("List Paragraph");
    expect(field.searchFromBottom).toEqual(true);

    field = doc.range.fields.at(2).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  Quote \\n", "\u200Eb )", field);
    expect(field.styleName).toEqual("Quote");
    expect(field.insertParagraphNumber).toEqual(true);

    field = doc.range.fields.at(3).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  Quote \\r", "\u200Eb )", field);
    expect(field.styleName).toEqual("Quote");
    expect(field.insertParagraphNumberInRelativeContext).toEqual(true);

    field = doc.range.fields.at(4).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  Quote \\w", "\u200E1.b )", field);
    expect(field.styleName).toEqual("Quote");
    expect(field.insertParagraphNumberInFullContext).toEqual(true);

    field = doc.range.fields.at(5).asFieldStyleRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldStyleRef, " STYLEREF  Quote \\w \\t", "\u200E1.b)", field);
    expect(field.styleName).toEqual("Quote");
    expect(field.insertParagraphNumberInFullContext).toEqual(true);
    expect(field.suppressNonDelimiters).toEqual(true);
  });


  test('FieldDate', () => {
    //ExStart
    //ExFor:FieldDate
    //ExFor:FieldDate.useLunarCalendar
    //ExFor:FieldDate.useSakaEraCalendar
    //ExFor:FieldDate.useUmAlQuraCalendar
    //ExFor:FieldDate.useLastFormat
    //ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If we want the text in the document always to display the correct date, we can use a DATE field.
    // Below are three types of cultural calendars that a DATE field can use to display a date.
    // 1 -  Islamic Lunar Calendar:
    let field = builder.insertField(aw.Fields.FieldType.FieldDate, true).asFieldDate();
    field.useLunarCalendar = true;
    expect(field.getFieldCode()).toEqual(" DATE  \\h");
    builder.writeln();

    // 2 -  Umm al-Qura calendar:
    field = builder.insertField(aw.Fields.FieldType.FieldDate, true).asFieldDate();
    field.useUmAlQuraCalendar = true;
    expect(field.getFieldCode()).toEqual(" DATE  \\u");
    builder.writeln();

    // 3 -  Indian National Calendar:
    field = builder.insertField(aw.Fields.FieldType.FieldDate, true).asFieldDate();
    field.useSakaEraCalendar = true;
    expect(field.getFieldCode()).toEqual(" DATE  \\s");
    builder.writeln();

    // Insert a DATE field and set its calendar type to the one last used by the host application.
    // In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
    field = builder.insertField(aw.Fields.FieldType.FieldDate, true).asFieldDate();
    field.useLastFormat = true;
    expect(field.getFieldCode()).toEqual(" DATE  \\l");
    builder.writeln();

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.DATE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.DATE.docx");

    field = doc.range.fields.at(0).asFieldDate();

    expect(field.type).toEqual(aw.Fields.FieldType.FieldDate);
    expect(field.useLunarCalendar).toEqual(true);
    expect(field.getFieldCode()).toEqual(" DATE  \\h");
    expect(new RegExp(String.raw`\d{1,2}[/]\d{1,2}[/]\d{4}`).test(doc.range.fields.at(0).result)).toEqual(true);

    field = doc.range.fields.at(1).asFieldDate();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, " DATE  \\u", moment(new Date()).format("D/MM/YYYY"), field);
    expect(field.useUmAlQuraCalendar).toEqual(true);

    field = doc.range.fields.at(2).asFieldDate();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, " DATE  \\s", moment(new Date()).format("D/MM/YYYY"), field);
    expect(field.useSakaEraCalendar).toEqual(true);

    field = doc.range.fields.at(3).asFieldDate();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDate, " DATE  \\l", moment(new Date()).format("D/MM/YYYY"), field);
    expect(field.useLastFormat).toEqual(true);
  });


  test.skip('FieldCreateDate: WORDSNET-17669 + System.Globalization.UmAlQuraCalendar', () => {
    //ExStart
    //ExFor:FieldCreateDate
    //ExFor:FieldCreateDate.useLunarCalendar
    //ExFor:FieldCreateDate.useSakaEraCalendar
    //ExFor:FieldCreateDate.useUmAlQuraCalendar
    //ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.writeln(" Date this document was created:");

    // We can use the CREATEDATE field to display the date and time of the creation of the document.
    // Below are three different calendar types according to which the CREATEDATE field can display the date/time.
    // 1 -  Islamic Lunar Calendar:
    builder.write("According to the Lunar Calendar - ");
    let field = builder.insertField(aw.Fields.FieldType.FieldCreateDate, true).asFieldCreateDate();
    field.useLunarCalendar = true;

    expect(field.getFieldCode()).toEqual(" CREATEDATE  \\h");

    // 2 -  Umm al-Qura calendar:
    builder.write("\nAccording to the Umm al-Qura Calendar - ");
    field = builder.insertField(aw.Fields.FieldType.FieldCreateDate, true).asFieldCreateDate();
    field.useUmAlQuraCalendar = true;

    expect(field.getFieldCode()).toEqual(" CREATEDATE  \\u");

    // 3 -  Indian National Calendar:
    builder.write("\nAccording to the Indian National Calendar - ");
    field = builder.insertField(aw.Fields.FieldType.FieldCreateDate, true).asFieldCreateDate();
    field.useSakaEraCalendar = true;

    expect(field.getFieldCode()).toEqual(" CREATEDATE  \\s");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.CREATEDATE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.CREATEDATE.docx");

    expect(doc.builtInDocumentProperties.createdTime).toEqual(new Date(2017, 12 - 1, 5, 9, 56, 0))

    let expectedDate = doc.builtInDocumentProperties.createdTime.AddHours(TimeZoneInfo.local.GetUtcOffset(DateTime.UtcNow).Hours);
    field = doc.range.fields.at(0).asFieldCreateDate();
    let umAlQuraCalendar = new UmAlQuraCalendar();

    TestUtil.verifyField(aw.Fields.FieldType.FieldCreateDate, " CREATEDATE  \\h",
      `${umAlQuraCalendar.GetMonth(expectedDate)}/${umAlQuraCalendar.GetDayOfMonth(expectedDate)}/${umAlQuraCalendar.GetYear(expectedDate)} ` +
      expectedDate.AddHours(1).ToString("hh:mm:ss tt"), field);
    expect(field.type).toEqual(aw.Fields.FieldType.FieldCreateDate);
    expect(field.useLunarCalendar).toEqual(true);
            
    field = doc.range.fields.at(1).asFieldCreateDate();

    TestUtil.VerifyField(aw.Fields.FieldType.FieldCreateDate, " CREATEDATE  \\u",
      `${umAlQuraCalendar.GetMonth(expectedDate)}/${umAlQuraCalendar.GetDayOfMonth(expectedDate)}/${umAlQuraCalendar.GetYear(expectedDate)} ` +
      expectedDate.AddHours(1).ToString("hh:mm:ss tt"), field);
    expect(field.type).toEqual(aw.Fields.FieldType.FieldCreateDate);
    expect(field.useUmAlQuraCalendar).toEqual(true);
  });


  test.skip('FieldSaveDate: WORDSNET-17669', () => {
    //ExStart
    //ExFor:BuiltInDocumentProperties.lastSavedTime
    //ExFor:FieldSaveDate
    //ExFor:FieldSaveDate.useLunarCalendar
    //ExFor:FieldSaveDate.useSakaEraCalendar
    //ExFor:FieldSaveDate.useUmAlQuraCalendar
    //ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.writeln(" Date this document was last saved:");

    // We can use the SAVEDATE field to display the last save operation's date and time on the document.
    // The save operation that these fields refer to is the manual save in an application like Microsoft Word,
    // not the document's Save method.
    // Below are three different calendar types according to which the SAVEDATE field can display the date/time.
    // 1 -  Islamic Lunar Calendar:
    builder.write("According to the Lunar Calendar - ");
    let field = builder.insertField(aw.Fields.FieldType.FieldSaveDate, true).asFieldSaveDate();
    field.useLunarCalendar = true;

    expect(field.getFieldCode()).toEqual(" SAVEDATE  \\h");

    // 2 -  Umm al-Qura calendar:
    builder.write("\nAccording to the Umm al-Qura calendar - ");
    field = builder.insertField(aw.Fields.FieldType.FieldSaveDate, true).asFieldSaveDate();
    field.useUmAlQuraCalendar = true;

    expect(field.getFieldCode()).toEqual(" SAVEDATE  \\u");

    // 3 -  Indian National calendar:
    builder.write("\nAccording to the Indian National calendar - ");
    field = builder.insertField(aw.Fields.FieldType.FieldSaveDate, true).asFieldSaveDate();
    field.useSakaEraCalendar = true;

    expect(field.getFieldCode()).toEqual(" SAVEDATE  \\s");

    // The SAVEDATE fields draw their date/time values from the LastSavedTime built-in property.
    // The document's Save method will not update this value, but we can still update it manually.
    doc.builtInDocumentProperties.lastSavedTime = Date.now();

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.SAVEDATE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SAVEDATE.docx");

    console.log(doc.builtInDocumentProperties.lastSavedTime);

    field = doc.range.fields.at(0).asFieldSaveDate();

    expect(field.type).toEqual(aw.Fields.FieldType.FieldSaveDate);
    expect(field.useLunarCalendar).toEqual(true);
    expect(field.getFieldCode()).toEqual(" SAVEDATE  \\h");

    expect(new RegExp("\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").test(field.result)).toEqual(true);

    field = doc.range.fields.at(1).asFieldSaveDate();

    expect(field.type).toEqual(aw.Fields.FieldType.FieldSaveDate);
    expect(field.useUmAlQuraCalendar).toEqual(true);
    expect(field.getFieldCode()).toEqual(" SAVEDATE  \\u");
    expect(new RegExp("\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").test(field.result)).toEqual(true);
  });


  test.skip('FieldBuilder: WORDSNODEJS-114', () => {
    //ExStart
    //ExFor:FieldBuilder
    //ExFor:FieldBuilder.addArgument(Int32)
    //ExFor:FieldBuilder.addArgument(FieldArgumentBuilder)
    //ExFor:FieldBuilder.addArgument(String)
    //ExFor:FieldBuilder.addArgument(Double)
    //ExFor:FieldBuilder.addArgument(FieldBuilder)
    //ExFor:FieldBuilder.addSwitch(String)
    //ExFor:FieldBuilder.addSwitch(String, Double)
    //ExFor:FieldBuilder.addSwitch(String, Int32)
    //ExFor:FieldBuilder.addSwitch(String, String)
    //ExFor:FieldBuilder.buildAndInsert(Paragraph)
    //ExFor:FieldArgumentBuilder
    //ExFor:FieldArgumentBuilder.#ctor
    //ExFor:FieldArgumentBuilder.addField(FieldBuilder)
    //ExFor:FieldArgumentBuilder.addText(String)
    //ExFor:FieldArgumentBuilder.addNode(Inline)
    //ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
    let doc = new aw.Document();

    // Below are three examples of field construction done using a field builder.
    // 1 -  Single field:
    // Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
    let builder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldSymbol);
    builder.addArgument(402);
    builder.addSwitch("\\f", "Arial");
    builder.addSwitch("\\s", 25);
    builder.addSwitch("\\u");
    let field = builder.buildAndInsert(doc.firstSection.body.firstParagraph);

    expect(field.getFieldCode()).toEqual(" SYMBOL 402 \\f Arial \\s 25 \\u ");

    // 2 -  Nested field:
    // Use a field builder to create a formula field used as an inner field by another field builder.
    let innerFormulaBuilder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldFormula);
    innerFormulaBuilder.addArgument(100);
    innerFormulaBuilder.addArgument("+");
    innerFormulaBuilder.addArgument(74);

    // Create another builder for another SYMBOL field, and insert the formula field
    // that we have created above into the SYMBOL field as its argument. 
    builder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldSymbol);
    builder.addArgument(innerFormulaBuilder);
    field = builder.buildAndInsert(doc.firstSection.body.appendParagraph(''));

    // The outer SYMBOL field will use the formula field result, 174, as its argument,
    // which will make the field display the ® (Registered Sign) symbol since its character number is 174.
    expect(field.getFieldCode()).toEqual(" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ");

    // 3 -  Multiple nested fields and arguments:
    // Now, we will use a builder to create an IF field, which displays one of two custom string values,
    // depending on the true/false value of its expression. To get a true/false value
    // that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
    // We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
    let leftExpression = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldFormula);
    leftExpression.addArgument(2);
    leftExpression.addArgument("+");
    leftExpression.addArgument(3);

    let rightExpression = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldFormula);
    rightExpression.addArgument(2.5);
    rightExpression.addArgument("*");
    rightExpression.addArgument(5.2);

    // Next, we will build two field arguments, which will serve as the true/false output strings for the IF field.
    // These arguments will reuse the output values of our numeric expressions.
    let trueOutput = new aw.Fields.FieldArgumentBuilder();
    trueOutput.addText("True, both expressions amount to ");
    trueOutput.addField(leftExpression);

    let falseOutput = new aw.Fields.FieldArgumentBuilder();
    falseOutput.addNode(new aw.Run(doc, "False, "));
    falseOutput.addField(leftExpression);
    falseOutput.addNode(new aw.Run(doc, " does not equal "));
    falseOutput.addField(rightExpression);

    // Finally, we will create one more field builder for the IF field and combine all of the expressions. 
    builder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf);
    builder.addArgument(leftExpression);
    builder.addArgument("=");
    builder.addArgument(rightExpression);
    builder.addArgument(trueOutput);
    builder.addArgument(falseOutput);
    field = builder.buildAndInsert(doc.firstSection.body.appendParagraph(''));

    expect(field.getFieldCode()).toEqual(" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 " +
                            "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                            "\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" ");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.SYMBOL.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SYMBOL.docx");

    let fieldSymbol = doc.range.fields.at(0).asFieldSymbol();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSymbol, " SYMBOL 402 \\f Arial \\s 25 \\u ", '', fieldSymbol);
    expect(fieldSymbol.displayResult).toEqual("ƒ");

    fieldSymbol = doc.range.fields.at(1).asFieldSymbol();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSymbol, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", '', fieldSymbol);
    expect(fieldSymbol.displayResult).toEqual("®");

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 100 + 74 ", "174", doc.range.fields.at(2));

    TestUtil.verifyField(aw.Fields.FieldType.FieldIf,
      " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
      "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
      "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
      "False, 5 does not equal 13", doc.range.fields.at(3));

    expect(() => TestUtil.fieldsAreNested(doc.range.fields.at(2), doc.range.fields.at(3))).toThrow("TODO: AssertionException");

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2 + 3 ", "5", doc.range.fields.at(4));
    TestUtil.fieldsAreNested(doc.range.fields.at(4), doc.range.fields.at(3));

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2.5 * 5.2 ", "13", doc.range.fields.at(5));
    TestUtil.fieldsAreNested(doc.range.fields.at(5), doc.range.fields.at(3));

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2 + 3 ", '', doc.range.fields.at(6));
    TestUtil.fieldsAreNested(doc.range.fields.at(6), doc.range.fields.at(3));

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2 + 3 ", "5", doc.range.fields.at(7));
    TestUtil.fieldsAreNested(doc.range.fields.at(7), doc.range.fields.at(3));

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2.5 * 5.2 ", "13", doc.range.fields.at(8));
    TestUtil.fieldsAreNested(doc.range.fields.at(8), doc.range.fields.at(3));
  });


  test('FieldAuthor', () => {
    //ExStart
    //ExFor:FieldAuthor
    //ExFor:FieldAuthor.authorName  
    //ExFor:FieldOptions.defaultDocumentAuthor
    //ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // AUTHOR fields source their results from the built-in document property called "Author".
    // If we create and save a document in Microsoft Word,
    // it will have our username in that property.
    // However, if we create a document programmatically using Aspose.words,
    // the "Author" property, by default, will be an empty string. 
    expect(doc.builtInDocumentProperties.author).toEqual('');

    // Set a backup author name for AUTHOR fields to use
    // if the "Author" property contains an empty string.
    doc.fieldOptions.defaultDocumentAuthor = "Joe Bloggs";

    builder.write("This document was created by ");
    let field = builder.insertField(aw.Fields.FieldType.FieldAuthor, true).asFieldAuthor();
    field.update();

    expect(field.getFieldCode()).toEqual(" AUTHOR ");
    expect(field.result).toEqual("Joe Bloggs");

    // Updating an AUTHOR field that contains a value
    // will apply that value to the "Author" built-in property.
    expect(doc.builtInDocumentProperties.author).toEqual("Joe Bloggs");

    // Changing this property, then updating the AUTHOR field will apply this value to the field.
    doc.builtInDocumentProperties.author = "John Doe";
    field.update();

    expect(field.getFieldCode()).toEqual(" AUTHOR ");
    expect(field.result).toEqual("John Doe");

    // If we update an AUTHOR field after changing its "Name" property,
    // then the field will display the new name and apply the new name to the built-in property.
    field.authorName = "Jane Doe";
    field.update();

    expect(field.getFieldCode()).toEqual(" AUTHOR  \"Jane Doe\"");
    expect(field.result).toEqual("Jane Doe");

    // AUTHOR fields do not affect the DefaultDocumentAuthor property.
    expect(doc.builtInDocumentProperties.author).toEqual("Jane Doe");
    expect(doc.fieldOptions.defaultDocumentAuthor).toEqual("Joe Bloggs");

    doc.save(base.artifactsDir + "Field.AUTHOR.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.AUTHOR.docx");

    expect(doc.fieldOptions.defaultDocumentAuthor).toBe(null);
    expect(doc.builtInDocumentProperties.author).toEqual("Jane Doe");

    field = doc.range.fields.at(0).asFieldAuthor();

    TestUtil.verifyField(aw.Fields.FieldType.FieldAuthor, " AUTHOR  \"Jane Doe\"", "Jane Doe", field);
    expect(field.authorName).toEqual("Jane Doe");
  });


  test('FieldDocVariable', () => {
    //ExStart
    //ExFor:FieldDocProperty
    //ExFor:FieldDocVariable
    //ExFor:FieldDocVariable.variableName
    //ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are two ways of using DOCPROPERTY fields.
    // 1 -  Display a built-in property:
    // Set a custom value for the "Category" built-in property, then insert a DOCPROPERTY field that references it.
    doc.builtInDocumentProperties.category = "My category";

    let fieldDocProperty = builder.insertField(" DOCPROPERTY Category ").asFieldDocProperty();
    fieldDocProperty.update();

    expect(fieldDocProperty.getFieldCode()).toEqual(" DOCPROPERTY Category ");
    expect(fieldDocProperty.result).toEqual("My category");

    builder.insertParagraph();

    // 2 -  Display a custom document variable:
    // Define a custom variable, then reference that variable with a DOCPROPERTY field.
    expect(doc.variables.count).toEqual(0);
    doc.variables.add("My variable", "My variable's value");

    let fieldDocVariable = builder.insertField(aw.Fields.FieldType.FieldDocVariable, true).asFieldDocVariable();
    fieldDocVariable.variableName = "My Variable";
    fieldDocVariable.update();

    expect(fieldDocVariable.getFieldCode()).toEqual(" DOCVARIABLE  \"My Variable\"");
    expect(fieldDocVariable.result).toEqual("My variable's value");

    doc.save(base.artifactsDir + "Field.DOCPROPERTY.DOCVARIABLE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.DOCPROPERTY.DOCVARIABLE.docx");

    expect(doc.builtInDocumentProperties.category).toEqual("My category");

    fieldDocProperty = doc.range.fields.at(0).asFieldDocProperty();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDocProperty, " DOCPROPERTY Category ", "My category", fieldDocProperty);

    fieldDocVariable = doc.range.fields.at(1).asFieldDocVariable();

    TestUtil.verifyField(aw.Fields.FieldType.FieldDocVariable, " DOCVARIABLE  \"My Variable\"", "My variable's value", fieldDocVariable);
    expect(fieldDocVariable.variableName).toEqual("My Variable");
  });


  test('FieldSubject', () => {
    //ExStart
    //ExFor:FieldSubject
    //ExFor:FieldSubject.text
    //ExSummary:Shows how to use the SUBJECT field.
    let doc = new aw.Document();

    // Set a value for the document's "Subject" built-in property.
    doc.builtInDocumentProperties.subject = "My subject";

    // Create a SUBJECT field to display the value of that built-in property.
    let builder = new aw.DocumentBuilder(doc);
    let field = builder.insertField(aw.Fields.FieldType.FieldSubject, true).asFieldSubject();
    field.update();

    expect(field.getFieldCode()).toEqual(" SUBJECT ");
    expect(field.result).toEqual("My subject");

    // If we give the SUBJECT field's Text property value and update it, the field will
    // overwrite the current value of the "Subject" built-in property with the value of its Text property,
    // and then display the new value.
    field.text = "My new subject";
    field.update();

    expect(field.getFieldCode()).toEqual(" SUBJECT  \"My new subject\"");
    expect(field.result).toEqual("My new subject");

    expect(doc.builtInDocumentProperties.subject).toEqual("My new subject");

    doc.save(base.artifactsDir + "Field.SUBJECT.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SUBJECT.docx");

    expect(doc.builtInDocumentProperties.subject).toEqual("My new subject");

    field = doc.range.fields.at(0).asFieldSubject();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSubject, " SUBJECT  \"My new subject\"", "My new subject", field);
    expect(field.text).toEqual("My new subject");
  });


  test('FieldComments', () => {
    //ExStart
    //ExFor:FieldComments
    //ExFor:FieldComments.text
    //ExSummary:Shows how to use the COMMENTS field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set a value for the document's "Comments" built-in property.
    doc.builtInDocumentProperties.comments = "My comment.";

    // Create a COMMENTS field to display the value of that built-in property.
    let field = builder.insertField(aw.Fields.FieldType.FieldComments, true).asFieldComments();
    field.update();

    expect(field.getFieldCode()).toEqual(" COMMENTS ");
    expect(field.result).toEqual("My comment.");

    // If we give the COMMENTS field's Text property value and update it, the field will
    // overwrite the current value of the "Comments" built-in property with the value of its Text property,
    // and then display the new value.
    field.text = "My overriding comment.";
    field.update();

    expect(field.getFieldCode()).toEqual(" COMMENTS  \"My overriding comment.\"");
    expect(field.result).toEqual("My overriding comment.");

    doc.save(base.artifactsDir + "Field.COMMENTS.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.COMMENTS.docx");

    expect(doc.builtInDocumentProperties.comments).toEqual("My overriding comment.");

    field = doc.range.fields.at(0).asFieldComments();

    TestUtil.verifyField(aw.Fields.FieldType.FieldComments, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field);
    expect(field.text).toEqual("My overriding comment.");
  });


  test('FieldFileSize', () => {
    //ExStart
    //ExFor:FieldFileSize
    //ExFor:FieldFileSize.isInKilobytes
    //ExFor:FieldFileSize.isInMegabytes
    //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
    let doc = new aw.Document(base.myDir + "Document.docx");

    expect(doc.builtInDocumentProperties.bytes).toEqual(18105);

    let builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.insertParagraph();

    // Below are three different units of measure
    // with which FILESIZE fields can display the document's file size.
    // 1 -  Bytes:
    let field = builder.insertField(aw.Fields.FieldType.FieldFileSize, true).asFieldFileSize();
    field.update();

    expect(field.getFieldCode()).toEqual(" FILESIZE ");
    expect(field.result).toEqual("18105");

    // 2 -  Kilobytes:
    builder.insertParagraph();
    field = builder.insertField(aw.Fields.FieldType.FieldFileSize, true).asFieldFileSize();
    field.isInKilobytes = true;
    field.update();

    expect(field.getFieldCode()).toEqual(" FILESIZE  \\k");
    expect(field.result).toEqual("18");

    // 3 -  Megabytes:
    builder.insertParagraph();
    field = builder.insertField(aw.Fields.FieldType.FieldFileSize, true).asFieldFileSize();
    field.isInMegabytes = true;
    field.update();

    expect(field.getFieldCode()).toEqual(" FILESIZE  \\m");
    expect(field.result).toEqual("0");

    // To update the values of these fields while editing in Microsoft Word,
    // we must first save the changes, and then manually update these fields.
    doc.save(base.artifactsDir + "Field.FILESIZE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.FILESIZE.docx");

    field = doc.range.fields.at(0).asFieldFileSize();

    TestUtil.verifyField(aw.Fields.FieldType.FieldFileSize, " FILESIZE ", "18105", field);

    // These fields will need to be updated to produce an accurate result.
    doc.updateFields();

    field = doc.range.fields.at(1).asFieldFileSize();

    TestUtil.verifyField(aw.Fields.FieldType.FieldFileSize, " FILESIZE  \\k", "13", field);
    expect(field.isInKilobytes).toEqual(true);

    field = doc.range.fields.at(2).asFieldFileSize();

    TestUtil.verifyField(aw.Fields.FieldType.FieldFileSize, " FILESIZE  \\m", "0", field);
    expect(field.isInMegabytes).toEqual(true);
  });


  test('FieldGoToButton', () => {
    //ExStart
    //ExFor:FieldGoToButton
    //ExFor:FieldGoToButton.displayText
    //ExFor:FieldGoToButton.location
    //ExSummary:Shows to insert a GOTOBUTTON field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
    // it will take the text cursor to the bookmark whose name the Location property references.
    let field = builder.insertField(aw.Fields.FieldType.FieldGoToButton, true).asFieldGoToButton();
    field.displayText = "My Button";
    field.location = "MyBookmark";

    expect(field.getFieldCode()).toEqual(" GOTOBUTTON  MyBookmark My Button");

    // Insert a valid bookmark for the field to reference.
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.startBookmark(field.location);
    builder.writeln("Bookmark text contents.");
    builder.endBookmark(field.location);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.GOTOBUTTON.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.GOTOBUTTON.docx");
    field = doc.range.fields.at(0).asFieldGoToButton();

    TestUtil.verifyField(aw.Fields.FieldType.FieldGoToButton, " GOTOBUTTON  MyBookmark My Button", '', field);
    expect(field.displayText).toEqual("My Button");
    expect(field.location).toEqual("MyBookmark");
  });


  /*  //ExStart
    //ExFor:FieldFillIn
    //ExFor:FieldFillIn.DefaultResponse
    //ExFor:FieldFillIn.PromptOnceOnMailMerge
    //ExFor:FieldFillIn.PromptText
    //ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
  test('FieldFillIn', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a FILLIN field. When we manually update this field in Microsoft Word,
    // it will prompt us to enter a response. The field will then display the response as text.
    let field = (FieldFillIn)builder.insertField(aw.Fields.FieldType.FieldFillIn, true);
    field.promptText = "Please enter a response:";
    field.defaultResponse = "A default response.";

    // We can also use these fields to ask the user for a unique response for each page
    // created during a mail merge done using Microsoft Word.
    field.promptOnceOnMailMerge = true;

    expect(field.getFieldCode()).toEqual(" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o");

    let mergeField = (FieldMergeField)builder.insertField(aw.Fields.FieldType.FieldMergeField, true);
    mergeField.fieldName = "MergeField";

    // If we perform a mail merge programmatically, we can use a custom prompt respondent
    // to automatically edit responses for FILLIN fields that the mail merge encounters.
    doc.fieldOptions.userPromptRespondent = new PromptRespondent();
    doc.mailMerge.execute(new [] { "MergeField" }, new object[] { "" });

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.FILLIN.docx");
    TestFieldFillIn(new aw.Document(base.artifactsDir + "Field.FILLIN.docx")); //ExSkip
  });


    /// <summary>
    /// Prepends a line to the default response of every FILLIN field during a mail merge.
    /// </summary>
  private class PromptRespondent : IFieldUserPromptRespondent
  {
    public string Respond(string promptText, string defaultResponse)
    {
      return "Response modified by PromptRespondent. " + defaultResponse;
    }
  }
    //ExEnd

  private void TestFieldFillIn(Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);

    expect(doc.range.fields.count).toEqual(1);

    let field = (FieldFillIn)doc.range.fields.at(0);

    TestUtil.VerifyField(aw.Fields.FieldType.FieldFillIn, " FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", 
      "Response modified by PromptRespondent. A default response.", field);
    expect(field.promptText).toEqual("Please enter a response:");
    expect(field.defaultResponse).toEqual("A default response.");
    expect(field.promptOnceOnMailMerge).toEqual(true);
  }*/

  test('FieldInfo', () => {
    //ExStart
    //ExFor:FieldInfo
    //ExFor:FieldInfo.infoType
    //ExFor:FieldInfo.newValue
    //ExSummary:Shows how to work with INFO fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Set a value for the "Comments" built-in property and then insert an INFO field to display that property's value.
    doc.builtInDocumentProperties.comments = "My comment";
    let field = builder.insertField(aw.Fields.FieldType.FieldInfo, true).asFieldInfo();
    field.infoType = "Comments";
    field.update();

    expect(field.getFieldCode()).toEqual(" INFO  Comments");
    expect(field.result).toEqual("My comment");

    builder.writeln();

    // Setting a value for the field's NewValue property and updating
    // the field will also overwrite the corresponding built-in property with the new value.
    field = builder.insertField(aw.Fields.FieldType.FieldInfo, true).asFieldInfo();
    field.infoType = "Comments";
    field.newValue = "New comment";
    field.update();

    expect(field.getFieldCode()).toEqual(" INFO  Comments \"New comment\"");
    expect(field.result).toEqual("New comment");
    expect(doc.builtInDocumentProperties.comments).toEqual("New comment");

    doc.save(base.artifactsDir + "Field.INFO.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.INFO.docx");

    expect(doc.builtInDocumentProperties.comments).toEqual("New comment");

    field = doc.range.fields.at(0).asFieldInfo();

    TestUtil.verifyField(aw.Fields.FieldType.FieldInfo, " INFO  Comments", "My comment", field);
    expect(field.infoType).toEqual("Comments");

    field = doc.range.fields.at(1).asFieldInfo();

    TestUtil.verifyField(aw.Fields.FieldType.FieldInfo, " INFO  Comments \"New comment\"", "New comment", field);
    expect(field.infoType).toEqual("Comments");
    expect(field.newValue).toEqual("New comment");
  });


  test('FieldMacroButton', () => {
    //ExStart
    //ExFor:Document.hasMacros
    //ExFor:FieldMacroButton
    //ExFor:FieldMacroButton.displayText
    //ExFor:FieldMacroButton.macroName
    //ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
    let doc = new aw.Document(base.myDir + "Macro.docm");
    let builder = new aw.DocumentBuilder(doc);

    expect(doc.hasMacros).toEqual(true);

    // Insert a MACROBUTTON field, and reference one of the document's macros by name in the MacroName property.
    let field = builder.insertField(aw.Fields.FieldType.FieldMacroButton, true).asFieldMacroButton();
    field.macroName = "MyMacro";
    field.displayText = "Double click to run macro: " + field.macroName;

    expect(field.getFieldCode()).toEqual(" MACROBUTTON  MyMacro Double click to run macro: MyMacro");

    // Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
    // We can find all other macros via View -> Macros (dropdown) -> View Macros.
    // In that menu, select "Word Commands" from the "Macros in:" drop down.
    // If our document contains a custom macro with the same name as a stock macro,
    // our macro will be the one that the MACROBUTTON field runs.
    builder.insertParagraph();
    field = builder.insertField(aw.Fields.FieldType.FieldMacroButton, true).asFieldMacroButton();
    field.macroName = "ViewZoom200";
    field.displayText = "Run " + field.macroName;

    expect(field.getFieldCode()).toEqual(" MACROBUTTON  ViewZoom200 Run ViewZoom200");

    // Save the document as a macro-enabled document type.
    doc.save(base.artifactsDir + "Field.MACROBUTTON.docm");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MACROBUTTON.docm");

    field = doc.range.fields.at(0).asFieldMacroButton();

    TestUtil.verifyField(aw.Fields.FieldType.FieldMacroButton, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", '', field);
    expect(field.macroName).toEqual("MyMacro");
    expect(field.displayText).toEqual("Double click to run macro: MyMacro");

    field = doc.range.fields.at(1).asFieldMacroButton();

    TestUtil.verifyField(aw.Fields.FieldType.FieldMacroButton, " MACROBUTTON  ViewZoom200 Run ViewZoom200", '', field);
    expect(field.macroName).toEqual("ViewZoom200");
    expect(field.displayText).toEqual("Run ViewZoom200");
  });


  test('FieldKeywords', () => {
    //ExStart
    //ExFor:FieldKeywords
    //ExFor:FieldKeywords.text
    //ExSummary:Shows to insert a KEYWORDS field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add some keywords, also referred to as "tags" in File Explorer.
    doc.builtInDocumentProperties.keywords = "Keyword1, Keyword2";

    // The KEYWORDS field displays the value of this property.
    let field = builder.insertField(aw.Fields.FieldType.FieldKeyword, true).asFieldKeywords();
    field.update();

    expect(field.getFieldCode()).toEqual(" KEYWORDS ");
    expect(field.result).toEqual("Keyword1, Keyword2");

    // Setting a value for the field's Text property,
    // and then updating the field will also overwrite the corresponding built-in property with the new value.
    field.text = "OverridingKeyword";
    field.update();

    expect(field.getFieldCode()).toEqual(" KEYWORDS  OverridingKeyword");
    expect(field.result).toEqual("OverridingKeyword");
    expect(doc.builtInDocumentProperties.keywords).toEqual("OverridingKeyword");

    doc.save(base.artifactsDir + "Field.KEYWORDS.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.KEYWORDS.docx");

    expect(doc.builtInDocumentProperties.keywords).toEqual("OverridingKeyword");

    field = doc.range.fields.at(0).asFieldKeywords();

    TestUtil.verifyField(aw.Fields.FieldType.FieldKeyword, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field);
    expect(field.text).toEqual("OverridingKeyword");
  });


  test('FieldNum', () => {
    //ExStart
    //ExFor:FieldPage
    //ExFor:FieldNumChars
    //ExFor:FieldNumPages
    //ExFor:FieldNumWords
    //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
    let doc = new aw.Document(base.myDir + "Paragraphs.docx");
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.FooterPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Center;

    // Below are three types of fields that we can use to track the size of our documents.
    // 1 -  Track the character count with a NUMCHARS field:
    let fieldNumChars = builder.insertField(aw.Fields.FieldType.FieldNumChars, true).asFieldNumChars();
    builder.writeln(" characters");

    // 2 -  Track the word count with a NUMWORDS field:
    let fieldNumWords = builder.insertField(aw.Fields.FieldType.FieldNumWords, true).asFieldNumWords();
    builder.writeln(" words");

    // 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
    // and the total number of pages in the document:
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Right;
    builder.write("Page ");
    let fieldPage = builder.insertField(aw.Fields.FieldType.FieldPage, true).asFieldPage();
    builder.write(" of ");
    let fieldNumPages = builder.insertField(aw.Fields.FieldType.FieldNumPages, true).asFieldNumPages();

    expect(fieldNumChars.getFieldCode()).toEqual(" NUMCHARS ");
    expect(fieldNumWords.getFieldCode()).toEqual(" NUMWORDS ");
    expect(fieldNumPages.getFieldCode()).toEqual(" NUMPAGES ");
    expect(fieldPage.getFieldCode()).toEqual(" PAGE ");

    // These fields will not maintain accurate values in real time
    // while we edit the document programmatically using Aspose.words, or in Microsoft Word.
    // We need to update them every we need to see an up-to-date value. 
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldNumChars, " NUMCHARS ", "6009", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldNumWords, " NUMWORDS ", "1054", doc.range.fields.at(1));

    TestUtil.verifyField(aw.Fields.FieldType.FieldPage, " PAGE ", "6", doc.range.fields.at(2));
    TestUtil.verifyField(aw.Fields.FieldType.FieldNumPages, " NUMPAGES ", "6", doc.range.fields.at(3));
  });


  test('FieldPrint', () => {
    //ExStart
    //ExFor:FieldPrint
    //ExFor:FieldPrint.postScriptGroup
    //ExFor:FieldPrint.printerInstructions
    //ExSummary:Shows to insert a PRINT field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("My paragraph");

    // The PRINT field can send instructions to the printer.
    let field = builder.insertField(aw.Fields.FieldType.FieldPrint, true).asFieldPrint();

    // Set the area for the printer to perform instructions over.
    // In this case, it will be the paragraph that contains our PRINT field.
    field.postScriptGroup = "para";

    // When we use a printer that supports PostScript to print our document,
    // this command will turn the entire area that we specified in "field.postScriptGroup" white.
    field.printerInstructions = "erasepage";

    expect(field.getFieldCode()).toEqual(" PRINT  erasepage \\p para");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.PRINT.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.PRINT.docx");

    field = doc.range.fields.at(0).asFieldPrint();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPrint, " PRINT  erasepage \\p para", '', field);
    expect(field.postScriptGroup).toEqual("para");
    expect(field.printerInstructions).toEqual("erasepage");
  });


  test('FieldPrintDate', () => {
    //ExStart
    //ExFor:FieldPrintDate
    //ExFor:FieldPrintDate.useLunarCalendar
    //ExFor:FieldPrintDate.useSakaEraCalendar
    //ExFor:FieldPrintDate.useUmAlQuraCalendar
    //ExSummary:Shows read PRINTDATE fields.
    let doc = new aw.Document(base.myDir + "Field sample - PRINTDATE.docx");

    // When a document is printed by a printer or printed as a PDF (but not exported to PDF),
    // PRINTDATE fields will display the print operation's date/time.
    // If no printing has taken place, these fields will display "0/0/0000".
    let field = doc.range.fields.at(0).asFieldPrintDate();

    expect(field.result).toEqual("3/25/2020 12:00:00 AM");
    expect(field.getFieldCode()).toEqual(" PRINTDATE ");

    // Below are three different calendar types according to which the PRINTDATE field
    // can display the date and time of the last printing operation.
    // 1 -  Islamic Lunar Calendar:
    field = doc.range.fields.at(1).asFieldPrintDate();

    expect(field.useLunarCalendar).toEqual(true);
    expect(field.result).toEqual("8/1/1441 12:00:00 AM");
    expect(field.getFieldCode()).toEqual(" PRINTDATE  \\h");

    field = doc.range.fields.at(2).asFieldPrintDate();

    // 2 -  Umm al-Qura calendar:
    expect(field.useUmAlQuraCalendar).toEqual(true);
    expect(field.result).toEqual("8/1/1441 12:00:00 AM");
    expect(field.getFieldCode()).toEqual(" PRINTDATE  \\u");

    field = doc.range.fields.at(3).asFieldPrintDate();

    // 3 -  Indian National Calendar:
    expect(field.useSakaEraCalendar).toEqual(true);
    expect(field.result).toEqual("1/5/1942 12:00:00 AM");
    expect(field.getFieldCode()).toEqual(" PRINTDATE  \\s");
    //ExEnd
  });


  test('FieldQuote', () => {
    //ExStart
    //ExFor:FieldQuote
    //ExFor:FieldQuote.text
    //ExFor:Document.updateFields
    //ExSummary:Shows to use the QUOTE field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a QUOTE field, which will display the value of its Text property.
    let field = builder.insertField(aw.Fields.FieldType.FieldQuote, true).asFieldQuote();
    field.text = "\"Quoted text\"";

    expect(field.getFieldCode()).toEqual(" QUOTE  \"\\\"Quoted text\\\"\"");

    // Insert a QUOTE field and nest a DATE field inside it.
    // DATE fields update their value to the current date every time we open the document using Microsoft Word.
    // Nesting the DATE field inside the QUOTE field like this will freeze its value
    // to the date when we created the document.
    builder.write("\nDocument creation date: ");
    field = builder.insertField(aw.Fields.FieldType.FieldQuote, true).asFieldQuote();
    builder.moveTo(field.separator);
    builder.insertField(aw.Fields.FieldType.FieldDate, true);

    expect(field.getFieldCode()).toEqual(" QUOTE \u0013 DATE \u0014" + moment(new Date()).format("D/MM/YYYY") + "\u0015");

    // Update all the fields to display their correct results.
    doc.updateFields();

    expect(doc.range.fields.at(0).result).toEqual("\"Quoted text\"");

    doc.save(base.artifactsDir + "Field.QUOTE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.QUOTE.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldQuote, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.range.fields.at(0));

    TestUtil.verifyField(aw.Fields.FieldType.FieldQuote, " QUOTE \u0013 DATE \u0014" + moment(new Date()).format("D/MM/YYYY") + "\u0015", moment(new Date()).format("D/MM/YYYY"), doc.range.fields.at(1));

  });


  //ExStart
  //ExFor:FieldNext
  //ExFor:FieldNextIf
  //ExFor:FieldNextIf.ComparisonOperator
  //ExFor:FieldNextIf.LeftExpression
  //ExFor:FieldNextIf.RightExpression
  //ExSummary:Shows how to use NEXT/NEXTIF fields to merge multiple rows into one page during a mail merge.
  test.skip('FieldNext: DataTable', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a data source for our mail merge with 3 rows.
    // A mail merge that uses this table would normally create a 3-page document.
    let table = new DataTable("Employees");
    table.columns.add("Courtesy Title");
    table.columns.add("First Name");
    table.columns.add("Last Name");
    table.rows.add("Mr.", "John", "Doe");
    table.rows.add("Mrs.", "Jane", "Cardholder");
    table.rows.add("Mr.", "Joe", "Bloggs");

    insertMergeFields(builder, "First row: ");

    // If we have multiple merge fields with the same FieldName,
    // they will receive data from the same row of the data source and display the same value after the merge.
    // A NEXT field tells the mail merge instantly to move down one row,
    // which means any MERGEFIELDs that follow the NEXT field will receive data from the next row.
    // Make sure never to try to skip to the next row while already on the last row.
    let fieldNext = builder.insertField(aw.Fields.FieldType.FieldNext, true).FieldNext();

    expect(fieldNext.getFieldCode()).toEqual(" NEXT ");

    // After the merge, the data source values that these MERGEFIELDs accept
    // will end up on the same page as the MERGEFIELDs above. 
    insertMergeFields(builder, "Second row: ");

    // A NEXTIF field has the same function as a NEXT field,
    // but it skips to the next row only if a statement constructed by the following 3 properties is true.
    let fieldNextIf = builder.insertField(aw.Fields.FieldType.FieldNextIf, true).asFieldNextIf();
    fieldNextIf.leftExpression = "5";
    fieldNextIf.rightExpression = "2 + 3";
    fieldNextIf.comparisonOperator = "=";

    expect(fieldNextIf.getFieldCode()).toEqual(" NEXTIF  5 = \"2 + 3\"");

    // If the comparison asserted by the above field is correct,
    // the following 3 merge fields will take data from the third row.
    // Otherwise, these fields will take data from row 2 again.
    insertMergeFields(builder, "Third row: ");

    doc.mailMerge.execute(table);

    // Our data source has 3 rows, and we skipped rows twice. 
    // Our output document will have 1 page with data from all 3 rows.
    doc.save(base.artifactsDir + "Field.NEXT.NEXTIF.docx");
    testFieldNext(doc); //ExSkip
  });


  /// <summary>
  /// Uses a document builder to insert MERGEFIELDs for a data source that contains columns named "Courtesy Title", "First Name" and "Last Name".
  /// </summary>
  function InsertMergeFields(builder, firstFieldTextBefore) {
    insertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ");
    insertMergeField(builder, "First Name", null, " ");
    insertMergeField(builder, "Last Name", null, null);
    builder.insertParagraph();
  }

  /// <summary>
  /// Uses a document builder to insert a MERRGEFIELD with specified properties.
  /// </summary>
  function insertMergeField(builder, fieldName, textBefore, textAfter) {
    let field = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    field.fieldName = fieldName;
    field.textBefore = textBefore;
    field.textAfter = textAfter;
  }
    //ExEnd

  function testFieldNext(doc) {
    doc = DocumentHelper.saveOpen(doc);

    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getText()).toEqual("First row: Mr. John Doe\r" +
                            "Second row: Mrs. Jane Cardholder\r" +
                            "Third row: Mr. Joe Bloggs\r\f");
  }

  //ExStart
  //ExFor:FieldNoteRef
  //ExFor:FieldNoteRef.BookmarkName
  //ExFor:FieldNoteRef.InsertHyperlink
  //ExFor:FieldNoteRef.InsertReferenceMark
  //ExFor:FieldNoteRef.InsertRelativePosition
  //ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
  test('FieldNoteRef', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a bookmark with a footnote that the NOTEREF field will reference.
    insertBookmarkWithFootnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1");

    // This NOTEREF field will display the number of the footnote inside the referenced bookmark.
    // Setting the InsertHyperlink property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
    expect(insertFieldNoteRef(builder, "MyBookmark2", true, false, false, "Hyperlink to Bookmark2, with footnote number ").getFieldCode()).toEqual(" NOTEREF  MyBookmark2 \\h");

    // When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
    // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
    expect(insertFieldNoteRef(builder, "MyBookmark1", true, true, false, "Bookmark1, with footnote number ").getFieldCode()).toEqual(" NOTEREF  MyBookmark1 \\h \\p");

    // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
    // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
    expect(insertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").getFieldCode()).toEqual(" NOTEREF  MyBookmark2 \\h \\p \\f");

    builder.insertBreak(aw.BreakType.PageBreak);
    insertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2");

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.NOTEREF.docx");
    testNoteRef(new aw.Document(base.artifactsDir + "Field.NOTEREF.docx")); //ExSkip
  });


  /// <summary>
  /// Uses a document builder to insert a NOTEREF field with specified properties.
  /// </summary>
  function insertFieldNoteRef(builder, bookmarkName, insertHyperlink, insertRelativePosition, insertReferenceMark, textBefore) {
    builder.write(textBefore);

    let field = builder.insertField(aw.Fields.FieldType.FieldNoteRef, true).asFieldNoteRef();
    field.bookmarkName = bookmarkName;
    field.insertHyperlink = insertHyperlink;
    field.insertRelativePosition = insertRelativePosition;
    field.insertReferenceMark = insertReferenceMark;
    builder.writeln();

    return field;
  }

  /// <summary>
  /// Uses a document builder to insert a named bookmark with a footnote at the end.
  /// </summary>
  function insertBookmarkWithFootnote(builder, bookmarkName, bookmarkText, footnoteText) {
    builder.startBookmark(bookmarkName);
    builder.write(bookmarkText);
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, footnoteText);
    builder.endBookmark(bookmarkName);
    builder.writeln();
  }
    //ExEnd

  function testNoteRef(doc) {
    let field = doc.range.fields.at(0).asFieldNoteRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldNoteRef, " NOTEREF  MyBookmark2 \\h", "2", field);
    expect(field.bookmarkName).toEqual("MyBookmark2");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(false);
    expect(field.insertReferenceMark).toEqual(false);

    field = doc.range.fields.at(1).asFieldNoteRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldNoteRef, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field);
    expect(field.bookmarkName).toEqual("MyBookmark1");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(true);
    expect(field.insertReferenceMark).toEqual(false);

    field = doc.range.fields.at(2).asFieldNoteRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldNoteRef, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field);
    expect(field.bookmarkName).toEqual("MyBookmark2");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(true);
    expect(field.insertReferenceMark).toEqual(true);
  }

  test('NoteRef', () => {
    //ExStart
    //ExFor:FieldNoteRef
    //ExSummary:Shows how to cross-reference footnotes with the NOTEREF field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("CrossReference: ");

    let field = builder.insertField(aw.Fields.FieldType.FieldNoteRef, false).asFieldNoteRef(); // <--- don't update field
    field.bookmarkName = "CrossRefBookmark";
    field.insertHyperlink = true;
    field.insertReferenceMark = true;
    field.insertRelativePosition = false;
    builder.writeln();

    builder.startBookmark("CrossRefBookmark");
    builder.write("Hello world!");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Cross referenced footnote.");
    builder.endBookmark("CrossRefBookmark");
    builder.writeln();

    doc.updateFields();

    // This field works only in older versions of Microsoft Word.
    doc.save(base.artifactsDir + "Field.NOTEREF.doc");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.NOTEREF.doc");
    field = doc.range.fields.at(0).asFieldNoteRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldNoteRef, " NOTEREF  CrossRefBookmark \\h \\f", "1", field);
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, null, "Cross referenced footnote.", doc.getFootnote(0, true));
  });


  //ExStart
  //ExFor:FieldPageRef
  //ExFor:FieldPageRef.BookmarkName
  //ExFor:FieldPageRef.InsertHyperlink
  //ExFor:FieldPageRef.InsertRelativePosition
  //ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
  test('FieldPageRef', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    insertAndNameBookmark(builder, "MyBookmark1");

    // Insert a PAGEREF field that displays what page a bookmark is on.
    // Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
    expect(insertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").getFieldCode()).toEqual(" PAGEREF  MyBookmark3 \\h");

    // We can use the \p flag to get the PAGEREF field to display
    // the bookmark's position relative to the position of the field.
    // Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
    expect(insertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").getFieldCode()).toEqual(" PAGEREF  MyBookmark1 \\h \\p");

    // Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
    expect(insertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").getFieldCode()).toEqual(" PAGEREF  MyBookmark2 \\h \\p");

    // Bookmark3 will be on a different page, so the field will display "on page 2".
    expect(insertFieldPageRef(builder, "MyBookmark3", true, true, "Bookmark3 is ").getFieldCode()).toEqual(" PAGEREF  MyBookmark3 \\h \\p");

    insertAndNameBookmark(builder, "MyBookmark2");
    builder.insertBreak(aw.BreakType.PageBreak);
    insertAndNameBookmark(builder, "MyBookmark3");

    doc.updatePageLayout();
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.PAGEREF.docx");
    testPageRef(new aw.Document(base.artifactsDir + "Field.PAGEREF.docx")); //ExSkip
  });


  /// <summary>
  /// Uses a document builder to insert a PAGEREF field and sets its properties.
  /// </summary>
  function insertFieldPageRef(builder, bookmarkName, insertHyperlink, insertRelativePosition, textBefore)
  {
    builder.write(textBefore);

    let field = builder.insertField(aw.Fields.FieldType.FieldPageRef, true).asFieldPageRef();
    field.bookmarkName = bookmarkName;
    field.insertHyperlink = insertHyperlink;
    field.insertRelativePosition = insertRelativePosition;
    builder.writeln();

    return field;
  }

  /// <summary>
  /// Uses a document builder to insert a named bookmark.
  /// </summary>
  function insertAndNameBookmark(builder, bookmarkName) {
    builder.startBookmark(bookmarkName);
    builder.writeln(`Contents of bookmark \"${bookmarkName}\".`);
    builder.endBookmark(bookmarkName);
  }
    //ExEnd

  function testPageRef(doc) {
    let field = doc.range.fields.at(0).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF  MyBookmark3 \\h", "2", field);
    expect(field.bookmarkName).toEqual("MyBookmark3");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(false);

    field = doc.range.fields.at(1).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF  MyBookmark1 \\h \\p", "above", field);
    expect(field.bookmarkName).toEqual("MyBookmark1");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(true);

    field = doc.range.fields.at(2).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF  MyBookmark2 \\h \\p", "below", field);
    expect(field.bookmarkName).toEqual("MyBookmark2");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(true);

    field = doc.range.fields.at(3).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF  MyBookmark3 \\h \\p", "on page 2", field);
    expect(field.bookmarkName).toEqual("MyBookmark3");
    expect(field.insertHyperlink).toEqual(true);
    expect(field.insertRelativePosition).toEqual(true);
  }

  //ExStart
  //ExFor:FieldRef
  //ExFor:FieldRef.BookmarkName
  //ExFor:FieldRef.IncludeNoteOrComment
  //ExFor:FieldRef.InsertHyperlink
  //ExFor:FieldRef.InsertParagraphNumber
  //ExFor:FieldRef.InsertParagraphNumberInFullContext
  //ExFor:FieldRef.InsertParagraphNumberInRelativeContext
  //ExFor:FieldRef.InsertRelativePosition
  //ExFor:FieldRef.NumberSeparator
  //ExFor:FieldRef.SuppressNonDelimiters
  //ExSummary:Shows how to insert REF fields to reference bookmarks.
  test('FieldRef', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("MyBookmark");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "MyBookmark footnote #1");
    builder.write("Text that will appear in REF field");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "MyBookmark footnote #2");
    builder.endBookmark("MyBookmark");
    builder.moveToDocumentStart();

    // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
    builder.listFormat.applyNumberDefault();
    builder.listFormat.listLevel.numberFormat = "> \u0000";

    // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
    let field = insertFieldRef(builder, "MyBookmark", "", "\n");
    field.includeNoteOrComment = true;
    field.insertHyperlink = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\f \\h");

    // Insert a REF field, and display whether the referenced bookmark is above or below it.
    field = insertFieldRef(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n");
    field.insertRelativePosition = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\p");

    // Display the list number of the bookmark as it appears in the document.
    field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n");
    field.insertParagraphNumber = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\n");

    // Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
    field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n");
    field.insertParagraphNumber = true;
    field.suppressNonDelimiters = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\n \\t");

    // Move down one list level.
    builder.listFormat.listLevelNumber++;
    builder.listFormat.listLevel.numberFormat = ">> \u0001";

    // Display the list number of the bookmark and the numbers of all the list levels above it.
    field = insertFieldRef(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n");
    field.insertParagraphNumberInFullContext = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\w");

    builder.insertBreak(aw.BreakType.PageBreak);

    // Display the list level numbers between this REF field, and the bookmark that it is referencing.
    field = insertFieldRef(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n");
    field.insertParagraphNumberInRelativeContext = true;

    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\r");

    // At the end of the document, the bookmark will show up as a list item here.
    builder.writeln("List level above bookmark");
    builder.listFormat.listLevelNumber++;
    builder.listFormat.listLevel.numberFormat = ">>> \u0002";

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.REF.docx");
    testFieldRef(new aw.Document(base.artifactsDir + "Field.REF.docx")); //ExSkip
  });


  /// <summary>
  /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
  /// </summary>
  function insertFieldRef(builder, bookmarkName, textBefore, textAfter) {
    builder.write(textBefore);
    let field = builder.insertField(aw.Fields.FieldType.FieldRef, true).asFieldRef();
    field.bookmarkName = bookmarkName;
    builder.write(textAfter);
    return field;
  }
  //ExEnd

  function testFieldRef(doc)
  {
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '', "MyBookmark footnote #1", doc.getFootnote(0, true));
    TestUtil.verifyFootnote(aw.Notes.FootnoteType.Footnote, true, '', "MyBookmark footnote #2", doc.getFootnote(1, true));

    let field = doc.range.fields.at(0).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\f \\h", 
      "Text that will appear in REF field", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.includeNoteOrComment).toEqual(true);
    expect(field.insertHyperlink).toEqual(true);

    field = doc.range.fields.at(1).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\p", "below", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.insertRelativePosition).toEqual(true);

    field = doc.range.fields.at(2).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\n", "\u200E>>> i", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.insertParagraphNumber).toEqual(true);
    expect(field.getFieldCode()).toEqual(" REF  MyBookmark \\n");
    expect(field.result).toEqual("\u200E>>> i");

    field = doc.range.fields.at(3).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\n \\t", "\u200Ei", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.insertParagraphNumber).toEqual(true);
    expect(field.suppressNonDelimiters).toEqual(true);

    field = doc.range.fields.at(4).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\w", "\u200E> 4>> c>>> i", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.insertParagraphNumberInFullContext).toEqual(true);

    field = doc.range.fields.at(5).asFieldRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark \\r", "\u200E>> c>>> i", field);
    expect(field.bookmarkName).toEqual("MyBookmark");
    expect(field.insertParagraphNumberInRelativeContext).toEqual(true);
  }

  test('FieldRD', () => {
    //ExStart
    //ExFor:FieldRD
    //ExFor:FieldRD.fileName
    //ExFor:FieldRD.isPathRelative
    //ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a document builder to insert a table of contents,
    // and then add one entry for the table of contents on the following page.
    builder.insertField(aw.Fields.FieldType.FieldTOC, true);
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.currentParagraph.paragraphFormat.styleName = "Heading 1";
    builder.writeln("TOC entry from within this document");

    // Insert an RD field, which references another local file system document in its FileName property.
    // The TOC will also now accept all headings from the referenced document as entries for its table.
    let field = builder.insertField(aw.Fields.FieldType.FieldRefDoc, true).asFieldRD();
    field.fileName = base.artifactsDir + "ReferencedDocument.docx";

    expect(field.getFieldCode()).toEqual(` RD  ${base.artifactsDir.replaceAll("\\","\\\\")}ReferencedDocument.docx`);

    // Create the document that the RD field is referencing and insert a heading. 
    // This heading will show up as an entry in the TOC field in our first document.
    let referencedDoc = new aw.Document();
    let refDocBuilder = new aw.DocumentBuilder(referencedDoc);
    refDocBuilder.currentParagraph.paragraphFormat.styleName = "Heading 1";
    refDocBuilder.writeln("TOC entry from referenced document");
    referencedDoc.save(base.artifactsDir + "ReferencedDocument.docx");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.RD.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.RD.docx");

    let fieldToc = doc.range.fields.at(0).asFieldToc();

    expect(fieldToc.result).toEqual("TOC entry from within this document\t\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r" +
                            "TOC entry from referenced document\t1\r");

    let fieldPageRef = doc.range.fields.at(1).asFieldPageRef();

    TestUtil.verifyField(aw.Fields.FieldType.FieldPageRef, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef);

    field = doc.range.fields.at(2).asFieldRD();

    TestUtil.verifyField(aw.Fields.FieldType.FieldRefDoc, ` RD  ${base.artifactsDir.replaceAll("\\","\\\\")}ReferencedDocument.docx`, '', field);
    expect(field.fileName).toEqual(base.artifactsDir.replaceAll("\\","\\\\") + "ReferencedDocument.docx");
    expect(field.isPathRelative).toEqual(false);
  });


  test.skip('SkipIf: DataTable', () => {
    //ExStart
    //ExFor:FieldSkipIf
    //ExFor:FieldSkipIf.comparisonOperator
    //ExFor:FieldSkipIf.leftExpression
    //ExFor:FieldSkipIf.rightExpression
    //ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a SKIPIF field. If the current row of a mail merge operation fulfills the condition
    // which the expressions of this field state, then the mail merge operation aborts the current row,
    // discards the current merge document, and then immediately moves to the next row to begin the next merge document.
    let fieldSkipIf = builder.insertField(aw.Fields.FieldType.FieldSkipIf, true).asFieldSkipIf();

    // Move the builder to the SKIPIF field's separator so we can place a MERGEFIELD inside the SKIPIF field.
    builder.moveTo(fieldSkipIf.separator);
    let fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Department";

    // The MERGEFIELD refers to the "Department" column in our data table. If a row from that table
    // has a value of "HR" in its "Department" column, then this row will fulfill the condition.
    fieldSkipIf.leftExpression = "=";
    fieldSkipIf.rightExpression = "HR";

    // Add content to our document, create the data source, and execute the mail merge.
    builder.moveToDocumentEnd();
    builder.write("Dear ");
    fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Name";
    builder.writeln(", ");

    // This table has three rows, and one of them fulfills the condition of our SKIPIF field. 
    // The mail merge will produce two pages.
    let table = new DataTable("Employees");
    table.columns.add("Name");
    table.columns.add("Department");
    table.rows.add("John Doe", "Sales");
    table.rows.add("Jane Doe", "Accounting");
    table.rows.add("John Cardholder", "HR");

    doc.mailMerge.execute(table);
    doc.save(base.artifactsDir + "Field.SKIPIF.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SKIPIF.docx");

    expect(doc.range.fields.count).toEqual(0);
    expect(doc.getText()).toEqual("Dear John Doe, \r" +
                            "\fDear Jane Doe, \r\f");
  });


  test('FieldSetRef', () => {
    //ExStart
    //ExFor:FieldRef
    //ExFor:FieldRef.bookmarkName
    //ExFor:FieldSet
    //ExFor:FieldSet.bookmarkName
    //ExFor:FieldSet.bookmarkText
    //ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Name bookmarked text with a SET field. 
    // This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
    let fieldSet = builder.insertField(aw.Fields.FieldType.FieldSet, false).asFieldSet();
    fieldSet.bookmarkName = "MyBookmark";
    fieldSet.bookmarkText = "Hello world!";
    fieldSet.update();

    expect(fieldSet.getFieldCode()).toEqual(" SET  MyBookmark \"Hello world!\"");

    // Refer to the bookmark by name in a REF field and display its contents.
    let fieldRef = builder.insertField(aw.Fields.FieldType.FieldRef, true).asFieldRef();
    fieldRef.bookmarkName = "MyBookmark";
    fieldRef.update();

    expect(fieldRef.getFieldCode()).toEqual(" REF  MyBookmark");
    expect(fieldRef.result).toEqual("Hello world!");

    doc.save(base.artifactsDir + "Field.SET.REF.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SET.REF.docx");

    expect(doc.range.bookmarks.at(0).text).toEqual("Hello world!");

    fieldSet = doc.range.fields.at(0).asFieldSet();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSet, " SET  MyBookmark \"Hello world!\"", "Hello world!", fieldSet);
    expect(fieldSet.bookmarkName).toEqual("MyBookmark");
    expect(fieldSet.bookmarkText).toEqual("Hello world!");

    TestUtil.verifyField(aw.Fields.FieldType.FieldRef, " REF  MyBookmark", "Hello world!", fieldRef);
    expect(fieldRef.result).toEqual("Hello world!");
  });


  test('FieldTemplate', () => {
    //ExStart
    //ExFor:FieldTemplate
    //ExFor:FieldTemplate.includeFullPath
    //ExFor:FieldOptions.templateName
    //ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // We can set a template name using by the fields. This property is used when the "doc.attachedTemplate" is empty.
    // If this property is empty the default template file name "Normal.dotm" is used.
    doc.fieldOptions.templateName = '';

    let field = builder.insertField(aw.Fields.FieldType.FieldTemplate, false).asFieldTemplate();
    expect(field.getFieldCode()).toEqual(" TEMPLATE ");

    builder.writeln();
    field = builder.insertField(aw.Fields.FieldType.FieldTemplate, false).asFieldTemplate();
    field.includeFullPath = true;

    expect(field.getFieldCode()).toEqual(" TEMPLATE  \\p");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TEMPLATE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.TEMPLATE.docx");

    field = doc.range.fields.at(0).asFieldTemplate();
    expect(field.getFieldCode()).toEqual(" TEMPLATE ");
    expect(field.result).toEqual("Normal.dotm");

    field = doc.range.fields.at(1).asFieldTemplate();
    expect(field.getFieldCode()).toEqual(" TEMPLATE  \\p");
    expect(field.result).toEqual("Normal.dotm");
  });


  test('FieldSymbol', () => {
    //ExStart
    //ExFor:FieldSymbol
    //ExFor:FieldSymbol.characterCode
    //ExFor:FieldSymbol.dontAffectsLineSpacing
    //ExFor:FieldSymbol.fontName
    //ExFor:FieldSymbol.fontSize
    //ExFor:FieldSymbol.isAnsi
    //ExFor:FieldSymbol.isShiftJis
    //ExFor:FieldSymbol.isUnicode
    //ExSummary:Shows how to use the SYMBOL field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Below are three ways to use a SYMBOL field to display a single character.
    // 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
    let field = builder.insertField(aw.Fields.FieldType.FieldSymbol, true).asFieldSymbol();

    // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
    field.characterCode = 0x00a9.toString();
    field.isAnsi = true;

    expect(field.getFieldCode()).toBe(' SYMBOL  169 \\\a');

    builder.writeln(" Line 1");

    // 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
    field = builder.insertField(aw.Fields.FieldType.FieldSymbol, true).asFieldSymbol();

    // In Unicode, the infinity symbol occupies the "221E" code.
    field.characterCode = 0x221E.toString();
    field.isUnicode = true;

    // Change the font of our symbol after using the Windows Character Map
    // to ensure that the font can represent that symbol.
    field.fontName = "Calibri";
    field.fontSize = "24";

    // We can set this flag for tall symbols to make them not push down the rest of the text on their line.
    field.dontAffectsLineSpacing = true;

    expect(field.getFieldCode()).toEqual(" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h");

    builder.writeln("Line 2");

    // 3 -  Add a SYMBOL field which displays the あ character,
    // with a font that supports Shift-JIS (Windows-932) codepage:
    field = builder.insertField(aw.Fields.FieldType.FieldSymbol, true).asFieldSymbol();
    field.fontName = "MS Gothic";
    field.characterCode = 0x82A0.toString();
    field.isShiftJis = true;

    expect(field.getFieldCode()).toEqual(" SYMBOL  33440 \\f \"MS Gothic\" \\j");

    builder.write("Line 3");

    doc.save(base.artifactsDir + "Field.SYMBOL.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SYMBOL.docx");

    field = doc.range.fields.at(0).asFieldSymbol();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSymbol, " SYMBOL  169 \\\a", '', field);
    expect(field.characterCode).toEqual(0x00a9.toString());
    expect(field.isAnsi).toEqual(true);
    expect(field.displayResult).toEqual("©");

    field = doc.range.fields.at(1).asFieldSymbol();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSymbol, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", '', field);
    expect(field.characterCode).toEqual(0x221E.toString());
    expect(field.fontName).toEqual("Calibri");
    expect(field.fontSize).toEqual("24");
    expect(field.isUnicode).toEqual(true);
    expect(field.dontAffectsLineSpacing).toEqual(true);
    expect(field.displayResult).toEqual("∞");

    field = doc.range.fields.at(2).asFieldSymbol();

    TestUtil.verifyField(aw.Fields.FieldType.FieldSymbol, " SYMBOL  33440 \\f \"MS Gothic\" \\j", '', field);
    expect(field.characterCode).toEqual(0x82A0.toString());
    expect(field.fontName).toEqual("MS Gothic");
    expect(field.isShiftJis).toEqual(true);
  });


  test('FieldTitle', () => {
    //ExStart
    //ExFor:FieldTitle
    //ExFor:FieldTitle.text
    //ExSummary:Shows how to use the TITLE field.
    let doc = new aw.Document();

    // Set a value for the "Title" built-in document property. 
    doc.builtInDocumentProperties.title = "My Title";

    // We can use the TITLE field to display the value of this property in the document.
    let builder = new aw.DocumentBuilder(doc);
    let field = builder.insertField(aw.Fields.FieldType.FieldTitle, false).asFieldTitle();
    field.update();

    expect(field.getFieldCode()).toEqual(" TITLE ");
    expect(field.result).toEqual("My Title");

    // Setting a value for the field's Text property,
    // and then updating the field will also overwrite the corresponding built-in property with the new value.
    builder.writeln();
    field = builder.insertField(aw.Fields.FieldType.FieldTitle, false).asFieldTitle();
    field.text = "My New Title";
    field.update();

    expect(field.getFieldCode()).toEqual(" TITLE  \"My New Title\"");
    expect(field.result).toEqual("My New Title");
    expect(doc.builtInDocumentProperties.title).toEqual("My New Title");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TITLE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.TITLE.docx");

    expect(doc.builtInDocumentProperties.title).toEqual("My New Title");

    field = doc.range.fields.at(0).asFieldTitle();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTitle, " TITLE ", "My New Title", field);

    field = doc.range.fields.at(1).asFieldTitle();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTitle, " TITLE  \"My New Title\"", "My New Title", field);
    expect(field.text).toEqual("My New Title");
  });


  //ExStart
  //ExFor:FieldToa
  //ExFor:FieldToa.BookmarkName
  //ExFor:FieldToa.EntryCategory
  //ExFor:FieldToa.EntrySeparator
  //ExFor:FieldToa.PageNumberListSeparator
  //ExFor:FieldToa.PageRangeSeparator
  //ExFor:FieldToa.RemoveEntryFormatting
  //ExFor:FieldToa.SequenceName
  //ExFor:FieldToa.SequenceSeparator
  //ExFor:FieldToa.UseHeading
  //ExFor:FieldToa.UsePassim
  //ExFor:FieldTA
  //ExFor:FieldTA.EntryCategory
  //ExFor:FieldTA.IsBold
  //ExFor:FieldTA.IsItalic
  //ExFor:FieldTA.LongCitation
  //ExFor:FieldTA.PageRangeBookmarkName
  //ExFor:FieldTA.ShortCitation
  //ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
  test('FieldTOA', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a TOA field, which will create an entry for each TA field in the document,
    // displaying long citations and page numbers for each entry.
    let fieldToa = builder.insertField(aw.Fields.FieldType.FieldTOA, false).asFieldToa();

    // Set the entry category for our table. This TOA will now only include TA fields
    // that have a matching value in their EntryCategory property.
    fieldToa.entryCategory = "1";

    // Moreover, the Table of Authorities category at index 1 is "Cases",
    // which will show up as our table's title if we set this variable to true.
    fieldToa.useHeading = true;

    // We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
    fieldToa.bookmarkName = "MyBookmark";

    // By default, a dotted line page-wide tab appears between the TA field's citation
    // and its page number. We can replace it with any text we put on this property.
    // Inserting a tab character will preserve the original tab.
    fieldToa.entrySeparator = " \t p.";

    // If we have multiple TA entries that share the same long citation,
    // all their respective page numbers will show up on one row.
    // We can use this property to specify a string that will separate their page numbers.
    fieldToa.pageNumberListSeparator = " & p. ";

    // We can set this to true to get our table to display the word "passim"
    // if there are five or more page numbers in one row.
    fieldToa.usePassim = true;

    // One TA field can refer to a range of pages.
    // We can specify a string here to appear between the start and end page numbers for such ranges.
    fieldToa.pageRangeSeparator = " to ";

    // The format from the TA fields will carry over into our table.
    // We can disable this by setting the RemoveEntryFormatting flag.
    fieldToa.removeEntryFormatting = true;
    builder.font.color = "#008000";
    builder.font.name = "Arial Black";

    expect(fieldToa.getFieldCode()).toEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f");

    builder.insertBreak(aw.BreakType.PageBreak);

    // This TA field will not appear as an entry in the TOA since it is outside
    // the bookmark's bounds that the TOA's BookmarkName property specifies.
    let fieldTA = insertToaEntry(builder, "1", "Source 1");

    expect(fieldTA.getFieldCode()).toEqual(" TA  \\c 1 \\l \"Source 1\"");

    // This TA field is inside the bookmark,
    // but the entry category does not match that of the table, so the TA field will not include it.
    builder.startBookmark("MyBookmark");
    fieldTA = insertToaEntry(builder, "2", "Source 2");

    // This entry will appear in the table.
    fieldTA = insertToaEntry(builder, "1", "Source 3");

    // A TOA table does not display short citations,
    // but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
    fieldTA.shortCitation = "S.3";

    expect(fieldTA.getFieldCode()).toEqual(" TA  \\c 1 \\l \"Source 3\" \\s S.3");

    // We can format the page number to make it bold/italic using the following properties.
    // We will still see these effects if we set our table to ignore formatting.
    fieldTA = insertToaEntry(builder, "1", "Source 2");
    fieldTA.isBold = true;
    fieldTA.isItalic = true;

    expect(fieldTA.getFieldCode()).toEqual(" TA  \\c 1 \\l \"Source 2\" \\b \\i");

    // We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
    // Note that this entry refers to the same source as the one above to share one row in our table.
    // This row will have the page number of the entry above and the page range of this entry,
    // with the table's page list and page number range separators between page numbers.
    fieldTA = insertToaEntry(builder, "1", "Source 3");
    fieldTA.pageRangeBookmarkName = "MyMultiPageBookmark";

    builder.startBookmark("MyMultiPageBookmark");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.endBookmark("MyMultiPageBookmark");

    expect(fieldTA.getFieldCode()).toEqual(" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark");

    // If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
    for (let i = 0; i < 5; i++)
    {
      insertToaEntry(builder, "1", "Source 4");
    }

    builder.endBookmark("MyBookmark");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.TOA.TA.docx");
    testFieldTOA(new aw.Document(base.artifactsDir + "Field.TOA.TA.docx")); //ExSkip
  });


  function insertToaEntry(builder,entryCategory, longCitation) {
    let field = builder.insertField(aw.Fields.FieldType.FieldTOAEntry, false).asFieldTA();
    field.entryCategory = entryCategory;
    field.longCitation = longCitation;

    builder.insertBreak(aw.BreakType.PageBreak);

    return field;
  }
  //ExEnd

  function testFieldTOA(doc) {
    let fieldTOA = doc.range.fields.at(0).asFieldToa();

    expect(fieldTOA.entryCategory).toEqual("1");
    expect(fieldTOA.useHeading).toEqual(true);
    expect(fieldTOA.bookmarkName).toEqual("MyBookmark");
    expect(fieldTOA.entrySeparator).toEqual(" \t p.");
    expect(fieldTOA.pageNumberListSeparator).toEqual(" & p. ");
    expect(fieldTOA.usePassim).toEqual(true);
    expect(fieldTOA.pageRangeSeparator).toEqual(" to ");
    expect(fieldTOA.removeEntryFormatting).toEqual(true);
    expect(fieldTOA.getFieldCode()).toEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f");
    expect(fieldTOA.result).toEqual("Cases\r" +
                            "Source 2 \t p.5\r" +
                            "Source 3 \t p.4 & p. 7 to 10\r" +
                            "Source 4 \t p.passim\r");

    let fieldTA = doc.range.fields.at(1).asFieldTA();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 1\"", '', fieldTA);
    expect(fieldTA.entryCategory).toEqual("1");
    expect(fieldTA.longCitation).toEqual("Source 1");

    fieldTA = doc.range.fields.at(2).asFieldTA();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 2 \\l \"Source 2\"", '', fieldTA);
    expect(fieldTA.entryCategory).toEqual("2");
    expect(fieldTA.longCitation).toEqual("Source 2");

    fieldTA = doc.range.fields.at(3).asFieldTA();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 3\" \\s S.3", '', fieldTA);
    expect(fieldTA.entryCategory).toEqual("1");
    expect(fieldTA.longCitation).toEqual("Source 3");
    expect(fieldTA.shortCitation).toEqual("S.3");

    fieldTA = doc.range.fields.at(4).asFieldTA();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 2\" \\b \\i", '', fieldTA);
    expect(fieldTA.entryCategory).toEqual("1");
    expect(fieldTA.longCitation).toEqual("Source 2");
    expect(fieldTA.isBold).toEqual(true);
    expect(fieldTA.isItalic).toEqual(true);

    fieldTA = doc.range.fields.at(5).asFieldTA();

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", '', fieldTA);
    expect(fieldTA.entryCategory).toEqual("1");
    expect(fieldTA.longCitation).toEqual("Source 3");
    expect(fieldTA.pageRangeBookmarkName).toEqual("MyMultiPageBookmark");

    for (let i = 6; i < 11; i++)
    {
      fieldTA = doc.range.fields.at(i).asFieldTA();

      TestUtil.verifyField(aw.Fields.FieldType.FieldTOAEntry, " TA  \\c 1 \\l \"Source 4\"", '', fieldTA);
      expect(fieldTA.entryCategory).toEqual("1");
      expect(fieldTA.longCitation).toEqual("Source 4");
    }
  }

  test('FieldAddIn', () => {
    //ExStart
    //ExFor:FieldAddIn
    //ExSummary:Shows how to process an ADDIN field.
    let doc = new aw.Document(base.myDir + "Field sample - ADDIN.docx");

    // Aspose.words does not support inserting ADDIN fields, but we can still load and read them.
    let field = doc.range.fields.at(0).asFieldAddIn();

    expect(field.getFieldCode()).toEqual(" ADDIN \"My value\" ");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    TestUtil.verifyField(aw.Fields.FieldType.FieldAddin, " ADDIN \"My value\" ", '', doc.range.fields.at(0));
  });


  test('FieldEditTime', () => {
    //ExStart
    //ExFor:FieldEditTime
    //ExSummary:Shows how to use the EDITTIME field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // The EDITTIME field will show, in minutes,
    // the time spent with the document open in a Microsoft Word window.
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.write("You've been editing this document for ");
    let field = builder.insertField(aw.Fields.FieldType.FieldEditTime, true).asFieldEditTime();
    builder.writeln(" minutes.");

    // This built in document property tracks the minutes. Microsoft Word uses this property
    // to track the time spent with the document open. We can also edit it ourselves.
    doc.builtInDocumentProperties.totalEditingTime = 10;
    field.update();

    expect(field.getFieldCode()).toEqual(" EDITTIME ");
    expect(field.result).toEqual("10");

    // The field does not update itself in real-time, and will also have to be
    // manually updated in Microsoft Word anytime we need an accurate value.
    doc.updateFields();
    doc.save(base.artifactsDir + "Field.EDITTIME.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.EDITTIME.docx");

    expect(doc.builtInDocumentProperties.totalEditingTime).toEqual(10);

    TestUtil.verifyField(aw.Fields.FieldType.FieldEditTime, " EDITTIME ", "10", doc.range.fields.at(0));
  });


  //ExStart
  //ExFor:FieldEQ
  //ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
  test('FieldEQ', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // An EQ field displays a mathematical equation consisting of one or many elements.
    // Each element takes the following form: [switch][options][arguments].
    // There may be one switch, and several possible options.
    // The arguments are a set of coma-separated values enclosed by round braces.

    // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
    // We will pass values 1 and 4 as arguments, and we will not use any options.
    // This field will display a fraction with 1 as the numerator and 4 as the denominator.
    let field = insertFieldEQ(builder, String.raw`\f(1,4)`);

    expect(field.getFieldCode()).toEqual(String.raw` EQ \f(1,4)`);

    // One EQ field may contain multiple elements placed sequentially.
    // We can also nest elements inside one another by placing the inner elements
    // inside the argument brackets of outer elements.
    // We can find the full list of switches, along with their uses here:
    // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

    // Below are applications of nine different EQ field switches that we can use to create different kinds of objects. 
    // 1 -  Array switch "\u0007", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
    insertFieldEQ(builder, String.raw`\a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)`);

    // 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
    // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
    insertFieldEQ(builder, String.raw`\b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))`);

    // 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
    insertFieldEQ(builder, String.raw`A \d \fo30 \li() B`);

    // 4 -  Formula consisting of multiple fractions:
    insertFieldEQ(builder, String.raw`\f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)`);

    // 5 -  Integral switch "\i", with a summation symbol:
    insertFieldEQ(builder, String.raw`\i \su(n=1,5,n)`);

    // 6 -  List switch "\l":
    insertFieldEQ(builder, String.raw`\l(1,1,2,3,n,8,13)`);

    // 7 -  Radical switch "\r", displaying a cubed root of x:
    insertFieldEQ(builder, String.raw`\r (3,x)`);

    // 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
    insertFieldEQ(builder, String.raw`\s \up8(Superscript) Text \s \do8(Subscript)`);

    // 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
    insertFieldEQ(builder, String.raw`\x \to \bo \le \ri(5)`);

    // Some more complex combinations.
    insertFieldEQ(builder, String.raw`\a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))`);
    insertFieldEQ(builder, String.raw`\i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)`);
    insertFieldEQ(builder, String.raw`\i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)`);

    doc.save(base.artifactsDir + "Field.EQ.docx");
    testFieldEQ(new aw.Document(base.artifactsDir + "Field.EQ.docx")); //ExSkip
  });


  /// <summary>
  /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
  /// </summary>
  function insertFieldEQ(builder, args) {
    let field = builder.insertField(aw.Fields.FieldType.FieldEquation, true).asFieldEQ();
    builder.moveTo(field.separator);
    builder.write(args);
    builder.moveTo(field.start.parentNode);

    builder.insertParagraph();
    return field;
  }
  //ExEnd

  function testFieldEQ(doc) {
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \f(1,4)`, '', doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)`, '', doc.range.fields.at(1));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))`, '', doc.range.fields.at(2));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ A \d \fo30 \li() B`, '', doc.range.fields.at(3));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)`, '', doc.range.fields.at(4));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \i \su(n=1,5,n)`, '', doc.range.fields.at(5));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \l(1,1,2,3,n,8,13)`, '', doc.range.fields.at(6));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \r (3,x)`, '', doc.range.fields.at(7));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \s \up8(Superscript) Text \s \do8(Subscript)`, '', doc.range.fields.at(8));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \x \to \bo \le \ri(5)`, '', doc.range.fields.at(9));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))`, '', doc.range.fields.at(10));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)`, '', doc.range.fields.at(11));
    TestUtil.verifyField(aw.Fields.FieldType.FieldEquation, String.raw` EQ \i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)`, '', doc.range.fields.at(12));
  }

  test('FieldEQAsOfficeMath', () => {
    //ExStart
    //ExFor:FieldEQ
    //ExFor:FieldEQ.asOfficeMath
    //ExSummary:Shows how to replace the EQ field with Office Math.
    let doc = new aw.Document(base.myDir + "Field sample - EQ.docx");
    let fieldEQ = Array.from(doc.range.fields).map(node => node.asFieldEQ()).filter(node => node != null)[0];

    let officeMath = fieldEQ.asOfficeMath();

    fieldEQ.start.parentNode.insertBefore(officeMath, fieldEQ.start);
    fieldEQ.remove();

    doc.save(base.artifactsDir + "Field.EQAsOfficeMath.docx");
    //ExEnd
  });


  test('FieldForms', () => {
    //ExStart
    //ExFor:FieldFormCheckBox
    //ExFor:FieldFormDropDown
    //ExFor:FieldFormText
    //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
    // These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.words.
    // In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
    let doc = new aw.Document(base.myDir + "Form fields.docx");

    let fieldFormCheckBox = doc.range.fields.at(1).asFieldFormCheckBox();
    expect(fieldFormCheckBox.getFieldCode()).toEqual(" FORMCHECKBOX \u0001");

    let fieldFormDropDown = doc.range.fields.at(2).asFieldFormDropDown();
    expect(fieldFormDropDown.getFieldCode()).toEqual(" FORMDROPDOWN \u0001");

    let fieldFormText = doc.range.fields.at(0).asFieldFormText();
    expect(fieldFormText.getFieldCode()).toEqual(" FORMTEXT \u0001");
    //ExEnd
  });


  test('FieldFormula', () => {
    //ExStart
    //ExFor:FieldFormula
    //ExSummary:Shows how to use the formula field to display the result of an equation.
    let doc = new aw.Document();

    // Use a field builder to construct a mathematical equation,
    // then create a formula field to display the equation's result in the document.
    let fieldBuilder = new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldFormula);
    fieldBuilder.addArgument(2);
    fieldBuilder.addArgument("*");
    fieldBuilder.addArgument(5);

    let field = fieldBuilder.buildAndInsert(doc.firstSection.body.firstParagraph).asFieldFormula();
    field.update();

    expect(field.getFieldCode()).toEqual(" = 2 * 5 ");
    expect(field.result).toEqual("10");

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.FORMULA.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.FORMULA.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 2 * 5 ", "10", doc.range.fields.at(0));
  });


  test('FieldLastSavedBy', () => {
    //ExStart
    //ExFor:FieldLastSavedBy
    //ExSummary:Shows how to use the LASTSAVEDBY field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
    // If we make a document programmatically, this property will be null, and we will need to assign a value. 
    doc.builtInDocumentProperties.lastSavedBy = "John Doe";

    // We can use the LASTSAVEDBY field to display the value of this property in the document.
    let field = builder.insertField(aw.Fields.FieldType.FieldLastSavedBy, true).asFieldLastSavedBy();

    expect(field.getFieldCode()).toEqual(" LASTSAVEDBY ");
    expect(field.result).toEqual("John Doe");

    doc.save(base.artifactsDir + "Field.LASTSAVEDBY.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.LASTSAVEDBY.docx");

    expect(doc.builtInDocumentProperties.lastSavedBy).toEqual("John Doe");
    TestUtil.verifyField(aw.Fields.FieldType.FieldLastSavedBy, " LASTSAVEDBY ", "John Doe", doc.range.fields.at(0));
  });


  test.skip('FieldMergeRec: DataTable', () => {
    //ExStart
    //ExFor:FieldMergeRec
    //ExFor:FieldMergeSeq
    //ExFor:FieldSkipIf
    //ExFor:FieldSkipIf.comparisonOperator
    //ExFor:FieldSkipIf.leftExpression
    //ExFor:FieldSkipIf.rightExpression
    //ExSummary:Shows how to use MERGEREC and MERGESEQ fields to the number and count mail merge records in a mail merge's output documents.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Dear ");
    let fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Name";
    builder.writeln(",");

    // A MERGEREC field will print the row number of the data being merged in every merge output document.
    builder.write("\nRow number of record in data source: ");
    let fieldMergeRec = builder.insertField(aw.Fields.FieldType.FieldMergeRec, true).asFieldMergeRec();

    expect(fieldMergeRec.getFieldCode()).toEqual(" MERGEREC ");

    // A MERGESEQ field will count the number of successful merges and print the current value on each respective page.
    // If a mail merge skips no rows and invokes no SKIP/SKIPIF/NEXT/NEXTIF fields, then all merges are successful.
    // The MERGESEQ and MERGEREC fields will display the same results of their mail merge was successful.
    builder.write("\nSuccessful merge number: ");
    let fieldMergeSeq = builder.insertField(aw.Fields.FieldType.FieldMergeSeq, true).asFieldMergeSeq();

    expect(fieldMergeSeq.getFieldCode()).toEqual(" MERGESEQ ");

    // Insert a SKIPIF field, which will skip a merge if the name is "John Doe".
    let fieldSkipIf = builder.insertField(aw.Fields.FieldType.FieldSkipIf, true).asFieldSkipIf();
    builder.moveTo(fieldSkipIf.separator);
    fieldMergeField = builder.insertField(aw.Fields.FieldType.FieldMergeField, true).asFieldMergeField();
    fieldMergeField.fieldName = "Name";
    fieldSkipIf.leftExpression = "=";
    fieldSkipIf.rightExpression = "John Doe";

    // Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column.
    // Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3.
    // On page 1, the MERGESEQ and MERGEREC fields will both display "1".
    // On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2".
    let table = new DataTable("Employees");
    table.columns.add("Name");
    table.rows.add("Jane Doe");
    table.rows.add("John Doe");
    table.rows.add("Joe Bloggs");

    doc.mailMerge.execute(table);
    doc.save(base.artifactsDir + "Field.MERGEREC.MERGESEQ.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.MERGEREC.MERGESEQ.docx");

    expect(doc.range.fields.count).toEqual(0);

    expect(doc.getText().trim()).toEqual("Dear Jane Doe,\r" +
                            "\r" +
                            "Row number of record in data source: 1\r" +
                            "Successful merge number: 1\fDear Joe Bloggs,\r" +
                            "\r" +
                            "Row number of record in data source: 3\r" +
                            "Successful merge number: 2");
  });


  test('FieldOcx', () => {
    //ExStart
    //ExFor:FieldOcx
    //ExSummary:Shows how to insert an OCX field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField(aw.Fields.FieldType.FieldOcx, true).asFieldOcx();

    expect(field.getFieldCode()).toEqual(" OCX ");
    //ExEnd

    TestUtil.verifyField(aw.Fields.FieldType.FieldOcx, " OCX ", '', field);
  });


  /*  //ExStart
    //ExFor:Field.Remove
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
  test('FieldPrivate', () => {
    // Open a Corel WordPerfect document which we have converted to .docx format.
    let doc = new aw.Document(base.myDir + "Field sample - PRIVATE.docx");

    // WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
    // Microsoft Word preserves PRIVATE fields during load/save operations,
    // but provides no functionality for them.
    let field = (FieldPrivate)doc.range.fields.at(0);

    expect(field.getFieldCode()).toEqual(" PRIVATE \"My value\" ");
    expect(field.type).toEqual(aw.Fields.FieldType.FieldPrivate);

    // We can also insert PRIVATE fields using a document builder.
    let builder = new aw.DocumentBuilder(doc);
    builder.insertField(aw.Fields.FieldType.FieldPrivate, true);

    // These fields are not a viable way of protecting sensitive information.
    // Unless backward compatibility with older versions of WordPerfect is essential,
    // we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
    expect(doc.range.fields.count).toEqual(2);

    let remover = new FieldPrivateRemover();
    doc.accept(remover);

    expect(remover.GetFieldsRemovedCount()).toEqual(2);
    expect(doc.range.fields.count).toEqual(0);
  });


    /// <summary>
    /// Removes all encountered PRIVATE fields.
    /// </summary>
  public class FieldPrivateRemover : DocumentVisitor
  {
    public FieldPrivateRemover()
    {
      mFieldsRemovedCount = 0;
    }

    public int GetFieldsRemovedCount()
    {
      return mFieldsRemovedCount;
    }

      /// <summary>
      /// Called when a FieldEnd node is encountered in the document.
      /// If the node belongs to a PRIVATE field, the entire field is removed.
      /// </summary>
    public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
    {
      if (fieldEnd.fieldType == aw.Fields.FieldType.FieldPrivate)
      {
        fieldEnd.getField().Remove();
        mFieldsRemovedCount++;
      }

      return aw.VisitorAction.Continue;
    }

    private int mFieldsRemovedCount;
  }
    //ExEnd*/

  test('FieldSection', () => {
    //ExStart
    //ExFor:FieldSection
    //ExFor:FieldSectionPages
    //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);
    builder.paragraphFormat.alignment = aw.ParagraphAlignment.Right;

    // A SECTION field displays the number of the section it is in.
    builder.write("Section ");
    let fieldSection = builder.insertField(aw.Fields.FieldType.FieldSection, true).asFieldSection();

    expect(fieldSection.getFieldCode()).toEqual(" SECTION ");

    // A PAGE field displays the number of the page it is in.
    builder.write("\nPage ");
    let fieldPage = builder.insertField(aw.Fields.FieldType.FieldPage, true).asFieldPage();

    expect(fieldPage.getFieldCode()).toEqual(" PAGE ");

    // A SECTIONPAGES field displays the number of pages that the section it is in spans across.
    builder.write(" of ");
    let fieldSectionPages = builder.insertField(aw.Fields.FieldType.FieldSectionPages, true).asFieldSectionPages();

    expect(fieldSectionPages.getFieldCode()).toEqual(" SECTIONPAGES ");

    // Move out of the header back into the main document and insert two pages.
    // All these pages will be in the first section. Our fields, which appear once every header,
    // will number the current/total pages of this section.
    builder.moveToDocumentEnd();
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertBreak(aw.BreakType.PageBreak);

    // We can insert a new section with the document builder like this.
    // This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
    builder.insertBreak(aw.BreakType.SectionBreakNewPage);

    // The PAGE field will keep counting pages across the whole document.
    // We can manually reset its count at each section to keep track of pages section-by-section.
    builder.currentSection.pageSetup.restartPageNumbering = true;
    builder.insertBreak(aw.BreakType.PageBreak);

    doc.updateFields();
    doc.save(base.artifactsDir + "Field.SECTION.SECTIONPAGES.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.SECTION.SECTIONPAGES.docx");

    TestUtil.verifyField(aw.Fields.FieldType.FieldSection, " SECTION ", "2", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldPage, " PAGE ", "2", doc.range.fields.at(1));
    TestUtil.verifyField(aw.Fields.FieldType.FieldSectionPages, " SECTIONPAGES ", "2", doc.range.fields.at(2));
  });


  //ExStart
  //ExFor:FieldTime
  //ExSummary:Shows how to display the current time using the TIME field.
  test('FieldTime', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // By default, time is displayed in the "h:mm am/pm" format.
    let field = insertFieldTime(builder, "");

    expect(field.getFieldCode()).toEqual(" TIME ");

    // We can use the \@ flag to change the format of our displayed time.
    field = insertFieldTime(builder, "\\@ HHmm");

    expect(field.getFieldCode()).toEqual(" TIME \\@ HHmm");

    // We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
    field = insertFieldTime(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"");

    expect(field.getFieldCode()).toEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"");

    doc.save(base.artifactsDir + "Field.TIME.docx");
    testFieldTime(new aw.Document(base.artifactsDir + "Field.TIME.docx")); //ExSkip
  });


  /// <summary>
  /// Use a document builder to insert a TIME field, insert a new paragraph and return the field.
  /// </summary>
  function insertFieldTime(builder, format)
  {
    let field = builder.insertField(aw.Fields.FieldType.FieldTime, true).asFieldTime();
    builder.moveTo(field.separator);
    builder.write(format);
    builder.moveTo(field.start.parentNode);

    builder.insertParagraph();
    return field;
  }
  //ExEnd

  function testFieldTime(doc) {
    let docLoadingTime = new Date();
    docLoadingTime.setSeconds(0);
    docLoadingTime.setMilliseconds(0);
    doc = DocumentHelper.saveOpen(doc);

    let field = doc.range.fields.at(0).asFieldTime();

    expect(field.getFieldCode()).toEqual(" TIME ");
    expect(field.type).toEqual(aw.Fields.FieldType.FieldTime);
    expect(moment(field.result, "LT").toISOString()).toBe(docLoadingTime.toISOString());

    field = doc.range.fields.at(1).asFieldTime();

    expect(field.getFieldCode()).toEqual(" TIME \\@ HHmm");
    expect(field.type).toEqual(aw.Fields.FieldType.FieldTime);
    expect(moment(field.result, "LT").toISOString()).toBe(docLoadingTime.toISOString());

    field = doc.range.fields.at(2).asFieldTime();

    expect(field.getFieldCode()).toEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"");
    expect(field.type).toEqual(aw.Fields.FieldType.FieldTime);
    expect(moment(field.result, "LT").toISOString()).toBe(docLoadingTime.toISOString());
  }

  test('BidiOutline', () => {
    //ExStart
    //ExFor:FieldBidiOutline
    //ExFor:FieldShape
    //ExFor:FieldShape.text
    //ExFor:ParagraphFormat.bidi
    //ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
    // but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
    // The following field will display ".1", the RTL equivalent of list number "1.".
    let field = builder.insertField(aw.Fields.FieldType.FieldBidiOutline, true).asFieldBidiOutline();
    builder.writeln("שלום");

    expect(field.getFieldCode()).toEqual(" BIDIOUTLINE ");

    // Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
    builder.insertField(aw.Fields.FieldType.FieldBidiOutline, true);
    builder.writeln("שלום");
    builder.insertField(aw.Fields.FieldType.FieldBidiOutline, true);
    builder.writeln("שלום");

    // Set the horizontal text alignment for every paragraph in the document to RTL.
    for (let para of doc.getChildNodes(aw.NodeType.Paragraph, true).toArray().map(node => node.asParagraph()))
    {
      para.paragraphFormat.bidi = true;
    }

    // If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
    // Otherwise, they will display "###".
    doc.save(base.artifactsDir + "Field.BIDIOUTLINE.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "Field.BIDIOUTLINE.docx");

    for (let fieldBidiOutline of doc.range.fields)
      TestUtil.verifyField(aw.Fields.FieldType.FieldBidiOutline, " BIDIOUTLINE ", '', fieldBidiOutline);
  });


  test('Legacy', () => {
    //ExStart
    //ExFor:FieldEmbed
    //ExFor:FieldShape
    //ExFor:FieldShape.text
    //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
    // Open a document that was created in Microsoft Word 2003.
    let doc = new aw.Document(base.myDir + "Legacy fields.doc");

    // If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
    // A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
    // An EMBED field has the same function, but for an embedded object,
    // such as a spreadsheet from an external Excel document.
    // However, these fields will not appear in the document's Fields collection.
    expect(doc.range.fields.count).toEqual(0);

    // These fields are supported only by old versions of Microsoft Word.
    // The document loading process will convert these fields into Shape objects,
    // which we can access in the document's node collection.
    let shapes = doc.getChildNodes(aw.NodeType.Shape, true);
    expect(shapes.count).toEqual(3);

    // The first Shape node corresponds to the SHAPE field in the input document,
    // which is the inline canvas for the AutoShape.
    let shape = shapes.at(0).asShape();
    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.Image);

    // The second Shape node is the AutoShape itself.
    shape = shapes.at(1).asShape();
    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.Can);

    // The third Shape is what was the EMBED field that contained the external spreadsheet.
    shape = shapes.at(2).asShape();
    expect(shape.shapeType).toEqual(aw.Drawing.ShapeType.OleObject);
    //ExEnd
  });


  test('SetFieldIndexFormat', () => {
    //ExStart
    //ExFor:FieldIndexFormat
    //ExFor:FieldOptions.fieldIndexFormat
    //ExSummary:Shows how to formatting FieldIndex fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("A");
    builder.insertBreak(aw.BreakType.LineBreak);
    builder.insertField("XE \"A\"");
    builder.write("B");

    builder.insertField(" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", null);

    doc.fieldOptions.fieldIndexFormat = aw.Fields.FieldIndexFormat.Fancy;
    doc.updateFields();

    doc.save(base.artifactsDir + "Field.SetFieldIndexFormat.docx");
    //ExEnd
  });


  /*  //ExStart
    //ExFor:ComparisonEvaluationResult.#ctor(bool)
    //ExFor:ComparisonEvaluationResult.#ctor(string)
    //ExFor:ComparisonEvaluationResult
    //ExFor:ComparisonEvaluationResult.ErrorMessage
    //ExFor:ComparisonEvaluationResult.Result
    //ExFor:ComparisonExpression
    //ExFor:ComparisonExpression.LeftExpression
    //ExFor:ComparisonExpression.ComparisonOperator
    //ExFor:ComparisonExpression.RightExpression
    //ExFor:FieldOptions.ComparisonExpressionEvaluator
    //ExFor:IComparisonExpressionEvaluator
    //ExFor:IComparisonExpressionEvaluator.Evaluate(Field,ComparisonExpression)
    //ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
  test.each([" IF {0} {1} {2} \"true argument\" \"false argument\" ", 1, null, "true argument",
    " IF {0} {1} {2} \"true argument\" \"false argument\" ", 0, null, "false argument",
    " IF {0} {1} {2} \"true argument\" \"false argument\" ", -1, "Custom Error", "Custom Error",
    " IF {0} {1} {2} \"true argument\" \"false argument\" ", -1, null, "true argument",
    " COMPARE {0} {1} {2} ", 1, null, "1",
    " COMPARE {0} {1} {2} ", 0, null, "0",
    " COMPARE {0} {1} {2} ", -1, "Custom Error", "Custom Error",
    " COMPARE {0} {1} {2} ", -1, null, "1"])('ConditionEvaluationExtensionPoint', (string fieldCode, sbyte comparisonResult, string comparisonError, string expectedResult) => {
    const string left = "\"left expression\"";
    const string @operator = "<>";
    const string right = "\"right expression\"";

    let builder = new aw.DocumentBuilder();

    // Field codes that we use in this example:
    // 1.   " IF {0} {1} {2} \"true argument\" \"false argument\" ".
    // 2.   " COMPARE {0} {1} {2} ".
    let field = builder.insertField(string.format(fieldCode, left, @operator, right), null);

    // If the "comparisonResult" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
    let result = comparisonResult != -1
      ? new aw.Fields.ComparisonEvaluationResult(comparisonResult == 1)
      : comparisonError != null ? new aw.Fields.ComparisonEvaluationResult(comparisonError) : null;

    let evaluator = new ComparisonExpressionEvaluator(result);
    builder.document.fieldOptions.comparisonExpressionEvaluator = evaluator;

    builder.document.updateFields();

    expect(field.result).toEqual(expectedResult);
    evaluator.AssertInvocationsCount(1).AssertInvocationArguments(0, left, @operator, right);
  });


    /// <summary>
    /// Comparison expressions evaluation for the FieldIf and FieldCompare.
    /// </summary>
  private class ComparisonExpressionEvaluator : IComparisonExpressionEvaluator
  {
    public ComparisonExpressionEvaluator(ComparisonEvaluationResult result)
    {
      mResult = result;
      if (mResult != null)
      {
        console.log(mResult.errorMessage);
        console.log(mResult.result);
      }
    }

    public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
    {
      mInvocations.add(new[]
      {
        expression.leftExpression,
        expression.comparisonOperator,
        expression.rightExpression
      });

      return mResult;
    }

    public ComparisonExpressionEvaluator AssertInvocationsCount(int expected)
    {
      expect(mInvocations.count).toEqual(expected);
      return this;
    }

    public ComparisonExpressionEvaluator AssertInvocationArguments(
      int invocationIndex,
      string expectedLeftExpression,
      string expectedComparisonOperator,
      string expectedRightExpression)
    {
      string.at(] arguments = mInvocations[invocationIndex);

      expect(arguments.at(0)).toEqual(expectedLeftExpression);
      expect(arguments.at(1)).toEqual(expectedComparisonOperator);
      expect(arguments.at(2)).toEqual(expectedRightExpression);

      return this;
    }

    private readonly ComparisonEvaluationResult mResult;
    private readonly List<string[]> mInvocations = new aw.Lists.List<string[]>();
  } 
    //ExEnd

  test('ComparisonExpressionEvaluatorNestedFields', () => {
    let document = new aw.Document();

    new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
      .AddArgument(
        new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
          .AddArgument(123)
          .AddArgument(">")
          .AddArgument(666)
          .AddArgument("left greater than right")
          .AddArgument("left less than right"))
      .AddArgument("<>")
      .AddArgument(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
        .AddArgument("left expression")
        .AddArgument("=")
        .AddArgument("right expression")
        .AddArgument("expression are equal")
        .AddArgument("expression are not equal"))
      .AddArgument(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
          .AddArgument(new aw.Fields.FieldArgumentBuilder()
            .AddText("#")
            .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldPage)))
          .AddArgument("=")
          .AddArgument(new aw.Fields.FieldArgumentBuilder()
            .AddText("#")
            .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldNumPages)))
          .AddArgument("the last page")
          .AddArgument("not the last page"))
      .AddArgument(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
          .AddArgument("unexpected")
          .AddArgument("=")
          .AddArgument("unexpected")
          .AddArgument("unexpected")
          .AddArgument("unexpected"))
      .BuildAndInsert(document.firstSection.body.firstParagraph);

    let evaluator = new ComparisonExpressionEvaluator(null);
    document.fieldOptions.comparisonExpressionEvaluator = evaluator;

    document.updateFields();

    evaluator
      .AssertInvocationsCount(4)
      .AssertInvocationArguments(0, "123", ">", "666")
      .AssertInvocationArguments(1, "\"left expression\"", "=", "\"right expression\"")
      .AssertInvocationArguments(2, "left less than right", "<>", "expression are not equal")
      .AssertInvocationArguments(3, "\"#1\"", "=", "\"#1\"");
  });


  test('ComparisonExpressionEvaluatorHeaderFooterFields', () => {
    let document = new aw.Document();
    let builder = new aw.DocumentBuilder(document);

    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.moveToHeaderFooter(aw.HeaderFooterType.HeaderPrimary);

    new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldIf)
      .AddArgument(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldPage))
      .AddArgument("=")
      .AddArgument(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldNumPages))
      .AddArgument(new aw.Fields.FieldArgumentBuilder()
        .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldPage))
        .AddText(" / ")
        .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldNumPages)))
      .AddArgument(new aw.Fields.FieldArgumentBuilder()
        .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldPage))
        .AddText(" / ")
        .AddField(new aw.Fields.FieldBuilder(aw.Fields.FieldType.FieldNumPages)))
      .BuildAndInsert(builder.currentParagraph);

    let evaluator = new ComparisonExpressionEvaluator(null);
    document.fieldOptions.comparisonExpressionEvaluator = evaluator;

    document.updateFields();

    evaluator
      .AssertInvocationsCount(3)
      .AssertInvocationArguments(0, "1", "=", "3")
      .AssertInvocationArguments(1, "2", "=", "3")
      .AssertInvocationArguments(2, "3", "=", "3");
  });


  //ExStart
  //ExFor:FieldOptions.FieldUpdatingCallback
  //ExFor:FieldOptions.FieldUpdatingProgressCallback
  //ExFor:IFieldUpdatingCallback
  //ExFor:IFieldUpdatingProgressCallback
  //ExFor:IFieldUpdatingProgressCallback.Notify(FieldUpdatingProgressArgs)
  //ExFor:FieldUpdatingProgressArgs
  //ExFor:FieldUpdatingProgressArgs.UpdateCompleted
  //ExFor:FieldUpdatingProgressArgs.TotalFieldsCount
  //ExFor:FieldUpdatingProgressArgs.UpdatedFieldsCount
  //ExFor:IFieldUpdatingCallback.FieldUpdating(Field)
  //ExFor:IFieldUpdatingCallback.FieldUpdated(Field)
  //ExSummary:Shows how to use callback methods during a field update.
  test('FieldUpdatingCallbackTest', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
    builder.insertField(" TIME ");
    builder.insertField(" REVNUM ");
    builder.insertField(" AUTHOR  \"John Doe\" ");
    builder.insertField(" SUBJECT \"My Subject\" ");
    builder.insertField(" QUOTE \"Hello world!\" ");

    let callback = new FieldUpdatingCallback();
    doc.fieldOptions.fieldUpdatingCallback = callback;

    doc.updateFields();

    expect(callback.FieldUpdatedCalls.contains("Updating John Doe")).toEqual(true);
  });


    /// <summary>
    /// Implement this interface if you want to have your own custom methods called during a field update.
    /// </summary>
  public class FieldUpdatingCallback : IFieldUpdatingCallback, IFieldUpdatingProgressCallback
  {
    public FieldUpdatingCallback()
    {
      FieldUpdatedCalls = new aw.Lists.List<string>();
    }

      /// <summary>
      /// A user defined method that is called just before a field is updated.
      /// </summary>
    void aw.Fields.IFieldUpdatingCallback.fieldUpdating(Field field)
    {
      if (field.type == aw.Fields.FieldType.FieldAuthor)
      {
        let fieldAuthor = (FieldAuthor) field;
        fieldAuthor.authorName = "Updating John Doe";
      }
    }

      /// <summary>
      /// A user defined method that is called just after a field is updated.
      /// </summary>
    void aw.Fields.IFieldUpdatingCallback.fieldUpdated(Field field)
    {
      FieldUpdatedCalls.add(field.result);
    }

    void aw.Fields.IFieldUpdatingProgressCallback.notify(FieldUpdatingProgressArgs args)
    {
      console.log(`${args.updateCompleted}/${args.totalFieldsCount}`);
      console.log(`${args.updatedFieldsCount}`);
    }

    public IList<string> FieldUpdatedCalls { get; }
  }
    //ExEnd
  */

  test('BibliographySources', () => {
    //ExStart:BibliographySources
    //GistId:eeeec1fbf118e95e7df3f346c91ed726
    //ExFor:Document.bibliography
    //ExFor:Bibliography
    //ExFor:Bibliography.sources
    //ExFor:Source
    //ExFor:Source.#ctor(string, SourceType)
    //ExFor:Source.title
    //ExFor:Source.abbreviatedCaseNumber
    //ExFor:Source.albumTitle
    //ExFor:Source.bookTitle
    //ExFor:Source.broadcaster
    //ExFor:Source.broadcastTitle
    //ExFor:Source.caseNumber
    //ExFor:Source.chapterNumber
    //ExFor:Source.city
    //ExFor:Source.comments
    //ExFor:Source.conferenceName
    //ExFor:Source.countryOrRegion
    //ExFor:Source.court
    //ExFor:Source.day
    //ExFor:Source.dayAccessed
    //ExFor:Source.department
    //ExFor:Source.distributor
    //ExFor:Source.doi
    //ExFor:Source.edition
    //ExFor:Source.guid
    //ExFor:Source.institution
    //ExFor:Source.internetSiteTitle
    //ExFor:Source.issue
    //ExFor:Source.journalName
    //ExFor:Source.lcid
    //ExFor:Source.medium
    //ExFor:Source.month
    //ExFor:Source.monthAccessed
    //ExFor:Source.numberVolumes
    //ExFor:Source.pages
    //ExFor:Source.patentNumber
    //ExFor:Source.periodicalTitle
    //ExFor:Source.productionCompany
    //ExFor:Source.publicationTitle
    //ExFor:Source.publisher
    //ExFor:Source.recordingNumber
    //ExFor:Source.refOrder
    //ExFor:Source.reporter
    //ExFor:Source.shortTitle
    //ExFor:Source.sourceType
    //ExFor:Source.standardNumber
    //ExFor:Source.stateOrProvince
    //ExFor:Source.station
    //ExFor:Source.tag
    //ExFor:Source.theater
    //ExFor:Source.thesisType
    //ExFor:Source.type
    //ExFor:Source.url
    //ExFor:Source.version
    //ExFor:Source.volume
    //ExFor:Source.year
    //ExFor:Source.yearAccessed
    //ExFor:Source.contributors
    //ExFor:SourceType
    //ExFor:Contributor
    //ExFor:ContributorCollection
    //ExFor:ContributorCollection.author
    //ExFor:ContributorCollection.artist
    //ExFor:ContributorCollection.bookAuthor
    //ExFor:ContributorCollection.compiler
    //ExFor:ContributorCollection.composer
    //ExFor:ContributorCollection.conductor
    //ExFor:ContributorCollection.counsel
    //ExFor:ContributorCollection.director
    //ExFor:ContributorCollection.editor
    //ExFor:ContributorCollection.interviewee
    //ExFor:ContributorCollection.interviewer
    //ExFor:ContributorCollection.inventor
    //ExFor:ContributorCollection.performer
    //ExFor:ContributorCollection.producer
    //ExFor:ContributorCollection.translator
    //ExFor:ContributorCollection.writer
    //ExFor:PersonCollection
    //ExFor:PersonCollection.count
    //ExFor:PersonCollection.item(Int32)
    //ExFor:Person.#ctor(string, string, string)
    //ExFor:Person
    //ExFor:Person.first
    //ExFor:Person.middle
    //ExFor:Person.last
    //ExSummary:Shows how to get bibliography sources available in the document.
    let document = new aw.Document(base.myDir + "Bibliography sources.docx");

    let bibliography = document.bibliography;
    expect(bibliography.sources.count).toEqual(12);

    let source = bibliography.sources.at(0);
    expect(source.title).toEqual("Book 0 (No LCID)");
    expect(source.sourceType).toEqual(aw.Bibliography.SourceType.Book);
    expect([...source.contributors].length).toEqual(3);
    expect(source.abbreviatedCaseNumber).toBe(null);
    expect(source.albumTitle).toBe(null);
    expect(source.bookTitle).toBe(null);
    expect(source.broadcaster).toBe(null);
    expect(source.broadcastTitle).toBe(null);
    expect(source.caseNumber).toBe(null);
    expect(source.chapterNumber).toBe(null);
    expect(source.comments).toBe(null);
    expect(source.conferenceName).toBe(null);
    expect(source.countryOrRegion).toBe(null);
    expect(source.court).toBe(null);
    expect(source.day).toBe(null);
    expect(source.dayAccessed).toBe(null);
    expect(source.department).toBe(null);
    expect(source.distributor).toBe(null);
    expect(source.doi).toBe(null);
    expect(source.edition).toBe(null);
    expect(source.guid).toBe(null);
    expect(source.institution).toBe(null);
    expect(source.internetSiteTitle).toBe(null);
    expect(source.issue).toBe(null);
    expect(source.journalName).toBe(null);
    expect(source.lcid).toBe(null);
    expect(source.medium).toBe(null);
    expect(source.month).toBe(null);
    expect(source.monthAccessed).toBe(null);
    expect(source.numberVolumes).toBe(null);
    expect(source.pages).toBe(null);
    expect(source.patentNumber).toBe(null);
    expect(source.periodicalTitle).toBe(null);
    expect(source.productionCompany).toBe(null);
    expect(source.publicationTitle).toBe(null);
    expect(source.publisher).toBe(null);
    expect(source.recordingNumber).toBe(null);
    expect(source.refOrder).toBe(null);
    expect(source.reporter).toBe(null);
    expect(source.shortTitle).toBe(null);
    expect(source.standardNumber).toBe(null);
    expect(source.stateOrProvince).toBe(null);
    expect(source.station).toBe(null);
    expect(source.tag).toEqual("BookNoLCID");
    expect(source.theater).toBe(null);
    expect(source.thesisType).toBe(null);
    expect(source.type).toBe(null);
    expect(source.url).toBe(null);
    expect(source.version).toBe(null);
    expect(source.volume).toBe(null);
    expect(source.year).toBe(null);
    expect(source.yearAccessed).toBe(null);

    // Also, you can create a new source.
    let newSource = new aw.Bibliography.Source("New source", aw.Bibliography.SourceType.Misc);

    let contributors = source.contributors;
    let authors = contributors.author.asPersonCollection();
    expect(authors.count).toEqual(2);

    let person = authors.at(0);
    expect(person.first).toEqual("Roxanne");
    expect(person.middle).toEqual("Brielle");
    expect(person.last).toEqual("Tejeda");
    //ExEnd:BibliographySources
  });

 
  test('BibliographyPersons', () => {
    //ExStart
    //ExFor:Person.#ctor(string, string, string)
    //ExFor:PersonCollection.#ctor
    //ExFor:PersonCollection.#ctor(Person[])
    //ExFor:PersonCollection.add(Person)
    //ExFor:PersonCollection.contains(Person)
    //ExFor:PersonCollection.clear
    //ExFor:PersonCollection.remove(Person)
    //ExFor:PersonCollection.removeAt(Int32)
    //ExSummary:Shows how to work with person collection.
    // Create a new person collection.
    let persons = new aw.Bibliography.PersonCollection();
    let person = new aw.Bibliography.Person("Roxanne", "Brielle", "Tejeda_updated");
    // Add new person to the collection.
    persons.add(person);
    expect(persons.count).toEqual(1);
    // Remove person from the collection if it exists.
    if (persons.contains(person))
      persons.remove(person);
    expect(persons.count).toEqual(0);

    // Create person collection with two persons.
    persons = new aw.Bibliography.PersonCollection([new aw.Bibliography.Person("Roxanne_1", "Brielle_1", "Tejeda_1"), new aw.Bibliography.Person("Roxanne_2", "Brielle_2", "Tejeda_2") ]);
    expect(persons.count).toEqual(2);
    // Remove person from the collection by the index.
    persons.removeAt(0);
    expect(persons.count).toEqual(1);
    // Remove all persons from the collection.
    persons.clear();
    expect(persons.count).toEqual(0);
    //ExEnd
  });


  test('CaptionlessTableOfFiguresLabel', () => {
    //ExStart
    //ExFor:FieldToc.captionlessTableOfFiguresLabel
    //ExSummary:Shows how to set the name of the sequence identifier.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let fieldToc = builder.insertField(aw.Fields.FieldType.FieldTOC, true).asFieldToc();
    fieldToc.captionlessTableOfFiguresLabel = "Test";

    expect(fieldToc.getFieldCode()).toEqual(" TOC  \\a Test");
    //ExEnd
  });

});
