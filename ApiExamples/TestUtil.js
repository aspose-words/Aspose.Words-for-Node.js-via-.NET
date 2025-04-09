// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const jimp = require("jimp");
var AdmZip = require('adm-zip');
var encoding = require('encoding-japanese');
var moment = require('moment');

/// <summary>
/// Checks whether a file at a specified filename contains a valid image with specified dimensions.
/// </summary>
/// <remarks>
/// Serves to check that an image file is valid and nonempty without looking up its file size.
/// </remarks>
/// <param name="expectedWidth">Expected width of the image, in pixels.</param>
/// <param name="expectedHeight">Expected height of the image, in pixels.</param>
/// <param name="filename">Local file system filename of the image file.</param>
async function verifyImage(expectedWidth, expectedHeight, filename) {
  const image = await jimp.Jimp.read(filename);
  expect(Math.abs(image.bitmap.width - expectedWidth)).toBeLessThanOrEqual(1);
  expect(Math.abs(image.bitmap.height - expectedHeight)).toBeLessThanOrEqual(1);
}


/// <summary>
/// Checks whether an image from the local file system contains any transparency.
/// </summary>
/// <param name="filename">Local file system filename of the image file.</param>
async function imageContainsTransparency(filename) {
  const image = await jimp.Jimp.read(filename);
  for (let x = 0; x < image.bitmap.width; x++) {
    for (let y = 0; y < image.bitmap.height; y++) {
      const pixel = jimp.intToRGBA(image.getPixelColor(x, y));
      if (pixel.a != 255) {
         return true;
      }
    }
  }

  return false;
}

/*  
    /// <summary>
    /// Checks whether an HTTP request sent to the specified address produces an expected web response. 
    /// </summary>
    /// <remarks>
    /// Serves as a notification of any URLs used in code examples becoming unusable in the future.
    /// </remarks>
    /// <param name="expectedHttpStatusCode">Expected result status code of a request HTTP "HEAD" method performed on the web address.</param>
    /// <param name="webAddress">URL where the request will be sent.</param>
  internal static async System.Threading.Tasks.Task VerifyWebResponseStatusCodeAsync(HttpStatusCode expectedHttpStatusCode, string webAddress)
  {
    var myClient = new System.Net.Http.HttpClient();
    var response = await myClient.GetAsync(webAddress);

    expect(response.StatusCode).toEqual(expectedHttpStatusCode);
  }

    /// <summary>
    /// Checks whether an SQL query performed on a database file stored in the local file system
    /// produces a result that resembles the contents of an Aspose.Words table.
    /// </summary>
    /// <param name="expectedResult">Expected result of the SQL query in the form of an Aspose.Words table.</param>
    /// <param name="dbFilename">Local system filename of a database file.</param>
    /// <param name="sqlQuery">Microsoft.Jet.OLEDB.4.0-compliant SQL query.</param>
  internal static void TableMatchesQueryResult(Table expectedResult, string dbFilename, string sqlQuery)
  {
    {
      connection.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbFilename};";
      connection.open();

      OleDbCommand command = connection.CreateCommand();
      command.CommandText = sqlQuery;
      OleDbDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

      let myDataTable = new DataTable();
      myDataTable.load(reader);

      expect(myDataTable.rows.count).toEqual(expectedResult.rows.count);
      expect(myDataTable.columns.count).toEqual(expectedResult.rows.at(0).Cells.count);

      for (let i = 0; i < myDataTable.rows.count; i++)
        for (let j = 0; j < myDataTable.columns.count; j++)
          Assert.AreEqual(expectedResult.rows.at(i).Cells.at(j).GetText().Replace(aw.ControlChar.cell, ''),
            myDataTable.rows.at(i)[j].ToString());
    }
  }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of every table produced by a list of consecutive SQL queries on a database.
    /// </summary>
    /// <remarks>
    /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    /// </remarks>
    /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    /// <param name="sqlQueries">List of SQL queries performed on the database all of whose results we expect to find in the document.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
  internal static void MailMergeMatchesQueryResultMultiple(string dbFilename, string[] sqlQueries, Document doc, bool onePagePerRow)
  {
    for (let query of sqlQueries)
      MailMergeMatchesQueryResult(dbFilename, query, doc, onePagePerRow);
  }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of a table produced by an SQL query on a database.
    /// </summary>
    /// <remarks>
    /// Currently, database types that cannot be represented by a string or a decimal are not checked for in the document.
    /// </remarks>
    /// <param name="dbFilename">Full local file system filename of a .mdb database file.</param>
    /// <param name="sqlQuery">SQL query performed on the database all of whose results we expect to find in the document.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
  internal static void MailMergeMatchesQueryResult(string dbFilename, string sqlQuery, Document doc, bool onePagePerRow)
  {
    List<string[]> expectedStrings = new aw.Lists.List<string[]>(); 
    string connectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + dbFilename;

    {
      let command = new OleDbCommand(sqlQuery, connection);
      command.CommandText = sqlQuery;

      try
      {
        connection.open();
        {
          while (reader.read())
          {
            string.at(] row = new string[reader.fieldCount);

            for (let i = 0; i < reader.fieldCount; i++)
              switch (reader.at(i))
              {
                case decimal d:
                  row.at(i) = d.toString("G29");
                  break;
                case string s:
                  row.at(i) = s.trim().Replace("\n", '');
                  break;
                default:
                  row.at(i) = '';
                  break;
              }

            expectedStrings.add(row);
          }
        }
      }
      catch (Exception ex)
      {
        console.log(ex.Message);
      }
    }

    MailMergeMatchesArray(expectedStrings.toArray(), doc, onePagePerRow);
  }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of every DataTable in a DataSet.
    /// </summary>
    /// <param name="dataSet">DataSet containing DataTables which contain values that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
  internal static void MailMergeMatchesDataSet(DataSet dataSet, Document doc, bool onePagePerRow)
  {
    for (let table of dataSet.tables)
      MailMergeMatchesDataTable(table, doc, onePagePerRow);
  }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of a DataTable.
    /// </summary>
    /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
  internal static void MailMergeMatchesDataTable(DataTable expectedResult, Document doc, bool onePagePerRow)
  {
    string.at(][) expectedStrings = new string.at(expectedResult.rows.count)[];

    for (let i = 0; i < expectedResult.rows.count; i++)
      expectedStrings.at(i) = Array.ConvertAll(expectedResult.rows.at(i).ItemArray, x => x.toString());

    MailMergeMatchesArray(expectedStrings, doc, onePagePerRow);
  }

    /// <summary>
    /// Checks whether a document produced during a mail merge contains every element of an array of arrays of strings.
    /// </summary>
    /// <remarks>
    /// Only suitable for rectangular arrays.
    /// </remarks>
    /// <param name="expectedResult">Values from the mail merge data source that we expect the document to contain.</param>
    /// <param name="doc">Document created during a mail merge.</param>
    /// <param name="onePagePerRow">True if the mail merge produced a document with one page per row in the data source.</param>
  internal static void MailMergeMatchesArray(string.at(][) expectedResult, Document doc, bool onePagePerRow)
  {
    try
    {
      if (onePagePerRow)
      {
        string.at(] docTextByPages = doc.getText().trim().Split(new[) { aw.ControlChar.pageBreak }, StringSplitOptions.RemoveEmptyEntries);

        for (let i = 0; i < expectedResult.length; i++)
          for (let j = 0; j < expectedResult.at(0).Length; j++)
            if (!docTextByPages[i].Contains(expectedResult.at(i)[j])) throw new ArgumentException(expectedResult.at(i)[j]);
      }
      else
      {
        string docText = doc.getText();

        for (let i = 0; i < expectedResult.length; i++)
          for (let j = 0; j < expectedResult.at(0).Length; j++)
            if (!docText.contains(expectedResult.at(i)[j])) throw new ArgumentException(expectedResult.at(i)[j]);

      }
    }
    catch (ArgumentException e)
    {
      Assert.Fail($"String \"{e.Message}\" not found in {(doc.originalFileName == null ? "a virtual document" : doc.originalFileName.split('\\').Last())}.");
    }
  }
*/    

  /// <summary>
  /// Checks whether a file inside a document's OOXML package contains a string.
  /// </summary>
  /// <remarks>
  /// If an output document does not have a testable value that can be found as a property in its object when loaded,
  /// the value can sometimes be found in the document's OOXML package. 
  /// </remarks>
  /// <param name="expected">The string we are looking for.</param>
  /// <param name="docFilename">Local file system filename of the document.</param>
  /// <param name="docPartFilename">Name of the file within the document opened as a .zip that is expected to contain the string.</param>
  function docPackageFileContainsString(expected, docFilename, docPartFilename) {
    const zip = new AdmZip(docFilename);
    const entry = zip.getEntries().find(e => e.entryName.endsWith(docPartFilename));
    if (!entry) {
      throw new Error(`Entry ${docPartFilename} not found in ${docFilename}`);
    }

    const entryContent = entry.getData().toString('utf8');
    expect(entryContent.includes(expected)).toEqual(true);
  }
  

/// <summary>
/// Checks whether a file in the local file system contains a string in its raw data.
/// </summary>
/// <param name="expected">The string we are looking for.</param>
/// <param name="filename">Local system filename of a file which, when read from the beginning, should contain the string.</param>
function fileContainsString(expected, filename) {
  let data = Array.from(fs.readFileSync(filename));
  let utf8Encode = new TextEncoder("utf-8");
  let expectData = Array.from(utf8Encode.encode(expected));
  expect(data).toEqual(expect.arrayContaining(expectData));
}

/// <summary>
/// Checks whether a file in the local file system doesn't contain a string in its raw data.
/// </summary>
/// <param name="expected">The string we are looking for.</param>
/// <param name="filename">Local system filename of a file which, when read from the beginning, should not contain the string.</param>
function fileNotContainString(expected, filename) {
  let data = Array.from(fs.readFileSync(filename));
  let utf8Encode = new TextEncoder("utf-8");
  let expectData = Array.from(utf8Encode.encode(expected));
  expect(data).toEqual(expect.not.arrayContaining(expectData));
}


/// <summary>
/// Checks whether a stream contains a string.
/// </summary>
/// <param name="expected">The string we are looking for.</param>
/// <param name="data">The array, when read from the beginning, should contain the string.</param>
function streamContainsString(expected, data) {
  let expectedSequence = Array.from(expected);
  expect(data).toEqual(expect.arrayContaining(expectedSequence));
}


/// <summary>
/// Checks whether values of properties of a field with a type related to date/time are equal to expected values.
/// </summary>
/// <remarks>
/// Used when comparing DateTime instances to Field.Result values parsed to DateTime, which may differ slightly. 
/// Give a delta value that's generous enough for any lower end system to pass, also a delta of zero is allowed.
/// </remarks>
/// <param name="expectedType">The FieldType that we expect the field to have.</param>
/// <param name="expectedFieldCode">The expected output value of GetFieldCode() being called on the field.</param>
/// <param name="expectedResult">The date/time that the field's result is expected to represent.</param>
/// <param name="field">The field that's being tested.</param>
/// <param name="delta">Margin of error for expectedResult.</param>
function verifyField(expectedType, expectedFieldCode, expectedResult, field, delta) {
  expect(field.type).toEqual(expectedType);
  expect(field.getFieldCode(true)).toEqual(expectedFieldCode);
  if (delta !== undefined) {
    let actual;
    expect (() => {
      // NZ Date format - "DD/MM/YYYY"
      actual = moment(field.result, ["DD/MM/YYYY", "DD.MM.YYYY", "hh:mm:ss"]);
    }).not.toThrow();
    expect(actual.isValid()).toEqual(true);

    verifyDate(expectedResult, actual, delta);
  } else {
    expect(field.result).toEqual(expectedResult);
  }
}

/// <summary>
/// Checks whether a DateTime matches an expected value, with a margin of error.
/// </summary>
/// <param name="expected">The date/time that we expect the result to be.</param>
/// <param name="actual">The DateTime object being tested.</param>
/// <param name="delta">Margin of error for expectedResult.</param>
function verifyDate(expected, actual, delta) {
  expect(expected - actual).toBeLessThanOrEqual(delta);
}

  /// <summary>
  /// Checks whether a field contains another complete field as a sibling within its nodes.
  /// </summary>
  /// <remarks>
  /// If two fields have the same immediate parent node and therefore their nodes are siblings,
  /// the FieldStart of the outer field appears before the FieldStart of the inner node,
  /// and the FieldEnd of the outer node appears after the FieldEnd of the inner node,
  /// then the inner field is considered to be nested within the outer field. 
  /// </remarks>
  /// <param name="innerField">The field that we expect to be fully within outerField.</param>
  /// <param name="outerField">The field that we to contain innerField.</param>
  function fieldsAreNested(innerField, outerField) {
    let innerFieldParent = innerField.start.parentNode;

    expect(innerFieldParent.referenceEquals(outerField.start.parentNode)).toEqual(true);
    expect(innerFieldParent.getChildNodes(aw.NodeType.Any, false).indexOf(innerField.start) >
           innerFieldParent.getChildNodes(aw.NodeType.Any, false).indexOf(outerField.start)).toEqual(true);
    expect(innerFieldParent.getChildNodes(aw.NodeType.Any, false).toArray().indexOf(innerField.end) <
           innerFieldParent.getChildNodes(aw.NodeType.Any, false).toArray().indexOf(outerField.end)).toEqual(true);
  }

  /// <summary>
  /// Checks whether a shape contains a valid image with specified dimensions.
  /// </summary>
  /// <remarks>
  /// Serves to check that an image file is valid and nonempty without looking up its data length.
  /// </remarks>
  /// <param name="expectedWidth">Expected width of the image, in pixels.</param>
  /// <param name="expectedHeight">Expected height of the image, in pixels.</param>
  /// <param name="expectedImageType">Expected format of the image.</param>
  /// <param name="imageShape">Shape that contains the image.</param>
  function verifyImageInShape(expectedWidth, expectedHeight, expectedImageType, imageShape) {
    expect(imageShape.hasImage).toEqual(true);
    expect(imageShape.imageData.imageType).toEqual(expectedImageType);
    expect(imageShape.imageData.imageSize.widthPixels).toEqual(expectedWidth);
    expect(imageShape.imageData.imageSize.heightPixels).toEqual(expectedHeight);
  }

  /// <summary>
  /// Checks whether values of a footnote's properties are equal to their expected values.
  /// </summary>
  /// <param name="expectedFootnoteType">Expected type of the footnote/endnote.</param>
  /// <param name="expectedIsAuto">Expected auto-numbered status of this footnote.</param>
  /// <param name="expectedReferenceMark">If "IsAuto" is false, then the footnote is expected to display this string instead of a number after referenced text.</param>
  /// <param name="expectedContents">Expected side comment provided by the footnote.</param>
  /// <param name="footnote">Footnote node in question.</param>
  function verifyFootnote(expectedFootnoteType, expectedIsAuto, expectedReferenceMark, expectedContents, footnote) {
    expect(footnote.footnoteType).toEqual(expectedFootnoteType);
    expect(footnote.isAuto).toEqual(expectedIsAuto);
    expect(footnote.referenceMark).toEqual(expectedReferenceMark);
    expect(footnote.toString(aw.SaveFormat.Text).trim()).toEqual(expectedContents);
  }

/// <summary>
/// Checks whether values of a list level's properties are equal to their expected values.
/// </summary>
/// <remarks>
/// Only necessary for list levels that have been explicitly created by the user.
/// </remarks>
/// <param name="expectedListFormat">Expected format for the list symbol.</param>
/// <param name="expectedNumberPosition">Expected indent for this level, usually growing larger with each level.</param>
/// <param name="expectedNumberStyle"></param>
/// <param name="listLevel">List level in question.</param>
function verifyListLevel(expectedListFormat, expectedNumberPosition, expectedNumberStyle, listLevel) {
  expect(listLevel.numberFormat).toEqual(expectedListFormat);
  expect(listLevel.numberPosition).toEqual(expectedNumberPosition);
  expect(listLevel.numberStyle).toEqual(expectedNumberStyle);
}

/// <summary>
/// Checks whether values of a tab stop's properties are equal to their expected values.
/// </summary>
/// <param name="expectedPosition">Expected position on the tab stop ruler, in points.</param>
/// <param name="expectedTabAlignment">Expected position where the position is measured from </param>
/// <param name="expectedTabLeader">Expected characters that pad the space between the start and end of the tab whitespace.</param>
/// <param name="isClear">Whether or no this tab stop clears any tab stops.</param>
/// <param name="tabStop">Tab stop that's being tested.</param>
function verifyTabStop(expectedPosition, expectedTabAlignment, expectedTabLeader, isClear, tabStop) {
  expect(tabStop.position).toEqual(expectedPosition);
  expect(tabStop.alignment).toEqual(expectedTabAlignment);
  expect(tabStop.leader).toEqual(expectedTabLeader);
  expect(tabStop.isClear).toEqual(isClear);
}

/*
    /// <summary>
    /// Copies from the current position in src stream till the end.
    /// Copies into the current position in dst stream.
    /// </summary>
  internal static void CopyStream(Stream srcStream, Stream dstStream)
  {
    if (srcStream == null)
      throw new ArgumentNullException("srcStream");
    if (dstStream == null)
      throw new ArgumentNullException("dstStream");

    byte.at(] buf = new byte[65536);
    while (true)
    {
      int bytesRead = srcStream.read(buf, 0, buf.length);
        // Read returns 0 when reached end of stream
        // Checking for negative too to make it conceptually close to Java
      if (bytesRead <= 0)
        break;
      dstStream.write(buf, 0, bytesRead);
    }
  }

    /// <summary>
    /// Checks whether values of a shape's properties are equal to their expected values.
    /// </summary>
    /// <remarks>
    /// All dimension measurements are in points.
    /// </remarks>
  internal static void VerifyShape(ShapeType expectedShapeType, string expectedName, double expectedWidth, double expectedHeight, double expectedTop, double expectedLeft, Shape shape)
  {
    Assert.multiple(() =>
    {
      expect(shape.shapeType).toEqual(expectedShapeType);
      expect(shape.name).toEqual(expectedName);
      expect(shape.width).toEqual(expectedWidth);
      expect(shape.height).toEqual(expectedHeight);
      expect(shape.top).toEqual(expectedTop);
      expect(shape.left).toEqual(expectedLeft);
    });
  }*/

  /// <summary>
  /// Checks whether values of properties of an editable range are equal to their expected values.
  /// </summary>
  function verifyEditableRange(expectedId, expectedEditorUser, expectedEditorGroup, editableRange) {
    expect(editableRange.id).toEqual(expectedId);
    expect(editableRange.singleUser).toEqual(expectedEditorUser);
    expect(editableRange.editorGroup).toEqual(expectedEditorGroup);
  }

/// <summary>
/// Get File's Encoding.
/// </summary>
function getEncoding(filename) {
  var fileBuffer = fs.readFileSync(filename);
  return encoding.detect(fileBuffer);
}

/// <summary>
/// Converts an image to a byte array.
/// </summary>
function imageToByteArray(imagePath) {
  return Array.from(fs.readFileSync(imagePath));
}

/// <summary>
/// Dumps byte array into a string.
/// </summary>
function dumpArray(data, start, count) {
  if (data == null)
    return "Null";

  const key = '0123456789ABCDEF'
  let dest = ""
  while (count > 0)
  {
    let leftByte = key[data[start] >> 4];
    let rightByte = key[data[start] & 15];
    dest += `${leftByte}${rightByte} `
    start++;
    count--;
  }
  return dest
}

/// <summary>
/// Checks whether values of a shape's properties are equal to their expected values.
/// </summary>
/// <remarks>
/// All dimension measurements are in points.
/// </remarks>
function verifyShape(expectedShapeType, expectedName, expectedWidth, expectedHeight, expectedTop, expectedLeft, shape) {
  expect(shape.shapeType).toEqual(expectedShapeType);
  expect(shape.name).toEqual(expectedName);
  expect(shape.width).toEqual(expectedWidth);
  expect(shape.height).toEqual(expectedHeight);
  expect(shape.top).toEqual(expectedTop);
  expect(shape.left).toEqual(expectedLeft);
}

/// <summary>
/// Checks whether values of properties of a textbox are equal to their expected values.
/// </summary>
/// <remarks>
/// All dimension measurements are in points.
/// </remarks>
function verifyTextBox(expectedLayoutFlow, expectedFitShapeToText, expectedTextBoxWrapMode, marginTop, marginBottom, marginLeft, marginRight, textBox) {
  expect(textBox.layoutFlow).toEqual(expectedLayoutFlow);
  expect(textBox.fitShapeToText).toEqual(expectedFitShapeToText);
  expect(textBox.textBoxWrapMode).toEqual(expectedTextBoxWrapMode);
  expect(textBox.internalMarginTop).toEqual(marginTop);
  expect(textBox.internalMarginBottom).toEqual(marginBottom);
  expect(textBox.internalMarginLeft).toEqual(marginLeft);
  expect(textBox.internalMarginRight).toEqual(marginRight);
}

module.exports = {
  verifyField,
  verifyDate,
  verifyImage,
  docPackageFileContainsString,
  verifyEditableRange,
  getEncoding,
  fieldsAreNested,
  verifyImageInShape,
  imageContainsTransparency,
  verifyListLevel,
  streamContainsString,
  fileContainsString,
  fileNotContainString,
  verifyFootnote,
  verifyTabStop,
  imageToByteArray,
  dumpArray,
  verifyShape,
  verifyTextBox
};
