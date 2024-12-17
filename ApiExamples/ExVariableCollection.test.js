// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;


describe("ExVariableCollection", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Primer', () => {
    //ExStart
    //ExFor:aw.Document.variables
    //ExFor:VariableCollection
    //ExFor:aw.VariableCollection.add
    //ExFor:aw.VariableCollection.clear
    //ExFor:aw.VariableCollection.contains
    //ExFor:aw.VariableCollection.count
    //ExFor:aw.VariableCollection.getEnumerator
    //ExFor:aw.VariableCollection.indexOfKey
    //ExFor:aw.VariableCollection.remove
    //ExFor:aw.VariableCollection.removeAt
    //ExSummary:Shows how to work with a document's variable collection.
    let doc = new aw.Document();
    let variables = doc.variables;

    // Every document has a collection of key/value pair variables, which we can add items to.
    variables.add("Home address", "123 Main St.");
    variables.add("City", "London");
    variables.add("Bedrooms", "3");

    expect(variables.count).toEqual(3);

    // We can display the values of variables in the document body using DOCVARIABLE fields.
    let builder = new aw.DocumentBuilder(doc);
    let field = builder.insertField(aw.Fields.FieldType.FieldDocVariable, true).asFieldDocVariable();
    field.variableName = "Home address";
    field.update();

    expect(field.result).toEqual("123 Main St.");

    // Assigning values to existing keys will update them.
    variables.add("Home address", "456 Queen St.");

    // We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
    expect(field.result).toEqual("123 Main St.");

    field.update();

    expect(field.result).toEqual("456 Queen St.");

    // Verify that the document variables with a certain name or value exist.
    expect(variables.contains("City")).toEqual(true);

    // The collection of variables automatically sorts variables alphabetically by name.
    expect(variables.indexOfKey("Bedrooms")).toEqual(0);
    expect(variables.indexOfKey("City")).toEqual(1);
    expect(variables.indexOfKey("Home address")).toEqual(2);

    // Enumerate over the collection of variables.
    for (var i = 0; i < variables.count; i++) {
      console.log(`Index: ${i}, Value: ${variables.at(i)}`);
    }

    // Below are three ways of removing document variables from a collection.
    // 1 -  By name:
    variables.remove("City");

    expect(variables.contains("City")).toEqual(false);

    // 2 -  By index:
    variables.removeAt(1);

    expect(variables.contains("Home address")).toEqual(false);

    // 3 -  Clear the whole collection at once:
    variables.clear();

    expect(variables.count).toEqual(0);
    //ExEnd
  });

});
