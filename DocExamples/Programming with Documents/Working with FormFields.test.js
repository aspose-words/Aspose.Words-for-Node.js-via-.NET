// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithFormFields", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('InsertFormFields', () => {
    //ExStart:InsertFormFields
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let items = ["One", "Two", "Three"];
    builder.insertComboBox("DropDown", items, 0);
    //ExEnd:InsertFormFields
  });

  test('FormFieldsWorkWithProperties', () => {
    //ExStart:FormFieldsWorkWithProperties
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document(base.myDir + "Form fields.docx");
    let formField = doc.range.formFields.at(3);
    if (formField.type == aw.Fields.FieldType.FieldFormTextInput)
        formField.result = "My name is " + formField.name;
    //ExEnd:FormFieldsWorkWithProperties
  });

  test('FormFieldsGetFormFieldsCollection', () => {
    //ExStart:FormFieldsGetFormFieldsCollection
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document(base.myDir + "Form fields.docx");

    let formFields = doc.range.formFields;
    //ExEnd:FormFieldsGetFormFieldsCollection
  });

  test('FormFieldsGetByName', () => {
    //ExStart:FormFieldsFontFormatting
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    //ExStart:FormFieldsGetByName
    //GistId:a317eda2c6381dd30c7eb70510e51d52
    let doc = new aw.Document(base.myDir + "Form fields.docx");
    let documentFormFields = doc.range.formFields;
    let formField1 = documentFormFields.at(3);
    let formField2 = documentFormFields.at("Text2");
    //ExEnd:FormFieldsGetByName
    formField1.font.size = 20;
    formField2.font.color = "#FF0000";
    //ExEnd:FormFieldsFontFormatting
  });

});