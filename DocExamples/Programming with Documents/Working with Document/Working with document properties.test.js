// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;


describe("WorkingWithDocumentProperties", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('GetVariables', () => {
    //ExStart:GetVariables
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document(base.myDir + "Document.docx");

    let variables = "";
    for (let entry of doc.variables) {
      let name = entry.key;
      let value = entry.value;
      if (variables === "")
        variables = "Name: " + name + "," + "Value: {1}" + value;
      else
        variables += "Name: " + name + "," + "Value: {1}" + value;
    }

    console.log("\nDocument has the following variables: " + variables);
    //ExEnd:GetVariables
  });

  test('EnumerateProperties', () => {
    //ExStart:EnumerateProperties
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document(base.myDir + "Properties.docx");

    console.log("1. Document name: {0}", doc.originalFileName);
    console.log("2. Built-in Properties");

    for (let prop of doc.builtInDocumentProperties) {
      console.log("{0} : {1}", prop.name, prop.value);
    }

    console.log("3. Custom Properties");

    for (let prop of doc.customDocumentProperties) {
      console.log("{0} : {1}", prop.name, prop.value);
    }
    //ExEnd:EnumerateProperties
  });

  test('AddCustomProperties', () => {
    //ExStart:AddCustomProperties
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document(base.myDir + "Properties.docx");

    let customDocumentProperties = doc.customDocumentProperties;

    if (customDocumentProperties["Authorized"] !== null) return;

    customDocumentProperties.add("Authorized", true);
    customDocumentProperties.add("Authorized By", "John Smith");
    customDocumentProperties.add("Authorized Date", new Date());
    customDocumentProperties.add("Authorized Revision", doc.builtInDocumentProperties.revisionNumber);
    customDocumentProperties.add("Authorized Amount", 123.45);
    //ExEnd:AddCustomProperties
  });

  test('RemoveCustomProperties', () => {
    //ExStart:RemoveCustomProperties
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document(base.myDir + "Properties.docx");
    doc.customDocumentProperties.remove("Authorized Date");
    //ExEnd:RemoveCustomProperties
  });

  test('RemovePersonalInformation', () => {
    //ExStart:RemovePersonalInformation
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document(base.myDir + "Properties.docx");
    doc.removePersonalInformation = true;

    doc.save(base.artifactsDir + "DocumentPropertiesAndVariables.RemovePersonalInformation.docx");
    //ExEnd:RemovePersonalInformation
  });

  test('ConfiguringLinkToContent', () => {
    //ExStart:ConfiguringLinkToContent
    //GistId:9bd62e688457850bceba59bc2c0ead99
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("MyBookmark");
    builder.writeln("Text inside a bookmark.");
    builder.endBookmark("MyBookmark");

    let customProperties = doc.customDocumentProperties;
    let customProperty = customProperties.addLinkToContent("Bookmark", "MyBookmark");
    customProperty = customProperties.at("Bookmark");

    let isLinkedToContent = customProperty.isLinkToContent;
    let linkSource = customProperty.linkSource;
    let customPropertyValue = customProperty.value;
    //ExEnd:ConfiguringLinkToContent
  });

  test('ConvertBetweenMeasurementUnits', () => {
    //ExStart:ConvertBetweenMeasurementUnits
    //GistId:0fd389e71a7e1ef74e6e48e94d37be5d
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let pageSetup = builder.pageSetup;
    pageSetup.topMargin = aw.ConvertUtil.inchToPoint(1.0);
    pageSetup.bottomMargin = aw.ConvertUtil.inchToPoint(1.0);
    pageSetup.leftMargin = aw.ConvertUtil.inchToPoint(1.5);
    pageSetup.rightMargin = aw.ConvertUtil.inchToPoint(1.5);
    pageSetup.headerDistance = aw.ConvertUtil.inchToPoint(0.2);
    pageSetup.footerDistance = aw.ConvertUtil.inchToPoint(0.2);
    //ExEnd:ConvertBetweenMeasurementUnits
  });

  test('UseControlCharacters', () => {
    //ExStart:UseControlCharacters
    //GistId:3d84715449d1d04a6029964ad5f2fdf0
    let text = "test\r";
    // Replace "\r" control character with "\r\n".
    let replace = text.replace(aw.ControlChar.cr, aw.ControlChar.crLf);
    //ExEnd:UseControlCharacters
  });
});
