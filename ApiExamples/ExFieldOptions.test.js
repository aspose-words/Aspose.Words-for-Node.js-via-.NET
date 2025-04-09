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

describe("ExFieldOptions", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CurrentUser', () => {
    //ExStart
    //ExFor:Document.updateFields
    //ExFor:FieldOptions.currentUser
    //ExFor:UserInformation
    //ExFor:UserInformation.name
    //ExFor:UserInformation.initials
    //ExFor:UserInformation.address
    //ExFor:UserInformation.defaultUser
    //ExSummary:Shows how to set user details, and display them using fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Create a UserInformation object and set it as the data source for fields that display user information.
    let userInformation = new aw.Fields.UserInformation();
    userInformation.name = "John Doe";
    userInformation.initials = "J. D.";
    userInformation.address = "123 Main Street";
    doc.fieldOptions.currentUser = userInformation;

    // Insert USERNAME, USERINITIALS, and USERADDRESS fields, which display values of
    // the respective properties of the UserInformation object that we have created above.
    expect(builder.insertField(" USERNAME ").result).toEqual(userInformation.name);
    expect(builder.insertField(" USERINITIALS ").result).toEqual(userInformation.initials);
    expect(builder.insertField(" USERADDRESS ").result).toEqual(userInformation.address);

    // The field options object also has a static default user that fields from all documents can refer to.
    aw.Fields.UserInformation.defaultUser.name = "Default User";
    aw.Fields.UserInformation.defaultUser.initials = "D. U.";
    aw.Fields.UserInformation.defaultUser.address = "One Microsoft Way";
    doc.fieldOptions.currentUser = aw.Fields.UserInformation.defaultUser;

    expect(builder.insertField(" USERNAME ").result).toEqual("Default User");
    expect(builder.insertField(" USERINITIALS ").result).toEqual("D. U.");
    expect(builder.insertField(" USERADDRESS ").result).toEqual("One Microsoft Way");

    doc.updateFields();
    doc.save(base.artifactsDir + "FieldOptions.currentUser.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FieldOptions.currentUser.docx");

    expect(doc.fieldOptions.currentUser).toBe(null);

    let fieldUserName = doc.range.fields.at(0).asFieldUserName();

    expect(fieldUserName.userName).toBe(null);
    expect(fieldUserName.result).toEqual("Default User");

    let fieldUserInitials = doc.range.fields.at(1).asFieldUserInitials();

    expect(fieldUserInitials.userInitials).toBe(null);
    expect(fieldUserInitials.result).toEqual("D. U.");

    let fieldUserAddress = doc.range.fields.at(2).asFieldUserAddress();

    expect(fieldUserAddress.userAddress).toBe(null);
    expect(fieldUserAddress.result).toEqual("One Microsoft Way");
  });


  test('FileName', () => {
    //ExStart
    //ExFor:FieldOptions.fileName
    //ExFor:FieldFileName
    //ExFor:FieldFileName.includeFullPath
    //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let builder = new aw.DocumentBuilder(doc);

    builder.moveToDocumentEnd();
    builder.writeln();

    // This FILENAME field will display the local system file name of the document we loaded.
    let field = builder.insertField(aw.Fields.FieldType.FieldFileName, true).asFieldFileName();
    field.update();

    expect(field.getFieldCode()).toEqual(" FILENAME ");
    expect(field.result).toEqual("Document.docx");

    builder.writeln();

    // By default, the FILENAME field shows the file's name, but not its full local file system path.
    // We can set a flag to make it show the full file path.
    field = builder.insertField(aw.Fields.FieldType.FieldFileName, true).asFieldFileName();
    field.includeFullPath = true;
    field.update();

    expect(field.result).toEqual(base.myDir + "Document.docx");

    // We can also set a value for this property to
    // override the value that the FILENAME field displays.
    doc.fieldOptions.fileName = "FieldOptions.FILENAME.docx";
    field.update();

    expect(field.getFieldCode()).toEqual(" FILENAME  \\p");
    expect(field.result).toEqual("FieldOptions.FILENAME.docx");

    doc.updateFields();
    doc.save(base.artifactsDir + doc.fieldOptions.fileName);
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FieldOptions.FILENAME.docx");

    expect(doc.fieldOptions.fileName).toBe(null);
    TestUtil.verifyField(aw.Fields.FieldType.FieldFileName, " FILENAME ", "FieldOptions.FILENAME.docx", doc.range.fields.at(0));
  });


  test('Bidi', () => {
    //ExStart
    //ExFor:FieldOptions.isBidiTextSupportedOnUpdate
    //ExSummary:Shows how to use FieldOptions to ensure that field updating fully supports bi-directional text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Ensure that any field operation involving right-to-left text is performs as expected.
    doc.fieldOptions.isBidiTextSupportedOnUpdate = true;

    // Use a document builder to insert a field that contains the right-to-left text.
    let comboBox = builder.insertComboBox("MyComboBox", ["עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים"], 0);
    comboBox.calculateOnExit = true;

    doc.updateFields();
    doc.save(base.artifactsDir + "FieldOptions.bidi.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FieldOptions.bidi.docx");

    expect(doc.fieldOptions.isBidiTextSupportedOnUpdate).toEqual(false);

    comboBox = doc.range.formFields.at(0);

    expect(comboBox.result).toEqual("עֶשְׂרִים");
  });


  test('LegacyNumberFormat', () => {
    //ExStart
    //ExFor:FieldOptions.legacyNumberFormat
    //ExSummary:Shows how enable legacy number formatting for fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let field = builder.insertField("= 2 + 3 \\# $##");

    expect(field.result).toEqual("$ 5");

    doc.fieldOptions.legacyNumberFormat = true;
    field.update();

    expect(field.result).toEqual("$5");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.fieldOptions.legacyNumberFormat).toEqual(false);
    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, "= 2 + 3 \\# $##", "$5", doc.range.fields.at(0));
  });


  test.skip('PreProcessCulture: unsupported culture (CultureInfo)', () => {
    //ExStart
    //ExFor:FieldOptions.preProcessCulture
    //ExSummary:Shows how to set the preprocess culture.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let builder = new aw.DocumentBuilder(doc);

    // Set the culture according to which some fields will format their displayed values.
    doc.fieldOptions.preProcessCulture = new CultureInfo("de-DE");

    let field = builder.insertField(" DOCPROPERTY CreateTime");

    // The DOCPROPERTY field will display its result formatted according to the preprocess culture
    // we have set to German. The field will display the date/time using the "dd.mm.yyyy hh:mm" format.
    expect(new RegExp(String.raw`\d{2}[.]\d{2}[.]\d{4} \d{2}[:]\d{2}`).test(field.result)).toEqual(true);

    doc.fieldOptions.preProcessCulture = CultureInfo.InvariantCulture;
    field.update();

    // After switching to the invariant culture, the DOCPROPERTY field will use the "mm/dd/yyyy hh:mm" format.
    expect(new RegExp(String.raw`\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}`).test(field.result)).toEqual(true);
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.fieldOptions.preProcessCulture).toBe(null);
    expect(new RegExp(String.raw`\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}`).test(doc.range.fields.at(0).result)).toEqual(true);
  });


  test('TableOfAuthorityCategories', () => {
    //ExStart
    //ExFor:FieldOptions.toaCategories
    //ExFor:ToaCategories
    //ExFor:ToaCategories.item(Int32)
    //ExFor:ToaCategories.defaultCategories
    //ExSummary:Shows how to specify a set of categories for TOA fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // TOA fields can filter their entries by categories defined in this collection.
    let toaCategories = new aw.Fields.ToaCategories();
    doc.fieldOptions.toaCategories = toaCategories;

    // This collection of categories comes with default values, which we can overwrite with custom values.
    expect(toaCategories.at(1)).toEqual("Cases");
    expect(toaCategories.at(2)).toEqual("Statutes");

    toaCategories.setAt(1, "My Category 1");
    toaCategories.setAt(2, "My Category 2");

    // We can always access the default values via this collection.
    expect(aw.Fields.ToaCategories.defaultCategories.at(1)).toEqual("Cases");
    expect(aw.Fields.ToaCategories.defaultCategories.at(2)).toEqual("Statutes");

    // Insert 2 TOA fields. TOA fields create an entry for each TA field in the document.
    // Use the "\c" switch to select the index of a category from our collection.
    //  With this switch, a TOA field will only pick up entries from TA fields that
    // also have a "\c" switch with a matching category index. Each TOA field will also display
    // the name of the category that its "\c" switch points to.
    builder.insertField("TOA \\c 1 \\h", null);
    builder.insertField("TOA \\c 2 \\h", null);
    builder.insertBreak(aw.BreakType.PageBreak);

    // Insert TOA entries across 2 categories. Our first TOA field will receive one entry,
    // from the second TA field whose "\c" switch also points to the first category.
    // The second TOA field will have two entries from the other two TA fields.
    builder.insertField("TA \\c 2 \\l \"entry 1\"");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertField("TA \\c 1 \\l \"entry 2\"");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.insertField("TA \\c 2 \\l \"entry 3\"");

    doc.updateFields();
    doc.save(base.artifactsDir + "FieldOptions.TOA.categories.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FieldOptions.TOA.categories.docx");

    expect(doc.fieldOptions.toaCategories).toBe(null);

    TestUtil.verifyField(aw.Fields.FieldType.FieldTOA, "TOA \\c 1 \\h", "My Category 1\rentry 2\t3\r", doc.range.fields.at(0));
    TestUtil.verifyField(aw.Fields.FieldType.FieldTOA, "TOA \\c 2 \\h",
      "My Category 2\r" +
      "entry 1\t2\r" +
      "entry 3\t4\r", doc.range.fields.at(1));
  });


  test.skip('UseInvariantCultureNumberFormat: unsupported culture (CultureInfo)', () => {
    //ExStart
    //ExFor:FieldOptions.useInvariantCultureNumberFormat
    //ExSummary:Shows how to format numbers according to the invariant culture.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    Thread.currentThread.CurrentCulture = new CultureInfo("de-DE");
    let field = builder.insertField(" = 1234567,89 \\# $#,###,###.##");
    field.update();

    // Sometimes, fields may not format their numbers correctly under certain cultures.
    expect(doc.fieldOptions.useInvariantCultureNumberFormat).toEqual(false);
    expect(field.result).toEqual("$1.234.567,89 ,     ");

    // To fix this, we could change the culture for the entire thread.
    // Another way to fix this is to set this flag,
    // which gets all fields to use the invariant culture when formatting numbers.
    // This way allows us to avoid changing the culture for the entire thread.
    doc.fieldOptions.useInvariantCultureNumberFormat = true;
    field.update();
    expect(field.result).toEqual("$1.234.567,89");
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);

    expect(doc.fieldOptions.useInvariantCultureNumberFormat).toEqual(false);
    TestUtil.verifyField(aw.Fields.FieldType.FieldFormula, " = 1234567,89 \\# $#,###,###.##", "$1.234.567,89", doc.range.fields.at(0));
  });


  /*//Commented
    //ExStart
    //ExFor:FieldOptions.FieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    //ExSummary:Shows how to specify a culture which parses date/time formatting for each field.
  test('DefineDateTimeFormatting', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField(aw.Fields.FieldType.FieldTime, true);

    doc.fieldOptions.fieldUpdateCultureSource = aw.Fields.FieldUpdateCultureSource.FieldCode;

    // Set a provider that returns a culture object specific to each field.
    doc.fieldOptions.fieldUpdateCultureProvider = new FieldUpdateCultureProvider();

    let fieldDate = (FieldTime)doc.range.fields.at(0);
    if (fieldDate.localeId != (int)aw.Loading.EditingLanguage.Russian)
      fieldDate.localeId = (int)aw.Loading.EditingLanguage.Russian;

    doc.save(base.artifactsDir + "FieldOptions.UpdateDateTimeFormatting.pdf");
  });


    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a field.
    /// </summary>
  private class FieldUpdateCultureProvider : IFieldUpdateCultureProvider
  {
      /// <summary>
      /// Returns a CultureInfo object to be used during the field's update.
      /// </summary>
    public CultureInfo GetCulture(string name, Field field)
    {
      switch (name)
      {
        case "ru-RU":
          let culture = new CultureInfo(name, false);
          DateTimeFormatInfo format = culture.dateTimeFormat;

          format.MonthNames = new[] { "месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8", "месяц 9", "месяц 10", "месяц 11", "месяц 12", "" };
          format.MonthGenitiveNames = format.MonthNames;
          format.AbbreviatedMonthNames = new[] { "мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10", "мес 11", "мес 12", "" };
          format.AbbreviatedMonthGenitiveNames = format.AbbreviatedMonthNames;

          format.DayNames = new[] { "день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4", "день недели 5", "день недели 6" };
          format.AbbreviatedDayNames = new[] { "день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6" };
          format.ShortestDayNames = new[] { "д7", "д1", "д2", "д3", "д4", "д5", "д6" };

          format.AMDesignator = "До полудня";
          format.PMDesignator = "После полудня";

          const string pattern = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
          format.LongDatePattern = pattern;
          format.LongTimePattern = pattern;
          format.ShortDatePattern = pattern;
          format.ShortTimePattern = pattern;

          return culture;
        case "en-US":
          return new CultureInfo(name, false);
        default:
          return null;
      }
    }
  }
  //ExEnd
  //EndCommented*/
});
