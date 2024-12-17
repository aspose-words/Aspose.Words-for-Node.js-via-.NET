// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');

describe("ExFormFields", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Create', () => {
    //ExStart
    //ExFor:FormField
    //ExFor:aw.Fields.FormField.result
    //ExFor:aw.Fields.FormField.type
    //ExFor:aw.Fields.FormField.name
    //ExSummary:Shows how to insert a combo box.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Please select a fruit: ");

    // Insert a combo box which will allow a user to choose an option from a collection of strings.
    let comboBox = builder.insertComboBox("MyComboBox", ["Apple", "Banana", "Cherry"], 0);

    expect(comboBox.name).toEqual("MyComboBox");
    expect(comboBox.type).toEqual(aw.Fields.FieldType.FieldFormDropDown);
    expect(comboBox.result).toEqual("Apple");

    // The form field will appear in the form of a "select" html tag.
    doc.save(base.artifactsDir + "FormFields.create.html");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FormFields.create.html");
    comboBox = doc.range.formFields.at(0);

    expect(comboBox.name).toEqual("MyComboBox");
    expect(comboBox.type).toEqual(aw.Fields.FieldType.FieldFormDropDown);
    expect(comboBox.result).toEqual("Apple");
  });


  test('TextInput', () => {
    //ExStart
    //ExFor:aw.DocumentBuilder.insertTextInput
    //ExSummary:Shows how to insert a text input form field.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Please enter text here: ");

    // Insert a text input field, which will allow the user to click it and enter text.
    // Assign some placeholder text that the user may overwrite and pass
    // a maximum text length of 0 to apply no limit on the form field's contents.
    builder.insertTextInput("TextInput1", aw.Fields.TextFormFieldType.Regular, "", "Placeholder text", 0);

    // The form field will appear in the form of an "input" html tag, with a type of "text".
    doc.save(base.artifactsDir + "FormFields.textInput.html");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "FormFields.textInput.html");

    let textInput = doc.range.formFields.at(0);

    expect(textInput.name).toEqual("TextInput1");
    expect(textInput.textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
    expect(textInput.textInputFormat).toEqual('');
    expect(textInput.result).toEqual("Placeholder text");
    expect(textInput.maxLength).toEqual(0);
  });


  test('DeleteFormField', () => {
    //ExStart
    //ExFor:aw.Fields.FormField.removeField
    //ExSummary:Shows how to delete a form field.
    let doc = new aw.Document(base.myDir + "Form fields.docx");

    let formField = doc.range.formFields.at(3);
    formField.removeField();
    //ExEnd

    let formFieldAfter = doc.range.formFields.at(3);

    expect(formFieldAfter).toBe(null);
  });


  test('DeleteFormFieldAssociatedWithBookmark', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.startBookmark("MyBookmark");
    builder.insertTextInput("TextInput1", aw.Fields.TextFormFieldType.Regular, "TestFormField", "SomeText", 0);
    builder.endBookmark("MyBookmark");

    doc = DocumentHelper.saveOpen(doc);

    let bookmarkBeforeDeleteFormField = doc.range.bookmarks;
    expect(bookmarkBeforeDeleteFormField.at(0).name).toEqual("MyBookmark");

    let formField = doc.range.formFields.at(0);
    formField.removeField();

    let bookmarkAfterDeleteFormField = doc.range.bookmarks;
    expect(bookmarkAfterDeleteFormField.at(0).name).toEqual("MyBookmark");
  });


  test('FormFieldFontFormatting', () => {
    //ExStart
    //ExFor:FormField
    //ExSummary:Shows how to formatting the entire FormField, including the field value.
    let doc = new aw.Document(base.myDir + "Form fields.docx");

    let formField = doc.range.formFields.at(0);
    formField.font.bold = true;
    formField.font.size = 24;
    formField.font.color = "#FF0000";

    formField.result = "Aspose.formField";

    doc = DocumentHelper.saveOpen(doc);

    let formFieldRun = doc.firstSection.body.firstParagraph.runs.at(1);

    expect(formFieldRun.text).toEqual("Aspose.formField");
    expect(formFieldRun.font.bold).toEqual(true);
    expect(formFieldRun.font.size).toEqual(24);
    expect(formFieldRun.font.color).toEqual("#FF0000");
    //ExEnd
  });


  /*  //ExStart
    //ExFor:FormField.Accept(DocumentVisitor)
    //ExFor:FormField.CalculateOnExit
    //ExFor:FormField.CheckBoxSize
    //ExFor:FormField.Checked
    //ExFor:FormField.Default
    //ExFor:FormField.DropDownItems
    //ExFor:FormField.DropDownSelectedIndex
    //ExFor:FormField.Enabled
    //ExFor:FormField.EntryMacro
    //ExFor:FormField.ExitMacro
    //ExFor:FormField.HelpText
    //ExFor:FormField.IsCheckBoxExactSize
    //ExFor:FormField.MaxLength
    //ExFor:FormField.OwnHelp
    //ExFor:FormField.OwnStatus
    //ExFor:FormField.SetTextInputValue(Object)
    //ExFor:FormField.StatusText
    //ExFor:FormField.TextInputDefault
    //ExFor:FormField.TextInputFormat
    //ExFor:FormField.TextInputType
    //ExFor:FormFieldCollection
    //ExFor:FormFieldCollection.Clear
    //ExFor:FormFieldCollection.Count
    //ExFor:FormFieldCollection.GetEnumerator
    //ExFor:FormFieldCollection.Item(Int32)
    //ExFor:FormFieldCollection.Item(String)
    //ExFor:FormFieldCollection.Remove(String)
    //ExFor:FormFieldCollection.RemoveAt(Int32)
    //ExFor:Range.FormFields
    //ExSummary:Shows how insert different kinds of form fields into a document, and process them with using a document visitor implementation.
  test('Visitor', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Use a document builder to insert a combo box.
    builder.write("Choose a value from this combo box: ");
    let comboBox = builder.insertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
    comboBox.calculateOnExit = true;
    expect(comboBox.dropDownItems.count).toEqual(3);
    expect(comboBox.dropDownSelectedIndex).toEqual(0);
    expect(comboBox.enabled).toEqual(true);

    builder.insertBreak(aw.BreakType.ParagraphBreak);

    // Use a document builder to insert a check box.
    builder.write("Click this check box to tick/untick it: ");
    let checkBox = builder.insertCheckBox("MyCheckBox", false, 50);
    checkBox.isCheckBoxExactSize = true;
    checkBox.helpText = "Right click to check this box";
    checkBox.ownHelp = true;
    checkBox.statusText = "Checkbox status text";
    checkBox.ownStatus = true;
    expect(checkBox.checkBoxSize).toEqual(50.0);
    expect(checkBox.checked).toEqual(false);
    expect(checkBox.default).toEqual(false);

    builder.insertBreak(aw.BreakType.ParagraphBreak);

    // Use a document builder to insert text input form field.
    builder.write("Enter text here: ");
    let textInput = builder.insertTextInput("MyTextInput", aw.Fields.TextFormFieldType.Regular, "", "Placeholder text", 50);
    textInput.entryMacro = "EntryMacro";
    textInput.exitMacro = "ExitMacro";
    textInput.textInputDefault = "Regular";
    textInput.textInputFormat = "FIRST CAPITAL";
    textInput.setTextInputValue("New placeholder text");
    expect(textInput.textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
    expect(textInput.maxLength).toEqual(50);

    // This collection contains all our form fields.
    let formFields = doc.range.formFields;
    expect(formFields.count).toEqual(3);

    // Fields display our form fields. We can see their field codes by opening this document
    // in Microsoft and pressing Alt + F9. These fields have no switches,
    // and members of the FormField object fully govern their form fields' content.
    expect(doc.range.fields.count).toEqual(3);
    expect(doc.range.fields.at(0).getFieldCode()).toEqual(" FORMDROPDOWN \u0001");
    expect(doc.range.fields.at(1).getFieldCode()).toEqual(" FORMCHECKBOX \u0001");
    expect(doc.range.fields.at(2).getFieldCode()).toEqual(" FORMTEXT \u0001");

    // Allow each form field to accept a document visitor.
    let formFieldVisitor = new FormFieldVisitor();

    using (IEnumerator<FormField> fieldEnumerator = formFields.getEnumerator())
      while (fieldEnumerator.moveNext())
        fieldEnumerator.current.accept(formFieldVisitor);

    console.log(formFieldVisitor.getText());

    doc.updateFields();
    doc.save(base.artifactsDir + "FormFields.Visitor.html");
    TestFormField(doc); //ExSkip
  });


    /// <summary>
    /// Visitor implementation that prints details of form fields that it visits. 
    /// </summary>
  public class FormFieldVisitor : DocumentVisitor
  {
    public FormFieldVisitor()
    {
      mBuilder = new StringBuilder();
    }

      /// <summary>
      /// Called when a FormField node is encountered in the document.
      /// </summary>
    public override VisitorAction VisitFormField(FormField formField)
    {
      AppendLine(formField.type + ": \"" + formField.name + "\"");
      AppendLine("\tStatus: " + (formField.enabled ? "Enabled" : "Disabled"));
      AppendLine("\tHelp Text:  " + formField.helpText);
      AppendLine("\tEntry macro name: " + formField.entryMacro);
      AppendLine("\tExit macro name: " + formField.exitMacro);

      switch (formField.type)
      {
        case aw.Fields.FieldType.FieldFormDropDown:
          AppendLine("\tDrop-down items count: " + formField.dropDownItems.count + ", default selected item index: " + formField.dropDownSelectedIndex);
          AppendLine("\tDrop-down items: " + string.Join(", ", formField.dropDownItems.toArray()));
          break;
        case aw.Fields.FieldType.FieldFormCheckBox:
          AppendLine("\tCheckbox size: " + formField.checkBoxSize);
          AppendLine("\t" + "Checkbox is currently: " + (formField.checked ? "checked, " : "unchecked, ") + "by default: " + (formField.default ? "checked" : "unchecked"));
          break;
        case aw.Fields.FieldType.FieldFormTextInput:
          AppendLine("\tInput format: " + formField.textInputFormat);
          AppendLine("\tCurrent contents: " + formField.result);
          break;
      }

        // Let the visitor continue visiting other nodes.
      return aw.VisitorAction.Continue;
    }

      /// <summary>
      /// Adds newline char-terminated text to the current output.
      /// </summary>
    private void AppendLine(string text)
    {
      mBuilder.append(text + '\n');
    }

      /// <summary>
      /// Gets the plain text of the document that was accumulated by the visitor.
      /// </summary>
    public string GetText()
    {
      return mBuilder.toString();
    }

    private readonly StringBuilder mBuilder;
  }
    //ExEnd

  private void TestFormField(Document doc)
  {
    doc = DocumentHelper.saveOpen(doc);
    let fields = doc.range.fields;
    expect(fields.count).toEqual(3);

    TestUtil.VerifyField(aw.Fields.FieldType.FieldFormDropDown, " FORMDROPDOWN \u0001", '', doc.range.fields.at(0));
    TestUtil.VerifyField(aw.Fields.FieldType.FieldFormCheckBox, " FORMCHECKBOX \u0001", '', doc.range.fields.at(1));
    TestUtil.VerifyField(aw.Fields.FieldType.FieldFormTextInput, " FORMTEXT \u0001", "Regular", doc.range.fields.at(2));

    let formFields = doc.range.formFields;
    expect(formFields.count).toEqual(3);

    expect(formFields.at(0).type).toEqual(aw.Fields.FieldType.FieldFormDropDown);
    expect(new.at(] { "One").toEqual("Two", "Three" }, formFields[0).dropDownItems);
    expect(formFields.at(0).calculateOnExit).toEqual(true);
    expect(formFields.at(0).dropDownSelectedIndex).toEqual(0);
    expect(formFields.at(0).enabled).toEqual(true);
    expect(formFields.at(0).result).toEqual("One");

    expect(formFields.at(1).type).toEqual(aw.Fields.FieldType.FieldFormCheckBox);
    expect(formFields.at(1).isCheckBoxExactSize).toEqual(true);
    expect(formFields.at(1).helpText).toEqual("Right click to check this box");
    expect(formFields.at(1).ownHelp).toEqual(true);
    expect(formFields.at(1).statusText).toEqual("Checkbox status text");
    expect(formFields.at(1).ownStatus).toEqual(true);
    expect(formFields.at(1).checkBoxSize).toEqual(50.0);
    expect(formFields.at(1).checked).toEqual(false);
    expect(formFields.at(1).default).toEqual(false);
    expect(formFields.at(1).result).toEqual("0");

    expect(formFields.at(2).type).toEqual(aw.Fields.FieldType.FieldFormTextInput);
    expect(formFields.at(2).entryMacro).toEqual("EntryMacro");
    expect(formFields.at(2).exitMacro).toEqual("ExitMacro");
    expect(formFields.at(2).textInputDefault).toEqual("Regular");
    expect(formFields.at(2).textInputFormat).toEqual("FIRST CAPITAL");
    expect(formFields.at(2).textInputType).toEqual(aw.Fields.TextFormFieldType.Regular);
    expect(formFields.at(2).maxLength).toEqual(50);
    expect(formFields.at(2).result).toEqual("Regular");
  }*/

  test('DropDownItemCollection', () => {
    //ExStart
    //ExFor:DropDownItemCollection
    //ExFor:aw.Fields.DropDownItemCollection.add(String)
    //ExFor:aw.Fields.DropDownItemCollection.clear
    //ExFor:aw.Fields.DropDownItemCollection.contains(String)
    //ExFor:aw.Fields.DropDownItemCollection.count
    //ExFor:aw.Fields.DropDownItemCollection.getEnumerator
    //ExFor:aw.Fields.DropDownItemCollection.indexOf(String)
    //ExFor:aw.Fields.DropDownItemCollection.insert(Int32, String)
    //ExFor:aw.Fields.DropDownItemCollection.item(Int32)
    //ExFor:aw.Fields.DropDownItemCollection.remove(String)
    //ExFor:aw.Fields.DropDownItemCollection.removeAt(Int32)
    //ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert a combo box, and then verify its collection of drop-down items.
    // In Microsoft Word, the user will click the combo box,
    // and then choose one of the items of text in the collection to display.
    let items = ["One", "Two", "Three"];
    let comboBoxField = builder.insertComboBox("DropDown", items, 0);
    let dropDownItems = comboBoxField.dropDownItems;

    expect(dropDownItems.count).toEqual(3);
    expect(dropDownItems.at(0)).toEqual("One");
    expect(dropDownItems.indexOf("Two")).toEqual(1);
    expect(dropDownItems.contains("Three")).toEqual(true);

    // There are two ways of adding a new item to an existing collection of drop-down box items.
    // 1 -  Append an item to the end of the collection:
    dropDownItems.add("Four");

    // 2 -  Insert an item before another item at a specified index:
    dropDownItems.insert(3, "Three and a half");

    expect(dropDownItems.count).toEqual(5);

    // Iterate over the collection and print every element.
    for (let dropDownCollectionCurrent of dropDownItems)
    {
      console.log(dropDownCollectionCurrent);
    }

    // There are two ways of removing elements from a collection of drop-down items.
    // 1 -  Remove an item with contents equal to the passed string:
    dropDownItems.remove("Four");

    // 2 -  Remove an item at an index:
    dropDownItems.removeAt(3);

    expect(dropDownItems.count).toEqual(3);
    expect(dropDownItems.contains("Three and a half")).toEqual(false);
    expect(dropDownItems.contains("Four")).toEqual(false);

    doc.save(base.artifactsDir + "FormFields.DropDownItemCollection.html");

    // Empty the whole collection of drop-down items.
    dropDownItems.clear();
    //ExEnd

    doc = DocumentHelper.saveOpen(doc);
    dropDownItems = doc.range.formFields.at(0).dropDownItems;

    expect(dropDownItems.count).toEqual(0);

    doc = new aw.Document(base.artifactsDir + "FormFields.DropDownItemCollection.html");
    dropDownItems = doc.range.formFields.at(0).dropDownItems;

    expect(dropDownItems.count).toEqual(3);
    expect(dropDownItems.at(0)).toEqual("One");
    expect(dropDownItems.at(1)).toEqual("Two");
    expect(dropDownItems.at(2)).toEqual("Three");
  });

});
