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


describe("ExRange", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('Replace', () => {
    //ExStart
    //ExFor:Range.replace(String, String)
    //ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Greetings, _FullName_!");

    // Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
    let replacementCount = doc.range.replace("_FullName_", "John Doe");

    expect(replacementCount).toEqual(1);
    expect(doc.getText().trim()).toEqual("Greetings, John Doe!");
    //ExEnd
  });


  test.each([false,
    true])('ReplaceMatchCase', (matchCase) => {
    //ExStart
    //ExFor:Range.replace(String, String, FindReplaceOptions)
    //ExFor:FindReplaceOptions
    //ExFor:FindReplaceOptions.matchCase
    //ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Ruby bought a ruby necklace.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
    // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
    options.matchCase = matchCase;

    doc.range.replace("Ruby", "Jade", options);

    expect(doc.getText().trim()).toEqual(matchCase ? "Jade bought a ruby necklace." : "Jade bought a Jade necklace.");
    //ExEnd
  });


  test.each([false,
    true])('ReplaceFindWholeWordsOnly', (findWholeWordsOnly) => {
    //ExStart
    //ExFor:Range.replace(String, String, FindReplaceOptions)
    //ExFor:FindReplaceOptions
    //ExFor:FindReplaceOptions.findWholeWordsOnly
    //ExSummary:Shows how to toggle standalone word-only find-and-replace operations. 
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Jackson will meet you in Jacksonville.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
    // Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
    options.findWholeWordsOnly = findWholeWordsOnly;

    doc.range.replace("Jackson", "Louis", options);

    expect(doc.getText().trim()).toEqual(
      findWholeWordsOnly ? "Louis will meet you in Jacksonville." : "Louis will meet you in Louisville." );
    //ExEnd
  });


  test.each([true,
    false])('IgnoreDeleted', (ignoreTextInsideDeleteRevisions) => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreDeleted
    //ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
 
    builder.writeln("Hello world!");
    builder.writeln("Hello again!");
 
    // Start tracking revisions and remove the second paragraph, which will create a delete revision.
    // That paragraph will persist in the document until we accept the delete revision.
    doc.startTrackRevisions("John Doe", Date.now());
    doc.firstSection.body.paragraphs.at(1).remove();
    doc.stopTrackRevisions();

    expect(doc.firstSection.body.paragraphs.at(1).isDeleteRevision).toEqual(true);

    // We can use a "FindReplaceOptions" object to modify the find and replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
    // operation to ignore paragraphs that are delete revisions.
    // Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
    // operation to also search for text inside delete revisions.
    options.ignoreDeleted = ignoreTextInsideDeleteRevisions;

    doc.range.replace("Hello", "Greetings", options);

    expect(doc.getText().trim()).toEqual(
      ignoreTextInsideDeleteRevisions ? "Greetings world!\rHello again!" : "Greetings world!\rGreetings again!");
    //ExEnd
  });


  test.each([true,
    false])('IgnoreInserted', (ignoreTextInsideInsertRevisions) => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreInserted
    //ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");

    // Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
    doc.startTrackRevisions("John Doe", Date.now());
    builder.writeln("Hello again!");
    doc.stopTrackRevisions();

    expect(doc.firstSection.body.paragraphs.at(1).isInsertRevision).toEqual(true);

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "IgnoreInserted" flag to "true" to get the find-and-replace
    // operation to ignore paragraphs that are insert revisions.
    // Set the "IgnoreInserted" flag to "false" to get the find-and-replace
    // operation to also search for text inside insert revisions.
    options.ignoreInserted = ignoreTextInsideInsertRevisions;

    doc.range.replace("Hello", "Greetings", options);

    expect(doc.getText().trim()).toEqual(
      ignoreTextInsideInsertRevisions
        ? "Greetings world!\rHello again!"
        : "Greetings world!\rGreetings again!");
    //ExEnd
  });


  test.each([true,
    false])('IgnoreFields', (ignoreTextInsideFields) => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreFields
    //ExSummary:Shows how to ignore text inside fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Hello world!");
    builder.insertField("QUOTE", "Hello again!");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "IgnoreFields" flag to "true" to get the find-and-replace
    // operation to ignore text inside fields.
    // Set the "IgnoreFields" flag to "false" to get the find-and-replace
    // operation to also search for text inside fields.
    options.ignoreFields = ignoreTextInsideFields;

    doc.range.replace("Hello", "Greetings", options);

    expect(doc.getText().trim()).toEqual(
      ignoreTextInsideFields
        ? "Greetings world!\r\u0013QUOTE\u0014Hello again!\u0015"
        : "Greetings world!\r\u0013QUOTE\u0014Greetings again!\u0015");
    //ExEnd
  });


  test.skip.each([true,
    false])('IgnoreFieldCodes - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', (ignoreFieldCodes) => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreFieldCodes
    //ExSummary:Shows how to ignore text inside field codes.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField("INCLUDETEXT", "Test IT!");

    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreFieldCodes = ignoreFieldCodes;;

    // Replace 'T' in document ignoring text inside field code or not.
    doc.range.replace(/*new Regex*/("T"), "*", options);
    console.log(doc.getText());

    Assert.AreEqual(
      ignoreFieldCodes
        ? "\u0013INCLUDETEXT\u0014*est I*!\u0015"
        : "\u0013INCLUDE*EX*\u0014*est I*!\u0015", doc.getText().trim());
    //ExEnd
  });


  test.each([true,
    false])('IgnoreFootnote', (isIgnoreFootnotes) => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreFootnotes
    //ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

    builder.insertParagraph();

    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder.insertFootnote(aw.Notes.FootnoteType.Endnote, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

    // Set the "IgnoreFootnotes" flag to "true" to get the find-and-replace
    // operation to ignore text inside footnotes.
    // Set the "IgnoreFootnotes" flag to "false" to get the find-and-replace
    // operation to also search for text inside footnotes.
    let options = new aw.Replacing.FindReplaceOptions();
    options.ignoreFootnotes = isIgnoreFootnotes;
    doc.range.replace("Lorem ipsum", "Replaced Lorem ipsum", options);
    //ExEnd

    let paragraphs = doc.firstSection.body.paragraphs.toArray();

    for (let para of paragraphs)
    {
      expect(para.runs.at(0).text).toEqual("Replaced Lorem ipsum");
    }

    let footnotes = doc.getChildNodes(aw.NodeType.Footnote, true);
    expect(footnotes.at(0).asFootnote().toString(aw.SaveFormat.Text).trim()).toEqual(
      isIgnoreFootnotes
        ? "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
        : "Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    expect(footnotes.at(1).asFootnote().toString(aw.SaveFormat.Text).trim()).toEqual(
      isIgnoreFootnotes
        ? "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
        : "Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
  });


  test('IgnoreShapes', () => {
    //ExStart
    //ExFor:FindReplaceOptions.ignoreShapes
    //ExSummary:Shows how to ignore shapes while replacing text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    builder.insertShape(aw.Drawing.ShapeType.Balloon, 200, 200);
    builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

    let findReplaceOptions = new aw.Replacing.FindReplaceOptions();
    findReplaceOptions.ignoreShapes = true;
    builder.document.range.replace("Lorem ipsum dolor sit amet, consectetur adipiscing elit.Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit.", findReplaceOptions);
    expect(builder.document.getText().trim()).toEqual("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
    //ExEnd
  });


  test('UpdateFieldsInRange', () => {
    //ExStart
    //ExFor:Range.updateFields
    //ExSummary:Shows how to update all the fields in a range.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.insertField(" DOCPROPERTY Category");
    builder.insertBreak(aw.BreakType.SectionBreakEvenPage);
    builder.insertField(" DOCPROPERTY Category");

    // The above DOCPROPERTY fields will display the value of this built-in document property.
    doc.builtInDocumentProperties.category = "MyCategory";

    // If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
    expect(doc.range.fields.at(0).result).toEqual('');
    expect(doc.range.fields.at(1).result).toEqual('');

    // Update all the fields that are in the range of the first section.
    doc.firstSection.range.updateFields();

    expect(doc.range.fields.at(0).result).toEqual("MyCategory");
    expect(doc.range.fields.at(1).result).toEqual('');
    //ExEnd
  });


  test('ReplaceWithString', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("This one is sad.");
    builder.writeln("That one is mad.");

    let options = new aw.Replacing.FindReplaceOptions();
    options.matchCase = false;
    options.findWholeWordsOnly = true;

    doc.range.replace("sad", "bad", options);

    doc.save(base.artifactsDir + "Range.ReplaceWithString.docx");
  });


  test.skip('ReplaceWithRegex - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', () => {
    //ExStart
    //ExFor:Range.replace(Regex, String)
    //ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("I decided to get the curtains in gray, ideal for the grey-accented room.");

    doc.range.replace(new Regex("gr(a|e)y"), "lavender");

    expect(doc.getText().trim()).toEqual("I decided to get the curtains in lavender, ideal for the lavender-accented room.");
    //ExEnd
  });


  //ExStart
  //ExFor:FindReplaceOptions.ReplacingCallback
  //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
  //ExFor:ReplacingArgs.Replacement
  //ExFor:IReplacingCallback
  //ExFor:IReplacingCallback.Replacing
  //ExFor:ReplacingArgs
  //ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
  test.skip('ReplaceWithCallback - TODO: WORDSNODEJS-107 - Add support of IReplacingCallback', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Our new location in New York City is opening tomorrow. " +
            "Hope to see all our NYC-based customers at the opening!");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set a callback that tracks any replacements that the "Replace" method will make.
    let logger = new TextFindAndReplacementLogger();
    options.replacingCallback = logger;

    doc.range.replace(new Regex("New York City|NYC"), "Washington", options);
            
    expect(doc.getText().trim()).toEqual("Our new location in (Old value:\"New York City\") Washington is opening tomorrow. " +
                            "Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!");

    expect(logger.GetLog().trim()).toEqual("\"New York City\" converted to \"Washington\" 20 characters into a Run node.\r\n" +
                            "\"NYC\" converted to \"Washington\" 42 characters into a Run node.");
  });

/*
    /// <summary>
    /// Maintains a log of every text replacement done by a find-and-replace operation
    /// and notes the original matched text's value.
    /// </summary>
  private class TextFindAndReplacementLogger : IReplacingCallback
  {
    ReplaceAction aw.Replacing.IReplacingCallback.replacing(ReplacingArgs args)
    {
      mLog.AppendLine(`\"${args.match.value}\" converted to \"${args.replacement}\" ` +
              `${args.matchOffset} characters into a ${args.matchNode.nodeType} node.`);

      args.replacement = `(Old value:\"${args.match.value}\") ${args.replacement}`;
      return aw.Replacing.ReplaceAction.Replace;
    }

    public string GetLog()
    {
      return mLog.toString();
    }

    private readonly StringBuilder mLog = new StringBuilder();
  }
    //ExEnd
*/    

  //ExStart
  //ExFor:FindReplaceOptions.ApplyFont
  //ExFor:FindReplaceOptions.ReplacingCallback
  //ExFor:ReplacingArgs.GroupIndex
  //ExFor:ReplacingArgs.GroupName
  //ExFor:ReplacingArgs.Match
  //ExFor:ReplacingArgs.MatchOffset
  //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
  test.skip('ConvertNumbersToHexadecimal - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.font.name = "Arial";
    builder.writeln("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n" +
            "123, 456, 789 and 17379.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
    options.applyFont.highlightColor = "#D3D3D3";

    let numberHexer = new NumberHexer();
    options.replacingCallback = numberHexer;

    let replacementCount = doc.range.replace(new Regex("[0-9]+"), "", options);

    console.log(numberHexer.GetLog());

    expect(replacementCount).toEqual(4);
    expect(doc.getText().trim()).toEqual("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r" +
                            "0x7B, 0x1C8, 0x315 and 0x43E3.");
    expect(doc.getChildNodes(aw.NodeType.Run, true).filter(r => r.asRun().font.highlightColor == "#D3D3D3")).toEqual(4);
  });

/*
    /// <summary>
    /// Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
    /// Maintains a log of every replacement.
    /// </summary>
  private class NumberHexer : IReplacingCallback
  {
    public ReplaceAction Replacing(ReplacingArgs args)
    {
      mCurrentReplacementNumber++;

      int number = Convert.ToInt32(args.match.value);

      args.replacement = `0x${number:X}`;

      mLog.AppendLine(`Match #${mCurrentReplacementNumber}`);
      mLog.AppendLine(`\tOriginal value:\t${args.match.value}`);
      mLog.AppendLine(`\tReplacement:\t${args.replacement}`);
      mLog.AppendLine(`\tOffset in parent ${args.matchNode.nodeType} node:\t${args.matchOffset}`);

      mLog.AppendLine(string.IsNullOrEmpty(args.groupName)
        ? `\tGroup index:\t${args.groupIndex}`
        : `\tGroup name:\t${args.groupName}`);

      return aw.Replacing.ReplaceAction.Replace;
    }

    public string GetLog()
    {
      return mLog.toString();
    }

    private int mCurrentReplacementNumber;
    private readonly StringBuilder mLog = new StringBuilder();
  }
    //ExEnd
*/    

  test('ApplyParagraphFormat', () => {
    //ExStart
    //ExFor:FindReplaceOptions.applyParagraphFormat
    //ExFor:Range.replace(String, String)
    //ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
    builder.writeln("This one will not!");
    builder.write("This one also will.");

    let paragraphs = doc.firstSection.body.paragraphs.toArray();

    expect(paragraphs.at(0).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Left);
    expect(paragraphs.at(1).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Left);
    expect(paragraphs.at(2).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Left);

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
    // that contains a match that the find-and-replace operation finds.
    options.applyParagraphFormat.alignment = aw.ParagraphAlignment.Right;

    // Replace every full stop that is right before a paragraph break with an exclamation point.
    var count = doc.range.replace(".&p", "!&p", options);

    expect(count).toEqual(2);
    paragraphs = doc.firstSection.body.paragraphs.toArray();
    expect(paragraphs.at(0).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Right);
    expect(paragraphs.at(1).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Left);
    expect(paragraphs.at(2).paragraphFormat.alignment).toEqual(aw.ParagraphAlignment.Right);
    expect(doc.getText().trim()).toEqual("Every paragraph that ends with a full stop like this one will be right aligned!\r" +
                            "This one will not!\r" +
                            "This one also will!");
    //ExEnd
  });


  test('DeleteSelection', () => {
    //ExStart
    //ExFor:Node.range
    //ExFor:Range.delete
    //ExSummary:Shows how to delete all the nodes from a range.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Add text to the first section in the document, and then add another section.
    builder.write("Section 1. ");
    builder.insertBreak(aw.BreakType.SectionBreakContinuous);
    builder.write("Section 2.");

    expect(doc.getText().trim()).toEqual("Section 1. \fSection 2.");

    // Remove the first section entirely by removing all the nodes
    // within its range, including the section itself.
    doc.sections.at(0).range.delete();

    expect(doc.sections.count).toEqual(1);
    expect(doc.getText().trim()).toEqual("Section 2.");
    //ExEnd
  });


  test('RangesGetText', () => {
    //ExStart
    //ExFor:Range
    //ExFor:Range.text
    //ExSummary:Shows how to get the text contents of all the nodes that a range covers.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.write("Hello world!");

    expect(doc.range.text.trim()).toEqual("Hello world!");
    //ExEnd
  });


  //ExStart
  //ExFor:FindReplaceOptions.UseLegacyOrder
  //ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
  test.skip.each([true,
    false])('UseLegacyOrder - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', (useLegacyOrder) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert three runs which we can search for using a regex pattern.
    // Place one of those runs inside a text box.
    builder.writeln("[tag 1]");
    let textBox = builder.insertShape(aw.Drawing.ShapeType.TextBox, 100, 50);
    builder.writeln("[tag 2]");
    builder.moveTo(textBox.firstParagraph);
    builder.write("[tag 3]");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Assign a custom callback to the "ReplacingCallback" property.
    let callback = new TextReplacementTracker();
    options.replacingCallback = callback;

    // If we set the "UseLegacyOrder" property to "true", the
    // find-and-replace operation will go through all the runs outside of a text box
    // before going through the ones inside a text box.
    // If we set the "UseLegacyOrder" property to "false", the
    // find-and-replace operation will go over all the runs in a range in sequential order.
    options.useLegacyOrder = useLegacyOrder;

    /* doc.range.replace(new Regex(@"\[tag \d*\]"), "", options); TODO

    Assert.AreEqual(useLegacyOrder ?
      new aw.Lists.List<string> { "[tag 1]", "[tag 3]", "[tag 2]" } :
      new aw.Lists.List<string> { "[tag 1]", "[tag 2]", "[tag 3]" }, callback.Matches); */
  });

/*
    /// <summary>
    /// Records the order of all matches that occur during a find-and-replace operation.
    /// </summary>
  private class TextReplacementTracker : IReplacingCallback
  {
    ReplaceAction aw.Replacing.IReplacingCallback.replacing(ReplacingArgs e)
    {
      Matches.add(e.match.value);
      return aw.Replacing.ReplaceAction.Replace;
    }

    public List<string> Matches { get; } = new aw.Lists.List<string>();
  }
    //ExEnd
 */

  test.skip.each([false,
    true])('UseSubstitutions - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', (useSubstitutions) => {
    //ExStart
    //ExFor:FindReplaceOptions.useSubstitutions
    //ExSummary:Shows how to replace the text with substitutions.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    builder.writeln("John sold a car to Paul.");
    builder.writeln("Jane sold a house to Joe.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Set the "UseSubstitutions" property to "true" to get
    // the find-and-replace operation to recognize substitution elements.
    // Set the "UseSubstitutions" property to "false" to ignore substitution elements.
    options.useSubstitutions = useSubstitutions;

    /* TODO
    let regex = new Regex(@"([A-z]+) sold a ([A-z]+) to ([A-z]+)");
    doc.range.replace(regex, @"$3 bought a $2 from $1", options);

    Assert.AreEqual(
      useSubstitutions
        ? "Paul bought a car from John.\rJoe bought a house from Jane."
        : "$3 bought a $2 from $1.\r$3 bought a $2 from $1.", doc.getText().trim()); */
    //ExEnd
  });


  //ExStart
  //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
  //ExFor:IReplacingCallback
  //ExFor:ReplaceAction
  //ExFor:IReplacingCallback.Replacing
  //ExFor:ReplacingArgs
  //ExFor:ReplacingArgs.MatchNode
  //ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
  test.skip('InsertDocumentAtReplace - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', () => {
    let mainDoc = new aw.Document(base.myDir + "Document insertion destination.docx");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();
    options.replacingCallback = new InsertDocumentAtReplaceHandler();

    mainDoc.range.replace(new Regex("\\[MY_DOCUMENT\\]"), "", options);
    mainDoc.save(base.artifactsDir + "InsertDocument.InsertDocumentAtReplace.docx");

    TestInsertDocumentAtReplace(new aw.Document(base.artifactsDir + "InsertDocument.InsertDocumentAtReplace.docx")); //ExSkip
  });

/*
  private class InsertDocumentAtReplaceHandler : IReplacingCallback
  {
    ReplaceAction aw.Replacing.IReplacingCallback.replacing(ReplacingArgs args)
    {
      let subDoc = new aw.Document(base.myDir + "Document.docx");

        // Insert a document after the paragraph containing the matched text.
      let para = (Paragraph)args.matchNode.parentNode;
      InsertDocument(para, subDoc);

        // Remove the paragraph with the matched text.
      para.remove();

      return aw.Replacing.ReplaceAction.Skip;
    }
  }

    /// <summary>
    /// Inserts all the nodes of another document after a paragraph or table.
    /// </summary>
  private static void InsertDocument(Node insertionDestination, Document docToInsert)
  {
    if (insertionDestination.nodeType == aw.NodeType.Paragraph || insertionDestination.nodeType == aw.NodeType.Table)
    {
      let dstStory = insertionDestination.parentNode;

      let importer =
        new aw.NodeImporter(docToInsert, insertionDestination.document, aw.ImportFormatMode.KeepSourceFormatting);

      foreach (Section srcSection in docToInsert.sections.OfType<Section>())
        for (let srcNode of srcSection.body)
        {
            // Skip the node if it is the last empty paragraph in a section.
          if (srcNode.nodeType == aw.NodeType.Paragraph)
          {
            let para = (Paragraph)srcNode;
            if (para.isEndOfSection && !para.hasChildNodes)
              continue;
          }

          let newNode = importer.importNode(srcNode, true);

          dstStory.insertAfter(newNode, insertionDestination);
          insertionDestination = newNode;
        }
    }
    else
    {
      throw new ArgumentException("The destination node must be either a paragraph or table.");
    }
  }
    //ExEnd

  private static void TestInsertDocumentAtReplace(Document doc)
  {
    expect(doc.firstSection.body.getText().trim()).toEqual("1) At text that can be identified by regex:\rHello World!\r" +
                            "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                            "3) At a bookmark:");
  }
*/
  //ExStart
  //ExFor:FindReplaceOptions.Direction
  //ExFor:FindReplaceDirection
  //ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
  test.skip.each([aw.Replacing.FindReplaceDirection.Backward,
    aw.Replacing.FindReplaceDirection.Forward])('Direction - TODO: WORDSNODEJS-106 - Add support of regex to doc.range.replace', (findReplaceDirection) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // Insert three runs which we can search for using a regex pattern.
    // Place one of those runs inside a text box.
    builder.writeln("Match 1.");
    builder.writeln("Match 2.");
    builder.writeln("Match 3.");
    builder.writeln("Match 4.");

    // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    let options = new aw.Replacing.FindReplaceOptions();

    // Assign a custom callback to the "ReplacingCallback" property.
    let callback = new TextReplacementRecorder();
    options.replacingCallback = callback;

    // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
    // operation to start from the end of the range, and traverse back to the beginning.
    // Set the "Direction" property to "FindReplaceDirection.Forward" to get the find-and-replace
    // operation to start from the beginning of the range, and traverse to the end.
    options.direction = findReplaceDirection;

    // TODO doc.range.replace(new Regex(@"Match \d*"), "Replacement", options);

    expect(doc.getText().trim()).toEqual("Replacement.\r" +
                            "Replacement.\r" +
                            "Replacement.\r" +
                            "Replacement.");

    switch (findReplaceDirection)
    {
      case aw.Replacing.FindReplaceDirection.Forward:
        expect(callback.Matches).toEqual(["Match 1", "Match 2", "Match 3", "Match 4"]);
        break;
      case aw.Replacing.FindReplaceDirection.Backward:
        expect(callback.Matches).toEqual(["Match 4", "Match 3", "Match 2", "Match 1"]);
        break;
    }
  });

/*
    /// <summary>
    /// Records all matches that occur during a find-and-replace operation in the order that they take place.
    /// </summary>
  private class TextReplacementRecorder : IReplacingCallback
  {
    ReplaceAction aw.Replacing.IReplacingCallback.replacing(ReplacingArgs e)
    {
      Matches.add(e.match.value);
      return aw.Replacing.ReplaceAction.Replace;
    }

    public List<string> Matches { get; } = new aw.Lists.List<string>();
  }
    //ExEnd
*/
});
