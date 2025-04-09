// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExHyphenation", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test.skip('Dictionary: CultureInfo', () => {
    //ExStart
    //ExFor:Hyphenation.isDictionaryRegistered(String)
    //ExFor:Hyphenation.registerDictionary(String, String)
    //ExFor:Hyphenation.unregisterDictionary(String)
    //ExSummary:Shows how to register a hyphenation dictionary.
    // A hyphenation dictionary contains a list of strings that define hyphenation rules for the dictionary's language.
    // When a document contains lines of text in which a word could be split up and continued on the next line,
    // hyphenation will look through the dictionary's list of strings for that word's substrings.
    // If the dictionary contains a substring, then hyphenation will split the word across two lines
    // by the substring and add a hyphen to the first half.
    // Register a dictionary file from the local file system to the "de-CH" locale.
    aw.Hyphenation.registerDictionary("de-CH", base.myDir + "hyph_de_CH.dic");

    expect(aw.Hyphenation.isDictionaryRegistered("de-CH")).toEqual(true);

    // Open a document containing text with a locale matching that of our dictionary,
    // and save it to a fixed-page save format. The text in that document will be hyphenated.
    let doc = new aw.Document(base.myDir + "German text.docx");

    expect(doc.firstSection.body.firstParagraph.runs.toArray().map(node => node.asRun()).All(r => r.font.localeId == new CultureInfo("de-CH").LCID)).toEqual(true);

    doc.save(base.artifactsDir + "Hyphenation.Dictionary.registered.pdf");

    // Re-load the document after un-registering the dictionary,
    // and save it to another PDF, which will not have hyphenated text.
    aw.Hyphenation.unregisterDictionary("de-CH");

    expect(aw.Hyphenation.isDictionaryRegistered("de-CH")).toEqual(false);

    doc = new aw.Document(base.myDir + "German text.docx");
    doc.save(base.artifactsDir + "Hyphenation.Dictionary.Unregistered.pdf");
    //ExEnd
  });


  /*//Commented
    //ExStart
    //ExFor:Hyphenation
    //ExFor:Hyphenation.Callback
    //ExFor:Hyphenation.RegisterDictionary(String, Stream)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.WarningCallback
    //ExFor:IHyphenationCallback
    //ExFor:IHyphenationCallback.RequestDictionary(String)
    //ExSummary:Shows how to open and register a dictionary from a file.
  test('RegisterDictionary', () => {
    // Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
    let warningInfoCollection = new aw.WarningInfoCollection();
    aw.Hyphenation.warningCallback = warningInfoCollection;

    // Register an English (US) hyphenation dictionary by stream.
    Stream dictionaryStream = new FileStream(base.myDir + "hyph_en_US.dic", FileMode.open);
    aw.Hyphenation.registerDictionary("en-US", dictionaryStream);

    expect(warningInfoCollection.count).toEqual(0);

    // Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
    let doc = new aw.Document(base.myDir + "German text.docx");

    // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
    // This callback will handle the automatic request for that dictionary.
    aw.Hyphenation.callback = new CustomHyphenationDictionaryRegister();

    // When we save the document, German hyphenation will take effect.
    doc.save(base.artifactsDir + "Hyphenation.registerDictionary.pdf");

    // This dictionary contains two identical patterns, which will trigger a warning.
    expect(warningInfoCollection.count).toEqual(1);
    expect(warningInfoCollection.at(0).warningType).toEqual(aw.WarningType.MinorFormattingLoss);
    expect(warningInfoCollection.at(0).source).toEqual(aw.WarningSource.Layout);
    expect(warningInfoCollection.at(0).description).toEqual("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
                            "Content can be wrapped differently.");

    aw.Hyphenation.warningCallback = null; //ExSkip
    aw.Hyphenation.unregisterDictionary("en-US"); //ExSkip
    aw.Hyphenation.callback = null; //ExSkip
  });


    /// <summary>
    /// Associates ISO language codes with local system filenames for hyphenation dictionary files.
    /// </summary>
  private class CustomHyphenationDictionaryRegister : IHyphenationCallback
  {
    public CustomHyphenationDictionaryRegister()
    {
      mHyphenationDictionaryFiles = new Dictionary<string, string>
      {
        { "en-US", base.myDir + "hyph_en_US.dic" },
        { "de-CH", base.myDir + "hyph_de_CH.dic" }
      };
    }

    public void RequestDictionary(string language)
    {
      Console.write("Hyphenation dictionary requested: " + language);

      if (aw.Hyphenation.isDictionaryRegistered(language))
      {
        console.log(", is already registered.");
        return;
      }

      if (mHyphenationDictionaryFiles.containsKey(language))
      {
        aw.Hyphenation.registerDictionary(language, mHyphenationDictionaryFiles.at(language));
        console.log(", successfully registered.");
        return;
      }

      console.log(", no respective dictionary file known by this Callback.");
    }

    private readonly Dictionary<string, string> mHyphenationDictionaryFiles;
  }
    //ExEnd
#endif
//EndCommented*/

});
