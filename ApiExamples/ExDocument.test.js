// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const MemoryStream = require('memorystream');
const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const DocumentHelper = require('./DocumentHelper');
const TestUtil = require('./TestUtil');

const fs = require('fs');
const path = require('path');


function testDocPackageCustomParts(parts)
{
  expect(parts.count).toEqual(3);

  expect(parts.at(0).name).toEqual("/payload/payload_on_package.test");
  expect(parts.at(0).contentType).toEqual("mytest/somedata");
  expect(parts.at(0).relationshipType).toEqual("http://mytest.payload.internal");
  expect(parts.at(0).isExternal).toEqual(false);
  expect(parts.at(0).data.length).toEqual(18);

  expect(parts.at(1).name).toEqual("http://www.aspose.com/Images/aspose-logo.jpg");
  expect(parts.at(1).contentType).toEqual("");
  expect(parts.at(1).relationshipType).toEqual("http://mytest.payload.external");
  expect(parts.at(1).isExternal).toEqual(true);
  expect(parts.at(1).data.length).toEqual(0);

  expect(parts.at(2).name).toEqual("http://www.aspose.com/Images/aspose-logo.jpg");
  expect(parts.at(2).contentType).toEqual("");
  expect(parts.at(2).relationshipType).toEqual("http://mytest.payload.external");
  expect(parts.at(2).isExternal).toEqual(true);
  expect(parts.at(2).data.length).toEqual(0);
}

describe('ExDocument', () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CreateSimpleDocument', () => {
    //ExStart:CreateSimpleDocument
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:DocumentBase.document
    //ExFor:Document.#ctor()
    //ExSummary:Shows how to create simple document.
    const doc = new aw.Document();
    // New Document objects by default come with the minimal set of nodes
    // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
    const section = new aw.Section(doc);
    doc.appendChild(section);
    const body = new aw.Body(doc);
    section.appendChild(body);
    const para = new aw.Paragraph(doc);
    body.appendChild(para);
    para.appendChild(new aw.Run(doc, "Hello world!"));
    //ExEnd:CreateSimpleDocument
    expect(doc.getText()).toEqual("\fHello world!\f");
  });

  test('Constructor', () => {
    //ExStart
    //ExFor:Document.#ctor()
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExSummary:Shows how to create and load documents.
    // There are two ways of creating a Document object using Aspose.Words.
    // 1 -  Create a blank document:
    var doc = new aw.Document();
    // New Document objects by default come with the minimal set of nodes
    // required to begin adding content such as text and shapes: a Section, a Body, and a Paragraph.
    doc.firstSection.body.firstParagraph.appendChild(new aw.Run(doc, "Hello world!"));
    // 2 -  Load a document that exists in the local file system:
    doc = new aw.Document(base.myDir + "Document.docx");
    // Loaded documents will have contents that we can access and edit.
    expect(doc.firstSection.body.firstParagraph.getText().trim()).toEqual("Hello World!");
    // Some operations that need to occur during loading, such as using a password to decrypt a document,
    // can be done by passing a LoadOptions object when loading the document.
    doc = new aw.Document(base.myDir + "Encrypted.docx", new aw.Loading.LoadOptions("docPassword"));
    expect(doc.firstSection.body.firstParagraph.getText().trim()).toEqual("Test encrypted document.");
    //ExEnd
  });

  test('LoadFromStream', () => {
    //ExStart
    //ExFor:Document.#ctor(Stream)
    //ExSummary:Shows how to load a document using a stream.
    const buffer = base.loadFileToBuffer(base.myDir + "Document.docx");
    const doc = new aw.Document(buffer);
    expect(doc.getText().trim()).toEqual("Hello World!\r\rHello Word!\r\r\rHello World!");
    //ExEnd
  });


  test('LoadFromWeb', async () => {
    //ExStart
    //ExFor:Document.#ctor(Stream)
    //ExSummary:Shows how to load a document from a URL.
    // Create a URL that points to a Microsoft Word document.
    const url = "https://filesamples.com/samples/document/docx/sample3.docx";

    const response = await fetch(url);
    const blob = await response.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const dataBytes = Buffer.from(arrayBuffer);    

    let doc = new aw.Document(dataBytes);

    // At this stage, we can read and edit the document's contents and then save it to the local file system.
    expect(doc.firstSection.body.paragraphs.at(3).getText().trim()).toEqual("There are eight section headings in this document. At the beginning, \"Sample Document\" is a level 1 heading. " +
                                    "The main section headings, such as \"Headings\" and \"Lists\" are level 2 headings. " +
                                    "The Tables section contains two sub-headings, \"Simple Table\" and \"Complex Table,\" which are both level 3 headings.");

    doc.save(base.artifactsDir + "Document.LoadFromWeb.docx");
    //ExEnd
  });


  test('ConvertToPdf', () => {
    //ExStart
    //ExFor:Document.#ctor(String)
    //ExFor:Document.save(String)
    //ExSummary:Shows how to open a document and convert it to .PDF.
    const doc = new aw.Document(base.myDir + "Document.docx");
    doc.save(base.artifactsDir + "Document.ConvertToPdf.pdf");
    //ExEnd
  });
  
  test('SaveToImageStream', async () => {
    //ExStart
    //ExFor:Document.save(Stream, SaveFormat)
    //ExSummary:Shows how to save a document to an image via stream, and then read the image from that stream.
    const doc = new aw.Document();
    const builder = new aw.DocumentBuilder(doc);
    builder.font.name = "Times New Roman";
    builder.font.size = 24;
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder.insertImage(base.imageDir + "Logo.jpg");

    let bmpFile = base.artifactsDir + "saveToImageStream.bmp";
    let writeStream = fs.createWriteStream(bmpFile);
    doc.save(writeStream, aw.SaveFormat.Bmp);
    await new Promise(resolve => writeStream.on("finish", resolve));

    // Read the stream back into an image.
    const bmp = require("bmp-js");
    let bmpData = fs.readFileSync(bmpFile);
    const image = bmp.decode(bmpData);
    expect(image.width).toEqual(816);
    expect(image.height).toEqual(1056);
    //ExEnd
  });

  test('DetectMobiDocumentFormat', () => {
    const info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Document.mobi");
    expect(info.LoadFormat).toEqual(aw.LoadFormat.mobi);
  });

  test('OpenFromStreamWithBaseUri', () => {
    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:LoadOptions.#ctor
    //ExFor:LoadOptions.baseUri
    //ExFor:ShapeBase.isImage
    //ExSummary:Shows how to open an HTML document with images from a stream using a base URI.
    const buffer = base.loadFileToBuffer(base.myDir + "Document.html");
    // Pass the URI of the base folder while loading it
    // so that any images with relative URIs in the HTML document can be found.
    const loadOptions = new aw.Loading.LoadOptions();
    loadOptions.baseUri = base.imageDir;
    const doc = new aw.Document(buffer, loadOptions);
    // Verify that the first shape of the document contains a valid image.
    const shape = doc.getShape(0, true);
    expect(shape.isImage).toEqual(true);
    expect(shape.imageData.imageBytes).not.toEqual(null);
    expect(aw.ConvertUtil.pointToPixel(shape.width)).toBeCloseTo(32, 2);
    expect(aw.ConvertUtil.pointToPixel(shape.height)).toBeCloseTo(32, 2);
    //ExEnd
  });

  test('InsertHtmlFromWebPage', async () => {
    //ExStart
    //ExFor:Document.#ctor(Stream, LoadOptions)
    //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
    //ExFor:LoadFormat
    //ExSummary:Shows how save a web page as a .docx file.
    const url = "https://products.aspose.com/words/";
    const response = await fetch(url);
    const blob = await response.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const dataBytes = Buffer.from(arrayBuffer);    

    let doc = new aw.Document(dataBytes);

    // At this stage, we can read and edit the document's contents and then save it to the local file system.
    expect(doc.getText().includes("HYPERLINK \"https://products.aspose.com/words/net/\" \\o \"Aspose.Words\"")).toBe(true) ; //ExSkip

    doc.save(base.artifactsDir + "Document.LoadFromWeb.docx");    
    //ExEnd
  });

  test('LoadEncrypted', () => {
    //ExStart
    //ExFor:Document.#ctor(Stream,LoadOptions)
    //ExFor:Document.#ctor(String,LoadOptions)
    //ExFor:LoadOptions
    //ExFor:LoadOptions.#ctor(String)
    //ExSummary:Shows how to load an encrypted Microsoft Word document.

    // Aspose.Words throw an exception if we try to open an encrypted document without its password.
    expect(() => new aw.Document(base.myDir + "Encrypted.docx")).toThrow("The document password is incorrect.");

    // When loading such a document, the password is passed to the document's constructor using a LoadOptions object.
    const options = new aw.Loading.LoadOptions("docPassword");

    // There are two ways of loading an encrypted document with a LoadOptions object.
    // 1 -  Load the document from the local file system by filename:
    const doc = new aw.Document(base.myDir + "Encrypted.docx", options);
    expect(doc.getText().trim()).toEqual("Test encrypted document."); //ExSkip

    // 2 -  Load the document from a stream:
    const stream = base.loadFileToBuffer(base.myDir + "Encrypted.docx");
    const doc2 = new aw.Document(stream, options);
    expect(doc2.getText().trim()).toEqual("Test encrypted document."); //ExSkip
    //ExEnd
  });

  test.skip('NotSupportedWarning - TODO: Inherence from IWarningCallback not supported.', () => {
    //ExStart
    //ExFor:WarningInfoCollection.count
    //ExFor:WarningInfoCollection.item(Int32)
    //ExSummary:Shows how to get warnings about unsupported formats.
    let warnings = new aw.WarningInfoCollection();
    let lo = new aw.Loading.LoadOptions();
    lo.warningCallback = warnings;
    let doc = new aw.Document(base.myDir + "FB2 document.fb2", lo);

    expect(warnings.at(0).description).toEqual("The original file load format is FB2, which is not supported by Aspose.words. The file is loaded as an XML document.");
    expect(warnings.count).toEqual(1);
    //ExEnd
  });

  test('TempFolder', () => {
    //ExStart
    //ExFor:LoadOptions.tempFolder
    //ExSummary:Shows how to load a document using temporary files.
    // Note that such an approach can reduce memory usage but degrades speed
    const loadOptions = new aw.Loading.LoadOptions();
    loadOptions.tempFolder = "C:\\Temp\\";

    // Ensure that the directory exists and load
    if (!fs.existsSync(loadOptions.tempFolder)){
      fs.mkdirSync(loadOptions.tempFolder);
    }    

    const doc = new aw.Document(base.myDir + "Document.docx", loadOptions);
    //ExEnd
  });

  test('ConvertToHtml', () => {
    //ExStart
    //ExFor:Document.save(String,SaveFormat)
    //ExFor:SaveFormat
    //ExSummary:Shows how to convert from DOCX to HTML format.
    const doc = new aw.Document(base.myDir + "Document.docx");

    doc.save(base.artifactsDir + "Document.ConvertToHtml.html", aw.SaveFormat.Html);
    //ExEnd
  });

  test('ConvertToMhtml', () => {
    const doc = new aw.Document(base.myDir + "Document.docx");
    doc.save(base.artifactsDir + "Document.ConvertToMhtml.mht");
  });

  test('ConvertToTxt', () => {
    const doc = new aw.Document(base.myDir + "Document.docx");
    doc.save(base.artifactsDir + "Document.ConvertToTxt.txt");
  });

  test('ConvertToEpub', () => {
    const doc = new aw.Document(base.myDir + "Rendering.docx");
    doc.save(base.artifactsDir + "Document.ConvertToEpub.epub");
  });

  test('SaveToStream', async () => {
    //ExStart
    //ExFor:Document.save(Stream,SaveFormat)
    //ExSummary:Shows how to save a document to a stream.
    const doc = new aw.Document(base.myDir + "Document.docx");

    const dstFile = base.artifactsDir + "saveToStream.docx";
    const dstStream = fs.createWriteStream(dstFile);
    doc.save(dstStream, aw.SaveFormat.Docx);
    await new Promise(resolve => dstStream.on("finish", resolve));
    const dstDoc = new aw.Document(dstFile);
    expect(dstDoc.getText().trim()).toEqual("Hello World!\r\rHello Word!\r\r\rHello World!");
    //ExEnd
  });

  test('AppendDocument', () => {
    //ExStart
    //ExFor:Document.AppendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to append a document to the end of another document.
    const srcDoc = new aw.Document();
    srcDoc.firstSection.body.appendParagraph("Source document text. ");

    const dstDoc = new aw.Document();
    dstDoc.firstSection.body.appendParagraph("Destination document text. ");

    // Append the source document to the destination document while preserving its formatting,
    // then save the source document to the local file system.
    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting);
    expect(dstDoc.sections.count).toEqual(2); //ExSkip

    dstDoc.save(base.artifactsDir + "Document.AppendDocument.docx");
    //ExEnd

    const outDocText = new aw.Document(base.artifactsDir + "Document.AppendDocument.docx").getText();

    expect(outDocText.startsWith(dstDoc.getText())).toEqual(true);
    expect(outDocText.endsWith(srcDoc.getText())).toEqual(true);;
  });

  // The file path used below does not point to an existing file.
  test('AppendDocumentFromAutomation', () => {
    const doc = new aw.Document();

    // We should call this method to clear this document of any existing content.
    doc.removeAllChildren();

    const recordCount = 5;
    for (var i = 1; i <= recordCount; i++)
    {
        const srcDoc = new aw.Document();

        expect(() => { new aw.Document("C:\\DetailsList.doc"); }).toThrow("Could not find file 'C:\\DetailsList.doc'.");

        // Append the source document at the end of the destination document.
        doc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles);

        // Automation required you to insert a new section break at this point, however, in Aspose.Words we
        // do not need to do anything here as the appended document is imported as separate sections already

        // Unlink all headers/footers in this section from the previous section headers/footers
        // if this is the second document or above being appended.
        if (i > 1) {
          expect(() => { doc.Sections[i].HeadersFooters.LinkToPrevious(false); })
            .toThrow(`Cannot read properties of undefined (reading '${i}')`);
        }
    }
  });

  test.each([true, false])('TestImportList', (isKeepSourceNumbering) => {
    //ExStart
    //ExFor:ImportFormatOptions.keepSourceNumbering
    //ExSummary:Shows how to import a document with numbered lists.
    const srcDoc = new aw.Document(base.myDir + "List source.docx");
    const dstDoc = new aw.Document(base.myDir + "List destination.docx");

    expect(dstDoc.lists.count).toEqual(4);

    const options = new aw.ImportFormatOptions();

    // If there is a clash of list styles, apply the list format of the source document.
    // Set the "KeepSourceNumbering" property to "false" to not import any list numbers into the destination document.
    // Set the "KeepSourceNumbering" property to "true" import all clashing
    // list style numbering with the same appearance that it had in the source document.
    options.keepSourceNumbering = isKeepSourceNumbering;

    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);
    dstDoc.updateListLabels();

    expect(dstDoc.lists.count).toEqual(isKeepSourceNumbering ? 5 : 4);
    //ExEnd
  });

  test('KeepSourceNumberingSameListIds', () => {
    //ExStart
    //ExFor:ImportFormatOptions.keepSourceNumbering
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how resolve a clash when importing documents that have lists with the same list definition identifier.
    const srcDoc = new aw.Document(base.myDir + "List with the same definition identifier - source.docx");
    const dstDoc = new aw.Document(base.myDir + "List with the same definition identifier - destination.docx");

    // Set the "KeepSourceNumbering" property to "true" to apply a different list definition ID
    // to identical styles as Aspose.Words imports them into destination documents.
    const importFormatOptions = new aw.ImportFormatOptions();
    importFormatOptions.keepSourceNumbering = true;

    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles, importFormatOptions);
    dstDoc.updateListLabels();
    //ExEnd

    const paraText = dstDoc.sections.at(1).body.lastParagraph.getText();

    expect(paraText.startsWith("13->13")).toEqual(true);
    expect(dstDoc.sections.at(1).body.lastParagraph.listLabel.labelString).toEqual("1.");
  });

  test('MergePastedLists', () => {
    //ExStart
    //ExFor:ImportFormatOptions.mergePastedLists
    //ExSummary:Shows how to merge lists from a documents.
    const srcDoc = new aw.Document(base.myDir + "List item.docx");
    const dstDoc = new aw.Document(base.myDir + "List destination.docx");

    const options = new aw.ImportFormatOptions();
    options.mergePastedLists = true;

    // Set the "MergePastedLists" property to "true" pasted lists will be merged with surrounding lists.
    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles, options);

    dstDoc.save(base.artifactsDir + "Document.MergePastedLists.docx");
    //ExEnd
  });

  test('ForceCopyStyles', () => {
    //ExStart
    //ExFor:ImportFormatOptions.forceCopyStyles
    //ExSummary:Shows how to copy source styles with unique names forcibly.
    // Both documents contain MyStyle1 and MyStyle2, MyStyle3 exists only in a source document.
    const srcDoc = new aw.Document(base.myDir + "Styles source.docx");
    const dstDoc = new aw.Document(base.myDir + "Styles destination.docx");

    const options = new aw.ImportFormatOptions();
    options.forceCopyStyles = true;
    dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.KeepSourceFormatting, options);

    const paras = dstDoc.sections.at(1).body.paragraphs;

    expect(paras.at(0).paragraphFormat.style.name).toEqual("MyStyle1_0");
    expect(paras.at(1).paragraphFormat.style.name).toEqual("MyStyle2_0");
    expect(paras.at(2).paragraphFormat.style.name).toEqual("MyStyle3");
    //ExEnd
  });

  test('AdjustSentenceAndWordSpacing', () => {
      //ExStart
      //ExFor:ImportFormatOptions.adjustSentenceAndWordSpacing
      //ExSummary:Shows how to adjust sentence and word spacing automatically.
      let srcDoc = new aw.Document();
      let dstDoc = new aw.Document();

      var builder = new aw.DocumentBuilder(srcDoc);
      builder.write("Dolor sit amet.");

      builder = new aw.DocumentBuilder(dstDoc);
      builder.write("Lorem ipsum.");

      const options = new aw.ImportFormatOptions();
      options.adjustSentenceAndWordSpacing = true;
      builder.insertDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles, options);

      expect(dstDoc.firstSection.body.firstParagraph.getText().trim()).toEqual("Lorem ipsum. Dolor sit amet.");
      //ExEnd
  });

  test('ValidateIndividualDocumentSignatures', () => {
    //ExStart
    //ExFor:CertificateHolder.certificate
    //ExFor:Document.digitalSignatures
    //ExFor:DigitalSignature
    //ExFor:DigitalSignatureCollection
    //ExFor:DigitalSignature.isValid
    //ExFor:DigitalSignature.comments
    //ExFor:DigitalSignature.signTime
    //ExFor:DigitalSignature.signatureType
    //ExSummary:Shows how to validate and display information about each signature in a document.
    const doc = new aw.Document(base.myDir + "Digitally signed.docx");

    for (var i = 0; i < doc.digitalSignatures.count; i++) {
      const signature = doc.digitalSignatures.at(i);
      console.log(`${signature.isValid ? "Valid" : "Invalid"} signature: `);
      console.log(`\tReason:\t${signature.comments}`);
      console.log(`\tType:\t${signature.signatureType}`);
      console.log(`\tSign time:\t${signature.signTime}`);
      console.log(`\r\n`);
    }
    //ExEnd

    expect(doc.digitalSignatures.count).toEqual(1);

    const digitalSig = doc.digitalSignatures.at(0);

    expect(digitalSig.isValid).toEqual(true);
    expect(digitalSig.comments).toEqual("Test Sign");
    expect(digitalSig.signatureType).toEqual(aw.DigitalSignatures.DigitalSignatureType.XmlDsig);
  });

  test.skip('DigitalSignature - X509Certificate2 type is not supported by Node.js', () => {
  });

  test('SignatureValue', () => {
    //ExStart
    //ExFor:DigitalSignature.signatureValue
    //ExSummary:Shows how to get a digital signature value from a digitally signed document.
    const doc = new aw.Document(base.myDir + "Digitally signed.docx");
    for (var i = 0; i < doc.digitalSignatures.count; i++) {
      const signature = doc.digitalSignatures.at(i);
      const buffer = Buffer.from(signature.signatureValue);
      const signatureValue = buffer.toString('base64');
      expect(signatureValue).toEqual("K1cVLLg2kbJRAzT5WK+m++G8eEO+l7S+5ENdjMxxTXkFzGUfvwxREuJdSFj9AbD" +
                                     "MhnGvDURv9KEhC25DDF1al8NRVR71TF3CjHVZXpYu7edQS5/yLw/k5CiFZzCp1+MmhOdYPcVO+Fm" +
                                     "+9fKr2iNLeyYB+fgEeZHfTqTFM2WwAqo=");
    }
    //ExEnd
  });

  test('AppendAllDocumentsInFolder', () => {
    //ExStart
    //ExFor:Document.appendDocument(Document, ImportFormatMode)
    //ExSummary:Shows how to append all the documents in a folder to the end of a template document.
    const dstDoc = new aw.Document();
    const builder = new aw.DocumentBuilder(dstDoc);
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Heading1;
    builder.writeln("Template Document");
    builder.paragraphFormat.styleIdentifier = aw.StyleIdentifier.Normal;
    builder.writeln("Some content here");
    expect(dstDoc.styles.count).toEqual(5); //ExSkip
    expect(dstDoc.sections.count).toEqual(1); //ExSkip
    // Append all unencrypted documents with the .doc extension
    // from our local file system directory to the base document.
    const docFiles = [];
    fs.readdirSync(base.myDir).forEach(file => {
      if (file.endsWith('.doc')) {
        docFiles.push(base.myDir + file); 
      }
    });
    for (var i = 0; i < docFiles.length; i++) {
      const fileName = docFiles[i];
      const info = aw.FileFormatUtil.detectFileFormat(fileName);
      if (info.isEncrypted)
          continue;
      const srcDoc = new aw.Document(fileName);
      dstDoc.appendDocument(srcDoc, aw.ImportFormatMode.UseDestinationStyles);
    }
    dstDoc.save(base.artifactsDir + "Document.AppendAllDocumentsInFolder.doc");
    //ExEnd
    expect(dstDoc.styles.count).toEqual(7);
    expect(dstDoc.sections.count).toEqual(10);
  });

  test('JoinRunsWithSameFormatting', () => {
    //ExStart
    //ExFor:Document.joinRunsWithSameFormatting
    //ExSummary:Shows how to join runs in a document to reduce unneeded runs.
    // Open a document that contains adjacent runs of text with identical formatting,
    // which commonly occurs if we edit the same paragraph multiple times in Microsoft Word.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    // If any number of these runs are adjacent with identical formatting,
    // then the document may be simplified.
    expect(doc.getChildNodes(aw.NodeType.Run, true).count).toEqual(317);
    // Combine such runs with this method and verify the number of run joins that will take place.
    expect(doc.joinRunsWithSameFormatting()).toEqual(121);
    // The number of joins and the number of runs we have after the join
    // should add up the number of runs we had initially.
    expect(doc.getChildNodes(aw.NodeType.Run, true).count).toEqual(196);
    //ExEnd
  });
              
  test('DefaultTabStop', () => {
    //ExStart
    //ExFor:Document.defaultTabStop
    //ExFor:ControlChar.tab
    //ExFor:ControlChar.tabChar
    //ExSummary:Shows how to set a custom interval for tab stop positions.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    // Set tab stops to appear every 72 points (1 inch).
    builder.document.defaultTabStop = 72;
    // Each tab character snaps the text after it to the next closest tab stop position.
    builder.writeln("Hello" + aw.ControlChar.tab + "World!");
    builder.writeln("Hello" + aw.ControlChar.tabChar + "World!");
    //ExEnd
    doc = DocumentHelper.saveOpen(doc);
    expect(doc.defaultTabStop).toEqual(72);
  });
      
  test('CloneDocument', () => {
    //ExStart
    //ExFor:Document.clone
    //ExSummary:Shows how to deep clone a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Hello world!");
    // Cloning will produce a new document with the same contents as the original,
    // but with a unique copy of each of the original document's nodes.
    let clone = doc.clone();
    expect(doc.firstSection.body.firstParagraph.runs.at(0).getText()).toEqual(
        clone.firstSection.body.firstParagraph.runs.at(0).text);
    //ExEnd
  });

  test('DocumentGetTextToString', () => {
    //ExStart
    //ExFor:CompositeNode.getText
    //ExFor:Node.toString(SaveFormat)
    //ExSummary:Shows the difference between calling the GetText and ToString methods on a node.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.insertField("MERGEFIELD Field");
    // GetText will retrieve the visible text as well as field codes and special characters.
    expect(doc.getText().trim()).toEqual("\u0013MERGEFIELD Field\u0014«Field»\u0015");
    // ToString will give us the document's appearance if saved to a passed save format.
    expect(doc.toString(aw.SaveFormat.Text).trim()).toEqual("«Field»");
    //ExEnd
  });

  test('ProtectUnprotect', () => {
    //ExStart
    //ExFor:Document.protect(ProtectionType,String)
    //ExFor:Document.protectionType
    //ExFor:Document.unprotect
    //ExFor:Document.unprotect(String)
    //ExSummary:Shows how to protect and unprotect a document.
    let doc = new aw.Document();
    doc.protect(aw.ProtectionType.ReadOnly, "password");
    expect(doc.protectionType).toEqual(aw.ProtectionType.ReadOnly);
    // If we open this document with Microsoft Word intending to edit it,
    // we will need to apply the password to get through the protection.
    doc.save(base.artifactsDir + "Document.protect.docx");
    // Note that the protection only applies to Microsoft Word users opening our document.
    // We have not encrypted the document in any way, and we do not need the password to open and edit it programmatically.
    let protectedDoc = new aw.Document(base.artifactsDir + "Document.protect.docx");
    expect(protectedDoc.protectionType).toEqual(aw.ProtectionType.ReadOnly);
    let builder = new aw.DocumentBuilder(protectedDoc);
    builder.writeln("Text added to a protected document.");
    expect(protectedDoc.range.text.trim()).toEqual("Text added to a protected document.");
    // There are two ways of removing protection from a document.
    // 1 - With no password:
    doc.unprotect();
    expect(doc.protectionType).toEqual(aw.ProtectionType.NoProtection);
    doc.protect(aw.ProtectionType.ReadOnly, "NewPassword");
    expect(doc.protectionType).toEqual(aw.ProtectionType.ReadOnly);
    doc.unprotect("WrongPassword");
    expect(doc.protectionType).toEqual(aw.ProtectionType.ReadOnly);
    // 2 - With the correct password:
    doc.unprotect("NewPassword");
    expect(doc.protectionType).toEqual(aw.ProtectionType.NoProtection);
    //ExEnd
  });
 
  /* getChildNodes not supported
  test('DocumentEnsureMinimum', () => {
    //ExStart
    //ExFor:Document.ensureMinimum
    //ExSummary:Shows how to ensure that a document contains the minimal set of nodes required for editing its contents.
    // A newly created document contains one child Section, which includes one child Body and one child Paragraph.
    // We can edit the document body's contents by adding nodes such as Runs or inline Shapes to that paragraph.
    let doc = new aw.Document();
    NodeCollection nodes = doc.getChildNodes(aw.NodeType.Any, true);
    expect(nodes[0].NodeType).toEqual(aw.NodeType.Section);
    expect(nodes[0].ParentNode).toEqual(doc);
    expect(nodes[1].NodeType).toEqual(aw.NodeType.Body);
    expect(nodes[1].ParentNode).toEqual(nodes[0]);
    expect(nodes[2].NodeType).toEqual(aw.NodeType.Paragraph);
    expect(nodes[2].ParentNode).toEqual(nodes[1]);
    // This is the minimal set of nodes that we need to be able to edit the document.
    // We will no longer be able to edit the document if we remove any of them.
    doc.removeAllChildren();
    expect(doc.getChildNodes(aw.NodeType.Any, true).Count).toEqual(0);
    // Call this method to make sure that the document has at least those three nodes so we can edit it again.
    doc.ensureMinimum();
    expect(nodes[0].NodeType).toEqual(aw.NodeType.Section);
    expect(nodes[1].NodeType).toEqual(aw.NodeType.Body);
    expect(nodes[2].NodeType).toEqual(aw.NodeType.Paragraph);
    ((Paragraph)nodes[2]).Runs.add(new aw.Run(doc, "Hello world!"));
    //ExEnd
    expect(doc.getText().trim()).toEqual("Hello world!");
  });  
  */
 
  test('RemoveMacrosFromDocument', () => {
    //ExStart
    //ExFor:Document.removeMacros
    //ExSummary:Shows how to remove all macros from a document.
    let doc = new aw.Document(base.myDir + "Macro.docm");
    expect(doc.hasMacros).toEqual(true);
    expect(doc.vbaProject.name).toEqual("Project");
    // Remove the document's VBA project, along with all its macros.
    doc.removeMacros();
    expect(doc.hasMacros).toEqual(false);
    expect(doc.vbaProject).toEqual(null);
    //ExEnd
  });

  test('GetPageCount', () => {
    //ExStart
    //ExFor:Document.pageCount
    //ExSummary:Shows how to count the number of pages in the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Page 1");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.write("Page 2");
    builder.insertBreak(aw.BreakType.PageBreak);
    builder.write("Page 3");
    // Verify the expected page count of the document.
    expect(doc.pageCount).toEqual(3);
    // Getting the PageCount property invoked the document's page layout to calculate the value.
    // This operation will not need to be re-done when rendering the document to a fixed page save format,
    // such as .pdf. So you can save some time, especially with more complex documents.
    doc.save(base.artifactsDir + "Document.GetPageCount.pdf");
    //ExEnd
  });

  test('GetUpdatedPageProperties', () => {
    //ExStart
    //ExFor:Document.updateWordCount()
    //ExFor:Document.updateWordCount(Boolean)
    //ExFor:BuiltInDocumentProperties.characters
    //ExFor:BuiltInDocumentProperties.words
    //ExFor:BuiltInDocumentProperties.paragraphs
    //ExFor:BuiltInDocumentProperties.lines
    //ExSummary:Shows how to update all list labels in a document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                    "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    builder.write("Ut enim ad minim veniam, " +
                    "quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");
    // Aspose.words does not track document metrics like these in real time.
    expect(doc.builtInDocumentProperties.characters).toEqual(0);
    expect(doc.builtInDocumentProperties.words).toEqual(0);
    expect(doc.builtInDocumentProperties.paragraphs).toEqual(1);
    expect(doc.builtInDocumentProperties.lines).toEqual(1);
    // To get accurate values for three of these properties, we will need to update them manually.
    doc.updateWordCount();
    expect(doc.builtInDocumentProperties.characters).toEqual(196);
    expect(doc.builtInDocumentProperties.words).toEqual(36);
    expect(doc.builtInDocumentProperties.paragraphs).toEqual(2);
    // For the line count, we will need to call a specific overload of the updating method.
    expect(doc.builtInDocumentProperties.lines).toEqual(1);
    doc.updateWordCount(true);
    expect(doc.builtInDocumentProperties.lines).toEqual(4);
    //ExEnd
  });
    
  /*//Commented
  test('TableStyleToDirectFormatting', () => {
    //ExStart
    //ExFor:CompositeNode.getChild
    //ExFor:Document.expandTableStylesToDirectFormatting
    //ExSummary:Shows how to apply the properties of a table's style directly to the table's elements.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    Table table = builder.startTable();
    builder.insertCell();
    builder.write("Hello world!");
    builder.endTable();
    TableStyle tableStyle = (TableStyle)doc.styles.add(aw.StyleType.Table, "MyTableStyle1");
    tableStyle.rowStripe = 3;
    tableStyle.cellSpacing = 5;
    tableStyle.shading.backgroundPatternColor = Color.AntiqueWhite;
    tableStyle.borders.color = Color.blue;
    tableStyle.borders.lineStyle = aw.LineStyle.DotDash;
    table.style = tableStyle;
    // This method concerns table style properties such as the ones we set above.
    doc.expandTableStylesToDirectFormatting();
    doc.save(base.artifactsDir + "Document.TableStyleToDirectFormatting.docx");
    //ExEnd
    TestUtil.DocPackageFileContainsString("<w:tblStyleRowBandSize w:val=\"3\" />",
        base.artifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
    TestUtil.DocPackageFileContainsString("<w:tblCellSpacing w:w=\"100\" w:type=\"dxa\" />",
        base.artifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
    TestUtil.DocPackageFileContainsString("<w:tblBorders><w:top w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:left w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:bottom w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:right w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideH w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /><w:insideV w:val=\"dotDash\" w:sz=\"2\" w:space=\"0\" w:color=\"0000FF\" /></w:tblBorders>",
        base.artifactsDir + "Document.TableStyleToDirectFormatting.docx", "document.xml");
  });
  //EndCommented*/

  test('GetOriginalFileInfo', () => {
    //ExStart
    //ExFor:Document.originalFileName
    //ExFor:Document.originalLoadFormat
    //ExSummary:Shows how to retrieve details of a document's load operation.
    let doc = new aw.Document(base.myDir + "Document.docx");
    expect(doc.originalFileName).toEqual(base.myDir + "Document.docx");
    expect(doc.originalLoadFormat).toEqual(aw.LoadFormat.Docx);
    //ExEnd
  });

  // [Description("WORDSNET-16099")]
  test('FootnoteColumns', () => {
    //ExStart
    //ExFor:FootnoteOptions
    //ExFor:FootnoteOptions.columns
    //ExSummary:Shows how to split the footnote section into a given number of columns.
    let doc = new aw.Document(base.myDir + "Footnotes and endnotes.docx");
    expect(doc.footnoteOptions.columns).toEqual(0);
    doc.footnoteOptions.columns = 2;
    doc.save(base.artifactsDir + "Document.FootnoteColumns.docx");
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.FootnoteColumns.docx");
    expect(doc.firstSection.pageSetup.footnoteOptions.columns).toEqual(2);
  });

  test('Compare', () => {
    //ExStart
    //ExFor:aw.Document.compare(Document, String, DateTime)
    //ExFor:aw.RevisionCollection.acceptAll
    //ExSummary:Shows how to compare documents.
    let docOriginal = new aw.Document();
    let builder = new aw.DocumentBuilder(docOriginal);
    builder.writeln("This is the original document.");
    let docEdited = new aw.Document();
    builder = new aw.DocumentBuilder(docEdited);
    builder.writeln("This is the edited document.");
    // Comparing documents with revisions will throw an exception.
    if (docOriginal.revisions.count == 0 && docEdited.revisions.count == 0)
        docOriginal.compare(docEdited, "authorName", new Date());
    // After the comparison, the original document will gain a new revision
    // for every element that is different in the edited document.
    expect(docOriginal.revisions.count).toEqual(2);
    for (var i = 0; i < docOriginal.revisions.count; i++)
    {
        //var r = docOriginal.revisions.at(i);
        //console.log(`Revision type: ${r.revisionType}, on a node of type "${r.parentNode.nodeType}"`);
        //console.log(`\tChanged text: "${r.parentNode.getText()}"`);
    }
    // Accepting these revisions will transform the original document into the edited document.
    docOriginal.revisions.acceptAll();
    expect(docEdited.getText()).toEqual(docOriginal.getText());
    //ExEnd
    docOriginal = DocumentHelper.saveOpen(docOriginal);
    expect(docOriginal.revisions.count).toEqual(0);
  });

  test('CompareDocumentWithRevisions', () => {
    let doc1 = new aw.Document();
    let builder = new aw.DocumentBuilder(doc1);
    builder.writeln("Hello world! This text is not a revision.");
    let docWithRevision = new aw.Document();
    builder = new aw.DocumentBuilder(docWithRevision);
    docWithRevision.startTrackRevisions("John Doe");
    builder.writeln("This is a revision.");
    expect(() => docWithRevision.compare(doc1, "John Doe", new Date(Date.now()))).toThrow("Compared documents must not have revisions.");
  });

  
  test.each([false, true])('IgnoreDmlUniqueId', (isIgnoreDmlUniqueId) => {
    //ExStart
    //ExFor:aw.Comparing.CompareOptions.ignoreDmlUniqueId
    //ExSummary:Shows how to compare documents ignoring DML unique ID.
    let docA = new aw.Document(base.myDir + "DML unique ID original.docx");
    let docB = new aw.Document(base.myDir + "DML unique ID compare.docx");
    // By default, Aspose.words do not ignore DML's unique ID, and the revisions count was 2.
    // If we are ignoring DML's unique ID, and revisions count were 0.
    let compareOptions = new aw.Comparing.CompareOptions();
    compareOptions.advancedOptions.ignoreDmlUniqueId = isIgnoreDmlUniqueId;
    docA.compare(docB, "Aspose.words", new Date(Date.now()), compareOptions);
    expect(docA.revisions.count).toEqual(isIgnoreDmlUniqueId ? 0 : 2);
    //ExEnd
  });

  test('RemoveExternalSchemaReferences', () => {
    //ExStart
    //ExFor:Document.removeExternalSchemaReferences
    //ExSummary:Shows how to remove all external XML schema references from a document.
    let doc = new aw.Document(base.myDir + "External XML schema.docx");
    doc.removeExternalSchemaReferences();
    //ExEnd
  });

  test.skip('UpdateThumbnail - TODO: Use JSSize for ThumbnailSize', () => {
    //ExStart
    //ExFor:Document.updateThumbnail()
    //ExFor:Document.updateThumbnail(ThumbnailGeneratingOptions)
    //ExFor:ThumbnailGeneratingOptions
    //ExFor:ThumbnailGeneratingOptions.generateFromFirstPage
    //ExFor:ThumbnailGeneratingOptions.thumbnailSize
    //ExSummary:Shows how to update a document's thumbnail.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");
    builder.insertImage(base.imageDir + "Logo.jpg");
    // There are two ways of setting a thumbnail image when saving a document to .epub.
    // 1 -  Use the document's first page:
    doc.updateThumbnail();
    doc.save(base.artifactsDir + "Document.updateThumbnail.firstPage.epub");
    // 2 -  Use the first image found in the document:
    let options = new aw.Rendering.ThumbnailGeneratingOptions();
    expect(options.thumbnailSize).toEqual(new aw.JSSize(600, 900));
    expect(options.generateFromFirstPage).toEqual(true); //ExSkip
    options.thumbnailSize = new aw.JSSize(400, 400);
    options.generateFromFirstPage = false;
    doc.updateThumbnail(options);
    doc.save(base.artifactsDir + "Document.updateThumbnail.FirstImage.epub");
    //ExEnd
  });

  test('HyphenationOptions', () => {
    //ExStart
    //ExFor:Document.hyphenationOptions
    //ExFor:HyphenationOptions
    //ExFor:HyphenationOptions.autoHyphenation
    //ExFor:HyphenationOptions.consecutiveHyphenLimit
    //ExFor:HyphenationOptions.hyphenationZone
    //ExFor:HyphenationOptions.hyphenateCaps
    //ExSummary:Shows how to configure automatic hyphenation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.font.size = 24;
    builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                    "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
    doc.hyphenationOptions.autoHyphenation = true;
    doc.hyphenationOptions.consecutiveHyphenLimit = 2;
    doc.hyphenationOptions.hyphenationZone = 720;
    doc.hyphenationOptions.hyphenateCaps = true;
    doc.save(base.artifactsDir + "Document.hyphenationOptions.docx");
    //ExEnd
    expect(doc.hyphenationOptions.autoHyphenation).toEqual(true);
    expect(doc.hyphenationOptions.consecutiveHyphenLimit).toEqual(2);
    expect(doc.hyphenationOptions.hyphenationZone).toEqual(720);
    expect(doc.hyphenationOptions.hyphenateCaps).toEqual(true);
    expect(DocumentHelper.compareDocs(base.artifactsDir + "Document.hyphenationOptions.docx",
        base.goldsDir + "Document.hyphenationOptions Gold.docx")).toEqual(true);
  });  

  test('HyphenationOptionsDefaultValues', () => {
    let doc = new aw.Document();
    let doc2 = DocumentHelper.saveOpen(doc);
    expect(doc2.hyphenationOptions.autoHyphenation).toEqual(false);
    expect(doc2.hyphenationOptions.consecutiveHyphenLimit).toEqual(0);
    expect(doc2.hyphenationOptions.hyphenationZone).toEqual(360);
    expect(doc2.hyphenationOptions.hyphenateCaps).toEqual(true);
  });

  test('HyphenationZoneException', () => {
    let doc = new aw.Document();
    expect(() => doc.hyphenationOptions.hyphenationZone = 0).toThrow("Specified argument was out of the range of valid values.");
  });

  test('OoxmlComplianceVersion', () => {
    //ExStart
    //ExFor:Document.compliance
    //ExSummary:Shows how to read a loaded document's Open Office XML compliance version.
    // The compliance version varies between documents created by different versions of Microsoft Word.
    let doc = new aw.Document(base.myDir + "Document.doc");
    expect(aw.Saving.OoxmlCompliance.Ecma376_2006).toEqual(doc.compliance);
    doc = new aw.Document(base.myDir + "Document.docx");
    expect(aw.Saving.OoxmlCompliance.Iso29500_2008_Transitional).toEqual(doc.compliance);
    //ExEnd
  });

  // WORDSNET-20342
  test('ImageSaveOptions', async () => {
    //ExStart
    //ExFor:Document.save(String, SaveOptions)
    //ExFor:SaveOptions.useAntiAliasing
    //ExFor:SaveOptions.useHighQualityRendering
    //ExSummary:Shows how to improve the quality of a rendered document with SaveOptions.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    let builder = new aw.DocumentBuilder(doc);
    builder.font.size = 60;
    builder.writeln("Some text.");
    let options = new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg);
    expect(options.useAntiAliasing).toEqual(false);
    expect(options.useHighQualityRendering).toEqual(false);
    doc.save(base.artifactsDir + "Document.ImageSaveOptions.default.jpg", options);
    options.useAntiAliasing = true;
    options.useHighQualityRendering = true;
    doc.save(base.artifactsDir + "Document.ImageSaveOptions.HighQuality.jpg", options);
    //ExEnd
    await TestUtil.verifyImage(794, 1122, base.artifactsDir + "Document.ImageSaveOptions.default.jpg");
    await TestUtil.verifyImage(794, 1122, base.artifactsDir + "Document.ImageSaveOptions.HighQuality.jpg");
  });

  test('Cleanup', () => {
    //ExStart
    //ExFor:Document.cleanup
    //ExSummary:Shows how to remove unused custom styles from a document.
    let doc = new aw.Document();
    doc.styles.add(aw.StyleType.List, "MyListStyle1");
    doc.styles.add(aw.StyleType.List, "MyListStyle2");
    doc.styles.add(aw.StyleType.Character, "MyParagraphStyle1");
    doc.styles.add(aw.StyleType.Character, "MyParagraphStyle2");
    // Combined with the built-in styles, the document now has eight styles.
    // A custom style counts as "used" while applied to some part of the document,
    // which means that the four styles we added are currently unused.
    expect(doc.styles.count).toEqual(8);
    // Apply a custom character style, and then a custom list style. Doing so will mark the styles as "used".
    let builder = new aw.DocumentBuilder(doc);
    builder.font.style = doc.styles.at("MyParagraphStyle1");
    builder.writeln("Hello world!");
    let list = doc.lists.add(doc.styles.at("MyListStyle1"));
    builder.listFormat.list = list;
    builder.writeln("Item 1");
    builder.writeln("Item 2");
    doc.cleanup();
    expect(doc.styles.count).toEqual(6);
    // Removing every node that a custom style is applied to marks it as "unused" again.
    // Run the Cleanup method again to remove them.
    doc.firstSection.body.removeAllChildren();
    doc.cleanup();
    expect(doc.styles.count).toEqual(4);
    //ExEnd
  });

  test('AutomaticallyUpdateStyles', () => {
    //ExStart
    //ExFor:Document.automaticallyUpdateStyles
    //ExSummary:Shows how to attach a template to a document.
    let doc = new aw.Document();
    // Microsoft Word documents by default come with an attached template called "Normal.dotm".
    // There is no default template for blank Aspose.words documents.
    expect(doc.attachedTemplate).toEqual("");
    // Attach a template, then set the flag to apply style changes
    // within the template to styles in our document.
    doc.attachedTemplate = base.myDir + "Business brochure.dotx";
    doc.automaticallyUpdateStyles = true;
    doc.save(base.artifactsDir + "Document.automaticallyUpdateStyles.docx");
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.automaticallyUpdateStyles.docx");
    expect(doc.automaticallyUpdateStyles).toEqual(true);
    expect(doc.attachedTemplate).toEqual(base.myDir + "Business brochure.dotx");
    expect(fs.existsSync(doc.attachedTemplate)).toEqual(true);
  });

  test('DefaultTemplate', () => {
    //ExStart
    //ExFor:Document.attachedTemplate
    //ExFor:Document.automaticallyUpdateStyles
    //ExFor:SaveOptions.createSaveOptions(String)
    //ExFor:SaveOptions.defaultTemplate
    //ExSummary:Shows how to set a default template for documents that do not have attached templates.
    let doc = new aw.Document();
    // Enable automatic style updating, but do not attach a template document.
    doc.automaticallyUpdateStyles = true;
    expect(doc.attachedTemplate).toEqual('');
    // Since there is no template document, the document had nowhere to track style changes.
    // Use a SaveOptions object to automatically set a template
    // if a document that we are saving does not have one.
    let options = aw.Saving.SaveOptions.createSaveOptions("Document.defaultTemplate.docx");
    options.defaultTemplate = base.myDir + "Business brochure.dotx";
    doc.save(base.artifactsDir + "Document.defaultTemplate.docx", options);
    //ExEnd
    expect(fs.existsSync(options.defaultTemplate)).toEqual(true);
  });

  test.skip('UseSubstitutions - TODO: Regex is not supported yet.', () => {
    //ExStart
    //ExFor:FindReplaceOptions.#ctor
    //ExFor:FindReplaceOptions.useSubstitutions
    //ExFor:FindReplaceOptions.legacyMode
    //ExSummary:Shows how to recognize and use substitutions within replacement patterns.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Jason gave money to Paul.");
    let regex = new Regex("([A-z]+) gave money to ([A-z]+)");
    let options = new aw.Replacing.FindReplaceOptions();
    options.useSubstitutions = true;
    // Using legacy mode does not support many advanced features, so we need to set it to 'false'.
    options.legacyMode = false;
    doc.range.replace(regex, "$2 took money from $1", options);
    expect(doc.getText()).toEqual("Paul took money from Jason.\f");
    //ExEnd
  });

  test('SetInvalidateFieldTypes', () => {
    //ExStart
    //ExFor:Document.normalizeFieldTypes
    //ExFor:Range.normalizeFieldTypes
    //ExSummary:Shows how to get the keep a field's type up to date with its field code.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    let field = builder.insertField("DATE", null);
    // Aspose.words automatically detects field types based on field codes.
    expect(field.type).toEqual(aw.Fields.FieldType.FieldDate);
    // Manually change the raw text of the field, which determines the field code.
    let fieldText = doc.firstSection.body.firstParagraph.getRun(0, true);
    expect(fieldText.text).toEqual("DATE");
    fieldText.text = "PAGE";
    // Changing the field code has changed this field to one of a different type,
    // but the field's type properties still display the old type.
    expect(field.getFieldCode()).toEqual("PAGE");
    expect(field.type).toEqual(aw.Fields.FieldType.FieldDate);
    expect(field.start.fieldType).toEqual(aw.Fields.FieldType.FieldDate);
    expect(field.separator.fieldType).toEqual(aw.Fields.FieldType.FieldDate);
    expect(field.end.fieldType).toEqual(aw.Fields.FieldType.FieldDate);
    // Update those properties with this method to display current value.
    doc.normalizeFieldTypes();
    expect(field.type).toEqual(aw.Fields.FieldType.FieldPage);
    expect(field.start.fieldType).toEqual(aw.Fields.FieldType.FieldPage);
    expect(field.separator.fieldType).toEqual(aw.Fields.FieldType.FieldPage);
    expect(field.end.fieldType).toEqual(aw.Fields.FieldType.FieldPage);
    //ExEnd
  });

  test.each([false, true])('LayoutOptionsHiddenText', (showHiddenText) => {
    //ExStart
    //ExFor:Document.layoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.showHiddenText
    //ExSummary:Shows how to hide text in a rendered output document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    expect(doc.layoutOptions.showHiddenText).toEqual(false);
    // Insert hidden text, then specify whether we wish to omit it from a rendered document.
    builder.writeln("This text is not hidden.");
    builder.font.hidden = true;
    builder.writeln("This text is hidden.");
    doc.layoutOptions.showHiddenText = showHiddenText;
    doc.save(base.artifactsDir + "Document.LayoutOptionsHiddenText.pdf");
    //ExEnd
  });

  /* 1374: #if !WORDS_AOT
  test.each[false,
    true])('UsePdfDocumentForLayoutOptionsHiddenText', (bool showHiddenText) => {
    LayoutOptionsHiddenText(showHiddenText);
    Aspose.pdf.document pdfDoc = new Aspose.pdf.document(base.artifactsDir + "Document.LayoutOptionsHiddenText.pdf");
    let textAbsorber = new TextAbsorber();
    textAbsorber.Visit(pdfDoc);
    Assert.AreEqual(showHiddenText ?
            $"This text is not hidden.{Environment.NewLine}This text is hidden." :
            "This text is not hidden.", textAbsorber.text);
  });
  1387: #endif  */
  
  test.each([false, true])('LayoutOptionsParagraphMarks', (showParagraphMarks) => {
    //ExStart
    //ExFor:Document.layoutOptions
    //ExFor:LayoutOptions
    //ExFor:LayoutOptions.showParagraphMarks
    //ExSummary:Shows how to show paragraph marks in a rendered output document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    expect(doc.layoutOptions.showParagraphMarks).toEqual(false);
    // Add some paragraphs, then enable paragraph marks to show the ends of paragraphs
    // with a pilcrow (¶) symbol when we render the document.
    builder.writeln("Hello world!");
    builder.writeln("Hello again!");
    doc.layoutOptions.showParagraphMarks = showParagraphMarks;
    doc.save(base.artifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
    //ExEnd
  });
  
  /* 1408: #if !WORDS_AOT
  test.each[false,
    true])('UsePdfDocumentForLayoutOptionsParagraphMarks', (bool showParagraphMarks) => {
    LayoutOptionsParagraphMarks(showParagraphMarks);
    Aspose.pdf.document pdfDoc = new Aspose.pdf.document(base.artifactsDir + "Document.LayoutOptionsParagraphMarks.pdf");
    let textAbsorber = new TextAbsorber();
    textAbsorber.Visit(pdfDoc);
    Assert.AreEqual(showParagraphMarks ?
            $"Hello world!¶{Environment.NewLine}Hello again!¶{Environment.NewLine}¶" :
            $"Hello world!{Environment.NewLine}Hello again!", textAbsorber.text.trim());
  });
  1421: #endif */
  
  test('UpdatePageLayout', () => {
    //ExStart
    //ExFor:StyleCollection.item(String)
    //ExFor:SectionCollection.item(Int32)
    //ExFor:Document.updatePageLayout
    //ExFor:Margins
    //ExFor:PageSetup.margins
    //ExSummary:Shows when to recalculate the page layout of the document.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    // Saving a document to PDF, to an image, or printing for the first time will automatically
    // cache the layout of the document within its pages.
    doc.save(base.artifactsDir + "Document.updatePageLayout.1.pdf");
    // Modify the document in some way.
    doc.styles.at("Normal").font.size = 6;
    doc.sections.at(0).pageSetup.orientation = aw.Orientation.Landscape;
    doc.sections.at(0).pageSetup.margins = aw.Margins.Mirrored;
    // In the current version of Aspose.words, modifying the document does not automatically rebuild
    // the cached page layout. If we wish for the cached layout
    // to stay up to date, we will need to update it manually.
    doc.updatePageLayout();
    doc.save(base.artifactsDir + "Document.updatePageLayout.2.pdf");
    //ExEnd
  });
  
  test('DocPackageCustomParts', () => {
    //ExStart
    //ExFor:CustomPart
    //ExFor:CustomPart.contentType
    //ExFor:CustomPart.relationshipType
    //ExFor:CustomPart.isExternal
    //ExFor:CustomPart.data
    //ExFor:CustomPart.name
    //ExFor:CustomPart.clone
    //ExFor:CustomPartCollection
    //ExFor:CustomPartCollection.add(CustomPart)
    //ExFor:CustomPartCollection.clear
    //ExFor:CustomPartCollection.clone
    //ExFor:CustomPartCollection.count
    //ExFor:CustomPartCollection.getEnumerator
    //ExFor:CustomPartCollection.item(Int32)
    //ExFor:CustomPartCollection.removeAt(Int32)
    //ExFor:Document.packageCustomParts
    //ExSummary:Shows how to access a document's arbitrary custom parts collection.
    let doc = new aw.Document(base.myDir + "Custom parts OOXML package.docx");
    expect(doc.packageCustomParts.count).toEqual(2);
    // Clone the second part, then add the clone to the collection.
    let clonedPart = doc.packageCustomParts.at(1).clone();
    doc.packageCustomParts.add(clonedPart);
    testDocPackageCustomParts(doc.packageCustomParts); //ExSkip
    expect(doc.packageCustomParts.count).toEqual(3);
    // Enumerate over the collection and print every part.
    const items = [...doc.packageCustomParts];  
    let index = 0;
    items.forEach((c) => {
      /*console.log(`Part index ${index}:`);
      console.log(`\tName:\t\t\t\t${c.name}`);
      console.log(`\tContent type:\t\t${c.contentType}`);
      console.log(`\tRelationship type:\t${c.relationshipType}`);
      console.log(c.isExternal ?
        "\tSourced from outside the document" :
        `\tStored within the document, length: ${c.data.length} bytes`);
      index++;*/
    });
    // We can remove elements from this collection individually, or all at once.
    doc.packageCustomParts.removeAt(2);
    expect(doc.packageCustomParts.count).toEqual(2);
    doc.packageCustomParts.clear();
    expect(doc.packageCustomParts.count).toEqual(0);
    //ExEnd
  });
  
  test.each([false, true])('ShadeFormData', (useGreyShading) => {
    //ExStart
    //ExFor:Document.shadeFormData
    //ExSummary:Shows how to apply gray shading to form fields.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    expect(doc.shadeFormData).toEqual(true);
    builder.write("Hello world! ");
    builder.insertTextInput("My form field", aw.Fields.TextFormFieldType.Regular, "",
        "Text contents of form field, which are shaded in grey by default.", 0);
    // We can turn the grey shading off, so the bookmarked text will blend in with the other text.
    doc.shadeFormData = useGreyShading;
    doc.save(base.artifactsDir + "Document.shadeFormData.docx");
    //ExEnd
  });
  
  test('VersionsCount', () => {
    //ExStart
    //ExFor:Document.versionsCount
    //ExSummary:Shows how to work with the versions count feature of older Microsoft Word documents.
    let doc = new aw.Document(base.myDir + "Versions.doc");
    // We can read this property of a document, but we cannot preserve it while saving.
    expect(doc.versionsCount).toEqual(4);
    doc.save(base.artifactsDir + "Document.versionsCount.doc");
    doc = new aw.Document(base.artifactsDir + "Document.versionsCount.doc");
    expect(doc.versionsCount).toEqual(0);
    //ExEnd
  });
  
  test('WriteProtection', () => {
    //ExStart
    //ExFor:Document.writeProtection
    //ExFor:WriteProtection
    //ExFor:WriteProtection.isWriteProtected
    //ExFor:WriteProtection.readOnlyRecommended
    //ExFor:WriteProtection.setPassword(String)
    //ExFor:WriteProtection.validatePassword(String)
    //ExSummary:Shows how to protect a document with a password.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world! This document is protected.");
    expect(doc.writeProtection.isWriteProtected).toEqual(false);
    expect(doc.writeProtection.readOnlyRecommended).toEqual(false);
    // Enter a password up to 15 characters in length, and then verify the document's protection status.
    doc.writeProtection.setPassword("MyPassword");
    doc.writeProtection.readOnlyRecommended = true;
    expect(doc.writeProtection.isWriteProtected).toEqual(true);
    expect(doc.writeProtection.validatePassword("MyPassword")).toEqual(true);
    // Protection does not prevent the document from being edited programmatically, nor does it encrypt the contents.
    doc.save(base.artifactsDir + "Document.writeProtection.docx");
    doc = new aw.Document(base.artifactsDir + "Document.writeProtection.docx");
    expect(doc.writeProtection.isWriteProtected).toEqual(true);
    builder = new aw.DocumentBuilder(doc);
    builder.moveToDocumentEnd();
    builder.writeln("Writing text in a protected document.");
    expect(doc.getText().trim()).toEqual("Hello world! This document is protected." +
      "\rWriting text in a protected document.");
    //ExEnd
    expect(doc.writeProtection.readOnlyRecommended).toEqual(true);
    expect(doc.writeProtection.validatePassword("MyPassword")).toEqual(true);
    expect(doc.writeProtection.validatePassword("wrongpassword")).toEqual(false);
  });
  
  test.each([false, true])('RemovePersonalInformation', (saveWithoutPersonalInfo) => {
    //ExStart
    //ExFor:Document.removePersonalInformation
    //ExSummary:Shows how to enable the removal of personal information during a manual save.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    // Insert some content with personal information.
    doc.builtInDocumentProperties.author = "John Doe";
    doc.builtInDocumentProperties.company = "Placeholder Inc.";
    doc.startTrackRevisions(doc.builtInDocumentProperties.author, Date.now());
    builder.write("Hello world!");
    doc.stopTrackRevisions();
    // This flag is equivalent to File -> Options -> Trust Center -> Trust Center Settings... ->
    // Privacy Options -> "Remove personal information from file properties on save" in Microsoft Word.
    doc.removePersonalInformation = saveWithoutPersonalInfo;
    // This option will not take effect during a save operation made using Aspose.words.
    // Personal data will be removed from our document with the flag set when we save it manually using Microsoft Word.
    doc.save(base.artifactsDir + "Document.removePersonalInformation.docx");
    doc = new aw.Document(base.artifactsDir + "Document.removePersonalInformation.docx");
    expect(doc.removePersonalInformation).toEqual(saveWithoutPersonalInfo);
    expect(doc.builtInDocumentProperties.author).toEqual("John Doe");
    expect(doc.builtInDocumentProperties.company).toEqual("Placeholder Inc.");
    expect(doc.revisions.at(0).author).toEqual("John Doe");
    //ExEnd
  });
  
  test('ShowComments', () => {
    //ExStart
    //ExFor:LayoutOptions.commentDisplayMode
    //ExFor:CommentDisplayMode
    //ExSummary:Shows how to show comments when saving a document to a rendered format.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.write("Hello world!");
    let comment = new aw.Comment(doc, "John Doe", "J.D.", Date.now());
    comment.setText("My comment.");
    builder.currentParagraph.appendChild(comment);
    // ShowInAnnotations is only available in Pdf1.7 and Pdf1.5 formats.
    // In other formats, it will work similarly to Hide.
    doc.layoutOptions.commentDisplayMode = aw.Layout.CommentDisplayMode.ShowInAnnotations;
    doc.save(base.artifactsDir + "Document.ShowCommentsInAnnotations.pdf");

    // Note that it's required to rebuild the document page layout (via Document.updatePageLayout() method)
    // after changing the Document.layoutOptions values.
    doc.layoutOptions.commentDisplayMode = aw.Layout.CommentDisplayMode.ShowInBalloons;
    doc.updatePageLayout();
    doc.save(base.artifactsDir + "Document.ShowCommentsInBalloons.pdf");
    //ExEnd
  });
  
  /* 1635: #if !WORDS_AOT
  test('UsePdfDocumentForShowComments', () => {
    ShowComments();
    Aspose.pdf.document pdfDoc = new Aspose.pdf.document(base.artifactsDir + "Document.ShowCommentsInBalloons.pdf");
    let textAbsorber = new TextAbsorber();
    textAbsorber.Visit(pdfDoc);
    Assert.AreEqual(
        "Hello world!                                                                    Commented [J.D.1]:  My comment.",
        textAbsorber.text);
  });
  1647: #endif*/
  
  test('CopyTemplateStylesViaDocument', () => {
    //ExStart
    //ExFor:Document.copyStylesFromTemplate(Document)
    //ExSummary:Shows how to copies styles from the template to a document via Document.
    let template = new aw.Document(base.myDir + "Rendering.docx");
    let target = new aw.Document(base.myDir + "Document.docx");
    expect(template.styles.count).toEqual(18);
    expect(target.styles.count).toEqual(12);
    target.copyStylesFromTemplate(template);
    expect(target.styles.count).toEqual(22);
    //ExEnd
  });
  
  test('CopyTemplateStylesViaDocumentNew', () => {
    //ExStart
    //ExFor:Document.copyStylesFromTemplate(Document)
    //ExFor:Document.copyStylesFromTemplate(String)
    //ExSummary:Shows how to copy styles from one document to another.
    // Create a document, and then add styles that we will copy to another document.
    let template = new aw.Document();
    let style = template.styles.add(aw.StyleType.Paragraph, "TemplateStyle1");
    style.font.name = "Times New Roman";
    style.font.color = "#000080"; //Color.Navy;
    style = template.styles.add(aw.StyleType.Paragraph, "TemplateStyle2");
    style.font.name = "Arial";
    style.font.color = "00BFff"; //Color.DeepSkyBlue;
    style = template.styles.add(aw.StyleType.Paragraph, "TemplateStyle3");
    style.font.name = "Courier New";
    style.font.color = "#4169e1"; //Color.RoyalBlue;
    expect(template.styles.count).toEqual(7);
    // Create a document which we will copy the styles to.
    let target = new aw.Document();
    // Create a style with the same name as a style from the template document and add it to the target document.
    style = target.styles.add(aw.StyleType.Paragraph, "TemplateStyle3");
    style.font.name = "Calibri";
    style.font.color = "#FFA500"; //Color.Orange;
    expect(target.styles.count).toEqual(5);
    // There are two ways of calling the method to copy all the styles from one document to another.
    // 1 -  Passing the template document object:
    target.copyStylesFromTemplate(template);
    // Copying styles adds all styles from the template document to the target
    // and overwrites existing styles with the same name.
    expect(target.styles.count).toEqual(7);
    expect(target.styles.at("TemplateStyle3").font.name).toEqual("Courier New");
    expect(target.styles.at("TemplateStyle3").font.color).toEqual("#4169E1");
    // 2 -  Passing the local system filename of a template document:
    target.copyStylesFromTemplate(base.myDir + "Rendering.docx");
    expect(target.styles.count).toEqual(21);
    //ExEnd
  });
  
  test('ReadMacrosFromExistingDocument', () => {
    //ExStart
    //ExFor:Document.vbaProject
    //ExFor:VbaModuleCollection
    //ExFor:VbaModuleCollection.count
    //ExFor:VbaModuleCollection.item(System.int32)
    //ExFor:VbaModuleCollection.item(System.string)
    //ExFor:VbaModuleCollection.remove
    //ExFor:VbaModule
    //ExFor:VbaModule.name
    //ExFor:VbaModule.sourceCode
    //ExFor:VbaProject
    //ExFor:VbaProject.name
    //ExFor:VbaProject.modules
    //ExFor:VbaProject.codePage
    //ExFor:VbaProject.isSigned
    //ExSummary:Shows how to access a document's VBA project information.
    let doc = new aw.Document(base.myDir + "VBA project.docm");
    // A VBA project contains a collection of VBA modules.
    let vbaProject = doc.vbaProject;
    expect(vbaProject.isSigned).toEqual(true);
    console.log(vbaProject.isSigned
        ? `Project name: ${vbaProject.name} signed; Project code page: ${vbaProject.codePage}; Modules count: ${vbaProject.modules.count}\n`
        : `Project name: ${vbaProject.name} not signed; Project code page: ${vbaProject.codePage}; Modules count: ${vbaProject.modules.count}\n`);
    let vbaModules = doc.vbaProject.modules;
    expect(vbaModules.count).toEqual(3);
    /*for (var i = 0; i < vbaModules.count; i++) {
        console.log(`Module name: ${vbaModules.at(i).name};\nModule code:\n${vbaModules.at(i).sourceCode}\n`);
    };*/
    // Set new source code for VBA module. You can access VBA modules in the collection either by index or by name.
    vbaModules.at(0).sourceCode = "Your VBA code...";
    vbaModules.at("Module1").sourceCode = "Your VBA code...";
    // Remove a module from the collection.
    vbaModules.remove(vbaModules.at(2));
    //ExEnd
    expect(vbaProject.name).toEqual("AsposeVBAtest");
    expect(vbaProject.modules.count).toEqual(2);
    expect(vbaProject.codePage).toEqual(1251);
    expect(vbaProject.isSigned).toEqual(false);
    expect(vbaModules.at(0).name).toEqual("ThisDocument");
    expect(vbaModules.at(0).sourceCode).toEqual("Your VBA code...");
    expect(vbaModules.at(1).name).toEqual("Module1");
    expect(vbaModules.at(1).sourceCode).toEqual("Your VBA code...");
  });
  
  test('SaveOutputParameters', () => {
    //ExStart
    //ExFor:SaveOutputParameters
    //ExFor:SaveOutputParameters.contentType
    //ExSummary:Shows how to access output parameters of a document's save operation.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");
    // After we save a document, we can access the Internet Media Type (MIME type) of the newly created output document.
    let parameters = doc.save(base.artifactsDir + "Document.SaveOutputParameters.doc");
    expect(parameters.contentType).toEqual("application/msword");
    // This property changes depending on the save format.
    parameters = doc.save(base.artifactsDir + "Document.SaveOutputParameters.pdf");
    expect(parameters.contentType).toEqual("application/pdf");
    //ExEnd
  });
  
  test('SubDocument', () => {
    //ExStart
    //ExFor:SubDocument
    //ExFor:SubDocument.nodeType
    //ExSummary:Shows how to access a master document's subdocument.
    let doc = new aw.Document(base.myDir + "Master document.docx");
    let subDocuments = doc.getChildNodes(aw.NodeType.SubDocument, true);
    expect(subDocuments.count).toEqual(1);
    // This node serves as a reference to an external document, and its contents cannot be accessed.
    subDocument = subDocuments.at(0).asSubDocument();
    expect(subDocument.isComposite).toEqual(false);
    //ExEnd
  });
  
  test.skip('CreateWebExtension - TODO: BaseWebExtensionCollection not supported yet.', () => {
    //ExStart
    //ExFor:BaseWebExtensionCollection`1.add(`0)
    //ExFor:BaseWebExtensionCollection`1.clear
    //ExFor:Document.webExtensionTaskPanes
    //ExFor:TaskPane
    //ExFor:TaskPane.dockState
    //ExFor:TaskPane.isVisible
    //ExFor:TaskPane.width
    //ExFor:TaskPane.isLocked
    //ExFor:TaskPane.webExtension
    //ExFor:TaskPane.row
    //ExFor:WebExtension
    //ExFor:WebExtension.id
    //ExFor:WebExtension.alternateReferences
    //ExFor:WebExtension.reference
    //ExFor:WebExtension.properties
    //ExFor:WebExtension.bindings
    //ExFor:WebExtension.isFrozen
    //ExFor:WebExtensionReference
    //ExFor:WebExtensionReference.id
    //ExFor:WebExtensionReference.version
    //ExFor:WebExtensionReference.storeType
    //ExFor:WebExtensionReference.store
    //ExFor:WebExtensionPropertyCollection
    //ExFor:WebExtensionBindingCollection
    //ExFor:WebExtensionProperty.#ctor(String, String)
    //ExFor:WebExtensionProperty.name
    //ExFor:WebExtensionProperty.value
    //ExFor:WebExtensionBinding.#ctor(String, WebExtensionBindingType, String)
    //ExFor:WebExtensionStoreType
    //ExFor:WebExtensionBindingType
    //ExFor:TaskPaneDockState
    //ExFor:TaskPaneCollection
    //ExFor:WebExtensionBinding.id
    //ExFor:WebExtensionBinding.appRef
    //ExFor:WebExtensionBinding.bindingType
    //ExSummary:Shows how to add a web extension to a document.
    let doc = new aw.Document();
    // Create task pane with "MyScript" add-in, which will be used by the document,
    // then set its default location.
    let myScriptTaskPane = new aw.WebExtensions.TaskPane();
    doc.webExtensionTaskPanes.add(myScriptTaskPane);
    myScriptTaskPane.dockState = aw.WebExtensions.TaskPaneDockState.Right;
    myScriptTaskPane.isVisible = true;
    myScriptTaskPane.width = 300;
    myScriptTaskPane.isLocked = true;
    // If there are multiple task panes in the same docking location, we can set this index to arrange them.
    myScriptTaskPane.row = 1;
    // Create an add-in called "MyScript Math Sample", which the task pane will display within.
    let webExtension = myScriptTaskPane.webExtension;
    // Set application store reference parameters for our add-in, such as the ID.
    webExtension.reference.id = "WA104380646";
    webExtension.reference.version = "1.0.0.0";
    webExtension.reference.storeType = aw.WebExtensions.WebExtensionStoreType.OMEX;
    //webExtension.reference.store = CultureInfo.CurrentCulture.name;
    webExtension.properties.add(new aw.WebExtensions.WebExtensionProperty("MyScript", "MyScript Math Sample"));
    webExtension.bindings.add(new aw.WebExtensions.WebExtensionBinding("MyScript", aw.WebExtensions.WebExtensionBindingType.Text, "104380646"));
    // Allow the user to interact with the add-in.
    webExtension.isFrozen = false;
    // We can access the web extension in Microsoft Word via Developer -> Add-ins.
    doc.save(base.artifactsDir + "Document.webExtension.docx");
    // Remove all web extension task panes at once like this.
    doc.webExtensionTaskPanes.clear();
    expect(doc.webExtensionTaskPanes.count).toEqual(0);
    doc = new aw.Document(base.artifactsDir + "Document.webExtension.docx");
            
    myScriptTaskPane = doc.webExtensionTaskPanes.at(0);
    expect(myScriptTaskPane.dockState).toEqual(aw.WebExtensions.TaskPaneDockState.Right);
    expect(myScriptTaskPane.isVisible).toEqual(true);
    expect(myScriptTaskPane.width).toEqual(300.0);
    expect(myScriptTaskPane.isLocked).toEqual(true);
    expect(myScriptTaskPane.row).toEqual(1);

    webExtension = myScriptTaskPane.webExtension;
    expect(webExtension.id).toEqual("");    
    expect(webExtension.reference.id).toEqual("WA104380646");
    expect(webExtension.reference.version).toEqual("1.0.0.0");
    expect(webExtension.reference.storeType).toEqual(aw.WebExtensions.WebExtensionStoreType.OMEX);
    //expect(webExtension.reference.store).toEqual(CultureInfo.CurrentCulture.name);
    expect(webExtension.properties.at(0).name).toEqual("MyScript");
    expect(webExtension.properties.at(0).value).toEqual("MyScript Math Sample");
    expect(webExtension.bindings.at(0).Id).toEqual("MyScript");
    expect(webExtension.bindings.at(0).bindingType).toEqual(aw.WebExtensions.WebExtensionBindingType.Text);
    expect(webExtension.bindings.at(0).appRef).toEqual("104380646");
    expect(webExtension.isFrozen).toEqual(false);
    //ExEnd
  });
  
  test('GetWebExtensionInfo', () => {
    //ExStart
    //ExFor:BaseWebExtensionCollection`1
    //ExFor:BaseWebExtensionCollection`1.getEnumerator
    //ExFor:BaseWebExtensionCollection`1.remove(Int32)
    //ExFor:BaseWebExtensionCollection`1.count
    //ExFor:BaseWebExtensionCollection`1.item(Int32)
    //ExSummary:Shows how to work with a document's collection of web extensions.
    let doc = new aw.Document(base.myDir + "Web extension.docx");
    expect(doc.webExtensionTaskPanes.count).toEqual(1);
    // Print all properties of the document's web extension.
    let webExtensionPropertyCollection = doc.webExtensionTaskPanes.at(0).webExtension.properties;
    for (var i = 0; i < webExtensionPropertyCollection.count; i++)
    {
      let webExtensionProperty = webExtensionPropertyCollection.at(i);
      console.log(`Binding name: ${webExtensionProperty.name}; Binding value: ${webExtensionProperty.value}`);
    }
    // Remove the web extension.
    doc.webExtensionTaskPanes.remove(0);
    expect(doc.webExtensionTaskPanes.count).toEqual(0);
    //ExEnd
  });
  
  test('EpubCover', () => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Hello world!");
    // When saving to .epub, some Microsoft Word document properties convert to .epub metadata.
    doc.builtInDocumentProperties.author = "John Doe";
    doc.builtInDocumentProperties.title = "My Book Title";
    // The thumbnail we specify here can become the cover image.
    // The thumbnail we specify here can become the cover image.
    let image = base.loadFileToArray(base.imageDir + "Transparent background logo.png");
    doc.builtInDocumentProperties.thumbnail =  image;
    doc.save(base.artifactsDir + "Document.EpubCover.epub");
  });
  
  test('TextWatermark', () => {
    //ExStart
    //ExFor:Document.watermark
    //ExFor:Watermark
    //ExFor:Watermark.setText(String)
    //ExFor:Watermark.setText(String, TextWatermarkOptions)
    //ExFor:Watermark.remove
    //ExFor:TextWatermarkOptions
    //ExFor:TextWatermarkOptions.fontFamily
    //ExFor:TextWatermarkOptions.fontSize
    //ExFor:TextWatermarkOptions.color
    //ExFor:TextWatermarkOptions.layout
    //ExFor:TextWatermarkOptions.isSemitrasparent
    //ExFor:WatermarkLayout
    //ExFor:WatermarkType
    //ExFor:Watermark.type
    //ExSummary:Shows how to create a text watermark.
    let doc = new aw.Document();
    // Add a plain text watermark.
    doc.watermark.setText("Aspose Watermark");
    // If we wish to edit the text formatting using it as a watermark,
    // we can do so by passing a TextWatermarkOptions object when creating the watermark.
    let textWatermarkOptions = new aw.TextWatermarkOptions();
    textWatermarkOptions.fontFamily = "Arial";
    textWatermarkOptions.fontSize = 36;
    textWatermarkOptions.color = "#000000"; //Color.black;
    textWatermarkOptions.layout = aw.WatermarkLayout.Diagonal;
    textWatermarkOptions.isSemitrasparent = false;
    doc.watermark.setText("Aspose Watermark", textWatermarkOptions);
    doc.save(base.artifactsDir + "Document.TextWatermark.docx");
    // We can remove a watermark from a document like this.
    if (doc.watermark.type == aw.WatermarkType.Text)
        doc.watermark.remove();
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.TextWatermark.docx");
    expect(doc.watermark.type).toEqual(aw.WatermarkType.Text);
  });
  
  test('ImageWatermark', () => {
    //ExStart
    //ExFor:Watermark.setImage(Image)
    //ExFor:Watermark.setImage(Image, ImageWatermarkOptions)
    //ExFor:Watermark.setImage(String, ImageWatermarkOptions)
    //ExFor:ImageWatermarkOptions
    //ExFor:ImageWatermarkOptions.scale
    //ExFor:ImageWatermarkOptions.isWashout
    //ExSummary:Shows how to create a watermark from an image in the local file system.
    let doc = new aw.Document();
    // Modify the image watermark's appearance with an ImageWatermarkOptions object,
    // then pass it while creating a watermark from an image file.
    let imageWatermarkOptions = new aw.ImageWatermarkOptions();
    imageWatermarkOptions.scale = 5;
    imageWatermarkOptions.isWashout = false;
    let image = new aw.JSImage(base.imageDir + "Logo.jpg");
    doc.watermark.setImage(image, imageWatermarkOptions);
    doc.save(base.artifactsDir + "Document.ImageWatermark.docx");
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.ImageWatermark.docx");
    expect(doc.watermark.type).toEqual(aw.WatermarkType.Image);
  });
  
  test.each([false, true])('SpellingAndGrammarErrors', (showErrors) => {
    //ExStart
    //ExFor:Document.showGrammaticalErrors
    //ExFor:Document.showSpellingErrors
    //ExSummary:Shows how to show/hide errors in the document.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    // Insert two sentences with mistakes that would be picked up
    // by the spelling and grammar checkers in Microsoft Word.
    builder.writeln("There is a speling error in this sentence.");
    builder.writeln("Their is a grammatical error in this sentence.");
    // If these options are enabled, then spelling errors will be underlined
    // in the output document by a jagged red line, and a double blue line will highlight grammatical mistakes.
    doc.showGrammaticalErrors = showErrors;
    doc.showSpellingErrors = showErrors;
    doc.save(base.artifactsDir + "Document.SpellingAndGrammarErrors.docx");
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.SpellingAndGrammarErrors.docx");
    expect(doc.showGrammaticalErrors).toEqual(showErrors);
    expect(doc.showSpellingErrors).toEqual(showErrors);
  });
  
  test('IgnorePrinterMetrics', () => {
    //ExStart
    //ExFor:LayoutOptions.ignorePrinterMetrics
    //ExSummary:Shows how to ignore 'Use printer metrics to lay out document' option.
    let doc = new aw.Document(base.myDir + "Rendering.docx");
    doc.layoutOptions.ignorePrinterMetrics = false;
    doc.save(base.artifactsDir + "Document.ignorePrinterMetrics.docx");
    //ExEnd
  });
  
  test('ExtractPages', () => {
    //ExStart
    //ExFor:Document.extractPages
    //ExSummary:Shows how to get specified range of pages from the document.
    let doc = new aw.Document(base.myDir + "Layout entities.docx");
    doc = doc.extractPages(0, 2);
    doc.save(base.artifactsDir + "Document.extractPages.docx");
    //ExEnd
    doc = new aw.Document(base.artifactsDir + "Document.extractPages.docx");
    expect(doc.pageCount).toEqual(2);
  });
  
  test.each([true, false])('SpellingOrGrammar', (checkSpellingGrammar) => {
    //ExStart
    //ExFor:Document.spellingChecked
    //ExFor:Document.grammarChecked
    //ExSummary:Shows how to set spelling or grammar verifying.
    let doc = new aw.Document();
    // The string with spelling errors.
    doc.firstSection.body.firstParagraph.runs.add(new aw.Run(doc, "The speeling in this documentz is all broked."));
    // Spelling/Grammar check start if we set properties to false.
    // We can see all errors in Microsoft Word via Review -> Spelling & Grammar.
    // Note that Microsoft Word does not start grammar/spell check automatically for DOC and RTF document format.
    doc.spellingChecked = checkSpellingGrammar;
    doc.grammarChecked = checkSpellingGrammar;
    doc.save(base.artifactsDir + "Document.SpellingOrGrammar.docx");
    //ExEnd
  });
  
  test('AllowEmbeddingPostScriptFonts - TODO: ', () => {
    //ExStart
    //ExFor:SaveOptions.allowEmbeddingPostScriptFonts
    //ExSummary:Shows how to save the document with PostScript font.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.font.name = "PostScriptFont";
    builder.writeln("Some text with PostScript font.");
    // Load the font with PostScript to use in the document.
    let otf = new aw.Fonts.MemoryFontSource(base.loadFileToArray(base.fontsDir + "AllegroOpen.otf"));
    doc.fontSettings = new aw.Fonts.FontSettings();
    doc.fontSettings.setFontsSources([otf]);
    // Embed TrueType fonts.
    doc.fontInfos.embedTrueTypeFonts = true;
    // Allow embedding PostScript fonts while embedding TrueType fonts.
    // Microsoft Word does not embed PostScript fonts, but can open documents with embedded fonts of this type.
    let saveOptions = aw.Saving.SaveOptions.createSaveOptions(aw.SaveFormat.Docx);
    saveOptions.allowEmbeddingPostScriptFonts = true;
    doc.save(base.artifactsDir + "Document.allowEmbeddingPostScriptFonts.docx", saveOptions);
    //ExEnd
  });
  
  test('Frameset', () => {
    //ExStart
    //ExFor:Document.frameset
    //ExFor:Frameset
    //ExFor:Frameset.frameDefaultUrl
    //ExFor:Frameset.isFrameLinkToFile
    //ExFor:Frameset.childFramesets
    //ExFor:FramesetCollection
    //ExFor:FramesetCollection.count
    //ExFor:FramesetCollection.item(Int32)
    //ExSummary:Shows how to access frames on-page.
    // Document contains several frames with links to other documents.
    let doc = new aw.Document(base.myDir + "Frameset.docx");
    // We can check the default URL (a web page URL or local document) or if the frame is an external resource.
    expect(doc.frameset.childFramesets.at(0).childFramesets.at(0).frameDefaultUrl)
      .toEqual("https://file-examples-com.github.io/uploads/2017/02/file-sample_100kB.docx");
    expect(doc.frameset.childFramesets.at(0).childFramesets.at(0).isFrameLinkToFile).toEqual(true);
    expect(doc.frameset.childFramesets.at(1).frameDefaultUrl).toEqual("Document.docx");
    expect(doc.frameset.childFramesets.at(1).isFrameLinkToFile).toEqual(false);
    // Change properties for one of our frames.
    doc.frameset.childFramesets.at(0).childFramesets.at(0).frameDefaultUrl =
        "https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx";
    doc.frameset.childFramesets.at(0).childFramesets.at(0).isFrameLinkToFile = false;
    //ExEnd
    doc = DocumentHelper.saveOpen(doc);
    expect(doc.frameset.childFramesets.at(0).childFramesets.at(0).frameDefaultUrl)
      .toEqual("https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/Examples/Data/Absolute%20position%20tab.docx");
    expect(doc.frameset.childFramesets.at(0).childFramesets.at(0).isFrameLinkToFile).toEqual(false);
  });
  
  test('OpenAzw', () => {
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Azw3 document.azw3");
    expect(info.loadFormat).toEqual(aw.LoadFormat.Azw3);
    let doc = new aw.Document(base.myDir + "Azw3 document.azw3");
    expect(doc.getText().includes("Hachette Book Group USA")).toEqual(true);
  });
  
  test('OpenEpub', () => {
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Epub document.epub");
    expect(info.loadFormat).toEqual(aw.LoadFormat.Epub);
    let doc = new aw.Document(base.myDir + "Epub document.epub");
    expect(doc.getText().includes("Down the Rabbit-Hole")).toEqual(true);
  });
  
  test('OpenXml', () => {
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Mail merge data - Customers.xml");
    expect(info.loadFormat).toEqual(aw.LoadFormat.Xml);
    let doc = new aw.Document(base.myDir + "Mail merge data - Purchase order.xml");
    expect(doc.getText().includes("Ellen Adams\r123 Maple Street")).toEqual(true);
  });
  
  test('MoveToStructuredDocumentTag', () => {
    //ExStart
    //ExFor:DocumentBuilder.moveToStructuredDocumentTag(int, int)
    //ExFor:DocumentBuilder.moveToStructuredDocumentTag(StructuredDocumentTag, int)
    //ExFor:DocumentBuilder.isAtEndOfStructuredDocumentTag
    //ExFor:DocumentBuilder.currentStructuredDocumentTag
    //ExSummary:Shows how to move cursor of DocumentBuilder inside a structured document tag.
    let doc = new aw.Document(base.myDir + "Structured document tags.docx");
    let builder = new aw.DocumentBuilder(doc);
    // There is a several ways to move the cursor:
    // 1 -  Move to the first character of structured document tag by index.
    builder.moveToStructuredDocumentTag(1, 1);
    // 2 -  Move to the first character of structured document tag by object.
    let tag = doc.getSdt(2, true);
    builder.moveToStructuredDocumentTag(tag, 1);
    builder.write(" New text.");
    expect(tag.getText().trim()).toEqual("R New text.ichText");
    // 3 -  Move to the end of the second structured document tag.
    builder.moveToStructuredDocumentTag(1, -1);
    expect(builder.isAtEndOfStructuredDocumentTag).toEqual(true);
    // Get currently selected structured document tag.
    builder.currentStructuredDocumentTag.color = "#008000";
    doc.save(base.artifactsDir + "Document.moveToStructuredDocumentTag.docx");
    //ExEnd
  });
  
  test('IncludeTextboxesFootnotesEndnotesInStat', () => {
    //ExStart
    //ExFor:Document.includeTextboxesFootnotesEndnotesInStat
    //ExSummary: Shows how to include or exclude textboxes, footnotes and endnotes from word count statistics.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Lorem ipsum");
    builder.insertFootnote(aw.Notes.FootnoteType.Footnote, "sit amet");
    // By default option is set to 'false'.
    doc.updateWordCount();
    // Words count without textboxes, footnotes and endnotes.
    expect(doc.builtInDocumentProperties.words).toEqual(2);
    doc.includeTextboxesFootnotesEndnotesInStat = true;
    doc.updateWordCount();
    // Words count with textboxes, footnotes and endnotes.
    expect(doc.builtInDocumentProperties.words).toEqual(4);
    //ExEnd
  });
  
  test('SetJustificationMode', () => {
    //ExStart
    //ExFor:Document.justificationMode
    //ExFor:JustificationMode
    //ExSummary:Shows how to manage character spacing control.
    let doc = new aw.Document(base.myDir + "Document.docx");
    let justificationMode = doc.justificationMode;
    if (justificationMode == aw.Settings.JustificationMode.Expand)
        doc.justificationMode = aw.Settings.JustificationMode.Compress;
    doc.save(base.artifactsDir + "Document.SetJustificationMode.docx");
    //ExEnd
  });
  
  test('PageIsInColor', () => {
    //ExStart
    //ExFor:PageInfo.colored
    //ExFor:Document.getPageInfo(Int32)
    //ExSummary:Shows how to check whether the page is in color or not.
    let doc = new aw.Document(base.myDir + "Document.docx");
    // Check that the first page of the document is not colored.
    expect(doc.getPageInfo(0).colored).toEqual(false);
    //ExEnd
  });
  
  test('InsertDocumentInline', () => {
    //ExStart:InsertDocumentInline
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:DocumentBuilder.insertDocumentInline(Document, ImportFormatMode, ImportFormatOptions)
    //ExSummary:Shows how to insert a document inline at the cursor position.
    let srcDoc = new aw.DocumentBuilder();
    srcDoc.write("[src content]");
    // Create destination document.
    let dstDoc = new aw.DocumentBuilder();
    dstDoc.write("Before ");
    dstDoc.insertNode(new aw.BookmarkStart(dstDoc.document, "src_place"));
    dstDoc.insertNode(new aw.BookmarkEnd(dstDoc.document, "src_place"));
    dstDoc.write(" after");
    expect(dstDoc.document.getText().trim()).toEqual("Before  after");
    // Insert source document into destination inline.
    dstDoc.moveToBookmark("src_place");
    dstDoc.insertDocumentInline(srcDoc.document, aw.ImportFormatMode.UseDestinationStyles, new aw.ImportFormatOptions());
    expect(dstDoc.document.getText().trim()).toEqual("Before [src content] after");
    //ExEnd:InsertDocumentInline
  });
  
  test.each([aw.SaveFormat.Doc, 
    aw.SaveFormat.Dot,
    aw.SaveFormat.Docx,
    aw.SaveFormat.Docm,
    aw.SaveFormat.Dotx,
    aw.SaveFormat.Dotm,
    aw.SaveFormat.FlatOpc,
    aw.SaveFormat.FlatOpcMacroEnabled,
    aw.SaveFormat.FlatOpcTemplate,
    aw.SaveFormat.FlatOpcTemplateMacroEnabled,
    aw.SaveFormat.Rtf,
    aw.SaveFormat.WordML,
    aw.SaveFormat.Pdf,
    aw.SaveFormat.Xps,
    aw.SaveFormat.XamlFixed,
    aw.SaveFormat.Svg,
    aw.SaveFormat.HtmlFixed,
    aw.SaveFormat.OpenXps,
    aw.SaveFormat.Ps,
    aw.SaveFormat.Pcl,
    aw.SaveFormat.Html,
    aw.SaveFormat.Mhtml,
    aw.SaveFormat.Epub,
    aw.SaveFormat.Azw3,
    aw.SaveFormat.Mobi,
    aw.SaveFormat.Odt,
    aw.SaveFormat.Ott,
    aw.SaveFormat.Text,
    aw.SaveFormat.XamlFlow,
    aw.SaveFormat.XamlFlowPack,
    aw.SaveFormat.Markdown,
    aw.SaveFormat.Xlsx,
    aw.SaveFormat.Tiff,
    aw.SaveFormat.Png,
    aw.SaveFormat.Bmp,
    aw.SaveFormat.Emf,
    aw.SaveFormat.Jpeg,
    aw.SaveFormat.Gif,
    aw.SaveFormat.Eps])('SaveDocumentToStream(%o)', (saveFormat) => {
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Lorem ipsum");
  
    let stream = new MemoryStream();  
    if (saveFormat == aw.SaveFormat.HtmlFixed)
    {
        let saveOptions = new aw.Saving.HtmlFixedSaveOptions();
        saveOptions.exportEmbeddedCss = true;
        saveOptions.exportEmbeddedFonts = true;
        saveOptions.saveFormat = saveFormat;
        doc.save(stream, saveOptions);
    }
    else if (saveFormat == aw.SaveFormat.XamlFixed)
    {
        let saveOptions = new aw.Saving.XamlFixedSaveOptions();
        saveOptions.resourcesFolder = base.artifactsDir;
        saveOptions.saveFormat = saveFormat;
        doc.save(stream, saveOptions);
    }
    else
        doc.save(stream, saveFormat);
  });
  
  test('HasMacros', () => {
    //ExStart:HasMacros
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:FileFormatInfo.hasMacros
    //ExSummary:Shows how to check VBA macro presence without loading document.
    let fileFormatInfo = aw.FileFormatUtil.detectFileFormat(base.myDir + "Macro.docm");
    expect(fileFormatInfo.hasMacros).toEqual(true);
    //ExEnd:HasMacros
  });

  
  test('PunctuationKerning', () => {
    //ExStart
    //ExFor:Document.punctuationKerning
    //ExSummary:Shows how to work with kerning applies to both Latin text and punctuation.
    let doc = new aw.Document(base.myDir + "Document.docx");
    expect(doc.punctuationKerning).toEqual(true);
    //ExEnd
  });


  test('RemoveBlankPages', () => {
    //ExStart
    //ExFor:Document.removeBlankPages
    //ExSummary:Shows how to remove blank pages from the document.
    let doc = new aw.Document(base.myDir + "Blank pages.docx");
    expect(doc.pageCount).toEqual(2);
    doc.removeBlankPages();
    doc.updatePageLayout();
    expect(doc.pageCount).toEqual(1);
    //ExEnd
  });


});
