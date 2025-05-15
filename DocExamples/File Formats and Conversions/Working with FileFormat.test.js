// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("WorkingWithFileFormat", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('DetectFileFormat', () => {
    //ExStart:CheckFormatCompatibility
    //GistId:eabbcbd1e117d4d628dfe4fd7c30321c
    const supportedDir = base.artifactsDir + "Supported";
    const unknownDir = base.artifactsDir + "Unknown";
    const encryptedDir = base.artifactsDir + "Encrypted";
    const pre97Dir = base.artifactsDir + "Pre97";

    // Create the directories if they do not already exist.
    if (!fs.existsSync(supportedDir))
      fs.mkdirSync(supportedDir);
    if (!fs.existsSync(unknownDir))
      fs.mkdirSync(unknownDir);
    if (!fs.existsSync(encryptedDir))
      fs.mkdirSync(encryptedDir);
    if (!fs.existsSync(pre97Dir))
      fs.mkdirSync(pre97Dir);

    //ExStart:GetFiles
    //GistId:eabbcbd1e117d4d628dfe4fd7c30321c
    let fileList = fs.readdirSync(base.myDir, { withFileTypes: true }).filter(
      f => f.isFile() && !f.name.endsWith("Corrupted document.docx")).map(f => f.name);
    for (let fileName of fileList) {
      console.log(fileName);
    //ExEnd:GetFiles

      let fullName = path.join(base.myDir, fileName);
      let info = aw.FileFormatUtil.detectFileFormat(fullName);

      // Display the document type
      switch (info.loadFormat) {
        case aw.LoadFormat.Doc:
          console.log("\tMicrosoft Word 97-2003 document.");
          break;
        case aw.LoadFormat.Dot:
          console.log("\tMicrosoft Word 97-2003 template.");
          break;
        case aw.LoadFormat.Docx:
          console.log("\tOffice Open XML WordprocessingML Macro-Free Document.");
          break;
        case aw.LoadFormat.Docm:
          console.log("\tOffice Open XML WordprocessingML Macro-Enabled Document.");
          break;
        case aw.LoadFormat.Dotx:
          console.log("\tOffice Open XML WordprocessingML Macro-Free Template.");
          break;
        case aw.LoadFormat.Dotm:
          console.log("\tOffice Open XML WordprocessingML Macro-Enabled Template.");
          break;
        case aw.LoadFormat.FlatOpc:
          console.log("\tFlat OPC document.");
          break;
        case aw.LoadFormat.Rtf:
          console.log("\tRTF format.");
          break;
        case aw.LoadFormat.WordML:
          console.log("\tMicrosoft Word 2003 WordprocessingML format.");
          break;
        case aw.LoadFormat.Html:
          console.log("\tHTML format.");
          break;
        case aw.LoadFormat.Mhtml:
          console.log("\tMHTML (Web archive) format.");
          break;
        case aw.LoadFormat.Odt:
          console.log("\tOpenDocument Text.");
          break;
        case aw.LoadFormat.Ott:
          console.log("\tOpenDocument Text Template.");
          break;
        case aw.LoadFormat.DocPreWord60:
          console.log("\tMS Word 6 or Word 95 format.");
          break;
        case aw.LoadFormat.Unknown:
          console.log("\tUnknown format.");
          break;
      }

      if (info.isEncrypted) {
        console.log("\tAn encrypted document.");
        fs.copyFileSync(fullName, path.join(encryptedDir, fileName));
      } else {
        switch (info.loadFormat) {
          case aw.LoadFormat.DocPreWord60:
            fs.copyFileSync(fullName, path.join(pre97Dir, fileName));
            break;
          case aw.LoadFormat.Unknown:
            fs.copyFileSync(fullName, path.join(unknownDir, fileName));
            break;
          default:
            fs.copyFileSync(fullName, path.join(supportedDir, fileName));
            break;
        }
      }
    }
    //ExEnd:CheckFormatCompatibility
  });


  test('DetectDocumentSignatures', () => {
    //ExStart:DetectDocumentSignatures
    //GistId:246abc8bf535665565cc872be9b805ac
    let fileName = path.join(base.myDir + "Digitally signed.docx");
    let info = aw.FileFormatUtil.detectFileFormat(fileName);

    if (info.hasDigitalSignature) {
      console.log(
        `Document ${fileName} has digital signatures, they will be lost if you open/save this document with Aspose.words.`);
    }
    //ExEnd:DetectDocumentSignatures
  });


  test('VerifyEncryptedDocument', () => {
    //ExStart:VerifyEncryptedDocument
    //GistId:50a58d2d88c2177a9a4888b5d0e4de81
    let info = aw.FileFormatUtil.detectFileFormat(base.myDir + "Encrypted.docx");
    console.log(info.isEncrypted);
    //ExEnd:VerifyEncryptedDocument
  });

});
