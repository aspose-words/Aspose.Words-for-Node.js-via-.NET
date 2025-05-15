// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const aw = require('@aspose/words');

class DocExampleBase {
    static oneTimeSetup() {
        this.setUnlimitedLicense();
        if (!fs.existsSync(this.artifactsDir)) {
            try {
              fs.mkdirSync(this.artifactsDir);
            } catch {
            };
        }
    }

    static oneTimeTearDown() {
    }

    static setUnlimitedLicense()
    {
        // This is where the test license is on my development machine.
        const testLicenseFileName = path.join(this.licenseDir, "Aspose.Words.NodeJs.NET.lic");
    
        // This shows how to use an Aspose.Words license when you have purchased one.
        // You don't have to specify full path as shown here. You can specify just the 
        // file name if you copy the license file into the same folder as your application
        // binaries or you add the license to your project as an embedded resource.
        if (fs.existsSync(testLicenseFileName)) {
          const wordsLicense = new aw.License();
          wordsLicense.setLicense(testLicenseFileName);
        } else {
            console.log(`License file ${testLicenseFileName} not found. Some tests may fail.`);
        }
    }

    static loadFileToBuffer(fileName)
    {
        return fs.readFileSync(fileName);
    }

    static loadFileToArray(fileName)
    {
        return [...Uint8Array.from(fs.readFileSync(fileName))];
    }

    /// <summary>
    /// Gets the path to the codebase directory.
    /// </summary>
    static mainDataDir;
    /// <summary>
    /// Gets the path to the license used by the code examples.
    /// </summary>
    static licenseDir;
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
     static artifactsDir;
    /// <summary>
    /// Gets the path to the documents used by the code examples. Ends with a back slash.
    /// </summary>
    static myDir;
    /// <summary>
    /// Gets the path to the images used by the code examples. Ends with a back slash.
    /// </summary>
    static imagesDir;
    /// <summary>
    /// Gets the path of the demo database. Ends with a back slash.
    /// </summary>
    static databaseDir;
    /// <summary>
    /// Gets the path of the free fonts. Ends with a back slash.
    /// </summary>
    static fontsDir;

    static emptyColor = "";
    
    static {
        this.codeBaseDir = __dirname;
        const dataDir =  path.join(this.codeBaseDir, "..", "Data");
        this.artifactsDir = path.join(dataDir, "Artifacts") + path.sep;
        this.licenseDir = path.join(dataDir, "License") + path.sep;
        this.myDir = dataDir + path.sep;
        this.imagesDir = path.join(dataDir, "Images") + path.sep;
        this.databaseDir = path.join(dataDir, "Database") + path.sep;
        this.fontsDir = path.join(dataDir, "MyFonts") + path.sep;
    }
}

module.exports = { DocExampleBase };