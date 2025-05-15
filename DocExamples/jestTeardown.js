// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const path = require('path');
const fs = require('fs');
const base = require('./DocExampleBase').DocExampleBase;

const teardown = () => {
    if (fs.existsSync(base.artifactsDir)) {
        fs.rmSync(base.artifactsDir, {recursive: true, force: true} );
    }
    console.log('\r\n======================================================================================');
    console.log(`Global teardown: Artifact dir '${base.artifactsDir}' has been removed.`);
    console.log(`To keep artifacts - edit jestTeardown.js file.`);
    console.log('==========================================================================================');
}
  
module.exports = teardown;