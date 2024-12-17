# Aspose.Words for Node.js via .NET examples and showcases

## Node.js Document Processing API
`Aspose.Words for Node.js` is a native library that offers developers a wealth of features to create, edit, and convert Word, PDF, Web documents, without the need for Microsoft Word environment to be installed on the system. This Node.js is a collection of classes and methods that rely on the Document Object Model (DOM), giving developers direct access to a document's inner workings at the element level. Using our product, Node.js developers can efficiently create complex documents and modify their formatting, layout, and content. This Node.js API is a reliable document processing solution for developers seeking a comprehensive instrument to streamline their document editing and document generation tasks; automate document-intensive business processes at scale; reduce manual intervention, errors, and delays.
The API is implemented as a native Node.js module, utilizing [Node-API](https://nodejs.org/api/n-api.html), which allows for maximum processing speed when handling large documents.

## Supported Formats

### Read and Write Formats
- **Microsoft Word┬о:** DOC, DOT, DOCX, DOTX, DOTM, FlatOpc, FlatOpcMacroEnabled, FlatOpcTemplate, FlatOpcTemplateMacroEnabled, RTF, Microsoft Word 2003 WordprocessingML
- **OpenDocument:** ODT, OTT
- **Web:** HTML, MHTML
- **Markdown:** MD
- **Text:** TXT
- **eBook:** AZW3, EPUB, MOBI, CHM

### Read-Only Formats
- **Microsoft Word┬о:** DocPreWord60
- **Other:** XML (XML Document)

### Write-Only Formats
- **Fixed Layout:** PDF, XPS, OpenXps
- **PostScript:** PS, EPS
- **Printer:** PCL
- **Markup:** XamlFixed, HtmlFixed, XamlFlow, XamlFlowPack
- **Image:** SVG, TIFF, PNG, BMP, JPEG, GIF, WEBP
- **Metafile:** EMF
- **Other:** XLSX

## Functionality
- Provides comprehensive document import and export with 35+ supported file formats. This allows developers to convert documents from one file format to another. For example, you can convert HTML to Word and Word to PDF documents with professional quality.
- Provides full access to all Word and OpenOffice document elements, including formatting properties and styling.
- Provides high-fidelity rendering of Word documents to PDF, JPG, PNG and other imaging formats.
- Provides the ability to print OpenOffice and Word documents programmatically.
- Provides a rich set of utility functions, you can use to split a document into parts, join documents together, compare documents, and much more.
- To become familiar with the most popular Aspose.Words functionality, please have a look at our [free online applications](https://products.aspose.app/words/family).


## Getting Started with Aspose.Words for Node.js
Simply execute `npm install @aspose/words` to get the latest version & try any of the following code snippets.

### Create a DOCX using Node.js
Aspose.Words for Node.js allows you to create a blank Word document and add content to the file.
```js
const aw = require('@aspose/words');

// Create a Word document.
var doc = new aw.Document();

// Use a DocumentBuilder instance to add content to the file.
var builder = new aw.DocumentBuilder(doc);

// Write a new paragraph to the document.
builder.writeln('This is an example of a Word document created in Node.js');

// Save it as a DOCX file. The output format is automatically determined by the filename extension.
doc.save('OutputWordDocument.docx');
```

### Convert a Word document to HTML with Node.js
You can convert Microsoft Word to PDF, XPS, Markdown, HTML, JPEG, TIFF, and other file formats. The following snippet demonstrates the conversion from DOCX to HTML:
```js
const aw = require('@aspose/words');

// Load a Word file from the local drive.
var doc = new aw.Document('InputWordDocument.docx');

// Save it to HTML format.
doc.save('OutputHtmlDocument.html');
```

## Licensing

### Evaluate Aspose.Words
You can use `Aspose.Words for Node.js` free of cost for evaluation. The evaluation version provides almost all functionality of the product with certain limitations. The same evaluation version becomes licensed when you purchase a license and add a couple of lines of code to apply the license.

If you want to test `Aspose.Words for Node.js` without evaluation version limitations, you can also request a 30 Day Temporary License. Please refer to [How to get a Temporary License?](https://purchase.aspose.com/temporary-license/)

### Evaluation Version Limitations 
Evaluation version of `Aspose.Words for Node.js` without the specified license provides full product functionality, but inserts an evaluative watermark at the top of the document upon loading and saving and limits the maximum document size to a few hundred paragraphs.

### About the License
The license is a plain text XML file that contains details such as the product name, number of developers it is licensed to, subscription expiry date and so on. The file is digitally signed, so donтАЩt modify the file. Even inadvertent addition of an extra line break into the file will invalidate it.

You need to set a license before utilizing `Aspose.Words for Node.js` if you want to avoid its evaluation limitations. It is only required to set a license once per application (or process).

### Apply License 
The easiest way to set a license, is to put the license file to the application folder and specify the file name without its path.
```js
// Instantiate an instance of License calss and set the license file through its path
const aw = require('@aspose/words');
const license = new aw.License();
license.setLicense("Aspose.Words.lic");
```

However, if the license is obtained from an external source (e.g., a database), you can use a [Buffer](https://nodejs.org/api/buffer.html#buffer):
```js
const aw = require('@aspose/words');
const fs = require('fs');

// For testing purposes, let's simply read the license data from a file.
const data = fs.readFileSync('Aspose.Words.lic');

const license = new aw.License();
license.setLicense(data);
```

## Supported Platforms
The first release only supports Microsoft Windows┬о x64 platform and Node.js 14.17.0 or higher.

[Product Page](https://products.aspose.com/words/nodejs-net/) | [Demos](https://products.aspose.app/words/family) | [Examples](https://github.com/alexey-butalov-aspose/Aspose.Words-for-Node.js-via-.NET/tree/main/ApiExamples) | [Blog](https://blog.aspose.com/category/words/) | [Search](https://search.aspose.com/) | [Free Support](https://forum.aspose.com/c/words) | [Temporary License](https://purchase.aspose.com/temporary-license)