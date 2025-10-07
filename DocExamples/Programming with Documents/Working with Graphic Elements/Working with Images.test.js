// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;

describe("WorkingWithImages", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('InsertBarcodeImage', () => {
    //ExStart:InsertBarcodeImage
    //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    // The number of pages the document should have
    const numPages = 4;
    // The document starts with one section, insert the barcode into this existing section
    insertBarcodeIntoFooter(builder, doc.firstSection, aw.HeaderFooterType.FooterPrimary);

    for (let i = 1; i < numPages; i++) {
      // Clone the first section and add it into the end of the document
      let cloneSection = doc.firstSection.clone();
      cloneSection.pageSetup.sectionStart = aw.SectionStart.NewPage;
      doc.appendChild(cloneSection);

      // Insert the barcode and other information into the footer of the section
      insertBarcodeIntoFooter(builder, cloneSection, aw.HeaderFooterType.FooterPrimary);
    }

    // Save the document as a PDF to disk
    // You can also save this directly to a stream
    doc.save(base.artifactsDir + "WorkingWithImages.InsertBarcodeImage.docx");
    //ExEnd:InsertBarcodeImage
  });

  //ExStart:InsertBarcodeIntoFooter
  //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
  function insertBarcodeIntoFooter(builder, section, footerType) {
    // Move to the footer type in the specific section.
    builder.moveToSection(section.document.indexOf(section));
    builder.moveToHeaderFooter(footerType);

    // Insert the barcode, then move to the next line and insert the ID along with the page number.
    // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
    builder.insertImage(base.imagesDir + "Barcode.png");
    builder.writeln();
    builder.write("1234567890");
    builder.insertField("PAGE");

    // Create a right-aligned tab at the right margin.
    let tabPos = section.pageSetup.pageWidth - section.pageSetup.rightMargin - section.pageSetup.leftMargin;
    builder.currentParagraph.paragraphFormat.tabStops.add(new aw.TabStop(tabPos, aw.TabAlignment.Right,
        aw.TabLeader.None));

    // Move to the right-hand side of the page and insert the page and page total.
    builder.write(aw.ControlChar.tab);
    builder.insertField("PAGE");
    builder.write(" of ");
    builder.insertField("NUMPAGES");
  }

  //ExEnd:InsertBarcodeIntoFooter

  test('CropImages', () => {
    //ExStart:CropImages
    //GistId:e2b8f833f9ab5de7c0598ddfd0ab1414
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let croppedImage = builder.insertImage(base.imagesDir + "Logo.jpg");

    let srcWidthPoints = croppedImage.width;
    let srcHeightPoints = croppedImage.height;

    croppedImage.width = aw.ConvertUtil.pixelToPoint(200);
    croppedImage.height = aw.ConvertUtil.pixelToPoint(200);

    let widthRatio = croppedImage.width / srcWidthPoints;
    let heightRatio = croppedImage.height / srcHeightPoints;

    if (widthRatio < 1) {
      croppedImage.imageData.cropRight = 1 - widthRatio;
    }

    if (heightRatio < 1) {
      croppedImage.imageData.cropBottom = 1 - heightRatio;
    }

    let leftToWidth = aw.ConvertUtil.pixelToPoint(100) / srcWidthPoints;
    let topToHeight = aw.ConvertUtil.pixelToPoint(90) / srcHeightPoints;

    croppedImage.imageData.cropLeft = leftToWidth;
    croppedImage.imageData.cropRight = croppedImage.imageData.cropRight - leftToWidth;

    croppedImage.imageData.cropTop = topToHeight;
    croppedImage.imageData.cropBottom = croppedImage.imageData.cropBottom - topToHeight;

    croppedImage.getShapeRenderer().save(base.artifactsDir + "WorkingWithImages.CropImages.jpg", new aw.Saving.ImageSaveOptions(aw.SaveFormat.Jpeg));
    //ExEnd:CropImages
  });
});
