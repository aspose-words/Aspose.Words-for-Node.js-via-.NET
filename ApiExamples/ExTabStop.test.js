// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;
const TestUtil = require('./TestUtil');


describe("ExTabStop", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('AddTabStops', () => {
    //ExStart
    //ExFor:TabStopCollection.add(TabStop)
    //ExFor:TabStopCollection.add(Double, TabAlignment, TabLeader)
    //ExSummary:Shows how to add custom tab stops to a document.
    let doc = new aw.Document();
    let paragraph = doc.getParagraph(0, true);

    // Below are two ways of adding tab stops to a paragraph's collection of tab stops via the "ParagraphFormat" property.
    // 1 -  Create a "TabStop" object, and then add it to the collection:
    let tabStop = new aw.TabStop(aw.ConvertUtil.inchToPoint(3), aw.TabAlignment.Left, aw.TabLeader.Dashes);
    paragraph.paragraphFormat.tabStops.add(tabStop);

    // 2 -  Pass the values for properties of a new tab stop to the "Add" method:
    paragraph.paragraphFormat.tabStops.add(aw.ConvertUtil.millimeterToPoint(100), aw.TabAlignment.Left,
      aw.TabLeader.Dashes);

    // Add tab stops at 5 cm to all paragraphs.
    for (var node of doc.getChildNodes(aw.NodeType.Paragraph, true))
    {
      var para = node.asParagraph();
      para.paragraphFormat.tabStops.add(aw.ConvertUtil.millimeterToPoint(50), aw.TabAlignment.Left,
        aw.TabLeader.Dashes);
    }

    // Every "tab" character takes the builder's cursor to the location of the next tab stop.
    let builder = new aw.DocumentBuilder(doc);
    builder.writeln("Start\tTab 1\tTab 2\tTab 3\tTab 4");

    doc.save(base.artifactsDir + "TabStopCollection.AddTabStops.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "TabStopCollection.AddTabStops.docx");
    let tabStops = doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops;

    TestUtil.verifyTabStop(141.75, aw.TabAlignment.Left, aw.TabLeader.Dashes, false, tabStops.at(0));
    TestUtil.verifyTabStop(216.0, aw.TabAlignment.Left, aw.TabLeader.Dashes, false, tabStops.at(1));
    TestUtil.verifyTabStop(283.45, aw.TabAlignment.Left, aw.TabLeader.Dashes, false, tabStops.at(2));
  });


  test('TabStopCollection', () => {
    //ExStart            
    //ExFor:TabStop.#ctor(Double)
    //ExFor:TabStop.#ctor(Double,TabAlignment,TabLeader)
    //ExFor:TabStop.equals(TabStop)
    //ExFor:TabStop.isClear
    //ExFor:TabStopCollection
    //ExFor:TabStopCollection.after(Double)
    //ExFor:TabStopCollection.before(Double)
    //ExFor:TabStopCollection.clear
    //ExFor:TabStopCollection.count
    //ExFor:TabStopCollection.equals(TabStopCollection)
    //ExFor:TabStopCollection.equals(Object)
    //ExFor:TabStopCollection.getHashCode
    //ExFor:TabStopCollection.item(Double)
    //ExFor:TabStopCollection.item(Int32)
    //ExSummary:Shows how to work with a document's collection of tab stops.
    let doc = new aw.Document();
    let builder = new aw.DocumentBuilder(doc);

    let tabStops = builder.paragraphFormat.tabStops;

    // 72 points is one "inch" on the Microsoft Word tab stop ruler.
    tabStops.add(new aw.TabStop(72.0));
    tabStops.add(new aw.TabStop(432.0, aw.TabAlignment.Right, aw.TabLeader.Dashes));

    expect(tabStops.count).toEqual(2);
    expect(tabStops.at(0).isClear).toEqual(false);
    expect(tabStops.at(0).equals(tabStops.at(1))).toEqual(false);

    // Every "tab" character takes the builder's cursor to the location of the next tab stop.
    builder.writeln("Start\tTab 1\tTab 2");

    let paragraphs = doc.firstSection.body.paragraphs;

    expect(paragraphs.count).toEqual(2);

    // Each paragraph gets its tab stop collection, which clones its values from the document builder's tab stop collection.
    expect(paragraphs.at(1).paragraphFormat.tabStops).toEqual(paragraphs.at(0).paragraphFormat.tabStops);

    // A tab stop collection can point us to TabStops before and after certain positions.
    expect(tabStops.before(100.0).position).toEqual(72.0);
    expect(tabStops.after(100.0).position).toEqual(432.0);

    // We can clear a paragraph's tab stop collection to revert to the default tabbing behavior.
    paragraphs.at(1).paragraphFormat.tabStops.clear();

    expect(paragraphs.at(1).paragraphFormat.tabStops.count).toEqual(0);

    doc.save(base.artifactsDir + "TabStopCollection.TabStopCollection.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "TabStopCollection.TabStopCollection.docx");
    tabStops = doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops;

    expect(tabStops.count).toEqual(2);
    TestUtil.verifyTabStop(72.0, aw.TabAlignment.Left, aw.TabLeader.None, false, tabStops.at(0));
    TestUtil.verifyTabStop(432.0, aw.TabAlignment.Right, aw.TabLeader.Dashes, false, tabStops.at(1));

    tabStops = doc.firstSection.body.paragraphs.at(1).paragraphFormat.tabStops;

    expect(tabStops.count).toEqual(0);
  });


  test('RemoveByIndex', () => {
    //ExStart
    //ExFor:TabStopCollection.removeByIndex
    //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
    let doc = new aw.Document();
    let tabStops = doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops;

    tabStops.add(aw.ConvertUtil.millimeterToPoint(30), aw.TabAlignment.Left, aw.TabLeader.Dashes);
    tabStops.add(aw.ConvertUtil.millimeterToPoint(60), aw.TabAlignment.Left, aw.TabLeader.Dashes);

    expect(tabStops.count).toEqual(2);

    // Remove the first tab stop.
    tabStops.removeByIndex(0);

    expect(tabStops.count).toEqual(1);

    doc.save(base.artifactsDir + "TabStopCollection.removeByIndex.docx");
    //ExEnd

    doc = new aw.Document(base.artifactsDir + "TabStopCollection.removeByIndex.docx");

    TestUtil.verifyTabStop(170.1, aw.TabAlignment.Left, aw.TabLeader.Dashes, false, doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops.at(0));
  });


  test('GetPositionByIndex', () => {
    //ExStart
    //ExFor:TabStopCollection.getPositionByIndex
    //ExSummary:Shows how to find a tab, stop by its index and verify its position.
    let doc = new aw.Document();
    let tabStops = doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops;

    tabStops.add(aw.ConvertUtil.millimeterToPoint(30), aw.TabAlignment.Left, aw.TabLeader.Dashes);
    tabStops.add(aw.ConvertUtil.millimeterToPoint(60), aw.TabAlignment.Left, aw.TabLeader.Dashes);

    // Verify the position of the second tab stop in the collection.
    expect(tabStops.getPositionByIndex(1)).toBeCloseTo(aw.ConvertUtil.millimeterToPoint(60), 1);
    //ExEnd
  });


  test('GetIndexByPosition', () => {
    //ExStart
    //ExFor:TabStopCollection.getIndexByPosition
    //ExSummary:Shows how to look up a position to see if a tab stop exists there and obtain its index.
    let doc = new aw.Document();
    let tabStops = doc.firstSection.body.paragraphs.at(0).paragraphFormat.tabStops;

    // Add a tab stop at a position of 30mm.
    tabStops.add(aw.ConvertUtil.millimeterToPoint(30), aw.TabAlignment.Left, aw.TabLeader.Dashes);

    // A result of "0" returned by "GetIndexByPosition" confirms that a tab stop
    // at 30mm exists in this collection, and it is at index 0.
    expect(tabStops.getIndexByPosition(aw.ConvertUtil.millimeterToPoint(30))).toEqual(0);

    // A "-1" returned by "GetIndexByPosition" confirms that
    // there is no tab stop in this collection with a position of 60mm.
    expect(tabStops.getIndexByPosition(aw.ConvertUtil.millimeterToPoint(60))).toEqual(-1);
    //ExEnd
  });


});
