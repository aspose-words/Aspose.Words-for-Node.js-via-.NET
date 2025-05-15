// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;

describe("ExBuildingBlocks", () => {

  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  /*//Commented
    //ExStart
    //ExFor:Document.GlossaryDocument
    //ExFor:BuildingBlock
    //ExFor:BuildingBlock.#ctor(GlossaryDocument)
    //ExFor:BuildingBlock.Accept(DocumentVisitor)
    //ExFor:BuildingBlock.AcceptStart(DocumentVisitor)
    //ExFor:BuildingBlock.AcceptEnd(DocumentVisitor)
    //ExFor:BuildingBlock.Behavior
    //ExFor:BuildingBlock.Category
    //ExFor:BuildingBlock.Description
    //ExFor:BuildingBlock.FirstSection
    //ExFor:BuildingBlock.Gallery
    //ExFor:BuildingBlock.Guid
    //ExFor:BuildingBlock.LastSection
    //ExFor:BuildingBlock.Name
    //ExFor:BuildingBlock.Sections
    //ExFor:BuildingBlock.Type
    //ExFor:BuildingBlockBehavior
    //ExFor:BuildingBlockType
    //ExSummary:Shows how to add a custom building block to a document.
  test('CreateAndInsert', () => {
    // A document's glossary document stores building blocks.
    let doc = new aw.Document();
    let glossaryDoc = new aw.BuildingBlocks.GlossaryDocument();
    doc.glossaryDocument = glossaryDoc;

    // Create a building block, name it, and then add it to the glossary document.
    let block = new aw.BuildingBlocks.BuildingBlock(glossaryDoc)
    {
      Name = "Custom Block"
    };

    glossaryDoc.appendChild(block);

    // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
    expect(block.guid.toString()).toEqual("00000000-0000-0000-0000-000000000000");

    block.guid = Guid.NewGuid();

    // The following properties categorize building blocks
    // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    expect(block.category).toEqual("(Empty Category)");
    expect(block.type).toEqual(aw.BuildingBlocks.BuildingBlockType.None);
    expect(block.gallery).toEqual(aw.BuildingBlocks.BuildingBlockGallery.All);
    expect(block.behavior).toEqual(aw.BuildingBlocks.BuildingBlockBehavior.Content);

    // Before we can add this building block to our document, we will need to give it some contents,
    // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
    let visitor = new BuildingBlockVisitor(glossaryDoc);
    // Visit start/end of the BuildingBlock.
    block.accept(visitor);

    // We can access the block that we just made from the glossary document.
    let customBlock = glossaryDoc.getBuildingBlock(aw.BuildingBlocks.BuildingBlockGallery.QuickParts,
      "My custom building blocks", "Custom Block");

    // The block itself is a section that contains the text.
    expect(customBlock.firstSection.body.firstParagraph.getText()).toEqual(`Text inside ${customBlock.name}\f`);
    expect(customBlock.lastSection).toEqual(customBlock.firstSection);
    Assert.DoesNotThrow(() => Guid.parse(customBlock.guid.toString())); //ExSkip
    expect(customBlock.category).toEqual("My custom building blocks");
    expect(customBlock.type).toEqual(aw.BuildingBlocks.BuildingBlockType.None);
    expect(customBlock.gallery).toEqual(aw.BuildingBlocks.BuildingBlockGallery.QuickParts);
    expect(customBlock.behavior).toEqual(aw.BuildingBlocks.BuildingBlockBehavior.Paragraph);

    // Now, we can insert it into the document as a new section.
    doc.appendChild(doc.importNode(customBlock.firstSection, true));

    // We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
    doc.save(base.artifactsDir + "BuildingBlocks.CreateAndInsert.dotx");
  });


    /// <summary>
    /// Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents.
    /// </summary>
  public class BuildingBlockVisitor : DocumentVisitor
  {
    public BuildingBlockVisitor(GlossaryDocument ownerGlossaryDoc)
    {
      mBuilder = new StringBuilder();
      mGlossaryDoc = ownerGlossaryDoc;
    }

    public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
    {
        // Configure the building block as a quick part, and add properties used by Building Blocks Organizer.
      block.behavior = aw.BuildingBlocks.BuildingBlockBehavior.Paragraph;
      block.category = "My custom building blocks";
      block.description =
        "Using this block in the Quick Parts section of word will place its contents at the cursor.";
      block.gallery = aw.BuildingBlocks.BuildingBlockGallery.QuickParts;

        // Add a section with text.
        // Inserting the block into the document will append this section with its child nodes at the location.
      let section = new aw.Section(mGlossaryDoc);
      block.appendChild(section);
      block.firstSection.ensureMinimum();

      let run = new aw.Run(mGlossaryDoc, "Text inside " + block.name);
      block.firstSection.body.firstParagraph.appendChild(run);

      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitBuildingBlockEnd(BuildingBlock block)
    {
      mBuilder.append("Visited " + block.name + "\r\n");
      return aw.VisitorAction.Continue;
    }

    private readonly StringBuilder mBuilder;
    private readonly GlossaryDocument mGlossaryDoc;
  }
    //ExEnd*/

  /*  //ExStart
    //ExFor:GlossaryDocument
    //ExFor:GlossaryDocument.Accept(DocumentVisitor)
    //ExFor:GlossaryDocument.AcceptStart(DocumentVisitor)
    //ExFor:GlossaryDocument.AcceptEnd(DocumentVisitor)
    //ExFor:GlossaryDocument.BuildingBlocks
    //ExFor:GlossaryDocument.FirstBuildingBlock
    //ExFor:GlossaryDocument.GetBuildingBlock(BuildingBlockGallery,String,String)
    //ExFor:GlossaryDocument.LastBuildingBlock
    //ExFor:BuildingBlockCollection
    //ExFor:BuildingBlockCollection.Item(Int32)
    //ExFor:BuildingBlockCollection.ToArray
    //ExFor:BuildingBlockGallery
    //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlock)
    //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlock)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(GlossaryDocument)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(GlossaryDocument)
    //ExSummary:Shows ways of accessing building blocks in a glossary document.
  test('GlossaryDocument', () => {
    let doc = new aw.Document();
    let glossaryDoc = new aw.BuildingBlocks.GlossaryDocument();

    let child1 = new aw.BuildingBlocks.BuildingBlock(glossaryDoc) { Name = "Block 1" };
    glossaryDoc.appendChild(child1);
    let child2 = new aw.BuildingBlocks.BuildingBlock(glossaryDoc) { Name = "Block 2" };
    glossaryDoc.appendChild(child2);
    let child3 = new aw.BuildingBlocks.BuildingBlock(glossaryDoc) { Name = "Block 3" };
    glossaryDoc.appendChild(child3);
    let child4 = new aw.BuildingBlocks.BuildingBlock(glossaryDoc) { Name = "Block 4" };
    glossaryDoc.appendChild(child4);
    let child5 = new aw.BuildingBlocks.BuildingBlock(glossaryDoc) { Name = "Block 5" };
    glossaryDoc.appendChild(child5);

    expect(glossaryDoc.buildingBlocks.count).toEqual(5);

    doc.glossaryDocument = glossaryDoc;

    // There are various ways of accessing building blocks.
    // 1 -  Get the first/last building blocks in the collection:
    expect(glossaryDoc.firstBuildingBlock.name).toEqual("Block 1");
    expect(glossaryDoc.lastBuildingBlock.name).toEqual("Block 5");

    // 2 -  Get a building block by index:
    expect(glossaryDoc.buildingBlocks.at(1).name).toEqual("Block 2");
    expect(glossaryDoc.buildingBlocks.toArray()[2].Name).toEqual("Block 3");

    // 3 -  Get the first building block that matches a gallery, name and category:
    expect(glossaryDoc.getBuildingBlock(aw.BuildingBlocks.BuildingBlockGallery.All, "(Empty Category)", "Block 4").Name).toEqual("Block 4");

    // We will do that using a custom visitor,
    // which will give every BuildingBlock in the GlossaryDocument a unique GUID
    let visitor = new GlossaryDocVisitor();
    // Visit start/end of the Glossary document.
    glossaryDoc.accept(visitor);
    // Visit only start of the Glossary document.
    glossaryDoc.acceptStart(visitor);
    // Visit only end of the Glossary document.
    glossaryDoc.acceptEnd(visitor);
    expect(visitor.GetDictionary().Count).toEqual(5);

    console.log(visitor.getText());

    // In Microsoft Word, we can access the building blocks via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    doc.save(base.artifactsDir + "BuildingBlocks.glossaryDocument.dotx"); 
  });


    /// <summary>
    /// Gives each building block in a visited glossary document a unique GUID.
    /// Stores the GUID-building block pairs in a dictionary.
    /// </summary>
  public class GlossaryDocVisitor : DocumentVisitor
  {
    public GlossaryDocVisitor()
    {
      mBlocksByGuid = new Dictionary<Guid, BuildingBlock>();
      mBuilder = new StringBuilder();
    }

    public string GetText()
    {
      return mBuilder.toString();
    }

    public Dictionary<Guid, BuildingBlock> GetDictionary()
    {
      return mBlocksByGuid;
    }

    public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossary)
    {
      mBuilder.AppendLine("Glossary document found!");
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossary)
    {
      mBuilder.AppendLine("Reached end of glossary!");
      mBuilder.AppendLine("BuildingBlocks found: " + mBlocksByGuid.count);
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitBuildingBlockStart(BuildingBlock block)
    {
      expect(block.guid.toString()).toEqual("00000000-0000-0000-0000-000000000000");
      block.guid = Guid.NewGuid();
      mBlocksByGuid.add(block.guid, block);
      return aw.VisitorAction.Continue;
    }

    public override VisitorAction VisitBuildingBlockEnd(BuildingBlock block)
    {
      mBuilder.AppendLine("\tVisited block \"" + block.name + "\"");
      mBuilder.AppendLine("\t Type: " + block.type);
      mBuilder.AppendLine("\t Gallery: " + block.gallery);
      mBuilder.AppendLine("\t Behavior: " + block.behavior);
      mBuilder.AppendLine("\t Description: " + block.description);

      return aw.VisitorAction.Continue;
    }

    private readonly Dictionary<Guid, BuildingBlock> mBlocksByGuid;
    private readonly StringBuilder mBuilder;
  }
  //ExEnd
  //Commented */
});
