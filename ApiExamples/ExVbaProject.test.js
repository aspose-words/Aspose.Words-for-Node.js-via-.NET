// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('./ApiExampleBase').ApiExampleBase;


describe("ExVbaProject", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });

  test('CreateNewVbaProject', () => {
    //ExStart
    //ExFor:VbaProject.#ctor
    //ExFor:VbaProject.name
    //ExFor:VbaModule.#ctor
    //ExFor:VbaModule.name
    //ExFor:VbaModule.type
    //ExFor:VbaModule.sourceCode
    //ExFor:VbaModuleCollection.add(VbaModule)
    //ExFor:VbaModuleType
    //ExSummary:Shows how to create a VBA project using macros.
    let doc = new aw.Document();

    // Create a new VBA project.
    let project = new aw.Vba.VbaProject();
    project.name = "Aspose.project";
    doc.vbaProject = project;

    // Create a new module and specify a macro source code.
    let module = new aw.Vba.VbaModule();
    module.name = "Aspose.Module";
    module.type = aw.Vba.VbaModuleType.ProceduralModule;
    module.sourceCode = "New source code";

    // Add the module to the VBA project.
    doc.vbaProject.modules.add(module);

    doc.save(base.artifactsDir + "VbaProject.CreateVBAMacros.docm");
    //ExEnd

    project = new aw.Document(base.artifactsDir + "VbaProject.CreateVBAMacros.docm").vbaProject;

    expect(project.name).toEqual("Aspose.project");

    let modules = doc.vbaProject.modules;

    expect(modules.count).toEqual(2);

    expect(modules.at(0).name).toEqual("ThisDocument");
    expect(modules.at(0).type).toEqual(aw.Vba.VbaModuleType.DocumentModule);
    expect(modules.at(0).sourceCode).toBe(null);

    expect(modules.at(1).name).toEqual("Aspose.Module");
    expect(modules.at(1).type).toEqual(aw.Vba.VbaModuleType.ProceduralModule);
    expect(modules.at(1).sourceCode).toEqual("New source code");
  });


  test('CloneVbaProject', () => {
    //ExStart
    //ExFor:VbaProject.clone
    //ExFor:VbaModule.clone
    //ExSummary:Shows how to deep clone a VBA project and module.
    let doc = new aw.Document(base.myDir + "VBA project.docm");
    let destDoc = new aw.Document();

    let copyVbaProject = doc.vbaProject.clone();
    destDoc.vbaProject = copyVbaProject;

    // In the destination document, we already have a module named "Module1"
    // because we cloned it along with the project. We will need to remove the module.
    let oldVbaModule = destDoc.vbaProject.modules.at("Module1");
    let copyVbaModule = doc.vbaProject.modules.at("Module1").clone();
    destDoc.vbaProject.modules.remove(oldVbaModule);
    destDoc.vbaProject.modules.add(copyVbaModule);

    destDoc.save(base.artifactsDir + "VbaProject.CloneVbaProject.docm");
    //ExEnd

    let originalVbaProject = new aw.Document(base.artifactsDir + "VbaProject.CloneVbaProject.docm").vbaProject;

    expect(originalVbaProject.name).toEqual(copyVbaProject.name);
    expect(originalVbaProject.codePage).toEqual(copyVbaProject.codePage);
    expect(originalVbaProject.isSigned).toEqual(copyVbaProject.isSigned);
    expect(originalVbaProject.modules.count).toEqual(copyVbaProject.modules.count);

    for (let i = 0; i < originalVbaProject.modules.count; i++)
    {
      expect(originalVbaProject.modules.at(i).name).toEqual(copyVbaProject.modules.at(i).name);
      expect(originalVbaProject.modules.at(i).type).toEqual(copyVbaProject.modules.at(i).type);
      expect(originalVbaProject.modules.at(i).sourceCode).toEqual(copyVbaProject.modules.at(i).sourceCode);
    }
  });


  //ExStart
  //ExFor:VbaReference
  //ExFor:VbaReference.Type
  //ExFor:VbaReference.LibId
  //ExFor:VbaReferenceCollection
  //ExFor:VbaReferenceCollection.Item(Int32)
  //ExFor:VbaReferenceCollection.Count
  //ExFor:VbaReferenceCollection.RemoveAt(int)
  //ExFor:VbaReferenceCollection.Remove(VbaReference)
  //ExFor:VbaReferenceType
  //ExFor:VbaProject.References
  //ExSummary:Shows how to get/remove an element from the VBA reference collection.
  test('RemoveVbaReference', () => {
    const brokenPath = "X:\\broken.dll";
    let doc = new aw.Document(base.myDir + "VBA project.docm");
            
    let references = doc.vbaProject.references;
    expect(references.count).toEqual(5);
            
    for (let i = references.count - 1; i >= 0; i--)
    {
      let reference = doc.vbaProject.references.at(i);
      let path = getLibIdPath(reference);
                
      if (path == brokenPath)
        references.removeAt(i);
    }
    expect(references.count).toEqual(4);

    references.remove(references.at(1));
    expect(references.count).toEqual(3);

    doc.save(base.artifactsDir + "VbaProject.RemoveVbaReference.docm"); 
  });


  /// <summary>
  /// Returns string representing LibId path of a specified reference. 
  /// </summary>
  function getLibIdPath(reference)
  {
    switch (reference.type)
    {
      case aw.Vba.VbaReferenceType.Registered:
      case aw.Vba.VbaReferenceType.Original:
      case aw.Vba.VbaReferenceType.Control:
        return getLibIdReferencePath(reference.libId);
      case aw.Vba.VbaReferenceType.Project:
        return getLibIdProjectPath(reference.libId);
      default:
        throw new Error("Unknown reference type.");
    }
  }

  /// <summary>
  /// Returns path from a specified identifier of an Automation type library.
  /// </summary>
  function getLibIdReferencePath(libIdReference)
  {
    if (libIdReference != null)
    {
      let refParts = libIdReference.split('#');
      if (refParts.length > 3)
        return refParts.at(3);
    }

    return "";
  }

  /// <summary>
  /// Returns path from a specified identifier of an Automation type library.
  /// </summary>
  function getLibIdProjectPath(libIdProject)
  {
    return libIdProject != null ? libIdProject.substring(3) : "";
  }
  //ExEnd


  test('IsProtected', () => {
    //ExStart:IsProtected
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:VbaProject.isProtected
    //ExSummary:Shows whether the VbaProject is password protected.
    let doc = new aw.Document(base.myDir + "Vba protected.docm");
    expect(doc.vbaProject.isProtected).toEqual(true);
    //ExEnd:IsProtected
  });

});
