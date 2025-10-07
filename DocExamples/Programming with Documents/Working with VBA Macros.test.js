// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../DocExampleBase').DocExampleBase;


describe("WorkingWithVba", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('CreateVbaProject', () => {
    //ExStart:CreateVbaProject
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document();

    let project = new aw.Vba.VbaProject();
    project.name = "AsposeProject";
    doc.vbaProject = project;

    // Create a new module and specify a macro source code.
    let module = new aw.Vba.VbaModule();
    module.name = "AsposeModule";
    module.type = aw.Vba.VbaModuleType.ProceduralModule;
    module.sourceCode = "New source code";

    // Add module to the VBA project.
    doc.vbaProject.modules.add(module);

    doc.save(base.artifactsDir + "WorkingWithVba.CreateVbaProject.docm");
    //ExEnd:CreateVbaProject
  });

  test('ReadVbaMacros', () => {
    //ExStart:ReadVbaMacros
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document(base.myDir + "VBA project.docm");

    if (doc.vbaProject != null) {
      let modules = doc.vbaProject.modules;
      for (let i = 0; i < modules.count; i++) {
        console.log(modules.at(i).sourceCode);
      }
    }
    //ExEnd:ReadVbaMacros
  });

  test('ModifyVbaMacros', () => {
    //ExStart:ModifyVbaMacros
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document(base.myDir + "VBA project.docm");

    let project = doc.vbaProject;

    let newSourceCode = "Test change source code";
    project.modules.at(0).sourceCode = newSourceCode;
    //ExEnd:ModifyVbaMacros

    doc.save(base.artifactsDir + "WorkingWithVba.ModifyVbaMacros.docm");
    //ExEnd:ModifyVbaMacros
  });

  test('CloneVbaProject', () => {
    //ExStart:CloneVbaProject
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document(base.myDir + "VBA project.docm");
    let destDoc = new aw.Document();
    destDoc.vbaProject = doc.vbaProject.clone();

    destDoc.save(base.artifactsDir + "WorkingWithVba.CloneVbaProject.docm");
    //ExEnd:CloneVbaProject
  });

  test('CloneVbaModule', () => {
    //ExStart:CloneVbaModule
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document(base.myDir + "VBA project.docm");
    let destDoc = new aw.Document();
    destDoc.vbaProject = new aw.Vba.VbaProject();

    let copyModule = doc.vbaProject.modules.at("Module1").clone();
    destDoc.vbaProject.modules.add(copyModule);

    destDoc.save(base.artifactsDir + "WorkingWithVba.CloneVbaModule.docm");
    //ExEnd:CloneVbaModule
  });

  test('RemoveVbaReferences', () => {
    //ExStart:RemoveVbaReferences
    //GistId:65a1b9bae9592a992d97821378084e93
    let doc = new aw.Document(base.myDir + "VBA project.docm");

    // Find and remove the reference with some LibId path.
    let brokenPath = "brokenPath.dll";
    let references = doc.vbaProject.references;
    for (let i = references.count - 1; i >= 0; i--) {
      let reference = doc.vbaProject.references.at(i);

      let path = getLibIdPath(reference);
      if (path == brokenPath)
        references.removeAt(i);
    }

    doc.save(base.artifactsDir + "WorkingWithVba.RemoveVbaReferences.docm");
    //ExEnd:RemoveVbaReferences
  });

  //ExStart:GetLibIdAndReferencePath
  //GistId:65a1b9bae9592a992d97821378084e93
  /// <summary>
  /// Returns string representing LibId path of a specified reference.
  /// </summary>
  function getLibIdPath(reference) {
    switch (reference.type) {
      case aw.Vba.VbaReferenceType.Registered:
      case aw.Vba.VbaReferenceType.Original:
      case aw.Vba.VbaReferenceType.Control:
        return getLibIdReferencePath(reference.libId);
      case aw.Vba.VbaReferenceType.Project:
        return getLibIdProjectPath(reference.libId);
      default:
        throw new Error("ArgumentOutOfRangeException");
    }
  }

  /// <summary>
  /// Returns path from a specified identifier of an Automation type library.
  /// </summary>
  /// <remarks>
  /// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference.
  /// </remarks>
  function getLibIdReferencePath(libIdReference) {
    if (libIdReference != null) {
      let refParts = libIdReference.split('#');
      if (refParts.length > 3)
        return refParts[3];
    }

    return "";
  }

  /// <summary>
  /// Returns path from a specified identifier of an Automation type library.
  /// </summary>
  /// <remarks>
  /// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference.
  /// </remarks>
  function getLibIdProjectPath(libIdProject) {
    return (libIdProject != null) ? libIdProject.substring(3) : "";
  }
  //ExEnd:GetLibIdAndReferencePath

});