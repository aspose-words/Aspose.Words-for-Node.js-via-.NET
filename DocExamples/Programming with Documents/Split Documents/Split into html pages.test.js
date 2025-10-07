// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;
const fs = require('fs');
const path = require('path');


describe("SplitIntoHtmlPages", () => {
  beforeAll(() => {
    base.oneTimeSetup();
  });

  afterAll(() => {
    base.oneTimeTearDown();
  });


  test('HtmlPages', () => {
    let srcFileName = base.myDir + "Big document.docx";

    let outDir = path.join(base.artifactsDir, "HtmlPages");
    if (!fs.existsSync(outDir)) {
      fs.mkdirSync(outDir, {recursive: true});
    }

    let w = new WordToHtmlConverter();
    w.execute(srcFileName, outDir);
  });

});

class WordToHtmlConverter {
  /// <summary>
  /// Performs the Word to HTML conversion.
  /// </summary>
  /// <param name="srcFileName">The MS Word file to convert.</param>
  /// This file needs to have a mail merge region called "TOC" defined and one mail merge field called "TocEntry".</param>
  /// <param name="dstDir">The output directory where to write HTML files.</param>
  execute(srcFileName, dstDir) {
    this.mDoc = new aw.Document(srcFileName);
    this.mDstDir = dstDir;

    let topicStartParas = this.selectTopicStarts();
    this.insertSectionBreaks(topicStartParas);
    this.saveHtmlTopics();
  }

  /// <summary>
  /// Selects heading paragraphs that must become topic starts.
  /// We can't modify them in this loop, so we need to remember them in an array first.
  /// </summary>
  selectTopicStarts() {
    let paras = this.mDoc.getChildNodes(aw.NodeType.Paragraph, true);
    let topicStartParas = [];

    for (let para of paras) {
      para = para.asParagraph();
      let style = para.paragraphFormat.styleIdentifier;
      if (style == aw.StyleIdentifier.Heading1) {
        topicStartParas.push(para);
      }
    }

    return topicStartParas;
  }

  //ExStart:InsertSectionBreaks
  //GistId:5331edc61a2137fd92565f1e0c953887
  /// <summary>
  /// Insert section breaks before the specified paragraphs.
  /// </summary>
  insertSectionBreaks(topicStartParas) {
    let builder = new aw.DocumentBuilder(this.mDoc);
    for (let para of topicStartParas) {
      let section = para.parentSection.asSection();

      // Insert section break if the paragraph is not at the beginning of a section already.
      if (!base.compareNodes(para, section.body.firstParagraph)) {
        builder.moveTo(para.firstChild);
        builder.insertBreak(aw.BreakType.SectionBreakNewPage);

        // This is the paragraph that was inserted at the end of the now old section.
        // We don't really need the extra paragraph, we just needed the section.
        section.body.lastParagraph.remove();
      }
    }
  }

  //ExEnd:InsertSectionBreaks

  /// <summary>
  /// Splits the current document into one topic per section and saves each topic
  /// as an HTML file. Returns a collection of Topic objects.
  /// </summary>
  saveHtmlTopics() {
    let topics = [];
    for (let sectionIdx = 0; sectionIdx < this.mDoc.sections.count; sectionIdx++) {
      let section = this.mDoc.sections.at(sectionIdx).asSection();

      let paraText = section.body.firstParagraph.getText();

      // Use the text of the heading paragraph to generate the HTML file name.
      let fileName = this.makeTopicFileName(paraText);
      if (fileName == "") {
        fileName = "UNTITLED SECTION " + sectionIdx;
      }

      fileName = path.join(this.mDstDir, fileName + ".html");

      // Use the text of the heading paragraph to generate the title for the TOC.
      let title = this.makeTopicTitle(paraText);
      if (title == "") {
        title = "UNTITLED SECTION " + sectionIdx;
      }

      let topic = new Topic(title, fileName);
      topics.push(topic);

      this.saveHtmlTopic(section, topic);
    }

    return topics;
  }

  /// <summary>
  /// Leaves alphanumeric characters, replaces white space with underscore
  /// And removes all other characters from a string.
  /// </summary>
  makeTopicFileName(paraText) {
    let b = "";
    for (let c of paraText) {
      if (/[a-zA-Z0-9]/.test(c)) {
        b += c;
      } else if (c == ' ') {
        b += '_';
      }
    }

    return b;
  }

  /// <summary>
  /// Removes the last character (which is a paragraph break character from the given string).
  /// </summary>
  makeTopicTitle(paraText) {
    return paraText.substring(0, paraText.length - 1);
  }

  /// <summary>
  /// Saves one section of a document as an HTML file.
  /// Any embedded images are saved as separate files in the same folder as the HTML file.
  /// </summary>
  saveHtmlTopic(section, topic) {
    let dummyDoc = new aw.Document();
    dummyDoc.removeAllChildren();
    dummyDoc.appendChild(dummyDoc.importNode(section, true, aw.ImportFormatMode.KeepSourceFormatting));

    dummyDoc.builtInDocumentProperties.title = topic.title;

    let saveOptions = new aw.Saving.HtmlSaveOptions();
    saveOptions.prettyFormat = true;
    saveOptions.allowNegativeIndent = true; // This is to allow headings to appear to the left of the main text.
    saveOptions.exportHeadersFootersMode = aw.Saving.ExportHeadersFootersMode.None;

    dummyDoc.save(topic.fileName, saveOptions);
  }
}

class Topic {
    constructor(title, fileName) {
        this.title = title;
        this.fileName = fileName;
    }
}
