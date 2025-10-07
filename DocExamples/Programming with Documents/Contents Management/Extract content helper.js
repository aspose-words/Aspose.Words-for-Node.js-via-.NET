// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

const aw = require('@aspose/words');
const base = require('../../DocExampleBase').DocExampleBase;

class ExtractContentHelper {
    static extractContent(startNode, endNode, isInclusive, copySection) {
        // First, check that the nodes passed to this method are valid for use.
        ExtractContentHelper.verifyParameterNodes(startNode, endNode);

        // Create a list to store the extracted nodes.
        let nodes = [];

        // If either marker is part of a comment, including the comment itself, we need to move the pointer
        // forward to the Comment Node found after the CommentRangeEnd node.
        if (endNode.nodeType === aw.NodeType.CommentRangeEnd && isInclusive) {
            let node = ExtractContentHelper.findNextNode(aw.NodeType.Comment, endNode.nextSibling);
            if (node != null)
                endNode = node;
        }

        // Keep a record of the original nodes passed to this method to split marker nodes if needed.
        let originalStartNode = startNode;
        let originalEndNode = endNode;

        // Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        // We will split the first and last nodes' content, depending if the marker nodes are inline.
        startNode = ExtractContentHelper.getAncestorInBody(startNode);
        endNode = ExtractContentHelper.getAncestorInBody(endNode);

        let isExtracting = true;
        let isStartingNode = true;
        // The current node we are extracting from the document.
        let currNode = startNode;

        // Begin extracting content. Process all block-level nodes and specifically split the first
        // and last nodes when needed, so paragraph formatting is retained.
        // Method is a little more complicated than a regular extractor as we need to factor
        // in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
        while (isExtracting) {
            if (copySection) {
                let section = currNode.getAncestor(aw.NodeType.Section);
                if (!nodes.some(o => o.range.text === section.range.text))
                    nodes.push(section.clone(true));
            }

            // Clone the current node and its children to obtain a copy.
            let cloneNode = currNode.clone(true);
            let isEndingNode = base.compareNodes(currNode, endNode);

            if (isStartingNode || isEndingNode) {
                // We need to process each marker separately, so pass it off to a separate method instead.
                // End should be processed at first to keep node indexes.
                if (isEndingNode) {
                    // !isStartingNode: don't add the node twice if the markers are the same node.
                    ExtractContentHelper.processMarker(cloneNode, nodes, originalEndNode, currNode, isInclusive,
                        false, !isStartingNode, false);
                    isExtracting = false;
                }

                // Conditional needs to be separate as the block level start and end markers, maybe the same node.
                if (isStartingNode) {
                    ExtractContentHelper.processMarker(cloneNode, nodes, originalStartNode, currNode, isInclusive,
                        true, true, false);
                    isStartingNode = false;
                }
            } else
                // Node is not a start or end marker, simply add the copy to the list.
                nodes.push(cloneNode);

            // Move to the next node and extract it. If the next node is null,
            // the rest of the content is found in a different section.
            if (currNode.nextSibling == null && isExtracting) {
                // Move to the next section.
                let nextSection = currNode.getAncestor(aw.NodeType.Section).nextSibling.asSection();
                currNode = nextSection.body.firstChild;
            } else
                // Move to the next node in the body.
                currNode = currNode.nextSibling;
        }

        // For compatibility with mode with inline bookmarks, add the next paragraph (empty).
        if (isInclusive && originalEndNode === endNode && !originalEndNode.isComposite)
            ExtractContentHelper.includeNextParagraph(endNode, nodes);

        // Return the nodes between the node markers.
        return nodes;
    }

    static verifyParameterNodes(startNode, endNode) {
        // The order in which these checks are done is important.
        if (startNode == null)
            throw new Error("Start node cannot be null");
        if (endNode == null)
            throw new Error("End node cannot be null");

        if (!base.compareNodes(startNode.document, endNode.document))
            throw new Error("Start node and end node must belong to the same document");

        if (startNode.getAncestor(aw.NodeType.Body) == null || endNode.getAncestor(aw.NodeType.Body) == null)
            throw new Error("Start node and end node must be a child or descendant of a body");

        // Check the end node is after the start node in the DOM tree.
        // First, check if they are in different sections, then if they're not,
        // check their position in the body of the same section.
        let startSection = startNode.getAncestor(aw.NodeType.Section).asSection();
        let endSection = endNode.getAncestor(aw.NodeType.Section).asSection();

        let startIndex = startSection.parentNode.indexOf(startSection);
        let endIndex = endSection.parentNode.indexOf(endSection);

        if (startIndex === endIndex) {
            if (startSection.body.indexOf(ExtractContentHelper.getAncestorInBody(startNode)) >
                endSection.body.indexOf(ExtractContentHelper.getAncestorInBody(endNode)))
                throw new Error("The end node must be after the start node in the body");
        } else if (startIndex > endIndex)
            throw new Error("The section of end node must be after the section start node");
    }

    static findNextNode(nodeType, fromNode) {
        if (fromNode == null || fromNode.nodeType === nodeType)
            return fromNode;

        if (fromNode.isComposite) {
            let node = ExtractContentHelper.findNextNode(nodeType, fromNode.asCompositeNode().firstChild);
            if (node != null)
                return node;
        }

        return ExtractContentHelper.findNextNode(nodeType, fromNode.nextSibling);
    }

    static processMarker(cloneNode, nodes, node, blockLevelAncestor, isInclusive, isStartMarker, canAdd, forceAdd) {
        // If we are dealing with a block-level node, see if it should be included and add it to the list.
        if (base.compareNodes(node, blockLevelAncestor)) {
            if (canAdd && isInclusive)
                nodes.push(cloneNode);
            return;
        }

        // cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor
        // is the node's ancestor that means it is a composite node.
        console.assert(cloneNode.isComposite);

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node.nodeType === aw.NodeType.FieldStart) {
            // If the marker is a start node and is not included, skip to the end of the field.
            // If the marker is an end node and is to be included, then move to the end field so the field will not be removed.
            if (isStartMarker && !isInclusive || !isStartMarker && isInclusive) {
                while (node.nextSibling != null && node.nodeType !== aw.NodeType.FieldEnd)
                    node = node.nextSibling;
            }
        }

        // Support a case if the marker node is on the third level of the document body or lower.
        let nodeBranch = ExtractContentHelper.fillSelfAndParents(node, blockLevelAncestor);

        // Process the corresponding node in our cloned node by index.
        let currentCloneNode = cloneNode;
        for (let i = nodeBranch.length - 1; i >= 0; i--) {
            let currentNode = nodeBranch.at(i);
            let nodeIndex = currentNode.parentNode.indexOf(currentNode);

            currentCloneNode = currentCloneNode.asCompositeNode().getChildNodes(aw.NodeType.Any, false).at(nodeIndex);

            ExtractContentHelper.removeNodesOutsideOfRange(currentCloneNode, isInclusive || (i > 0), isStartMarker);
        }

        // After processing, the composite node may become empty if it has doesn't include it.
        if (canAdd && (forceAdd || cloneNode.asCompositeNode().hasChildNodes))
            nodes.push(cloneNode);
    }

    static removeNodesOutsideOfRange(markerNode, isInclusive, isStartMarker) {
        let isProcessing = true;
        let isRemoving = isStartMarker;
        let nextNode = markerNode.parentNode.firstChild;

        while (isProcessing && nextNode != null) {
            let currentNode = nextNode;
            let isSkip = false;

            if (base.compareNodes(currentNode, markerNode)) {
                if (isStartMarker) {
                    isProcessing = false;
                    if (isInclusive)
                        isRemoving = false;
                } else {
                    isRemoving = true;
                    if (isInclusive)
                        isSkip = true;
                }
            }

            nextNode = nextNode.nextSibling;
            if (isRemoving && !isSkip)
                currentNode.remove();
        }
    }

    static fillSelfAndParents(node, tillNode) {
        let list = [];
        let currentNode = node;

        while (!base.compareNodes(currentNode, tillNode)) {
            list.push(currentNode);
            currentNode = currentNode.parentNode;
        }

        return list;
    }

    static includeNextParagraph(node, nodes) {
        let paragraph = ExtractContentHelper.findNextNode(aw.NodeType.Paragraph, node.nextSibling);
        if (paragraph != null) {
            paragraph = paragraph.asParagraph();
            // Move to the first child to include paragraphs without content.
            let markerNode = paragraph.hasChildNodes ? paragraph.firstChild : paragraph;
            let rootNode = ExtractContentHelper.getAncestorInBody(paragraph);

            ExtractContentHelper.processMarker(rootNode.clone(true), nodes, markerNode, rootNode,
                markerNode === paragraph, false, true, true);
        }
    }

    static getAncestorInBody(startNode) {
        while (startNode.parentNode.nodeType !== aw.NodeType.Body)
            startNode = startNode.parentNode;
        return startNode;
    }

    //ExStart:GenerateDocument
    //GistId:433f5122fe18fdc24a406528b70b0020
    static generateDocument(srcDoc, nodes) {
        let dstDoc = new aw.Document();

        let importedSection = nodes.some(node => node.nodeType == aw.NodeType.Section) ? null : dstDoc.firstSection;
        if (importedSection == null)
            dstDoc.firstSection.remove();

        // Import each node from the list into the new document. Keep the original formatting of the node.
        let importer = new aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.KeepSourceFormatting);
        for (let node of nodes) {
            if (node.nodeType === aw.NodeType.Section) {
                // Import a section from the source document.
                let srcSection = node.asSection();
                importedSection = importer.importNode(srcSection, false).asSection();
                importedSection.appendChild(importer.importNode(srcSection.body, false));
                for (let hf of srcSection.headersFooters) {
                    hf = hf.asHeaderFooter();
                    importedSection.headersFooters.add(importer.importNode(hf, true));
                }

                dstDoc.appendChild(importedSection);
            } else {
                let importNode = importer.importNode(node, true);
                importedSection.body.appendChild(importNode);
            }
        }

        return dstDoc;
    }
    //ExEnd:GenerateDocument
}

module.exports = { ExtractContentHelper };