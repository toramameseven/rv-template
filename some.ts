import { MessageType, ShowMessage } from "./common";
import * as fs from "fs";
import * as Path from "path";
import {
  Bookmark,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  Indent,
  InternalHyperlink,
  Paragraph,
  ParagraphChild,
  patchDocument,
  PatchDocumentOptions,
  PatchType,
  Table,
  TableCell,
  TableRow,
  TextDirection,
  TextRun,
  VerticalAlign,
} from "docx";

const _sp = "\t";
//
const NodeType = {
  non: "non",
  section: "section",
  heading: "heading",
  OderList: "OderList",
  NormalList: "NormalList",
  //marked
  title: "title",
  subTitle: "subTitle",
  paragraph: "paragraph",
  list: "list",
  listitem: "listitem",
  code: "code",
  blockquote: "blockquote",
  table: "table",
  tablerow: "tablerow",
  tablecell: "tablecell",
  text: "text",
  image: "image",
  link: "link",
  html: "html",

  // word down
  author: "author",
  date: "date",
  division: "division",
  docxEngine: "docxEngine",
  docxTemplate: "docxTemplate",
  pageSetup: "pageSetup",
  toc: "toc",

  crossRef: "crossRef",
  property: "property",
  clearContent: "clearContent",
  docNumber: "docNumber",
  indentPlus: "indentPlus",
  indentMinus: "indentMinus",
  endParagraph: "endParagraph",
  newLine: "newLine",
  newPage: "newPage",
  htmlWdCommand: "htmlWdCommand",
  hr: "hr",
  // table
  cols: "cols",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",
} as const;
type NodeType = (typeof NodeType)[keyof typeof NodeType];

const DocStyle = {
  "1": "1",
  Body1: "body1",
  nList1: "nList1",
  nList2: "nList2",
  nList3: "nList3",
  numList1: "numList1",
  numList2: "numList2",
  numList3: "numList3",
  code: "code",
  Error: "Error",
} as const;

type DocStyle = (typeof DocStyle)[keyof typeof DocStyle];

interface MyNode {
  createNode: () => Paragraph | Table;
}

interface MyParagraph {
  createParagraph: () => Paragraph | Table;
}

class DocParagraph implements MyParagraph {
  nodeType: NodeType;
  isFlush: boolean;
  indent: number;
  children: ParagraphChild[] = [];
  docStyle: DocStyle;
  constructor(
    nodeType: NodeType = NodeType.non,
    docStyle: DocStyle = DocStyle.Body1,
    childe: ParagraphChild = new TextRun("")
  ) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.indent = 0;
    this.children = [childe];
    this.docStyle = docStyle;
  }
  createParagraph() {
    const docxR = new Paragraph({
      children: this.children,
      style: this.docStyle,
    });
    return docxR;
  }
  reset() {
    this.nodeType = "non";
    this.isFlush = false;
  }
  addChilde(s: string | ParagraphChild) {
    const ss = typeof s === "string" ? new TextRun(s) : s;
    this.children.push(ss);
  }
}

class DocJsNoeBase implements MyNode {
  nodeType: NodeType;
  isFlush: boolean;
  indent: number;
  text: string;
  //children: ParagraphChild[];
  createNode() {
    const docxR = new Paragraph({
      text: "",
      children: [
        new TextRun("My awesome text here for my university dissertation"),
        new TextRun("Foo Bar"),
        new TextRun({
          text: "this.text",
          style: "this.style",
        }),
      ],
    });
    return docxR;
  }
  reset() {
    this.nodeType = "non";
    this.isFlush = false;
  }
  constructor(text: string = "", nodeType: NodeType = NodeType.non) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.indent = 0;
    this.text = text;
  }
  addText(s: string) {
    this.text += s;
  }
}

class DocJsHeading extends DocJsNoeBase {
  heading: string;
  anchor: string;
  constructor(title: string, heading: string, anchor: string) {
    super(title, "heading");
    super.nodeType = "heading";
    this.heading = heading;
    this.anchor = anchor;
  }
  createNode() {
    super.reset();
    const docxR = new Paragraph({
      style: this.heading,
      children: [
        new Bookmark({
          id: this.anchor,
          children: [new TextRun(this.text)],
        }),
      ],
    });

    return docxR;
  }
}

class DocJsText extends DocJsNoeBase {
  style: string;

  constructor(text: string, style = "body1") {
    super(text);
    super.nodeType = "text";
    this.style = style;
  }

  createNode() {
    super.reset();
    const docxR = new Paragraph({
      text: super.text,
      style: this.style,
    });
    return docxR;
  }
}

class DocJsList extends DocJsNoeBase {
  style: string;
  constructor(listType: NodeType, listOrder: number) {
    super();
    super.nodeType = listType;
    this.style = this.createListType(listType, listOrder);
  }

  createListType(listType: NodeType, listOrder: number) {
    if (listType === NodeType.NormalList) {
      return `nList${listOrder}`;
    }
    if (listType === NodeType.OderList) {
      return `numList${listOrder}`;
    }

    return "error";
  }

  createNode() {
    super.reset();
    const docxR = new Paragraph({
      text: this.text,
      style: this.style,
    });
    return docxR;
  }
}

class DocJsLink extends DocJsNoeBase {
  //text: string;
  style: string;
  ref: string;
  constructor(ref: string) {
    super();
    super.nodeType = NodeType.link;
    //this.text = "";
    this.style = "";
    this.ref = ref;
  }

  createNode() {
    super.reset();
    const link = new InternalHyperlink({
      children: [
        new TextRun({
          text: this.text,
          style: "Hyperlink",
        }),
      ],
      anchor: this.ref,
    });

    const docxR = new Paragraph({
      children: [link],
    });

    return docxR;
  }
}

export async function wdToDocxJsOr(
  wd: string,
  docxTemplatePath: string,
  docxOutPath: string
): Promise<void> {
  let patches: (Paragraph | Table)[] = [];

  const lines = wd.split(/\r?\n/);
  let currentNodes: DocJsNoeBase[];
  currentNodes = [new DocJsNoeBase("", NodeType.non)];

  for (let i = 0; i < lines.length; i++) {
    currentNodes = resolveCommand(lines[i], currentNodes);
    if (currentNodes[currentNodes.length - 1].isFlush) {
      const filtered = currentNodes.filter((j) => j.nodeType !== NodeType.non);

      filtered.forEach((item) => {
        console.log("===========>" + item.nodeType);
        patches.push(item.createNode());
      });

      currentNodes = [new DocJsNoeBase("", NodeType.non)];
    }
  }
  await createDocxPatch(patches, docxTemplatePath, docxOutPath);
}

export async function wdToDocxJs(
  wd: string,
  docxTemplatePath: string,
  docxOutPath: string
): Promise<void> {
  let patches: (Paragraph | Table)[] = [];

  const lines = wd.split(/\r?\n/);
  let currentNodes: DocParagraph;
  currentNodes = new DocParagraph(NodeType.text);

  for (let i = 0; i < lines.length; i++) {
    currentNodes = resolveCommandEx(lines[i], currentNodes);
    if (currentNodes.isFlush) {
      const p = currentNodes.createParagraph();
      console.dir(p);
      patches.push(p);
      currentNodes = new DocParagraph(NodeType.text);
    }

  }

  await createDocxPatch(patches, docxTemplatePath, docxOutPath);
}

// ############################################################

function createListType(listType: NodeType, listOrder: number) {
  if (listType === NodeType.NormalList) {
    return `nList${listOrder}` as DocStyle;
  }
  if (listType === NodeType.OderList) {
    return `numList${listOrder}` as DocStyle;
  }
  return DocStyle.Error;
}

function resolveCommandEx(line: string, nodes: DocParagraph) {
  const words = line.split(_sp);
  let current: DocParagraph;
  const nodeType = words[0] as NodeType;
  let style: DocStyle;
  let childe: ParagraphChild;
  switch (nodeType) {
    case "section":
      // section 2 Heading2 heading2
      childe = new Bookmark({
        id: words[3],
        children: [new TextRun(words[2])],
      });
      current = new DocParagraph(nodeType, words[1] as DocStyle, childe);
      //nodes.push(current);
      return current;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(NodeType.NormalList, style);
      //nodes.push(current);
      return current;
      break;
    case NodeType.OderList:
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(NodeType.OderList, style);
      return current;
      break;
    case "code":
      childe = new TextRun(words[1]);
      current = new DocParagraph(nodeType, "code", childe);
      return current;
      break;
    case NodeType.link:
      childe = new InternalHyperlink({
        children: [
          new TextRun({
            text: words[3],
          }),
        ],
        anchor: words[1],
      });
      childe = new ExternalHyperlink({
        children: [
          new TextRun({
            text: words[3],
          }),
        ],
        link: words[1],
      });
      nodes.addChilde(childe);
      return nodes;
      break;
    case "text":
      //nodes.nodeType = NodeType.text;
      nodes.addChilde(words[1]);
      return nodes;
      break;
    case "newLine":
      nodes.isFlush = true;
      return nodes;
    default:
      return nodes;
  }
}

function resolveCommand(line: string, nodes: DocJsNoeBase[]) {
  const words = line.split(_sp);
  let currentNode: DocJsNoeBase;
  switch (words[0] as NodeType) {
    case "section":
      currentNode = new DocJsHeading(words[2], words[1], words[3]);
      nodes.push(currentNode);
      return nodes;
      break;
    case "code":
      currentNode = new DocJsHeading(words[1], "code", "");
      nodes.push(currentNode);
      return nodes;
      break;
    case "text":
      if (nodes[nodes.length - 1].nodeType !== NodeType.non) {
        nodes[nodes.length - 1].addText(words[1]);
        return nodes;
      } else {
        currentNode = new DocJsText(words[1], "body1");
        nodes.push(currentNode);
        return nodes;
      }
      break;
    case "NormalList":
      currentNode = new DocJsList(NodeType.NormalList, parseInt(words[1]));
      nodes.push(currentNode);
      return nodes;
      break;
    case NodeType.OderList:
      currentNode = new DocJsList(NodeType.OderList, parseInt(words[1]));
      nodes.push(currentNode);
      return nodes;
      break;
    case NodeType.link:
      currentNode = new DocJsLink(words[1]);
      nodes.push(currentNode);
      return nodes;
      break;
    case "newLine":
      nodes[nodes.length - 1].isFlush = true;
      return nodes;
    default:
      return nodes;
  }
}

export async function createDocxPatch(
  children: (Paragraph | Table)[],
  docxTemplatePath: string,
  docxOutPath: string
) {
  // console.dir(children);
  const patchDoc = await patchDocument(fs.readFileSync(docxTemplatePath), {
    patches: {
      paragraphReplace: {
        type: PatchType.DOCUMENT,
        children: children,
      },
    },
  });
  fs.writeFileSync(docxOutPath, patchDoc);
}
