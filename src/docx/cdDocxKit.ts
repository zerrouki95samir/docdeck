/* Custom tiptap wiring for the docx WYSIWYG editor:
   - paragraph & list custom attrs (styleName/align/numId/ilvl/pageBreak/opaqueXml)
   - fontSize + fontFamily + color operate on textStyle marks
   - tab-stop atom node (\t in OOXML)
   - inline image with OOXML attrs (relId/wpx/hpx)
   - link mark with relId (external relationships) */

import { Extension, Node, mergeAttributes } from '@tiptap/react';
import type { CommandProps } from '@tiptap/core';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import TextStyle from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import FontFamily from '@tiptap/extension-font-family';
import Highlight from '@tiptap/extension-highlight';
import Link from '@tiptap/extension-link';
import Image from '@tiptap/extension-image';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import Table from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import { StarterKit } from '@tiptap/starter-kit';
import { CD_PARA, CD_TAB } from './blocksToTiptap';

const paraAttrs = {
  styleName: {
    default: null as string | null,
    parseHTML: (el: HTMLElement) => el.getAttribute('data-cd-style'),
    renderHTML: (attrs: Record<string, unknown>) =>
      attrs.styleName ? { 'data-cd-style': attrs.styleName } : {},
  },
  numId: {
    default: null as number | null,
    parseHTML: (el: HTMLElement) => el.getAttribute('data-cd-num') && parseInt(el.getAttribute('data-cd-num') || '', 10),
    renderHTML: (attrs: Record<string, unknown>) =>
      attrs.numId != null ? { 'data-cd-num': String(attrs.numId) } : {},
  },
  ilvl: { default: 0 },
  pageBreak: {
    default: false,
    parseHTML: (el: HTMLElement) => el.hasAttribute('data-cd-pbreak'),
    renderHTML: (attrs: Record<string, unknown>) => (attrs.pageBreak ? { 'data-cd-pbreak': '' } : {}),
  },
  opaqueXml: { default: null as string | null },
  opaqueTag: { default: null as string | null },
};

export const CdAttrs = Extension.create({
  name: 'cdAttrs',
  addGlobalAttributes() {
    return [
      {
        types: [CD_PARA],
        attributes: paraAttrs,
      },
      {
        types: ['bulletList', 'orderedList'],
        attributes: {
          numId: {
            default: null as number | null,
            parseHTML: (el: HTMLElement) => el.getAttribute('data-cd-num') && parseInt(el.getAttribute('data-cd-num') || '', 10),
            renderHTML: (attrs: Record<string, unknown>) =>
              attrs.numId != null ? { 'data-cd-num': String(attrs.numId) } : {},
          },
          ilvl: { default: 0 },
        },
      },
    ];
  },
});

/** set/unset a fontSize global attribute on textStyle marks */
declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    fontSize: {
      setFontSize: (size: string) => ReturnType;
      unsetFontSize: () => ReturnType;
    };
  }
}

export const FontSize = Extension.create({
  name: 'fontSize',
  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          fontSize: {
            default: null as string | null,
            parseHTML: (el: HTMLElement) => el.style.fontSize || null,
            renderHTML: (attrs: Record<string, unknown>) =>
              attrs.fontSize ? { style: `font-size: ${attrs.fontSize}` } : {},
          },
        },
      },
    ];
  },
  addCommands() {
    return {
      setFontSize:
        (size: string) =>
        ({ chain }: CommandProps) =>
          chain().extendMarkRange('textStyle').setMark('textStyle', { fontSize: size }).run(),
      unsetFontSize:
        () =>
        ({ chain }: CommandProps) =>
          chain().extendMarkRange('textStyle').setMark('textStyle', { fontSize: null }).run(),
    };
  },
});

export const CdTab = Node.create({
  name: CD_TAB,
  inline: true,
  group: 'inline',
  atom: true,
  renderHTML() {
    return ['span', { class: 'cd-tab' }];
  },
  parseHTML() {
    return [{ tag: 'span.cd-tab' }];
  },
});

export const CdImage = Image.extend({
  inline: () => true as never,
  addAttributes() {
    return {
      ...this.parent?.(),
      relId: { default: null as string | null },
      wpx: { default: null },
      hpx: { default: null },
    };
  },
  renderHTML({ HTMLAttributes, node }) {
    const w = (node.attrs.wpx as number) || 200;
    return [
      'img',
      mergeAttributes(HTMLAttributes, {
        'data-cd-rel': (node.attrs.relId as string) || '',
        width: w,
        class: 'cd-img',
      }),
    ];
  },
}).configure({ inline: true, allowBase64: true });

export const CdLink = Link.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      relId: { default: null as string | null },
    };
  },
}).configure({ openOnClick: false, autolink: true, linkOnPaste: true });

export function buildDocxExtensions() {
  return [
    StarterKit.configure({
      heading: { levels: [1, 2, 3] },
      blockquote: false,
      codeBlock: false,
      horizontalRule: false,
    }),
    Underline,
    TextStyle,
    Color,
    FontFamily,
    FontSize,
    Highlight.configure({ multicolor: true }),
    Subscript,
    Superscript,
    CdLink,
    TextAlign.configure({ types: ['heading', CD_PARA] }),
    CdAttrs,
    CdTab,
    CdImage,
    Table.configure({ resizable: true }),
    TableRow,
    TableCell,
    TableHeader,
  ];
}
