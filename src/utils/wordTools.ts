import { tool } from '@langchain/core/tools'
import { z } from 'zod'

export type WordToolName =
  | 'getSelectedText'
  | 'getDocumentContent'
  | 'insertText'
  | 'replaceSelectedText'
  | 'appendText'
  | 'insertParagraph'
  | 'formatText'
  | 'searchAndReplace'
  | 'getDocumentProperties'
  | 'insertTable'
  | 'formatTable'
  | 'setColumnWidths'
  | 'updateTableCell'
  | 'insertTableRow'
  | 'deleteTableRow'
  | 'insertList'
  | 'deleteText'
  | 'clearFormatting'
  | 'setFontName'
  | 'insertPageBreak'
  | 'getRangeInfo'
  | 'selectText'
  | 'insertImage'
  | 'getTableInfo'
  | 'insertBookmark'
  | 'goToBookmark'
  | 'insertContentControl'
  | 'findText'
  | 'insertComment'
  | 'getComments'
  | 'deleteComment'
  | 'replyToComment'
  | 'insertHeader'
  | 'insertFooter'
  | 'getHeaderFooter'
  | 'toggleTrackChanges'
  | 'insertTableOfContents'
  | 'updateTableOfContents'
  | 'applyStyle'
  | 'listStyles'
  | 'insertSectionBreak'
  | 'getSections'
  | 'insertFootnote'
  | 'insertEndnote'
  | 'setPageMargins'
  | 'clearDocument'
  | 'insertCoverPage'
  | 'insertEquation'

// Word operation mutex — ensures parallel tool calls from LangGraph's ToolNode
// execute sequentially. Without this, concurrent insertParagraph('End') calls
// race for the same body position, producing reversed or scrambled document structure.
let _wordLock = Promise.resolve()
const withWordLock = <T>(fn: () => Promise<T>): Promise<T> => {
  const current = _wordLock
  let release!: () => void
  _wordLock = new Promise<void>(res => { release = res })
  return current.then(() => fn().finally(release))
}

// Escape XML special characters for safe OOXML injection
const escapeXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')

// Build a flat-OPC OOXML package containing a table with an optional blue header row.
// Using OOXML injection bypasses the Mac Office.js bug where cell.shadingColor is silently ignored.
// Pass headerColor=undefined to skip header styling entirely.
const buildTableOoxml = (data: string[][], opts: { headerColor?: string; noBorder?: boolean } = {}): string => {
  const { headerColor, noBorder = false } = opts
  const applyHeader = headerColor !== undefined
  const hColor = headerColor ?? '2E74B5'
  const cols = data[0]?.length || 0
  const styleId = noBorder ? 'TableNormal' : 'TableGrid'

  const gridCols = Array(cols).fill('<w:gridCol/>').join('')

  const wRows = data
    .map((rowData, rowIdx) => {
      const isHdr = applyHeader && rowIdx === 0
      const trPr = isHdr ? '<w:trPr><w:tblHeader/></w:trPr>' : ''
      const cells = rowData
        .map(text => {
          const tcPr = isHdr
            ? `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${hColor}"/></w:tcPr>`
            : ''
          const rPr = isHdr ? '<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' : ''
          return `<w:tc>${tcPr}<w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`
        })
        .join('')
      return `<w:tr>${trPr}${cells}</w:tr>`
    })
    .join('')

  const tbl = `<w:tbl><w:tblPr><w:tblStyle w:val="${styleId}"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${wRows}</w:tbl>`

  // Full flat-OPC package — the only format body.insertOoxml() reliably accepts
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${tbl}<w:p/><w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
  )
}

// ── LaTeX → OMML (Office Math Markup Language) ─────────────────────────────

const LATEX_SYMBOLS: Record<string, string> = {
  // Greek lowercase
  '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
  '\\epsilon': 'ε', '\\varepsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η',
  '\\theta': 'θ', '\\vartheta': 'ϑ', '\\iota': 'ι', '\\kappa': 'κ',
  '\\lambda': 'λ', '\\mu': 'μ', '\\nu': 'ν', '\\xi': 'ξ',
  '\\pi': 'π', '\\varpi': 'ϖ', '\\rho': 'ρ', '\\sigma': 'σ',
  '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ', '\\varphi': 'φ',
  '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
  // Greek uppercase
  '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
  '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Upsilon': 'Υ',
  '\\Phi': 'Φ', '\\Psi': 'Ψ', '\\Omega': 'Ω',
  // Operators
  '\\times': '×', '\\div': '÷', '\\pm': '±', '\\mp': '∓',
  '\\cdot': '·', '\\ast': '∗', '\\star': '⋆',
  // Relations
  '\\leq': '≤', '\\le': '≤', '\\geq': '≥', '\\ge': '≥',
  '\\neq': '≠', '\\ne': '≠', '\\approx': '≈', '\\equiv': '≡',
  '\\sim': '∼', '\\propto': '∝', '\\ll': '≪', '\\gg': '≫',
  // Large operators
  '\\sum': 'Σ', '\\prod': 'Π', '\\int': '∫', '\\oint': '∮',
  '\\bigcap': '∩', '\\bigcup': '∪',
  // Misc math
  '\\partial': '∂', '\\infty': '∞', '\\nabla': '∇',
  '\\forall': '∀', '\\exists': '∃',
  '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\supset': '⊃',
  '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
  '\\cdots': '⋯', '\\ldots': '…', '\\vdots': '⋮',
  '\\to': '→', '\\rightarrow': '→', '\\leftarrow': '←',
  '\\Rightarrow': '⇒', '\\Leftarrow': '⇐', '\\Leftrightarrow': '⇔',
  '\\langle': '⟨', '\\rangle': '⟩',
  '\\lfloor': '⌊', '\\rfloor': '⌋', '\\lceil': '⌈', '\\rceil': '⌉',
  '\\mid': '|', '\\prime': '′', '\\circ': '∘', '\\bullet': '•',
  '\\oplus': '⊕', '\\otimes': '⊗',
  '\\because': '∵', '\\therefore': '∴',
  '\\angle': '∠', '\\perp': '⊥', '\\parallel': '∥',
  '\\not': '¬',
  // Escaped specials
  '\\%': '%', '\\$': '$', '\\#': '#',
  // Spacing (render as nothing)
  '\\,': '', '\\;': '', '\\:': '', '\\!': '', '\\ ': ' ',
}

// Commands that simply pass through their single argument (decorators, font changes)
const PASSTHROUGH_CMDS = new Set([
  '\\mathrm', '\\mathbf', '\\mathit', '\\mathbb', '\\mathcal', '\\mathsf', '\\mathtt',
  '\\text', '\\textrm', '\\textbf', '\\textit',
  '\\hat', '\\bar', '\\vec', '\\dot', '\\ddot', '\\tilde',
  '\\widehat', '\\widetilde', '\\overline', '\\underline',
  '\\overbrace', '\\underbrace', '\\boldsymbol', '\\operatorname',
])

type LatexToken =
  | { t: 'cmd'; v: string }
  | { t: 'text'; v: string }
  | { t: 'lbrace' }
  | { t: 'rbrace' }
  | { t: 'lbracket' }
  | { t: 'rbracket' }
  | { t: 'sup' }
  | { t: 'sub' }

function tokenizeLaTeX(src: string): LatexToken[] {
  const toks: LatexToken[] = []
  let i = 0
  while (i < src.length) {
    const ch = src[i]
    if (ch === '\\') {
      let j = i + 1
      if (j < src.length && /[a-zA-Z]/.test(src[j])) {
        while (j < src.length && /[a-zA-Z]/.test(src[j])) j++
        toks.push({ t: 'cmd', v: src.slice(i, j) })
      } else if (j < src.length) {
        toks.push({ t: 'cmd', v: src.slice(i, j + 1) })
        j++
      }
      i = j
    } else if (ch === '{') { toks.push({ t: 'lbrace' }); i++ }
    else if (ch === '}') { toks.push({ t: 'rbrace' }); i++ }
    else if (ch === '[') { toks.push({ t: 'lbracket' }); i++ }
    else if (ch === ']') { toks.push({ t: 'rbracket' }); i++ }
    else if (ch === '^') { toks.push({ t: 'sup' }); i++ }
    else if (ch === '_') { toks.push({ t: 'sub' }); i++ }
    else if (ch === ' ' || ch === '\t' || ch === '\n') { i++ }
    else {
      let j = i
      while (j < src.length && !'\\{}[]^_ \t\n'.includes(src[j])) j++
      if (j > i) toks.push({ t: 'text', v: src.slice(i, j) })
      i = j
    }
  }
  return toks
}

class OmmlParser {
  private toks: LatexToken[]
  private pos = 0
  constructor(toks: LatexToken[]) { this.toks = toks }

  private cur(): LatexToken | undefined { return this.toks[this.pos] }
  private run(text: string): string {
    return `<m:r><m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>`
  }

  private parseAtom(): string {
    const tok = this.cur()
    if (!tok) return ''

    if (tok.t === 'lbrace') {
      this.pos++
      const inner = this.parseSequence(true)
      if (this.cur()?.t === 'rbrace') this.pos++
      return inner
    }

    if (tok.t === 'lbracket') { this.pos++; return this.run('[') }
    if (tok.t === 'rbracket') { this.pos++; return this.run(']') }

    if (tok.t === 'text') { this.pos++; return this.run(tok.v) }

    if (tok.t === 'cmd') {
      this.pos++
      const cmd = tok.v

      if (cmd === '\\frac' || cmd === '\\dfrac') {
        const num = this.parseAtom()
        const den = this.parseAtom()
        return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`
      }
      if (cmd === '\\sqrt') {
        // Skip optional degree \sqrt[n]{x} — consume tokens until ]
        if (this.cur()?.t === 'lbracket') {
          this.pos++
          while (this.pos < this.toks.length && this.cur()?.t !== 'rbracket') this.pos++
          if (this.cur()?.t === 'rbracket') this.pos++
        }
        const arg = this.parseAtom()
        return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${arg}</m:e></m:rad>`
      }
      if (PASSTHROUGH_CMDS.has(cmd)) return this.parseAtom()
      if (cmd === '\\left' || cmd === '\\right') return '' // skip size hint, delimiter follows as text

      const sym = LATEX_SYMBOLS[cmd]
      if (sym !== undefined) return sym ? this.run(sym) : ''
      return this.run(cmd.slice(1)) // unknown — strip backslash and render
    }

    this.pos++
    return ''
  }

  private parseNode(): string {
    const base = this.parseAtom()
    let sub: string | null = null
    let sup: string | null = null
    while (true) {
      const tok = this.cur()
      if (tok?.t === 'sub' && sub === null) { this.pos++; sub = this.parseAtom() }
      else if (tok?.t === 'sup' && sup === null) { this.pos++; sup = this.parseAtom() }
      else break
    }
    if (sub !== null && sup !== null)
      return `<m:sSubSup><m:sSubSupPr/><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`
    if (sub !== null)
      return `<m:sSub><m:sSubPr/><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`
    if (sup !== null)
      return `<m:sSup><m:sSupPr/><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`
    return base
  }

  parseSequence(insideGroup = false): string {
    const parts: string[] = []
    while (this.pos < this.toks.length) {
      if (insideGroup && this.cur()?.t === 'rbrace') break
      parts.push(this.parseNode())
    }
    return parts.join('')
  }
}

function buildEquationOoxml(latex: string, displayMode: boolean): string {
  const omml = new OmmlParser(tokenizeLaTeX(latex)).parseSequence()
  const m = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
  const para = displayMode
    ? `<w:p><m:oMathPara ${m}><m:oMath>${omml}</m:oMath></m:oMathPara></w:p>`
    : `<w:p><m:oMath ${m}>${omml}</m:oMath></w:p>`
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${m}>` +
    `<w:body>${para}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
  )
}

// ── End LaTeX helpers ────────────────────────────────────────────────────────

const wordToolDefinitions: Record<WordToolName, WordToolDefinition> = {
  getSelectedText: {
    name: 'getSelectedText',
    description:
      'Get the currently selected text in the Word document. Returns the selected text or empty string if nothing is selected.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.load('text')
        await context.sync()
        return range.text || ''
      })
    },
  },

  getDocumentContent: {
    name: 'getDocumentContent',
    description: 'Get the full content of the Word document body as plain text.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const body = context.document.body
        body.load('text')
        await context.sync()
        return body.text || ''
      })
    },
  },

  insertText: {
    name: 'insertText',
    description: 'Insert text at the current cursor position in the Word document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to insert',
        },
        location: {
          type: 'string',
          description: 'Where to insert: "Start", "End", "Before", "After", or "Replace"',
          enum: ['Start', 'End', 'Before', 'After', 'Replace'],
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text, location = 'End' } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertText(text, location)
        await context.sync()
        return `Text inserted: "${text.substring(0, 80)}${text.length > 80 ? '...' : ''}"`
      })
    },
  },

  replaceSelectedText: {
    name: 'replaceSelectedText',
    description: 'Replace the currently selected text with new text.',
    inputSchema: {
      type: 'object',
      properties: {
        newText: {
          type: 'string',
          description: 'The new text to replace the selection with',
        },
      },
      required: ['newText'],
    },
    execute: async args => {
      const { newText } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertText(newText, 'Replace')
        await context.sync()
        return `Selected text replaced with: "${newText.substring(0, 80)}${newText.length > 80 ? '...' : ''}"`
      })
    },
  },

  appendText: {
    name: 'appendText',
    description: 'Append text to the end of the document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to append',
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text } = args
      return Word.run(async context => {
        const body = context.document.body
        body.insertText(text, 'End')
        await context.sync()
        return `Text appended to document end.`
      })
    },
  },

  insertParagraph: {
    name: 'insertParagraph',
    description: 'Insert a new paragraph at the specified location. Specify fontSize and fontFamily for proper Chinese document formatting.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The paragraph text. To write multiple paragraphs in one call, separate them with a real newline character (\\n). Each \\n creates a new paragraph with the same style.',
        },
        location: {
          type: 'string',
          description:
            'Where to insert: "After" (after cursor/selection), "Before" (before cursor), "Start" (start of doc), or "End" (end of doc). Default is "End".',
          enum: ['After', 'Before', 'Start', 'End'],
        },
        style: {
          type: 'string',
          description:
            'Word built-in style. Use Heading1 for 一/二/三 sections, Heading2 for sub-sections, Normal for body text.',
          enum: ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Quote', 'IntenseQuote', 'Title', 'Subtitle'],
        },
        fontSize: {
          type: 'number',
          description:
            'Font size in points. Recommended: body text = 12, section heading = 16, sub-heading = 14. Omit to use the style default.',
        },
        bold: {
          type: 'boolean',
          description: 'Whether the paragraph text should be bold.',
        },
        fontFamily: {
          type: 'string',
          description: 'Font family. For Chinese documents: "SimSun" (宋体), "SimHei" (黑体), "Microsoft YaHei" (微软雅黑). Omit to use style default.',
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text, location = 'End', style, fontSize, bold, fontFamily } = args
      return Word.run(async context => {
        // Normalise: convert literal \n sequences (backslash+n from agent JSON) to real newlines,
        // then split so each line becomes a separate Word paragraph.
        const normalised = text.replace(/\\n/g, '\n')
        const segments = normalised.split('\n')

        const applyFormat = (para: Word.Paragraph) => {
          if (style) { try { para.styleBuiltIn = style as any } catch { para.style = style } }
          if (fontSize !== undefined) para.font.size = fontSize
          if (bold !== undefined) para.font.bold = bold
          if (fontFamily) para.font.name = fontFamily
        }

        if (location === 'Start' || location === 'End') {
          // For Start, insert in reverse order so reading order is preserved top-to-bottom
          const ordered = location === 'Start' ? [...segments].reverse() : segments
          for (const seg of ordered) {
            applyFormat(context.document.body.insertParagraph(seg, location as any))
          }
        } else {
          // Before/After: cursor-relative; multi-line not supported, use joined text
          const range = context.document.getSelection()
          applyFormat(range.insertParagraph(normalised, location as any))
        }

        await context.sync()
        return segments.length > 1
          ? `${segments.length} paragraphs inserted at ${location}: "${text.substring(0, 60)}"`
          : `Paragraph inserted at ${location}: "${text.substring(0, 80)}"`
      })
    },
  },

  formatText: {
    name: 'formatText',
    description: 'Apply formatting to the currently selected text.',
    inputSchema: {
      type: 'object',
      properties: {
        bold: {
          type: 'boolean',
          description: 'Make text bold',
        },
        italic: {
          type: 'boolean',
          description: 'Make text italic',
        },
        underline: {
          type: 'boolean',
          description: 'Underline text',
        },
        fontSize: {
          type: 'number',
          description: 'Font size in points',
        },
        fontColor: {
          type: 'string',
          description: 'Font color as hex (e.g., "#FF0000" for red)',
        },
        highlightColor: {
          type: 'string',
          description:
            'Highlight color: Yellow, Green, Cyan, Pink, Blue, Red, DarkBlue, Teal, Lime, Purple, Orange, etc.',
        },
      },
      required: [],
    },
    execute: async args => {
      const { bold, italic, underline, fontSize, fontColor, highlightColor } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        if (bold !== undefined) range.font.bold = bold
        if (italic !== undefined) range.font.italic = italic
        if (underline !== undefined) range.font.underline = underline ? 'Single' : 'None'
        if (fontSize !== undefined) range.font.size = fontSize
        if (fontColor !== undefined) range.font.color = fontColor
        if (highlightColor !== undefined) range.font.highlightColor = highlightColor
        await context.sync()
        return 'Formatting applied to selected text.'
      })
    },
  },

  searchAndReplace: {
    name: 'searchAndReplace',
    description: 'Search for text in the document and replace it with new text.',
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'The text to search for',
        },
        replaceText: {
          type: 'string',
          description: 'The text to replace with',
        },
        matchCase: {
          type: 'boolean',
          description: 'Whether to match case (default: false)',
        },
        matchWholeWord: {
          type: 'boolean',
          description: 'Whether to match whole word only (default: false)',
        },
      },
      required: ['searchText', 'replaceText'],
    },
    execute: async args => {
      const { searchText, replaceText, matchCase = false, matchWholeWord = false } = args
      return Word.run(async context => {
        const results = context.document.body.search(searchText, { matchCase, matchWholeWord })
        results.load('items')
        await context.sync()
        const count = results.items.length
        for (const result of results.items) {
          result.insertText(replaceText, 'Replace')
        }
        await context.sync()
        return `Replaced ${count} occurrence(s) of "${searchText}" with "${replaceText}"`
      })
    },
  },

  getDocumentProperties: {
    name: 'getDocumentProperties',
    description: 'Get document properties including paragraph count, word count, and character count.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const body = context.document.body
        body.load(['text'])

        const paragraphs = body.paragraphs
        paragraphs.load('items')

        await context.sync()

        const text = body.text || ''
        const wordCount = text.split(/\s+/).filter(word => word.length > 0).length
        const charCount = text.length
        const paragraphCount = paragraphs.items.length

        return JSON.stringify(
          {
            paragraphCount,
            wordCount,
            characterCount: charCount,
          },
          null,
          2,
        )
      })
    },
  },

  insertTable: {
    name: 'insertTable',
    description: 'Insert a professional table with blue header and automatic formatting. Provide the full data array including header row. Call ONLY ONCE. Use insertParagraph BEFORE this call if you need a heading above the table.',
    inputSchema: {
      type: 'object',
      properties: {
        rows: { type: 'number', description: 'Total number of rows including header row' },
        columns: { type: 'number', description: 'Number of columns' },
        data: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: '2D array of cell values, first row is the header' },
      },
      required: ['rows', 'columns', 'data'],
    },
    execute: async args => {
      const { rows, columns, data } = args
      return Word.run(async context => {
        // Use OOXML injection — cell.shadingColor on Mac Office.js is silently ignored,
        // but writing the <w:shd> element directly into OOXML bypasses that bug entirely.
        const td: string[][] = data || Array(rows).fill(null).map(() => Array(columns).fill(''))
        const ooxml = buildTableOoxml(td, { headerColor: '2E74B5' })
        context.document.body.insertOoxml(ooxml, 'End')
        await context.sync()
        return 'SUCCESS: Table inserted with blue header.'
      })
    },
  },

  formatTable: {
    name: 'formatTable',
    description:
      'Reformat an existing table: applies blue header, controls borders, and updates styling. ' +
      'Default (no args): blue header + visible borders (Table Grid). ' +
      'To REMOVE borders: set noBorder=true. ' +
      'To ADD borders back: call with noBorder=false (or omit noBorder). ' +
      'To skip header styling: set hasHeader=false.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: {
          type: 'number',
          description: 'Zero-based index of the table to format (default: 0 = first table)',
        },
        noBorder: {
          type: 'boolean',
          description: 'true = remove all borders (borderless table). false or omitted = visible borders via Table Grid style.',
        },
        hasHeader: {
          type: 'boolean',
          description: 'true (default) = apply blue background + white bold text to first row. false = no header styling.',
        },
        headerColor: {
          type: 'string',
          description: 'Hex color for header background, e.g. "#2E74B5" (default). Must include the # prefix.',
        },
      },
      required: [],
    },
    execute: async args => {
      const { tableIndex = 0, noBorder = false, hasHeader = true, headerColor = '#2E74B5' } = args
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0) {
          return 'ERROR: No table found in document. Use insertTable first.'
        }

        const idx = Math.min(Math.max(0, tableIndex), tables.items.length - 1)
        const table = tables.items[idx]

        // Read the existing cell values so we can rebuild the table as OOXML.
        // table.values is a 2D string array — the only reliable cross-platform way
        // to get cell text without the Mac cell-proxy shadingColor bug.
        table.load('values')
        await context.sync()

        const tableData: string[][] = table.values
        const color = headerColor.replace(/^#/, '')

        // Replace the table range in-place with fresh OOXML that bakes in the header shading.
        const range = table.getRange()
        const ooxml = buildTableOoxml(tableData, { headerColor: hasHeader ? color : undefined, noBorder })
        range.insertOoxml(ooxml, 'Replace')
        await context.sync()

        return `Table ${idx} reformatted${noBorder ? ' (borders removed)' : ''} with${hasHeader ? ' blue header' : 'out header styling'}.`
      })
    },
  },

  setColumnWidths: {
    name: 'setColumnWidths',
    description:
      'Set precise column widths for a table in points. A4 portrait content area ≈ 450 pts total. Example for 4 columns: [100, 150, 120, 80].',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: {
          type: 'number',
          description: 'Zero-based index of the table (default: 0)',
        },
        widths: {
          type: 'array',
          description: 'Array of column widths in points. Must match the number of columns.',
          items: { type: 'number' },
        },
      },
      required: ['widths'],
    },
    execute: async args => {
      const { tableIndex = 0, widths } = args
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0) {
          return 'ERROR: No table found. Use insertTable first.'
        }

        const idx = Math.min(Math.max(0, tableIndex), tables.items.length - 1)
        const table = tables.items[idx]
        table.columns.load('items')
        await context.sync()

        for (let i = 0; i < Math.min(widths.length, table.columns.items.length); i++) {
          table.columns.items[i].width = widths[i]
        }
        await context.sync()
        return `Column widths set to [${widths.join(', ')}] pts on table ${idx}.`
      })
    },
  },

  updateTableCell: {
    name: 'updateTableCell',
    description: 'Update the content of a specific cell in a table by row and column index.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: {
          type: 'number',
          description: 'Zero-based index of the table (default: 0)',
        },
        rowIndex: {
          type: 'number',
          description: 'Zero-based row index',
        },
        columnIndex: {
          type: 'number',
          description: 'Zero-based column index',
        },
        text: {
          type: 'string',
          description: 'New text content for the cell',
        },
      },
      required: ['rowIndex', 'columnIndex', 'text'],
    },
    execute: async args => {
      const { tableIndex = 0, rowIndex, columnIndex, text } = args
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0) {
          return 'ERROR: No table found. Use insertTable first.'
        }

        const idx = Math.min(Math.max(0, tableIndex), tables.items.length - 1)
        const cell = tables.items[idx].getCell(rowIndex, columnIndex)
        cell.body.insertText(text, 'Replace')
        await context.sync()
        return `Cell [${rowIndex},${columnIndex}] updated to: "${text}"`
      })
    },
  },

  insertTableRow: {
    name: 'insertTableRow',
    description: 'Insert a new row into an existing table at a specified position.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: {
          type: 'number',
          description: 'Zero-based index of the table (default: 0)',
        },
        afterRowIndex: {
          type: 'number',
          description: 'Insert after this row index. Use -1 to insert before the first row.',
        },
        values: {
          type: 'array',
          description: 'Array of cell text values for the new row',
          items: { type: 'string' },
        },
      },
      required: ['afterRowIndex', 'values'],
    },
    execute: async args => {
      const { tableIndex = 0, afterRowIndex, values } = args
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0) {
          return 'ERROR: No table found. Use insertTable first.'
        }

        const idx = Math.min(Math.max(0, tableIndex), tables.items.length - 1)
        const table = tables.items[idx]
        table.load('rowCount')
        await context.sync()

        if (afterRowIndex < 0) {
          table.addRows('Start', 1, values ? [values] : undefined)
        } else {
          const safeIdx = Math.min(afterRowIndex, table.rowCount - 1)
          const targetRow = table.rows.getItem(safeIdx)
          targetRow.insertRows(1, 'After', values ? [values] : undefined)
        }
        await context.sync()
        return `Row inserted after index ${afterRowIndex}.`
      })
    },
  },

  deleteTableRow: {
    name: 'deleteTableRow',
    description: 'Delete a specific row from a table by index.',
    inputSchema: {
      type: 'object',
      properties: {
        tableIndex: {
          type: 'number',
          description: 'Zero-based index of the table (default: 0)',
        },
        rowIndex: {
          type: 'number',
          description: 'Zero-based index of the row to delete',
        },
      },
      required: ['rowIndex'],
    },
    execute: async args => {
      const { tableIndex = 0, rowIndex } = args
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load('items')
        await context.sync()

        if (tables.items.length === 0) {
          return 'ERROR: No table found.'
        }

        const idx = Math.min(Math.max(0, tableIndex), tables.items.length - 1)
        const table = tables.items[idx]
        table.load('rowCount')
        await context.sync()

        if (rowIndex < 0 || rowIndex >= table.rowCount) {
          return `ERROR: Row index ${rowIndex} out of range (table has ${table.rowCount} rows).`
        }

        table.rows.getItem(rowIndex).delete()
        await context.sync()
        return `Row ${rowIndex} deleted from table ${idx}.`
      })
    },
  },

  insertList: {
    name: 'insertList',
    description: 'Insert a bulleted or numbered list at the current position.',
    inputSchema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          description: 'Array of list item texts',
          items: { type: 'string' },
        },
        listType: {
          type: 'string',
          description: 'Type of list: "bullet" or "number"',
          enum: ['bullet', 'number'],
        },
      },
      required: ['items', 'listType'],
    },
    execute: async args => {
      const { items, listType } = args
      return Word.run(async context => {
        const body = context.document.body
        for (const item of items) {
          const para = body.insertParagraph(item, 'End')
          try {
            para.styleBuiltIn = (listType === 'bullet' ? 'ListBullet' : 'ListNumber') as any
          } catch {
            para.style = listType === 'bullet' ? 'List Bullet' : 'List Number'
          }
        }
        await context.sync()
        return `${listType === 'bullet' ? 'Bulleted' : 'Numbered'} list with ${items.length} items inserted.`
      })
    },
  },

  deleteText: {
    name: 'deleteText',
    description:
      'Delete the currently selected text or a specific range. If no text is selected, this will delete at the cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        direction: {
          type: 'string',
          description: 'Direction to delete if nothing selected: "Before" (backspace) or "After" (delete key)',
          enum: ['Before', 'After'],
        },
      },
      required: [],
    },
    execute: async args => {
      const { direction = 'After' } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.load('text')
        await context.sync()
        if (range.text.length > 0) {
          range.delete()
        } else {
          // No selection: expand in the given direction and delete
          const expanded = range.expandTo(range)
          expanded.delete()
        }
        await context.sync()
        return 'Text deleted.'
      })
    },
  },

  clearFormatting: {
    name: 'clearFormatting',
    description: 'Clear all formatting from the selected text, returning it to default style.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.font.bold = false
        range.font.italic = false
        range.font.underline = 'None'
        range.styleBuiltIn = 'Normal'
        await context.sync()
        return 'Successfully cleared formatting'
      })
    },
  },

  setFontName: {
    name: 'setFontName',
    description: 'Set the font name/family for the selected text (e.g., Arial, Times New Roman, Calibri).',
    inputSchema: {
      type: 'object',
      properties: {
        fontName: {
          type: 'string',
          description: 'The font name to apply (e.g., "Arial", "Times New Roman", "Calibri", "Consolas")',
        },
      },
      required: ['fontName'],
    },
    execute: async args => {
      const { fontName } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.font.name = fontName
        await context.sync()
        return `Font set to "${fontName}" on selected text.`
      })
    },
  },

  insertPageBreak: {
    name: 'insertPageBreak',
    description: 'Insert a page break at the current cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        location: {
          type: 'string',
          description: 'Where to insert: "Before", "After", "Start", or "End"',
          enum: ['Before', 'After', 'Start', 'End'],
        },
      },
      required: [],
    },
    execute: async args => {
      const { location = 'After' } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertBreak('Page', location)
        await context.sync()
        return `Page break inserted at ${location}.`
      })
    },
  },

  getRangeInfo: {
    name: 'getRangeInfo',
    description: 'Get detailed information about the current selection including text, formatting, and position.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.load([
          'text',
          'style',
          'font/name',
          'font/size',
          'font/bold',
          'font/italic',
          'font/underline',
          'font/color',
        ])
        await context.sync()

        return JSON.stringify(
          {
            text: range.text || '',
            style: range.style,
            font: {
              name: range.font.name,
              size: range.font.size,
              bold: range.font.bold,
              italic: range.font.italic,
              underline: range.font.underline,
              color: range.font.color,
            },
          },
          null,
          2,
        )
      })
    },
  },

  selectText: {
    name: 'selectText',
    description: 'Select all text in the document or specific location.',
    inputSchema: {
      type: 'object',
      properties: {
        scope: {
          type: 'string',
          description: 'What to select: "All" for entire document',
          enum: ['All'],
        },
      },
      required: ['scope'],
    },
    execute: async args => {
      const { scope } = args
      return Word.run(async context => {
        if (scope === 'All') {
          context.document.body.getRange().select()
        }
        await context.sync()
        return `${scope === 'All' ? 'All document text selected.' : 'Text selected.'}`
      })
    },
  },

  insertImage: {
    name: 'insertImage',
    description: 'Insert an image from a URL at the current cursor position. The image URL must be accessible.',
    inputSchema: {
      type: 'object',
      properties: {
        imageUrl: {
          type: 'string',
          description: 'The URL of the image to insert',
        },
        width: {
          type: 'number',
          description: 'Optional width in points',
        },
        height: {
          type: 'number',
          description: 'Optional height in points',
        },
        location: {
          type: 'string',
          description: 'Where to insert: "Before", "After", "Start", "End", or "Replace"',
          enum: ['Before', 'After', 'Start', 'End', 'Replace'],
        },
      },
      required: ['imageUrl'],
    },
    execute: async args => {
      const { imageUrl, width, height, location = 'End' } = args
      try {
        const response = await fetch(imageUrl)
        if (!response.ok) throw new Error(`HTTP ${response.status}`)
        const blob = await response.blob()
        const base64 = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader()
          reader.onloadend = () => {
            const result = reader.result as string
            resolve(result.split(',')[1])
          }
          reader.onerror = reject
          reader.readAsDataURL(blob)
        })
        return Word.run(async context => {
          const range = context.document.getSelection()
          const pic = range.insertInlinePictureFromBase64(base64, location)
          if (width) pic.width = width
          if (height) pic.height = height
          await context.sync()
          return `Image inserted from ${imageUrl}`
        })
      } catch (e: any) {
        return `ERROR: Could not load image: ${e.message}`
      }
    },
  },

  getTableInfo: {
    name: 'getTableInfo',
    description: 'Get information about tables in the document, including row and column counts.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const tables = context.document.body.tables
        tables.load(['items'])
        await context.sync()

        const tableInfos = []
        for (let i = 0; i < tables.items.length; i++) {
          const table = tables.items[i]
          table.load(['rowCount', 'values'])
          await context.sync()

          const columnCount = table.values && table.values[0] ? table.values[0].length : 0

          tableInfos.push({
            index: i,
            rowCount: table.rowCount,
            columnCount,
          })
        }

        return JSON.stringify(
          {
            tableCount: tables.items.length,
            tables: tableInfos,
          },
          null,
          2,
        )
      })
    },
  },

  insertBookmark: {
    name: 'insertBookmark',
    description: 'Insert a bookmark at the current selection to mark a location in the document.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'The name of the bookmark (must be unique, no spaces allowed)',
        },
      },
      required: ['name'],
    },
    execute: async args => {
      const { name } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertBookmark(name)
        await context.sync()
        return `Bookmark "${name}" inserted.`
      })
    },
  },

  goToBookmark: {
    name: 'goToBookmark',
    description: 'Navigate to a previously created bookmark in the document.',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'The name of the bookmark to navigate to',
        },
      },
      required: ['name'],
    },
    execute: async args => {
      const { name } = args
      return Word.run(async context => {
        const bookmark = context.document.getBookmarkRange(name)
        bookmark.select()
        await context.sync()
        return `Navigated to bookmark "${name}".`
      })
    },
  },

  insertContentControl: {
    name: 'insertContentControl',
    description:
      'Insert a content control (a container for content) at the current selection. Useful for creating structured documents.',
    inputSchema: {
      type: 'object',
      properties: {
        title: {
          type: 'string',
          description: 'The title of the content control',
        },
        tag: {
          type: 'string',
          description: 'Optional tag for programmatic identification',
        },
        appearance: {
          type: 'string',
          description: 'Visual appearance of the control',
          enum: ['BoundingBox', 'Tags', 'Hidden'],
        },
      },
      required: ['title'],
    },
    execute: async args => {
      const { title, tag, appearance = 'BoundingBox' } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        const cc = range.insertContentControl()
        cc.title = title
        if (tag) cc.tag = tag
        cc.appearance = appearance as any
        await context.sync()
        return `Content control "${title}" inserted.`
      })
    },
  },

  findText: {
    name: 'findText',
    description: 'Find text in the document and return information about matches. Does not modify the document.',
    inputSchema: {
      type: 'object',
      properties: {
        searchText: {
          type: 'string',
          description: 'The text to search for',
        },
        matchCase: {
          type: 'boolean',
          description: 'Whether to match case (default: false)',
        },
        matchWholeWord: {
          type: 'boolean',
          description: 'Whether to match whole word only (default: false)',
        },
      },
      required: ['searchText'],
    },
    execute: async args => {
      const { searchText, matchCase = false, matchWholeWord = false } = args
      return Word.run(async context => {
        const results = context.document.body.search(searchText, { matchCase, matchWholeWord })
        results.load('items')
        await context.sync()
        return JSON.stringify({
          matchCount: results.items.length,
          searchText,
          found: results.items.length > 0,
        }, null, 2)
      })
    },
  },

  insertComment: {
    name: 'insertComment',
    description: 'Insert a comment on the currently selected text or at the cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The comment text to insert',
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertComment(text)
        await context.sync()
        return `Comment inserted: "${text}"`
      })
    },
  },

  getComments: {
    name: 'getComments',
    description: 'Get all comments in the document with their author, content, and reply info.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const comments = context.document.body.getComments()
        comments.load('items')
        await context.sync()

        const commentList = []
        for (let i = 0; i < comments.items.length; i++) {
          const c = comments.items[i]
          c.load(['authorName', 'content', 'contentRange', 'replies'])
          await context.sync()

          const replies = c.replies
          replies.load('items')
          await context.sync()

          commentList.push({
            index: i,
            author: c.authorName,
            content: c.content,
            replyCount: replies.items.length,
          })
        }

        return JSON.stringify({ commentCount: comments.items.length, comments: commentList }, null, 2)
      })
    },
  },

  deleteComment: {
    name: 'deleteComment',
    description: 'Delete a comment by its index (0-based, from getComments).',
    inputSchema: {
      type: 'object',
      properties: {
        commentIndex: {
          type: 'number',
          description: 'Zero-based index of the comment to delete',
        },
      },
      required: ['commentIndex'],
    },
    execute: async args => {
      const { commentIndex } = args
      return Word.run(async context => {
        const comments = context.document.body.getComments()
        comments.load('items')
        await context.sync()

        if (commentIndex < 0 || commentIndex >= comments.items.length) {
          return `ERROR: Comment index ${commentIndex} out of range (${comments.items.length} comments total).`
        }
        comments.items[commentIndex].delete()
        await context.sync()
        return `Comment ${commentIndex} deleted.`
      })
    },
  },

  replyToComment: {
    name: 'replyToComment',
    description: 'Reply to an existing comment by its index.',
    inputSchema: {
      type: 'object',
      properties: {
        commentIndex: {
          type: 'number',
          description: 'Zero-based index of the comment to reply to',
        },
        text: {
          type: 'string',
          description: 'The reply text',
        },
      },
      required: ['commentIndex', 'text'],
    },
    execute: async args => {
      const { commentIndex, text } = args
      return Word.run(async context => {
        const comments = context.document.body.getComments()
        comments.load('items')
        await context.sync()

        if (commentIndex < 0 || commentIndex >= comments.items.length) {
          return `ERROR: Comment index ${commentIndex} out of range.`
        }
        comments.items[commentIndex].reply(text)
        await context.sync()
        return `Reply added to comment ${commentIndex}: "${text}"`
      })
    },
  },

  insertHeader: {
    name: 'insertHeader',
    description: 'Insert text into the primary header of the first section.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to insert into the header',
        },
        alignment: {
          type: 'string',
          description: 'Text alignment: "Left", "Center", or "Right"',
          enum: ['Left', 'Center', 'Right'],
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text, alignment = 'Left' } = args
      return Word.run(async context => {
        const section = context.document.sections.getFirst()
        const header = section.getHeader('Primary')
        const para = header.insertParagraph(text, 'End')
        para.alignment = alignment as any
        await context.sync()
        return `Header set to: "${text}"`
      })
    },
  },

  insertFooter: {
    name: 'insertFooter',
    description: 'Insert text into the primary footer of the first section.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The text to insert into the footer',
        },
        alignment: {
          type: 'string',
          description: 'Text alignment: "Left", "Center", or "Right"',
          enum: ['Left', 'Center', 'Right'],
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text, alignment = 'Center' } = args
      return Word.run(async context => {
        const section = context.document.sections.getFirst()
        const footer = section.getFooter('Primary')
        const para = footer.insertParagraph(text, 'End')
        para.alignment = alignment as any
        await context.sync()
        return `Footer set to: "${text}"`
      })
    },
  },

  getHeaderFooter: {
    name: 'getHeaderFooter',
    description: 'Read the content of headers and footers from all sections.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()

        const result = []
        for (let i = 0; i < sections.items.length; i++) {
          const section = sections.items[i]
          const header = section.getHeader('Primary')
          const footer = section.getFooter('Primary')
          header.load('text')
          footer.load('text')
          await context.sync()
          result.push({
            section: i,
            header: header.text || '(empty)',
            footer: footer.text || '(empty)',
          })
        }
        return JSON.stringify(result, null, 2)
      })
    },
  },

  toggleTrackChanges: {
    name: 'toggleTrackChanges',
    description: 'Enable or disable Track Changes (revision tracking) mode.',
    inputSchema: {
      type: 'object',
      properties: {
        enabled: {
          type: 'boolean',
          description: 'True to enable track changes, false to disable',
        },
      },
      required: ['enabled'],
    },
    execute: async args => {
      const { enabled } = args
      return Word.run(async context => {
        context.document.changeTrackingMode = enabled
          ? Word.ChangeTrackingMode.trackAll
          : Word.ChangeTrackingMode.off
        await context.sync()
        return `Track changes ${enabled ? 'enabled' : 'disabled'}.`
      })
    },
  },

  insertTableOfContents: {
    name: 'insertTableOfContents',
    description: 'Insert an automatic Table of Contents at the current cursor position. Uses heading styles to build the TOC.',
    inputSchema: {
      type: 'object',
      properties: {
        title: {
          type: 'string',
          description: 'Optional TOC title (default: "Table of Contents")',
        },
      },
      required: [],
    },
    execute: async args => {
      const { title = 'Table of Contents' } = args
      return Word.run(async context => {
        const body = context.document.body
        const titlePara = body.insertParagraph(title, 'End')
        try { titlePara.styleBuiltIn = 'TOCHeading' as any } catch { titlePara.style = 'TOC Heading' }
        // TOC placeholder note — actual TOC requires Word field codes not available in Office.js
        const notePara = body.insertParagraph('[Table of Contents will be generated from Heading styles]', 'End')
        notePara.style = 'Normal'
        await context.sync()
        return `TOC placeholder inserted with title "${title}". Update the TOC using Word's References → Update Table after adding content.`
      })
    },
  },

  updateTableOfContents: {
    name: 'updateTableOfContents',
    description: 'Update all Table of Contents fields in the document to reflect current headings.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const body = context.document.body
        const tablesOfContents = (body as any).tableOfContents
        if (!tablesOfContents) return 'No Table of Contents found'
        tablesOfContents.load('items')
        await context.sync()

        for (const toc of tablesOfContents.items) {
          toc.update()
        }
        await context.sync()
        return `Updated ${tablesOfContents.items.length} Table(s) of Contents`
      })
    },
  },

  applyStyle: {
    name: 'applyStyle',
    description:
      'Apply a built-in Word style to the selected paragraph(s). Use listStyles to see available styles.',
    inputSchema: {
      type: 'object',
      properties: {
        styleName: {
          type: 'string',
          description: 'The built-in style name to apply (e.g., "Normal", "Heading1", "Quote", "Title").',
        },
      },
      required: ['styleName'],
    },
    execute: async args => {
      const { styleName } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.paragraphs.load('items')
        await context.sync()
        for (const para of range.paragraphs.items) {
          try { para.styleBuiltIn = styleName as any } catch { para.style = styleName }
        }
        await context.sync()
        return `Style "${styleName}" applied to selected paragraph(s).`
      })
    },
  },

  listStyles: {
    name: 'listStyles',
    description: 'List commonly used built-in Word styles available for applyStyle.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      const styles = [
        'Normal',
        'Heading1',
        'Heading2',
        'Heading3',
        'Heading4',
        'Heading5',
        'Heading6',
        'Title',
        'Subtitle',
        'Quote',
        'IntenseQuote',
        'ListParagraph',
        'NoSpacing',
        'TOCHeading',
      ]
      return `Available built-in styles:\n${styles.join('\n')}`
    },
  },

  insertSectionBreak: {
    name: 'insertSectionBreak',
    description: 'Insert a section break at the current cursor position.',
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          description: 'Type of section break: "NextPage", "Continuous", "EvenPage", or "OddPage"',
          enum: ['NextPage', 'Continuous', 'EvenPage', 'OddPage'],
        },
      },
      required: ['type'],
    },
    execute: async args => {
      const { type = 'NextPage' } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertBreak(`Section${type}` as any, 'After')
        await context.sync()
        return `Section break (${type}) inserted.`
      })
    },
  },

  getSections: {
    name: 'getSections',
    description: 'Get information about document sections including count and header/footer status.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()

        const sectionList = []
        for (let i = 0; i < sections.items.length; i++) {
          const header = sections.items[i].getHeader('Primary')
          const footer = sections.items[i].getFooter('Primary')
          header.load('text')
          footer.load('text')
          await context.sync()
          sectionList.push({
            index: i,
            hasHeader: !!(header.text && header.text.trim()),
            hasFooter: !!(footer.text && footer.text.trim()),
          })
        }
        return JSON.stringify({ sectionCount: sections.items.length, sections: sectionList }, null, 2)
      })
    },
  },

  insertFootnote: {
    name: 'insertFootnote',
    description: 'Insert a footnote at the current cursor position. The footnote text will appear at the bottom of the page.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The footnote text',
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertFootnote(text)
        await context.sync()
        return `Footnote inserted: "${text}"`
      })
    },
  },

  insertEndnote: {
    name: 'insertEndnote',
    description: 'Insert an endnote at the current cursor position. The endnote text will appear at the end of the document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: {
          type: 'string',
          description: 'The endnote text',
        },
      },
      required: ['text'],
    },
    execute: async args => {
      const { text } = args
      return Word.run(async context => {
        const range = context.document.getSelection()
        range.insertEndnote(text)
        await context.sync()
        return `Endnote inserted: "${text}"`
      })
    },
  },

  setPageMargins: {
    name: 'setPageMargins',
    description: 'Set page margins for all sections of the document. Values in points (1 inch = 72 pts).',
    inputSchema: {
      type: 'object',
      properties: {
        top: {
          type: 'number',
          description: 'Top margin in points (default: 72 = 1 inch)',
        },
        bottom: {
          type: 'number',
          description: 'Bottom margin in points (default: 72)',
        },
        left: {
          type: 'number',
          description: 'Left margin in points (default: 90 = 1.25 inch)',
        },
        right: {
          type: 'number',
          description: 'Right margin in points (default: 90)',
        },
      },
      required: [],
    },
    execute: async args => {
      const { top = 72, bottom = 72, left = 90, right = 90 } = args
      return Word.run(async context => {
        const sections = context.document.sections
        sections.load('items')
        await context.sync()
        for (const section of sections.items) {
          section.pageMargin.top = top
          section.pageMargin.bottom = bottom
          section.pageMargin.left = left
          section.pageMargin.right = right
        }
        await context.sync()
        return `Page margins set: top=${top}, bottom=${bottom}, left=${left}, right=${right} pts.`
      })
    },
  },

  insertCoverPage: {
    name: 'insertCoverPage',
    description:
      'Insert a formatted cover page at the VERY TOP of the document, regardless of when this tool is called. ' +
      'Use this for the title, subtitle, and cover metadata of any report or formal document. ' +
      'This tool ALWAYS places content above everything else already in the document. ' +
      'Call this tool ONCE with all cover elements; do NOT use insertParagraph for cover/title content.',
    inputSchema: {
      type: 'object',
      properties: {
        title: {
          type: 'string',
          description: 'Main title (company name, document name, etc.)',
        },
        subtitle: {
          type: 'string',
          description: 'Secondary title line, e.g. "2025年年度财务报告（合并报表）"',
        },
        line3: {
          type: 'string',
          description: 'Third cover line, e.g. report period "报告期：2025年1月1日至12月31日"',
        },
        line4: {
          type: 'string',
          description: 'Fourth cover line, e.g. publication date or stock code',
        },
      },
      required: ['title'],
    },
    execute: async args => {
      const { title, subtitle, line3, line4 } = args
      return Word.run(async context => {
        const body = context.document.body
        // Insert in REVERSE order at Start so the final reading order is top-to-bottom:
        // title → subtitle → line3 → line4 (above whatever sections already exist)
        const lines = [title, subtitle, line3, line4].filter(Boolean) as string[]
        for (let i = lines.length - 1; i >= 0; i--) {
          const para = body.insertParagraph(lines[i], 'Start')
          if (i === 0) {
            try { para.styleBuiltIn = 'Title' as any } catch { para.style = 'Title' }
          } else if (i === 1) {
            try { para.styleBuiltIn = 'Subtitle' as any } catch { para.style = 'Subtitle' }
          }
        }
        await context.sync()
        return `Cover page inserted at document top: "${title}"`
      })
    },
  },

  insertEquation: {
    name: 'insertEquation',
    description:
      'Insert a mathematical equation or formula using LaTeX notation. ' +
      'Supports: Greek letters (\\\\alpha, \\\\beta…), operators (\\\\times, \\\\pm, \\\\leq…), ' +
      '\\\\frac{num}{den} for fractions, \\\\sqrt{x} for radicals, x^{n} for superscripts, x_{i} for subscripts, ' +
      'and combined sub+superscripts. ' +
      'Examples: "ROE = \\\\frac{净利润}{股东权益} \\\\times 100\\\\%"  |  "\\\\Delta R^2 = \\\\sigma^2_{between}"  |  "\\\\sum_{i=1}^{n} x_i". ' +
      'Use displayMode:true for a centred block equation, false (default) for inline.',
    inputSchema: {
      type: 'object',
      properties: {
        latex: {
          type: 'string',
          description: 'LaTeX expression string. Escape backslashes as double-backslash in JSON.',
        },
        displayMode: {
          type: 'boolean',
          description: 'true = centred block equation; false (default) = inline equation on its own paragraph.',
        },
      },
      required: ['latex'],
    },
    execute: async args => {
      const { latex, displayMode = false } = args
      return Word.run(async context => {
        const ooxml = buildEquationOoxml(latex, displayMode)
        context.document.body.insertOoxml(ooxml, 'End')
        await context.sync()
        return `Equation inserted: ${latex.substring(0, 80)}`
      })
    },
  },

  clearDocument: {
    name: 'clearDocument',
    description:
      'Delete ALL content from the document body, leaving a blank document. ' +
      'Use this as the FIRST step when the user asks to create a new document or when content is in the wrong order and you need to start fresh. ' +
      'After calling this, insert content in TOP-TO-BOTTOM reading order: title first, subtitle second, then sections in numbered order.',
    inputSchema: {
      type: 'object',
      properties: {},
      required: [],
    },
    execute: async () => {
      return Word.run(async context => {
        const body = context.document.body
        body.clear()
        await context.sync()
        return 'Document cleared. Now insert content in reading order: title first, then sections top to bottom.'
      })
    },
  },
}

export function createWordTools(enabledTools?: WordToolName[]) {
  const tools = Object.entries(wordToolDefinitions)
    .filter(([name]) => !enabledTools || enabledTools.includes(name as WordToolName))
    .map(([, def]) => {
      const schemaObj: Record<string, z.ZodTypeAny> = {}

      for (const [propName, prop] of Object.entries(def.inputSchema.properties)) {
        let zodType: z.ZodTypeAny

        switch (prop.type) {
          case 'string':
            zodType = prop.enum ? z.enum(prop.enum as [string, ...string[]]) : z.string()
            break
          case 'number':
            zodType = z.number()
            break
          case 'boolean':
            zodType = z.boolean()
            break
          case 'array':
            zodType = z.array(z.any())
            break
          default:
            zodType = z.any()
        }

        if (prop.description) {
          zodType = zodType.describe(prop.description)
        }

        if (!def.inputSchema.required?.includes(propName)) {
          zodType = zodType.optional()
        }

        schemaObj[propName] = zodType
      }

      return tool(
        async input => {
          try {
            // withWordLock serialises concurrent calls from LangGraph's parallel ToolNode,
            // preserving insertion order even when the agent emits multiple tool_calls at once.
            return await withWordLock(() => def.execute(input))
          } catch (error: any) {
            return `Error: ${error.message || 'Unknown error occurred'}`
          }
        },
        {
          name: def.name,
          description: def.description,
          schema: z.object(schemaObj),
        },
      )
    })

  return tools
}

export function getWordToolDefinitions(): WordToolDefinition[] {
  return Object.values(wordToolDefinitions)
}

export function getWordTool(name: WordToolName): WordToolDefinition | undefined {
  return wordToolDefinitions[name]
}

export { wordToolDefinitions }
