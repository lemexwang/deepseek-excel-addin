// Standalone test for the LaTeX → OMML converter logic
// Extracted from src/utils/wordTools.ts (pure functions only, no Word.run)

const escapeXml = s =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
   .replace(/"/g, '&quot;').replace(/'/g, '&apos;')

const LATEX_SYMBOLS = {
  '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
  '\\epsilon': 'ε', '\\varepsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η',
  '\\theta': 'θ', '\\vartheta': 'ϑ', '\\iota': 'ι', '\\kappa': 'κ',
  '\\lambda': 'λ', '\\mu': 'μ', '\\nu': 'ν', '\\xi': 'ξ',
  '\\pi': 'π', '\\varpi': 'ϖ', '\\rho': 'ρ', '\\sigma': 'σ',
  '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ', '\\varphi': 'φ',
  '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
  '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
  '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Upsilon': 'Υ',
  '\\Phi': 'Φ', '\\Psi': 'Ψ', '\\Omega': 'Ω',
  '\\times': '×', '\\div': '÷', '\\pm': '±', '\\mp': '∓',
  '\\cdot': '·', '\\ast': '∗', '\\star': '⋆',
  '\\leq': '≤', '\\le': '≤', '\\geq': '≥', '\\ge': '≥',
  '\\neq': '≠', '\\ne': '≠', '\\approx': '≈', '\\equiv': '≡',
  '\\sim': '∼', '\\propto': '∝', '\\ll': '≪', '\\gg': '≫',
  '\\sum': 'Σ', '\\prod': 'Π', '\\int': '∫', '\\oint': '∮',
  '\\bigcap': '∩', '\\bigcup': '∪',
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
  '\\%': '%', '\\$': '$', '\\#': '#',
  '\\,': '', '\\;': '', '\\:': '', '\\!': '', '\\ ': ' ',
}

const PASSTHROUGH_CMDS = new Set([
  '\\mathrm', '\\mathbf', '\\mathit', '\\mathbb', '\\mathcal', '\\mathsf', '\\mathtt',
  '\\text', '\\textrm', '\\textbf', '\\textit',
  '\\hat', '\\bar', '\\vec', '\\dot', '\\ddot', '\\tilde',
  '\\widehat', '\\widetilde', '\\overline', '\\underline',
  '\\overbrace', '\\underbrace', '\\boldsymbol', '\\operatorname',
])

function tokenizeLaTeX(src) {
  const toks = []
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
  constructor(toks) { this.toks = toks; this.pos = 0 }
  cur() { return this.toks[this.pos] }
  run(text) { return `<m:r><m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>` }

  parseAtom() {
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
        // Skip optional degree [n]
        if (this.cur()?.t === 'lbracket') {
          this.pos++
          while (this.pos < this.toks.length && this.cur()?.t !== 'rbracket') this.pos++
          if (this.cur()?.t === 'rbracket') this.pos++
        }
        const arg = this.parseAtom()
        return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${arg}</m:e></m:rad>`
      }
      if (PASSTHROUGH_CMDS.has(cmd)) return this.parseAtom()
      if (cmd === '\\left' || cmd === '\\right') return ''
      const sym = LATEX_SYMBOLS[cmd]
      if (sym !== undefined) return sym ? this.run(sym) : ''
      return this.run(cmd.slice(1))
    }
    this.pos++
    return ''
  }

  parseNode() {
    const base = this.parseAtom()
    let sub = null, sup = null
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

  parseSequence(insideGroup = false) {
    const parts = []
    while (this.pos < this.toks.length) {
      if (insideGroup && this.cur()?.t === 'rbrace') break
      parts.push(this.parseNode())
    }
    return parts.join('')
  }
}

function latexToOmml(latex) {
  return new OmmlParser(tokenizeLaTeX(latex)).parseSequence()
}

// ── Test runner ─────────────────────────────────────────────────────────────

let passed = 0, failed = 0, warnings = 0

function check(label, latex, assertions) {
  let omml
  try {
    omml = latexToOmml(latex)
  } catch (e) {
    console.error(`❌ CRASH  [${label}]: ${e.message}`)
    failed++
    return
  }

  const results = []
  for (const [desc, test] of Object.entries(assertions)) {
    const ok = test(omml)
    results.push({ desc, ok })
  }

  const allOk = results.every(r => r.ok)
  if (allOk) {
    console.log(`✅ PASS   [${label}]`)
    passed++
  } else {
    console.log(`❌ FAIL   [${label}]: ${latex}`)
    for (const r of results.filter(r => !r.ok)) {
      console.log(`         ✗ ${r.desc}`)
    }
    console.log(`         OMML: ${omml.substring(0, 300)}`)
    failed++
  }
}

function warn(label, latex, issue) {
  console.log(`⚠️  WARN   [${label}]: ${issue}`)
  warnings++
}

console.log('\n=== LaTeX → OMML Parser Tests ===\n')

// T01: Simple Greek letters
check('T01 Greek letters', '\\alpha + \\beta = \\gamma', {
  'contains α': o => o.includes('α'),
  'contains β': o => o.includes('β'),
  'contains γ': o => o.includes('γ'),
  'plus sign preserved': o => o.includes('+'),
})

// T02: Basic fraction
check('T02 Basic fraction', '\\frac{a}{b}', {
  'has m:f element': o => o.includes('<m:f>'),
  'has m:num': o => o.includes('<m:num>'),
  'has m:den': o => o.includes('<m:den>'),
  'numerator is a': o => o.includes('<m:num><m:r><m:t xml:space="preserve">a</m:t></m:r></m:num>'),
  'denominator is b': o => o.includes('<m:den><m:r><m:t xml:space="preserve">b</m:t></m:r></m:den>'),
})

// T03: Fraction with Chinese text (critical for financial reports)
check('T03 Chinese fraction', '\\frac{净利润}{股东权益}', {
  'has fraction': o => o.includes('<m:f>'),
  'numerator has 净利润': o => o.includes('净利润'),
  'denominator has 股东权益': o => o.includes('股东权益'),
})

// T04: Financial formula ROE
check('T04 ROE formula', 'ROE = \\frac{净利润}{平均股东权益} \\times 100\\%', {
  'has ROE': o => o.includes('ROE'),
  'has fraction': o => o.includes('<m:f>'),
  'has × operator': o => o.includes('×'),
  'has 100': o => o.includes('100'),
  'has % sign': o => o.includes('%'),
})

// T05: Superscript
check('T05 Superscript', 'R^{2}', {
  'has sSup': o => o.includes('<m:sSup>'),
  'base is R': o => o.includes('<m:e><m:r><m:t xml:space="preserve">R</m:t></m:r></m:e>'),
  'sup is 2': o => o.includes('<m:sup><m:r><m:t xml:space="preserve">2</m:t></m:r></m:sup>'),
})

// T06: Subscript
check('T06 Subscript', 'x_{i}', {
  'has sSub': o => o.includes('<m:sSub>'),
  'base is x': o => o.includes('>x<'),
  'sub is i': o => o.includes('>i<'),
})

// T07: Combined sub+sup
check('T07 Combined sub+sup', 'x_{i}^{n}', {
  'has sSubSup': o => o.includes('<m:sSubSup>'),
  'has base x': o => o.includes('>x<'),
  'has sub i': o => o.includes('<m:sub>'),
  'has sup n': o => o.includes('<m:sup>'),
})

// T08: Sum with limits
check('T08 Sum with limits', '\\sum_{i=1}^{n} x_i', {
  'has Σ': o => o.includes('Σ'),
  'has sSubSup for sum': o => o.includes('<m:sSubSup>'),
  'has i=1 in sub': o => o.includes('i=1'),
  'has n in sup': o => o.includes('>n<'),
  'has x_i as sSub': o => o.includes('<m:sSub>'),
})

// T09: Square root
check('T09 Square root', '\\sqrt{x^2 + y^2}', {
  'has m:rad': o => o.includes('<m:rad>'),
  'has degHide': o => o.includes('m:degHide'),
  'has x': o => o.includes('>x<'),
  'has y': o => o.includes('>y<'),
})

// T10: sqrt with optional degree [n] — previously BROKEN, now fixed
check('T10 sqrt[n]{x}', '\\sqrt[3]{x}', {
  'has m:rad': o => o.includes('<m:rad>'),
  'radicand is x (not [3])': o => {
    // The [3] degree should be skipped; radicand should be x
    const eContent = o.match(/<m:e>(.*?)<\/m:e>/)?.[1] || ''
    return eContent.includes('>x<') && !eContent.includes('[3]')
  },
  '[3] not in radicand': o => {
    const eContent = o.match(/<m:e>(.*?)<\/m:e>/)?.[1] || ''
    return !eContent.includes('[3]')
  },
})

// T11: \dfrac alias
check('T11 dfrac alias', '\\dfrac{a}{b}', {
  'has fraction': o => o.includes('<m:f>'),
  'same as frac': o => o.includes('<m:num>') && o.includes('<m:den>'),
})

// T12: Delta and subscript (financial)
check('T12 Delta with subscript', '\\Delta R = R_{2024} - R_{2023}', {
  'has Δ': o => o.includes('Δ'),
  'has R': o => o.includes('>R<'),
  'has 2024': o => o.includes('2024'),
  'has 2023': o => o.includes('2023'),
  'has sSub': o => o.includes('<m:sSub>'),
})

// T13: Passthrough commands (\mathrm, \overline)
check('T13 Passthrough \\overline', '\\overline{净利润}', {
  'renders content': o => o.includes('净利润'),
  'no crash on overline': o => typeof o === 'string',
})

// T14: \left \right delimiters
check('T14 \\left \\right', '\\left( \\frac{a}{b} \\right)', {
  'has opening paren': o => o.includes('('),
  'has closing paren': o => o.includes(')'),
  'has fraction': o => o.includes('<m:f>'),
})

// T15: Superscript on closing paren \left(\frac{a}{b}\right)^2
check('T15 Paren with power', '\\left(\\frac{a}{b}\\right)^{2}', {
  'has fraction': o => o.includes('<m:f>'),
  'has superscript 2': o => o.includes('>2<'),
  'has parens': o => o.includes('(') && o.includes(')'),
})

// T16: No-brace single-char subscript
check('T16 No-brace subscript', 'x_i^n', {
  'has sub or sSubSup': o => o.includes('<m:sSub') || o.includes('<m:sSubSup'),
  'has sup or sSubSup': o => o.includes('<m:sSup') || o.includes('<m:sSubSup'),
})

// T17: Escaped percent in financial formula
check('T17 Escaped percent', 'CAGR = \\left(\\frac{V_f}{V_i}\\right)^{\\frac{1}{n}} - 1', {
  'no crash': o => typeof o === 'string' && o.length > 0,
  'has CAGR': o => o.includes('CAGR'),
  'has fraction': o => o.includes('<m:f>'),
})

// T18: Empty input
check('T18 Empty input', '', {
  'returns empty string': o => o === '',
})

// T19: Unknown command — fallback to text
check('T19 Unknown command', '\\unknowncmd', {
  'strips backslash': o => o.includes('unknowncmd') && !o.includes('\\unknowncmd'),
})

// T20: XML special chars in text
check('T20 XML escaping', '\\frac{a < b}{c & d}', {
  'escapes < in numerator': o => o.includes('&lt;'),
  'escapes & in denominator': o => o.includes('&amp;'),
})

// T21: Nested fraction
check('T21 Nested fraction', '\\frac{\\frac{a}{b}}{c}', {
  'outer fraction': o => (o.match(/<m:f>/g) || []).length >= 2,
  'inner fraction in numerator': o => {
    const numContent = o.match(/<m:num>(.*?)<\/m:num>/)?.[1] || ''
    return numContent.includes('<m:f>')
  },
})

// T22: Spacing commands produce nothing
check('T22 Spacing commands', 'a\\,b\\;c\\!d', {
  'contains a b c d': o => o.includes('>a<') && o.includes('>b<') && o.includes('>c<') && o.includes('>d<'),
})

// T23: Full financial balance sheet formula
check('T23 Asset formula', '总资产 = 总负债 + 股东权益', {
  'has 总资产': o => o.includes('总资产'),
  'has +': o => o.includes('+'),
  'has 股东权益': o => o.includes('股东权益'),
})

// T24: Growth rate formula
check('T24 Growth rate', '增长率 = \\frac{本期 - 上期}{上期} \\times 100\\%', {
  'has 增长率': o => o.includes('增长率'),
  'has fraction': o => o.includes('<m:f>'),
  'has ×': o => o.includes('×'),
  'has %': o => o.includes('%'),
})

// T25: OMML namespace check in final OOXML
{
  const ooxml = buildEquation('R^2', false)
  const label = 'T25 OOXML wrapper'
  const checks = {
    'has pkg:package': ooxml.includes('<pkg:package'),
    'has word/document.xml': ooxml.includes('word/document.xml'),
    'has m:oMath': ooxml.includes('<m:oMath'),
    'has math namespace': ooxml.includes('http://schemas.openxmlformats.org/officeDocument/2006/math'),
    'starts with xml declaration': ooxml.startsWith('<?xml'),
  }
  const allOk = Object.values(checks).every(Boolean)
  if (allOk) { console.log(`✅ PASS   [${label}]`); passed++ }
  else {
    console.log(`❌ FAIL   [${label}]`)
    for (const [k, v] of Object.entries(checks)) if (!v) console.log(`         ✗ ${k}`)
    failed++
  }
}

function buildEquation(latex, displayMode) {
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

// T26: Display mode OOXML
{
  const ooxml = buildEquation('\\frac{a}{b}', true)
  const label = 'T26 Display mode'
  const ok = ooxml.includes('<m:oMathPara')
  if (ok) { console.log(`✅ PASS   [${label}]`); passed++ }
  else { console.log(`❌ FAIL   [${label}]: missing m:oMathPara`); failed++ }
}

// ── Summary ─────────────────────────────────────────────────────────────────
console.log(`\n${'─'.repeat(45)}`)
console.log(`Results: ${passed} passed, ${failed} failed, ${warnings} warnings`)
if (failed > 0) process.exit(1)
