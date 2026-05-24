// Comprehensive consistency tests — no Word/browser runtime needed
// Covers: i18n, tool registration, API providers, OOXML builder, dead code

import { readFileSync } from 'fs'

const ROOT = '/Users/alice/Developer/deepseek-word-addin/src'
const read = f => readFileSync(f, 'utf8')

let passed = 0, failed = 0
const results = []

function check(label, fn) {
  try {
    const issues = fn()
    if (!issues || issues.length === 0) {
      console.log(`✅ PASS   ${label}`)
      passed++
    } else {
      console.log(`❌ FAIL   ${label}`)
      issues.forEach(i => console.log(`         ✗ ${i}`))
      failed++
    }
  } catch (e) {
    console.log(`💥 ERROR  ${label}: ${e.message}`)
    failed++
  }
}

// ── 1. i18n completeness ───────────────────────────────────────────────────
console.log('\n=== 1. i18n Completeness ===\n')

const en   = JSON.parse(read(`${ROOT}/i18n/locales/en.json`))
const zhcn = JSON.parse(read(`${ROOT}/i18n/locales/zh-cn.json`))
const enKeys   = new Set(Object.keys(en))
const zhcnKeys = new Set(Object.keys(zhcn))

check('All en keys exist in zh-cn', () =>
  [...enKeys].filter(k => !zhcnKeys.has(k)).map(k => `en has "${k}", missing from zh-cn`)
)

check('All zh-cn keys exist in en', () =>
  [...zhcnKeys].filter(k => !enKeys.has(k)).map(k => `zh-cn has "${k}", missing from en`)
)

// New tools added in this session must have i18n entries
const NEW_TOOLS = ['clearDocument', 'insertCoverPage', 'insertEquation']
check('New tools have wordTool_X i18n keys', () => {
  const issues = []
  for (const t of NEW_TOOLS) {
    if (!enKeys.has(`wordTool_${t}`))   issues.push(`en missing wordTool_${t}`)
    if (!zhcnKeys.has(`wordTool_${t}`)) issues.push(`zh-cn missing wordTool_${t}`)
    if (!enKeys.has(`wordTool_${t}_desc`))   issues.push(`en missing wordTool_${t}_desc`)
    if (!zhcnKeys.has(`wordTool_${t}_desc`)) issues.push(`zh-cn missing wordTool_${t}_desc`)
  }
  return issues
})

// ── 2. Tool name registration ──────────────────────────────────────────────
console.log('\n=== 2. Tool Name Registration ===\n')

const wordToolsSrc = read(`${ROOT}/utils/wordTools.ts`)
const homePageSrc  = read(`${ROOT}/pages/HomePage.vue`)

// Extract WordToolName union members
const unionMatches = [...wordToolsSrc.matchAll(/\| '([a-zA-Z]+)'/g)]
const unionNames   = new Set(unionMatches.map(m => m[1]))

// Extract allWordToolNames array — find the array block
const arrayMatch = homePageSrc.match(/const allWordToolNames[^=]*=\s*\[([^\]]+)\]/)
const arrayBlock = arrayMatch?.[1] ?? ''
const registeredNames = new Set([...arrayBlock.matchAll(/'([a-zA-Z]+)'/g)].map(m => m[1]))

check('Every WordToolName is registered in allWordToolNames', () =>
  [...unionNames].filter(n => !registeredNames.has(n))
    .map(n => `"${n}" in WordToolName union but NOT in allWordToolNames`)
)

check('Every allWordToolNames entry exists in WordToolName union', () =>
  [...registeredNames].filter(n => !unionNames.has(n))
    .map(n => `"${n}" in allWordToolNames but NOT in WordToolName union`)
)

// Every registered tool name should have a wordToolDefinitions entry
check('Every registered tool has a wordToolDefinitions entry', () => {
  const issues = []
  for (const name of registeredNames) {
    // Check for "name: {" or "name," pattern in wordToolDefinitions object
    const inDefs = wordToolsSrc.includes(`  ${name}: {`)
    if (!inDefs) issues.push(`"${name}" registered but no entry in wordToolDefinitions`)
  }
  return issues
})

// Every registered tool should have i18n wordTool_X key
check('Every registered tool has wordTool_X i18n key in en.json', () =>
  [...registeredNames]
    .filter(n => !enKeys.has(`wordTool_${n}`))
    .map(n => `"${n}" registered but en.json missing wordTool_${n}`)
)

// ── 3. API provider consistency ────────────────────────────────────────────
console.log('\n=== 3. API Provider Consistency ===\n')

const constantSrc   = read(`${ROOT}/utils/constant.ts`)
const unionApiSrc   = read(`${ROOT}/api/union.ts`)

// Providers declared in availableAPIs
const apiMatch  = constantSrc.match(/export const availableAPIs[^=]*=\s*\{([^}]+)\}/)
const apiBlock  = apiMatch?.[1] ?? ''
const declaredAPIs = new Set([...apiBlock.matchAll(/(\w+):/g)].map(m => m[1]))

// Providers in ModelCreators
const creatorMatches = [...unionApiSrc.matchAll(/^\s{2}(\w+):\s*\(/gm)]
const modelCreators  = new Set(creatorMatches.map(m => m[1]))

// Providers in providerCapabilities
const capabilityMatches = [...unionApiSrc.matchAll(/^\s{2}(\w+):\s*\{/gm)]
const capabilityProviders = new Set(capabilityMatches.map(m => m[1]))

check('Every declared API has a ModelCreator', () =>
  [...declaredAPIs].filter(p => !modelCreators.has(p))
    .map(p => `"${p}" in availableAPIs but missing from ModelCreators`)
)

check('Every declared API has providerCapabilities entry', () =>
  [...declaredAPIs].filter(p => !capabilityProviders.has(p))
    .map(p => `"${p}" in availableAPIs but missing from providerCapabilities`)
)

check('No orphan ModelCreators (not in availableAPIs)', () =>
  [...modelCreators].filter(p => !declaredAPIs.has(p))
    .map(p => `"${p}" in ModelCreators but NOT in availableAPIs`)
)

// ── 4. Dead code detection ─────────────────────────────────────────────────
console.log('\n=== 4. Dead Code Detection ===\n')

check('constant.ts has no orphan agentPrompt/standardPrompt stubs', () => {
  const issues = []
  if (constantSrc.includes('export const agentPrompt'))
    issues.push('constant.ts exports agentPrompt stub — dead code, never imported')
  if (constantSrc.includes('export const standardPrompt'))
    issues.push('constant.ts exports standardPrompt stub — dead code, never imported')
  return issues
})

// Check that no file imports agentPrompt from constant.ts
check('No file accidentally imports agentPrompt from constant.ts', () => {
  // Already confirmed via grep, but encode as test
  const files = [
    `${ROOT}/pages/SettingsPage.vue`,
    `${ROOT}/api/union.ts`,
    `${ROOT}/utils/common.ts`,
  ]
  const issues = []
  for (const f of files) {
    const src = read(f)
    if (src.includes("'agentPrompt'") || src.match(/import.*agentPrompt.*constant/))
      issues.push(`${f.split('/').pop()} imports agentPrompt from constant`)
  }
  return issues
})

// ── 5. OOXML table builder ─────────────────────────────────────────────────
console.log('\n=== 5. OOXML Table Builder ===\n')

// Inline the pure buildTableOoxml logic for testing
const escapeXml = s =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
   .replace(/"/g, '&quot;').replace(/'/g, '&apos;')

function buildTableOoxml(data, opts = {}) {
  const { headerColor, noBorder = false } = opts
  const applyHeader = headerColor !== undefined
  const hColor = headerColor ?? '2E74B5'
  const cols = data[0]?.length || 0
  const styleId = noBorder ? 'TableNormal' : 'TableGrid'
  const gridCols = Array(cols).fill('<w:gridCol/>').join('')
  const wRows = data.map((rowData, rowIdx) => {
    const isHdr = applyHeader && rowIdx === 0
    const trPr = isHdr ? '<w:trPr><w:tblHeader/></w:trPr>' : ''
    const cells = rowData.map(text => {
      const tcPr = isHdr ? `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${hColor}"/></w:tcPr>` : ''
      const rPr  = isHdr ? '<w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>' : ''
      return `<w:tc>${tcPr}<w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:tc>`
    }).join('')
    return `<w:tr>${trPr}${cells}</w:tr>`
  }).join('')
  const tbl = `<w:tbl><w:tblPr><w:tblStyle w:val="${styleId}"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${wRows}</w:tbl>`
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

function checkOoxml(label, data, opts, assertions) {
  try {
    const xml = buildTableOoxml(data, opts)
    const issues = []
    for (const [desc, test] of Object.entries(assertions)) {
      if (!test(xml)) issues.push(desc)
    }
    if (issues.length === 0) { console.log(`✅ PASS   OOXML: ${label}`); passed++ }
    else {
      console.log(`❌ FAIL   OOXML: ${label}`)
      issues.forEach(i => console.log(`         ✗ ${i}`))
      failed++
    }
  } catch (e) {
    console.log(`💥 ERROR  OOXML: ${label}: ${e.message}`); failed++
  }
}

checkOoxml('Blue header 3x2 table', [
  ['Name', 'Value'],
  ['A', '100'],
  ['B', '200'],
], { headerColor: '2E74B5' }, {
  'has pkg:package wrapper':    x => x.includes('<pkg:package'),
  'has word/document.xml part': x => x.includes('word/document.xml'),
  'has w:tbl element':          x => x.includes('<w:tbl>'),
  'has TableGrid style':        x => x.includes('w:val="TableGrid"'),
  'header has blue fill':       x => x.includes('w:fill="2E74B5"'),
  'header row has tblHeader':   x => x.includes('<w:tblHeader/>'),
  'header text is white+bold':  x => x.includes('<w:b/>') && x.includes('w:val="FFFFFF"'),
  'data rows have no shd':      x => { const rows = x.split('<w:tr>'); return rows.length > 2 && !rows[2].includes('w:fill') },
  'cell text "Name" present':   x => x.includes('>Name<'),
  'cell text "100" present':    x => x.includes('>100<'),
})

checkOoxml('No-border table', [['A', 'B']], { noBorder: true }, {
  'uses TableNormal style': x => x.includes('w:val="TableNormal"'),
  'no TableGrid style':     x => !x.includes('w:val="TableGrid"'),
})

checkOoxml('No header styling', [['H1', 'H2'], ['D1', 'D2']], {}, {
  'no blue fill (headerColor undefined)': x => !x.includes('w:fill='),
  'no tblHeader':                         x => !x.includes('<w:tblHeader/>'),
})

checkOoxml('XML escaping in cells', [['a < b', 'x & y'], ['<tag>', '"quote"']], { headerColor: '2E74B5' }, {
  '< escaped in header': x => x.includes('a &lt; b'),
  '& escaped in header': x => x.includes('x &amp; y'),
  '< escaped in data':   x => x.includes('&lt;tag&gt;'),
  '" escaped in data':   x => x.includes('&quot;quote&quot;'),
})

checkOoxml('Empty cells produce valid XML', [['', '', ''], ['', '', '']], { headerColor: '2E74B5' }, {
  'has w:tbl':                x => x.includes('<w:tbl>'),
  'three gridCol elements':   x => (x.match(/<w:gridCol\/>/g) || []).length === 3,
  'two rows (tr elements)':   x => (x.match(/<w:tr>/g) || []).length === 2,
})

checkOoxml('Chinese text in cells', [['指标', '金额（万元）'], ['营业收入', '12,345']], { headerColor: '2E74B5' }, {
  'Chinese header text preserved': x => x.includes('指标') && x.includes('金额（万元）'),
  'Chinese data text preserved':   x => x.includes('营业收入'),
})

checkOoxml('Large table (10 rows x 5 cols)', [
  ...Array(10).fill(null).map((_, r) => Array(5).fill(null).map((_, c) => `R${r}C${c}`))
], { headerColor: '2E74B5' }, {
  'has 10 rows':             x => (x.match(/<w:tr>/g) || []).length === 10,
  'has 5 gridCol elements':  x => (x.match(/<w:gridCol\/>/g) || []).length === 5,
  'first row has header':    x => x.includes('<w:tblHeader/>'),
})

// ── 6. settingPreset vs enum consistency ───────────────────────────────────
console.log('\n=== 6. Settings Key Consistency ===\n')

const presetSrc = read(`${ROOT}/utils/settingPreset.ts`)
const enumSrc   = read(`${ROOT}/utils/enum.ts`)

// Extract Setting_Names array from settingPreset.ts
const settingNamesMatch = presetSrc.match(/export const Setting_Names = \[([\s\S]*?)\] as const/)
const settingNamesBlock = settingNamesMatch?.[1] ?? ''
const settingNames = new Set([...settingNamesBlock.matchAll(/'([a-zA-Z]+)'/g)].map(m => m[1]))

// Extract settingPreset keys
const presetKeysMatch = presetSrc.match(/export const settingPreset = \{([\s\S]*?)\} as const/)
const presetBlock = presetKeysMatch?.[1] ?? ''
const presetKeys = new Set([...presetBlock.matchAll(/^\s{2}(\w+):/gm)].map(m => m[1]))

check('Every Setting_Name has a settingPreset entry', () =>
  [...settingNames].filter(n => !presetKeys.has(n))
    .map(n => `"${n}" in Setting_Names but NOT in settingPreset`)
)

check('Every settingPreset entry is in Setting_Names', () =>
  [...presetKeys].filter(n => !settingNames.has(n))
    .map(n => `"${n}" in settingPreset but NOT in Setting_Names`)
)

// Extract enum keys from localStorageKey
const enumKeysMatch = enumSrc.match(/export const localStorageKey[^=]*=\s*\{([\s\S]*?)\}/)
const enumBlock = enumKeysMatch?.[1] ?? ''
const enumKeys = new Set([...enumBlock.matchAll(/^\s{2}(\w+):/gm)].map(m => m[1]))

check('localStorageKey enum has required agent key', () =>
  ['agentMaxIterations'].filter(k => !enumKeys.has(k))
    .map(k => `localStorageKey missing "${k}"`)
)

// ── 7. insertParagraph \n normalisation ────────────────────────────────────
console.log('\n=== 7. insertParagraph \\n Normalisation ===\n')

// Replicate the normalisation logic from wordTools.ts
function normaliseText(text) {
  return text.replace(/\\n/g, '\n')
}
function splitSegments(text) {
  return normaliseText(text).split('\n')
}

check('Literal \\\\n becomes real newline', () => {
  const result = normaliseText('hello\\nworld')
  return result === 'hello\nworld' ? [] : [`Expected real newline, got: ${JSON.stringify(result)}`]
})

check('Double \\\\n\\\\n becomes two newlines', () => {
  const result = normaliseText('a\\n\\nb')
  return result === 'a\n\nb' ? [] : [`Got: ${JSON.stringify(result)}`]
})

check('Single-line text produces 1 segment', () => {
  const segs = splitSegments('no newlines here')
  return segs.length === 1 ? [] : [`Expected 1 segment, got ${segs.length}`]
})

check('Two-segment split on literal \\\\n', () => {
  const segs = splitSegments('line one\\nline two')
  return segs.length === 2 && segs[0] === 'line one' && segs[1] === 'line two'
    ? [] : [`Segments: ${JSON.stringify(segs)}`]
})

check('Three-segment split on \\\\n\\\\n\\\\n', () => {
  const segs = splitSegments('a\\nb\\nc')
  return segs.length === 3 ? [] : [`Expected 3, got ${segs.length}: ${JSON.stringify(segs)}`]
})

check('Real newlines in input pass through unchanged', () => {
  const segs = splitSegments('a\nb')
  return segs.length === 2 ? [] : [`Real newline should split too, got ${segs.length} segments`]
})

check('Empty string produces 1 empty segment', () => {
  const segs = splitSegments('')
  return segs.length === 1 && segs[0] === '' ? [] : [`Got: ${JSON.stringify(segs)}`]
})

check('Mixed content: list items with \\\\n prefix', () => {
  const raw = 'Title\\n1. Item one\\n2. Item two\\n3. Item three'
  const segs = splitSegments(raw)
  return segs.length === 4 && segs[0] === 'Title' && segs[3] === '3. Item three'
    ? [] : [`Segments: ${JSON.stringify(segs)}`]
})

// ── Summary ────────────────────────────────────────────────────────────────
console.log(`\n${'─'.repeat(50)}`)
console.log(`Results: ${passed} passed, ${failed} failed`)
if (failed > 0) process.exit(1)
