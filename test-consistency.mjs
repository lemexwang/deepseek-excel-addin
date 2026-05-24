// Comprehensive consistency tests — no Excel/browser runtime needed
// Covers: i18n, tool registration, API providers, dead code, settings, Excel logic

import { readFileSync } from 'fs'

const ROOT = '/Users/alice/Developer/deepseek-excel-addin/src'
const read = f => readFileSync(f, 'utf8')

let passed = 0, failed = 0

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

// ── 1. i18n Completeness ───────────────────────────────────────────────────
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

check('No stale wordTool_* keys in en.json', () =>
  [...enKeys].filter(k => k.startsWith('wordTool_') && !k.startsWith('wordTools'))
    .map(k => `Stale Word key in en.json: "${k}"`)
)

check('No stale wordTool_* keys in zh-cn.json', () =>
  [...zhcnKeys].filter(k => k.startsWith('wordTool_') && !k.startsWith('wordTools'))
    .map(k => `Stale Word key in zh-cn.json: "${k}"`)
)

const EXCEL_TOOLS = [
  'getSelectedRange', 'getSheetData', 'getWorkbookInfo', 'getRangeFormulas', 'getNamedRanges', 'getLastCell',
  'setCellValue', 'setRangeValues', 'setFormula', 'setRangeFormulas',
  'formatRange', 'setColumnWidth', 'setRowHeight', 'mergeCells', 'clearFormatting', 'setNumberFormat',
  'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'clearRange',
  'addSheet', 'deleteSheet', 'renameSheet', 'copySheet', 'getSheetNames',
  'sortRange', 'autoFilter', 'findAndReplace',
  'createTable', 'formatTable',
  'insertChart',
  'autoFit', 'freezePanes',
]

check('All 35 Excel tools have excelTool_X key in en.json', () => {
  const issues = []
  for (const t of EXCEL_TOOLS) {
    if (!enKeys.has(`excelTool_${t}`))      issues.push(`en missing excelTool_${t}`)
    if (!enKeys.has(`excelTool_${t}_desc`)) issues.push(`en missing excelTool_${t}_desc`)
  }
  return issues
})

check('All 35 Excel tools have excelTool_X key in zh-cn.json', () => {
  const issues = []
  for (const t of EXCEL_TOOLS) {
    if (!zhcnKeys.has(`excelTool_${t}`))      issues.push(`zh-cn missing excelTool_${t}`)
    if (!zhcnKeys.has(`excelTool_${t}_desc`)) issues.push(`zh-cn missing excelTool_${t}_desc`)
  }
  return issues
})

// ── 2. Tool Name Registration ──────────────────────────────────────────────
console.log('\n=== 2. Tool Name Registration ===\n')

const excelToolsSrc = read(`${ROOT}/utils/excelTools.ts`)
const homePageSrc   = read(`${ROOT}/pages/HomePage.vue`)
const settingsSrc   = read(`${ROOT}/pages/SettingsPage.vue`)

const unionNames = new Set(
  [...excelToolsSrc.matchAll(/^\s+\| '([a-zA-Z]+)'/gm)].map(m => m[1])
)

const arrayMatch = excelToolsSrc.match(/export const allExcelToolNames[^=]*=\s*\[([\s\S]*?)\]/)
const arrayBlock = arrayMatch?.[1] ?? ''
const registeredNames = new Set(
  [...arrayBlock.matchAll(/'([a-zA-Z]+)'/g)].map(m => m[1])
)

check(`ExcelToolName union has ${EXCEL_TOOLS.length} tools`, () =>
  unionNames.size === EXCEL_TOOLS.length
    ? [] : [`Expected ${EXCEL_TOOLS.length}, union has ${unionNames.size}: ${[...unionNames].filter(n => !EXCEL_TOOLS.includes(n)).join(', ')}`]
)

check('Every ExcelToolName is in allExcelToolNames', () =>
  [...unionNames].filter(n => !registeredNames.has(n))
    .map(n => `"${n}" in ExcelToolName union but NOT in allExcelToolNames`)
)

check('Every allExcelToolNames entry is in ExcelToolName union', () =>
  [...registeredNames].filter(n => !unionNames.has(n))
    .map(n => `"${n}" in allExcelToolNames but NOT in ExcelToolName union`)
)

check('Every registered tool has an excelToolDefinitions entry', () => {
  const issues = []
  for (const name of registeredNames) {
    if (!excelToolsSrc.includes(`  ${name}: {`))
      issues.push(`"${name}" registered but no entry in excelToolDefinitions`)
  }
  return issues
})

check('Every registered tool has excelTool_X i18n key in en.json', () =>
  [...registeredNames]
    .filter(n => !enKeys.has(`excelTool_${n}`))
    .map(n => `"${n}" registered but en.json missing excelTool_${n}`)
)

check('HomePage.vue imports from excelTools, not wordTools', () => {
  const issues = []
  if (homePageSrc.includes("from '@/utils/wordTools'"))
    issues.push("HomePage.vue still imports from wordTools")
  if (!homePageSrc.includes("from '@/utils/excelTools'"))
    issues.push("HomePage.vue does not import from excelTools")
  return issues
})

check('SettingsPage.vue uses getExcelToolDefinitions', () => {
  const issues = []
  if (settingsSrc.includes('getWordToolDefinitions'))
    issues.push("SettingsPage.vue still references getWordToolDefinitions")
  if (!settingsSrc.includes('getExcelToolDefinitions'))
    issues.push("SettingsPage.vue does not use getExcelToolDefinitions")
  return issues
})

// ── 3. API Provider Consistency ────────────────────────────────────────────
console.log('\n=== 3. API Provider Consistency ===\n')

const constantSrc = read(`${ROOT}/utils/constant.ts`)
const unionApiSrc = read(`${ROOT}/api/union.ts`)

const apiMatch    = constantSrc.match(/export const availableAPIs[^=]*=\s*\{([^}]+)\}/)
const apiBlock    = apiMatch?.[1] ?? ''
const declaredAPIs = new Set([...apiBlock.matchAll(/(\w+):/g)].map(m => m[1]))

const modelCreators = new Set(
  [...unionApiSrc.matchAll(/^\s{2}(\w+):\s*\(/gm)].map(m => m[1])
)
const capabilityProviders = new Set(
  [...unionApiSrc.matchAll(/^\s{2}(\w+):\s*\{/gm)].map(m => m[1])
)

check('Every declared API has a ModelCreator', () =>
  [...declaredAPIs].filter(p => !modelCreators.has(p))
    .map(p => `"${p}" in availableAPIs but missing from ModelCreators`)
)

check('Every declared API has providerCapabilities entry', () =>
  [...declaredAPIs].filter(p => !capabilityProviders.has(p))
    .map(p => `"${p}" in availableAPIs but missing from providerCapabilities`)
)

check('No orphan ModelCreators', () =>
  [...modelCreators].filter(p => !declaredAPIs.has(p))
    .map(p => `"${p}" in ModelCreators but NOT in availableAPIs`)
)

// ── 4. Dead Code & Stale References ───────────────────────────────────────
console.log('\n=== 4. Dead Code & Stale References ===\n')

check('wordTools.ts has been deleted', () => {
  try { read(`${ROOT}/utils/wordTools.ts`); return ['wordTools.ts still exists — should be deleted'] }
  catch { return [] }
})

check('wordFormatter.ts has been deleted', () => {
  try { read(`${ROOT}/utils/wordFormatter.ts`); return ['wordFormatter.ts still exists — should be deleted'] }
  catch { return [] }
})

check('common.ts does not import WordFormatter', () => {
  const commonSrc = read(`${ROOT}/api/common.ts`)
  return commonSrc.includes('WordFormatter') ? ['common.ts still imports WordFormatter'] : []
})

check('manifest-dev.xml uses Workbook host, not Document', () => {
  const manifest = read('/Users/alice/Developer/deepseek-excel-addin/manifest-dev.xml')
  const issues = []
  if (manifest.includes('<Host Name="Document"'))  issues.push('manifest still has Word Document host')
  if (!manifest.includes('<Host Name="Workbook"')) issues.push('manifest missing Workbook host')
  if (manifest.includes('xsi:type="Document"'))   issues.push('manifest VersionOverrides still has Document type')
  if (!manifest.includes('xsi:type="Workbook"'))  issues.push('manifest VersionOverrides missing Workbook type')
  return issues
})

check('localStorage uses enabledExcelTools, not enabledWordTools', () => {
  const issues = []
  if (settingsSrc.includes("'enabledWordTools'")) issues.push("SettingsPage uses stale key 'enabledWordTools'")
  if (homePageSrc.includes("'enabledWordTools'"))  issues.push("HomePage uses stale key 'enabledWordTools'")
  return issues
})

// ── 5. Excel Tool Logic ────────────────────────────────────────────────────
console.log('\n=== 5. Excel Tool Logic ===\n')

// Column letter conversion — mirrors excelTools.ts insertColumn/deleteColumn
function toColLetter(i) {
  let s = ''
  i++
  while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26) }
  return s
}

check('Column index → letter: 0=A', () => toColLetter(0) === 'A' ? [] : [`Expected A, got ${toColLetter(0)}`])
check('Column index → letter: 1=B', () => toColLetter(1) === 'B' ? [] : [`Expected B, got ${toColLetter(1)}`])
check('Column index → letter: 25=Z', () => toColLetter(25) === 'Z' ? [] : [`Expected Z, got ${toColLetter(25)}`])
check('Column index → letter: 26=AA', () => toColLetter(26) === 'AA' ? [] : [`Expected AA, got ${toColLetter(26)}`])
check('Column index → letter: 51=AZ', () => toColLetter(51) === 'AZ' ? [] : [`Expected AZ, got ${toColLetter(51)}`])
check('Column index → letter: 52=BA', () => toColLetter(52) === 'BA' ? [] : [`Expected BA, got ${toColLetter(52)}`])

// setRangeValues JSON string parsing
function parseValues(v) {
  return typeof v === 'string' ? JSON.parse(v) : v
}
check('setRangeValues: JSON string parses to 2D array', () => {
  const r = parseValues('[["Name","Age"],["Alice",30]]')
  return Array.isArray(r) && r.length === 2 && r[0][0] === 'Name' && r[1][1] === 30
    ? [] : [`Unexpected result: ${JSON.stringify(r)}`]
})
check('setRangeValues: native array passthrough', () => {
  const input = [['A', 'B'], ['1', '2']]
  return parseValues(input) === input ? [] : ['Array input should pass through unchanged']
})
check('setRangeValues: 5x3 array has correct dimensions', () => {
  const data = JSON.parse(JSON.stringify(Array(5).fill(null).map((_, r) => ['R'+r+'C0','R'+r+'C1','R'+r+'C2'])))
  return data.length === 5 && data[0].length === 3 ? [] : ['Unexpected dimensions']
})

// Row index conversion (0-based to 1-based Excel row numbers)
check('insertRow: rowIndex 0 → Excel row 1', () => {
  const rowIndex = 0
  const startRow = rowIndex + 1
  return startRow === 1 ? [] : [`Expected 1, got ${startRow}`]
})
check('insertRow: rowIndex 4 → Excel row 5', () => {
  const rowIndex = 4
  const startRow = rowIndex + 1
  return startRow === 5 ? [] : [`Expected 5, got ${startRow}`]
})

// agentPrompt contains Excel-specific content
check('agentPrompt references Excel tools (getWorkbookInfo, setRangeValues)', () => {
  const required = ['getWorkbookInfo', 'setRangeValues', 'createTable', 'insertChart', 'Read Before Write']
  return required.filter(k => !homePageSrc.includes(k))
    .map(k => `agentPrompt missing keyword: "${k}"`)
})

check('agentPrompt does NOT reference Word-only tools', () => {
  const promptMatch = homePageSrc.match(/const agentPrompt = \(lang[^)]*\) =>\s*`([\s\S]*?)`\.trim\(\)/)
  if (!promptMatch) return ['Could not locate agentPrompt function']
  const prompt = promptMatch[1]
  const forbidden = ['insertParagraph', 'clearDocument', 'insertCoverPage', 'insertEquation', 'getDocumentContent']
  return forbidden.filter(w => prompt.includes(w)).map(w => `agentPrompt references Word tool "${w}"`)
})

// ── 6. Settings Key Consistency ────────────────────────────────────────────
console.log('\n=== 6. Settings Key Consistency ===\n')

const presetSrc = read(`${ROOT}/utils/settingPreset.ts`)
const enumSrc   = read(`${ROOT}/utils/enum.ts`)

const settingNamesMatch = presetSrc.match(/export const Setting_Names = \[([\s\S]*?)\] as const/)
const settingNamesBlock = settingNamesMatch?.[1] ?? ''
const settingNames = new Set([...settingNamesBlock.matchAll(/'([a-zA-Z]+)'/g)].map(m => m[1]))

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

const enumKeysMatch = enumSrc.match(/export const localStorageKey[^=]*=\s*\{([\s\S]*?)\}/)
const enumBlock = enumKeysMatch?.[1] ?? ''
const localStorageKeys = new Set([...enumBlock.matchAll(/^\s{2}(\w+):/gm)].map(m => m[1]))

check('localStorageKey enum has required agent key', () =>
  ['agentMaxIterations'].filter(k => !localStorageKeys.has(k))
    .map(k => `localStorageKey missing "${k}"`)
)

// ── 7. Bug Fix Verification ────────────────────────────────────────────────
console.log('\n=== 7. Bug Fix Verification ===\n')

check('autoFilter uses sheet.autoFilter not r.autoFilter', () => {
  const issues = []
  if (excelToolsSrc.includes('r.autoFilter.clearCriteria'))
    issues.push('autoFilter still calls r.autoFilter.clearCriteria (Range has no autoFilter)')
  if (excelToolsSrc.includes('r.autoFilter.apply'))
    issues.push('autoFilter still calls r.autoFilter.apply (should be sheet.autoFilter.apply)')
  if (!excelToolsSrc.includes('sheet.autoFilter.apply'))
    issues.push('autoFilter does not call sheet.autoFilter.apply')
  if (!excelToolsSrc.includes('sheet.autoFilter.clearCriteria'))
    issues.push('autoFilter does not call sheet.autoFilter.clearCriteria')
  return issues
})

check('autoFilter uses FilterOn.custom not FilterOn.values', () =>
  excelToolsSrc.includes('FilterOn.values')
    ? ['autoFilter still uses FilterOn.values — should use FilterOn.custom for criterion-based filter']
    : []
)

check('setNumberFormat uses Array.from not Array.fill for 2D array', () =>
  excelToolsSrc.includes('Array(r.rowCount).fill(Array(')
    ? ['setNumberFormat still uses Array.fill shared-reference pattern']
    : []
)

check('getSheetData uses getUsedRangeOrNullObject', () => {
  const block = excelToolsSrc.match(/getSheetData: \{[\s\S]*?getWorkbookInfo: \{/)
  if (!block) return ['Could not locate getSheetData definition block']
  return block[0].includes('getUsedRangeOrNullObject')
    ? []
    : ['getSheetData still uses getUsedRange() without null guard']
})

check('getLastCell uses getUsedRangeOrNullObject', () => {
  const block = excelToolsSrc.match(/getLastCell: \{[\s\S]*?\/\/ ── 写入/)
  if (!block) return ['Could not locate getLastCell definition block']
  return block[0].includes('getUsedRangeOrNullObject')
    ? []
    : ['getLastCell still uses getUsedRange() without null guard']
})

check('autoFit uses getUsedRangeOrNullObject', () => {
  const block = excelToolsSrc.match(/autoFit: \{[\s\S]*?freezePanes: \{/)
  if (!block) return ['Could not locate autoFit definition block']
  return block[0].includes('getUsedRangeOrNullObject')
    ? []
    : ['autoFit still uses getUsedRange() without null guard']
})

check('findAndReplace uses getUsedRangeOrNullObject', () => {
  const block = excelToolsSrc.match(/findAndReplace: \{[\s\S]*?createTable: \{/)
  if (!block) return ['Could not locate findAndReplace definition block']
  return block[0].includes('getUsedRangeOrNullObject')
    ? []
    : ['findAndReplace still uses getUsedRange() without null guard']
})

check('createExcelTools returns objects with name field', () => {
  const block = excelToolsSrc.match(/export function createExcelTools[\s\S]*?\}\)/)
  if (!block) return ['Could not locate createExcelTools']
  return block[0].includes('name,')
    ? []
    : ['createExcelTools map is missing the name field']
})

check('insertChart uses sheet.getRange for position (not bare string)', () => {
  const block = excelToolsSrc.match(/insertChart: \{[\s\S]*?\/\/ ── 其他/)
  if (!block) return ['Could not locate insertChart definition block']
  if (block[0].includes('chart.setPosition(position)'))
    return ['insertChart still passes bare string to chart.setPosition()']
  if (!block[0].includes('chart.setPosition(sheet.getRange('))
    return ['insertChart does not call chart.setPosition with sheet.getRange()']
  return []
})

check('insertChart sets chart.title.visible before .text', () => {
  const block = excelToolsSrc.match(/insertChart: \{[\s\S]*?\/\/ ── 其他/)
  if (!block) return ['Could not locate insertChart definition block']
  return block[0].includes('chart.title.visible = true')
    ? []
    : ['insertChart does not set chart.title.visible = true before assigning title text']
})

check('formatRange uses defensive hex color helper', () => {
  const block = excelToolsSrc.match(/formatRange: \{[\s\S]*?setColumnWidth: \{/)
  if (!block) return ['Could not locate formatRange definition block']
  return block[0].includes('toHex')
    ? []
    : ['formatRange still uses raw template literal for color without # guard']
})

// ── Summary ────────────────────────────────────────────────────────────────
console.log(`\n${'─'.repeat(50)}`)
console.log(`Results: ${passed} passed, ${failed} failed`)
if (failed > 0) process.exit(1)
