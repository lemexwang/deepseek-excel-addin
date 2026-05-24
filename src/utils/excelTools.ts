export type ExcelToolName =
  // 读取
  | 'getSelectedRange'
  | 'getSheetData'
  | 'getWorkbookInfo'
  | 'getRangeFormulas'
  | 'getNamedRanges'
  | 'getLastCell'
  // 写入
  | 'setCellValue'
  | 'setRangeValues'
  | 'setFormula'
  | 'setRangeFormulas'
  // 格式化
  | 'formatRange'
  | 'setColumnWidth'
  | 'setRowHeight'
  | 'mergeCells'
  | 'clearFormatting'
  | 'setNumberFormat'
  // 结构操作
  | 'insertRow'
  | 'deleteRow'
  | 'insertColumn'
  | 'deleteColumn'
  | 'clearRange'
  // 工作表管理
  | 'addSheet'
  | 'deleteSheet'
  | 'renameSheet'
  | 'copySheet'
  | 'getSheetNames'
  // 数据操作
  | 'sortRange'
  | 'autoFilter'
  | 'findAndReplace'
  // Excel 表格
  | 'createTable'
  | 'formatTable'
  // 图表
  | 'insertChart'
  // 其他
  | 'autoFit'
  | 'freezePanes'

// Mutex — prevents LangGraph ToolNode parallel calls from racing on the same workbook
let _excelLock = Promise.resolve()
const withExcelLock = <T>(fn: () => Promise<T>): Promise<T> => {
  const current = _excelLock
  let release!: () => void
  _excelLock = new Promise<void>(res => { release = res })
  return current.then(() => fn().finally(release))
}

interface ExcelToolDefinition {
  description: string
  parameters: {
    type: 'object'
    properties: Record<string, { type: string; description: string; enum?: string[] }>
    required?: string[]
  }
  execute: (args: Record<string, any>) => Promise<string>
}

const getSheet = (ctx: Excel.RequestContext, sheetName?: string) =>
  sheetName
    ? ctx.workbook.worksheets.getItem(sheetName)
    : ctx.workbook.worksheets.getActiveWorksheet()

const excelToolDefinitions: Record<ExcelToolName, ExcelToolDefinition> = {
  // ── 读取 ─────────────────────────────────────────────────────────────────

  getSelectedRange: {
    description: 'Get the currently selected range with its values, formulas, row/column count, and address.',
    parameters: { type: 'object', properties: {} },
    execute: async () =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const range = ctx.workbook.getSelectedRange()
          range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount', 'numberFormat'])
          await ctx.sync()
          return JSON.stringify({
            address: range.address,
            values: range.values,
            formulas: range.formulas,
            rows: range.rowCount,
            cols: range.columnCount,
          })
        })
      ),
  },

  getSheetData: {
    description: 'Get all used data from a worksheet. Prefer getSelectedRange when the user has a selection.',
    parameters: {
      type: 'object',
      properties: {
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
    },
    execute: async ({ sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheet = getSheet(ctx, sheetName)
          const used = sheet.getUsedRange()
          used.load(['address', 'values', 'rowCount', 'columnCount'])
          await ctx.sync()
          return JSON.stringify({
            address: used.address,
            values: used.values,
            rows: used.rowCount,
            cols: used.columnCount,
          })
        })
      ),
  },

  getWorkbookInfo: {
    description: 'Get all worksheet names and the active sheet name. Call this first when exploring an unknown workbook.',
    parameters: { type: 'object', properties: {} },
    execute: async () =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheets = ctx.workbook.worksheets
          sheets.load('items/name')
          const active = ctx.workbook.worksheets.getActiveWorksheet()
          active.load('name')
          await ctx.sync()
          return JSON.stringify({
            sheets: sheets.items.map(s => s.name),
            activeSheet: active.name,
          })
        })
      ),
  },

  getRangeFormulas: {
    description: 'Get formulas and computed values from a specific range.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range, e.g. "A1:C10".' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          r.load(['formulas', 'values'])
          await ctx.sync()
          return JSON.stringify({ range, formulas: r.formulas, values: r.values })
        })
      ),
  },

  getNamedRanges: {
    description: 'List all named ranges defined in the workbook.',
    parameters: { type: 'object', properties: {} },
    execute: async () =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const names = ctx.workbook.names
          names.load('items/name,items/formula')
          await ctx.sync()
          return JSON.stringify(names.items.map(n => ({ name: n.name, formula: n.formula })))
        })
      ),
  },

  getLastCell: {
    description: 'Get the address of the last used cell in a worksheet. Useful for determining data boundaries.',
    parameters: {
      type: 'object',
      properties: {
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
    },
    execute: async ({ sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const last = getSheet(ctx, sheetName).getUsedRange().getLastCell()
          last.load('address')
          await ctx.sync()
          return last.address
        })
      ),
  },

  // ── 写入 ─────────────────────────────────────────────────────────────────

  setCellValue: {
    description: 'Write a single value to a cell.',
    parameters: {
      type: 'object',
      properties: {
        cell: { type: 'string', description: 'A1-notation cell address, e.g. "B3".' },
        value: { type: 'string', description: 'Value to write (string, number, or boolean as string).' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['cell', 'value'],
    },
    execute: async ({ cell, value, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          getSheet(ctx, sheetName).getRange(cell).values = [[value]]
          await ctx.sync()
          return `Set ${cell} = ${value}`
        })
      ),
  },

  setRangeValues: {
    description: 'Write a 2D array of values to a range. Preferred over repeated setCellValue calls.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range, e.g. "A1:D5".' },
        values: { type: 'string', description: 'JSON string of a 2D array (rows × cols), e.g. [["Name","Age"],["Alice",30]].' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range', 'values'],
    },
    execute: async ({ range, values, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const data = typeof values === 'string' ? JSON.parse(values) : values
          getSheet(ctx, sheetName).getRange(range).values = data
          await ctx.sync()
          return `Set ${range} with ${data.length} rows × ${data[0]?.length ?? 0} cols`
        })
      ),
  },

  setFormula: {
    description: 'Write a formula to a cell. Formula must start with "=".',
    parameters: {
      type: 'object',
      properties: {
        cell: { type: 'string', description: 'A1-notation cell address, e.g. "E2".' },
        formula: { type: 'string', description: 'Excel formula starting with "=", e.g. "=SUM(A1:A10)".' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['cell', 'formula'],
    },
    execute: async ({ cell, formula, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          getSheet(ctx, sheetName).getRange(cell).formulas = [[formula]]
          await ctx.sync()
          return `Set formula ${cell} = ${formula}`
        })
      ),
  },

  setRangeFormulas: {
    description: 'Write formulas to a range of cells in bulk.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range, e.g. "E2:E10".' },
        formulas: { type: 'string', description: 'JSON string of a 2D array of formula strings, e.g. [["=B2/C2"],["=B3/C3"]].' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range', 'formulas'],
    },
    execute: async ({ range, formulas, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const data = typeof formulas === 'string' ? JSON.parse(formulas) : formulas
          getSheet(ctx, sheetName).getRange(range).formulas = data
          await ctx.sync()
          return `Set formulas in ${range}`
        })
      ),
  },

  // ── 格式化 ───────────────────────────────────────────────────────────────

  formatRange: {
    description: 'Apply formatting to a range: bold, italic, font size, font color, fill color, alignment, wrap text.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range.' },
        bold: { type: 'boolean', description: 'Set bold.' },
        italic: { type: 'boolean', description: 'Set italic.' },
        fontSize: { type: 'number', description: 'Font size in points.' },
        fontColor: { type: 'string', description: 'Font color as hex without #, e.g. "FF0000" for red.' },
        fillColor: { type: 'string', description: 'Fill/background color as hex without #, e.g. "2E74B5" for blue.' },
        horizontalAlignment: { type: 'string', description: 'Alignment: "Left", "Center", "Right".' },
        wrapText: { type: 'boolean', description: 'Wrap text in cells.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, bold, italic, fontSize, fontColor, fillColor, horizontalAlignment, wrapText, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          if (bold !== undefined) r.format.font.bold = bold
          if (italic !== undefined) r.format.font.italic = italic
          if (fontSize !== undefined) r.format.font.size = fontSize
          if (fontColor !== undefined) r.format.font.color = `#${fontColor}`
          if (fillColor !== undefined) r.format.fill.color = `#${fillColor}`
          if (horizontalAlignment !== undefined) r.format.horizontalAlignment = horizontalAlignment as Excel.HorizontalAlignment
          if (wrapText !== undefined) r.format.wrapText = wrapText
          await ctx.sync()
          return `Formatted ${range}`
        })
      ),
  },

  setColumnWidth: {
    description: 'Set the width of one or more columns.',
    parameters: {
      type: 'object',
      properties: {
        column: { type: 'string', description: 'Column letter or range, e.g. "A" or "A:C".' },
        width: { type: 'number', description: 'Column width in points.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['column', 'width'],
    },
    execute: async ({ column, width, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const col = column.includes(':') ? column : `${column}:${column}`
          getSheet(ctx, sheetName).getRange(col).format.columnWidth = width
          await ctx.sync()
          return `Set column ${column} width to ${width}`
        })
      ),
  },

  setRowHeight: {
    description: 'Set the height of a row.',
    parameters: {
      type: 'object',
      properties: {
        row: { type: 'number', description: '1-based row number.' },
        height: { type: 'number', description: 'Row height in points.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['row', 'height'],
    },
    execute: async ({ row, height, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          getSheet(ctx, sheetName).getRange(`${row}:${row}`).format.rowHeight = height
          await ctx.sync()
          return `Set row ${row} height to ${height}`
        })
      ),
  },

  mergeCells: {
    description: 'Merge cells in a range.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range to merge.' },
        across: { type: 'boolean', description: 'If true, merge each row independently. Default false (merge all into one).' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, across = false, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          getSheet(ctx, sheetName).getRange(range).merge(across)
          await ctx.sync()
          return `Merged ${range}`
        })
      ),
  },

  clearFormatting: {
    description: 'Clear all formatting (fill, font, borders) from a range.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          r.clear(Excel.ClearApplyTo.formats)
          await ctx.sync()
          return `Cleared formatting in ${range}`
        })
      ),
  },

  setNumberFormat: {
    description: 'Set number format for a range. Examples: "0.00%" for percentage, "#,##0" for comma integer, "yyyy-mm-dd" for date, "@" for text.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range.' },
        format: { type: 'string', description: 'Excel number format string.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range', 'format'],
    },
    execute: async ({ range, format, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          r.load('rowCount,columnCount')
          await ctx.sync()
          r.numberFormat = Array(r.rowCount).fill(Array(r.columnCount).fill(format))
          await ctx.sync()
          return `Set number format in ${range} to "${format}"`
        })
      ),
  },

  // ── 结构操作 ─────────────────────────────────────────────────────────────

  insertRow: {
    description: 'Insert one or more rows above a given row index (0-based).',
    parameters: {
      type: 'object',
      properties: {
        rowIndex: { type: 'number', description: '0-based row index. Rows will be inserted above this row.' },
        count: { type: 'number', description: 'Number of rows to insert. Default 1.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['rowIndex'],
    },
    execute: async ({ rowIndex, count = 1, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const startRow = rowIndex + 1
          const endRow = startRow + count - 1
          getSheet(ctx, sheetName).getRange(`${startRow}:${endRow}`).insert(Excel.InsertShiftDirection.down)
          await ctx.sync()
          return `Inserted ${count} row(s) above row ${startRow}`
        })
      ),
  },

  deleteRow: {
    description: 'Delete one or more rows starting at a given row index (0-based).',
    parameters: {
      type: 'object',
      properties: {
        rowIndex: { type: 'number', description: '0-based row index of the first row to delete.' },
        count: { type: 'number', description: 'Number of rows to delete. Default 1.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['rowIndex'],
    },
    execute: async ({ rowIndex, count = 1, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const startRow = rowIndex + 1
          const endRow = startRow + count - 1
          getSheet(ctx, sheetName).getRange(`${startRow}:${endRow}`).delete(Excel.DeleteShiftDirection.up)
          await ctx.sync()
          return `Deleted ${count} row(s) starting at row ${startRow}`
        })
      ),
  },

  insertColumn: {
    description: 'Insert one or more columns to the left of a given column index (0-based, A=0).',
    parameters: {
      type: 'object',
      properties: {
        columnIndex: { type: 'number', description: '0-based column index (A=0, B=1, ...). Columns inserted to the left.' },
        count: { type: 'number', description: 'Number of columns to insert. Default 1.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['columnIndex'],
    },
    execute: async ({ columnIndex, count = 1, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const toColLetter = (i: number) => {
            let s = ''
            i++
            while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26) }
            return s
          }
          const startCol = toColLetter(columnIndex)
          const endCol = toColLetter(columnIndex + count - 1)
          getSheet(ctx, sheetName).getRange(`${startCol}:${endCol}`).insert(Excel.InsertShiftDirection.right)
          await ctx.sync()
          return `Inserted ${count} column(s) to the left of ${startCol}`
        })
      ),
  },

  deleteColumn: {
    description: 'Delete one or more columns starting at a given column index (0-based, A=0).',
    parameters: {
      type: 'object',
      properties: {
        columnIndex: { type: 'number', description: '0-based column index (A=0, B=1, ...).' },
        count: { type: 'number', description: 'Number of columns to delete. Default 1.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['columnIndex'],
    },
    execute: async ({ columnIndex, count = 1, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const toColLetter = (i: number) => {
            let s = ''
            i++
            while (i > 0) { s = String.fromCharCode(64 + (i % 26 || 26)) + s; i = Math.floor((i - 1) / 26) }
            return s
          }
          const startCol = toColLetter(columnIndex)
          const endCol = toColLetter(columnIndex + count - 1)
          getSheet(ctx, sheetName).getRange(`${startCol}:${endCol}`).delete(Excel.DeleteShiftDirection.left)
          await ctx.sync()
          return `Deleted ${count} column(s) starting at ${startCol}`
        })
      ),
  },

  clearRange: {
    description: 'Clear contents, formats, or everything from a range.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range.' },
        clearType: { type: 'string', description: '"contents" (default), "formats", or "all".', enum: ['contents', 'formats', 'all'] },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, clearType = 'contents', sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          const applyTo =
            clearType === 'formats' ? Excel.ClearApplyTo.formats :
            clearType === 'all' ? Excel.ClearApplyTo.all :
            Excel.ClearApplyTo.contents
          r.clear(applyTo)
          await ctx.sync()
          return `Cleared ${clearType} from ${range}`
        })
      ),
  },

  // ── 工作表管理 ───────────────────────────────────────────────────────────

  addSheet: {
    description: 'Add a new worksheet.',
    parameters: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Name for the new worksheet.' },
        position: { type: 'number', description: '0-based position index. Omit to add at end.' },
      },
      required: ['name'],
    },
    execute: async ({ name, position }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const newSheet = ctx.workbook.worksheets.add(name)
          if (position !== undefined) newSheet.position = position
          await ctx.sync()
          return `Added sheet "${name}"`
        })
      ),
  },

  deleteSheet: {
    description: 'Delete a worksheet by name.',
    parameters: {
      type: 'object',
      properties: {
        sheetName: { type: 'string', description: 'Name of the worksheet to delete.' },
      },
      required: ['sheetName'],
    },
    execute: async ({ sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          ctx.workbook.worksheets.getItem(sheetName).delete()
          await ctx.sync()
          return `Deleted sheet "${sheetName}"`
        })
      ),
  },

  renameSheet: {
    description: 'Rename a worksheet.',
    parameters: {
      type: 'object',
      properties: {
        oldName: { type: 'string', description: 'Current worksheet name.' },
        newName: { type: 'string', description: 'New worksheet name.' },
      },
      required: ['oldName', 'newName'],
    },
    execute: async ({ oldName, newName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          ctx.workbook.worksheets.getItem(oldName).name = newName
          await ctx.sync()
          return `Renamed sheet "${oldName}" → "${newName}"`
        })
      ),
  },

  copySheet: {
    description: 'Copy a worksheet and give the copy a new name.',
    parameters: {
      type: 'object',
      properties: {
        sourceName: { type: 'string', description: 'Name of the worksheet to copy.' },
        newName: { type: 'string', description: 'Name for the copied worksheet.' },
      },
      required: ['sourceName', 'newName'],
    },
    execute: async ({ sourceName, newName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const copy = ctx.workbook.worksheets.getItem(sourceName).copy(Excel.WorksheetPositionType.end)
          copy.name = newName
          await ctx.sync()
          return `Copied sheet "${sourceName}" → "${newName}"`
        })
      ),
  },

  getSheetNames: {
    description: 'List all worksheet names in the workbook.',
    parameters: { type: 'object', properties: {} },
    execute: async () =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheets = ctx.workbook.worksheets
          sheets.load('items/name')
          await ctx.sync()
          return JSON.stringify(sheets.items.map(s => s.name))
        })
      ),
  },

  // ── 数据操作 ─────────────────────────────────────────────────────────────

  sortRange: {
    description: 'Sort a range by a specified column.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range to sort.' },
        columnIndex: { type: 'number', description: '0-based column index within the range to sort by.' },
        ascending: { type: 'boolean', description: 'Sort ascending (true, default) or descending (false).' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range', 'columnIndex'],
    },
    execute: async ({ range, columnIndex, ascending = true, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          getSheet(ctx, sheetName).getRange(range).sort.apply([{ key: columnIndex, ascending }])
          await ctx.sync()
          return `Sorted ${range} by column index ${columnIndex} (${ascending ? 'ascending' : 'descending'})`
        })
      ),
  },

  autoFilter: {
    description: 'Apply or clear AutoFilter on a range. Omit criteria to show all rows.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range (should include header row).' },
        columnIndex: { type: 'number', description: '0-based column index within the range to filter.' },
        criteria: { type: 'string', description: 'Filter value. Omit to apply filter without criteria.' },
        clear: { type: 'boolean', description: 'If true, clear all filter criteria and show all rows.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, columnIndex = 0, criteria, clear = false, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const r = getSheet(ctx, sheetName).getRange(range)
          if (clear) {
            r.autoFilter.clearCriteria()
          } else if (criteria) {
            r.autoFilter.apply(r, columnIndex, { criterion1: criteria, filterOn: Excel.FilterOn.values })
          } else {
            r.autoFilter.apply(r, columnIndex)
          }
          await ctx.sync()
          return clear
            ? `Cleared filter on ${range}`
            : `Applied filter on column ${columnIndex} of ${range}${criteria ? ` for "${criteria}"` : ''}`
        })
      ),
  },

  findAndReplace: {
    description: 'Find and replace text in the active worksheet.',
    parameters: {
      type: 'object',
      properties: {
        findText: { type: 'string', description: 'Text to find.' },
        replaceText: { type: 'string', description: 'Replacement text.' },
        matchCase: { type: 'boolean', description: 'Case-sensitive match. Default false.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['findText', 'replaceText'],
    },
    execute: async ({ findText, replaceText, matchCase = false, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const count = getSheet(ctx, sheetName)
            .getUsedRange()
            .replaceAll(findText, replaceText, { matchCase, matchEntireCellContents: false })
          await ctx.sync()
          return `Replaced "${findText}" with "${replaceText}" (${count.value} occurrence(s))`
        })
      ),
  },

  // ── Excel 表格 ───────────────────────────────────────────────────────────

  createTable: {
    description: 'Convert a range to an Excel Table object with optional headers.',
    parameters: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'A1-notation range including headers.' },
        hasHeaders: { type: 'boolean', description: 'Whether the first row is a header row. Default true.' },
        tableName: { type: 'string', description: 'Optional name for the table.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['range'],
    },
    execute: async ({ range, hasHeaders = true, tableName, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const table = getSheet(ctx, sheetName).tables.add(range, hasHeaders)
          if (tableName) table.name = tableName
          await ctx.sync()
          return `Created table "${tableName ?? 'Table'}" from ${range}`
        })
      ),
  },

  formatTable: {
    description: 'Apply a built-in style to an Excel Table. Styles: TableStyleMedium2 (recommended), TableStyleLight1–21, TableStyleDark1–11.',
    parameters: {
      type: 'object',
      properties: {
        tableName: { type: 'string', description: 'Table name. Omit to use the first table on the sheet.' },
        style: { type: 'string', description: 'Table style name. Default "TableStyleMedium2".' },
        showBandedRows: { type: 'boolean', description: 'Show alternating row colors.' },
        showHeaderRow: { type: 'boolean', description: 'Show or hide the header row.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
    },
    execute: async ({ tableName, style = 'TableStyleMedium2', showBandedRows, showHeaderRow, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheet = getSheet(ctx, sheetName)
          const table = tableName
            ? ctx.workbook.tables.getItem(tableName)
            : sheet.tables.getItemAt(0)
          table.style = style
          if (showBandedRows !== undefined) table.showBandedRows = showBandedRows
          if (showHeaderRow !== undefined) table.showHeaders = showHeaderRow
          await ctx.sync()
          return `Applied style "${style}" to table`
        })
      ),
  },

  // ── 图表 ─────────────────────────────────────────────────────────────────

  insertChart: {
    description: 'Insert a chart from a data range. chartType options: "ColumnClustered", "Line", "Pie", "Bar", "Area", "XYScatter".',
    parameters: {
      type: 'object',
      properties: {
        dataRange: { type: 'string', description: 'A1-notation range containing chart data (including headers).' },
        chartType: { type: 'string', description: 'Chart type: ColumnClustered, Line, Pie, Bar, Area, XYScatter.', enum: ['ColumnClustered', 'Line', 'Pie', 'Bar', 'Area', 'XYScatter'] },
        title: { type: 'string', description: 'Chart title.' },
        position: { type: 'string', description: 'Cell address for chart top-left corner, e.g. "E2". Default "E2".' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['dataRange', 'chartType'],
    },
    execute: async ({ dataRange, chartType, title, position = 'E2', sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheet = getSheet(ctx, sheetName)
          const typeMap: Record<string, Excel.ChartType> = {
            ColumnClustered: Excel.ChartType.columnClustered,
            Line: Excel.ChartType.line,
            Pie: Excel.ChartType.pie,
            Bar: Excel.ChartType.barClustered,
            Area: Excel.ChartType.area,
            XYScatter: Excel.ChartType.xyscatter,
          }
          const chart = sheet.charts.add(
            typeMap[chartType] ?? Excel.ChartType.columnClustered,
            sheet.getRange(dataRange),
            Excel.ChartSeriesBy.auto
          )
          if (title) chart.title.text = title
          chart.setPosition(position)
          await ctx.sync()
          return `Inserted ${chartType} chart from ${dataRange}`
        })
      ),
  },

  // ── 其他 ─────────────────────────────────────────────────────────────────

  autoFit: {
    description: 'Auto-fit column widths and/or row heights to content.',
    parameters: {
      type: 'object',
      properties: {
        target: { type: 'string', description: '"columns", "rows", or "both".', enum: ['columns', 'rows', 'both'] },
        range: { type: 'string', description: 'A1-notation range. Omit to auto-fit the entire used range.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
      required: ['target'],
    },
    execute: async ({ target, range, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheet = getSheet(ctx, sheetName)
          const r = range ? sheet.getRange(range) : sheet.getUsedRange()
          if (target === 'columns' || target === 'both') r.format.autofitColumns()
          if (target === 'rows' || target === 'both') r.format.autofitRows()
          await ctx.sync()
          return `AutoFit ${target} in ${range ?? 'used range'}`
        })
      ),
  },

  freezePanes: {
    description: 'Freeze rows/columns at a cell, freeze only the first row, or unfreeze all panes.',
    parameters: {
      type: 'object',
      properties: {
        cell: { type: 'string', description: 'Freeze at this cell (e.g. "B2" freezes row 1 and column A). Omit to freeze only row 1.' },
        unfreeze: { type: 'boolean', description: 'If true, unfreeze all panes.' },
        sheetName: { type: 'string', description: 'Worksheet name. Omit to use the active sheet.' },
      },
    },
    execute: async ({ cell, unfreeze = false, sheetName }) =>
      withExcelLock(() =>
        Excel.run(async ctx => {
          const sheet = getSheet(ctx, sheetName)
          if (unfreeze) {
            sheet.freezePanes.unfreeze()
          } else if (cell) {
            sheet.freezePanes.freezeAt(sheet.getRange(cell))
          } else {
            sheet.freezePanes.freezeRows(1)
          }
          await ctx.sync()
          return unfreeze ? 'Unfroze all panes' : `Froze panes at ${cell ?? 'row 1'}`
        })
      ),
  },
}

// ── Exports ───────────────────────────────────────────────────────────────────

export const allExcelToolNames: ExcelToolName[] = [
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

export function createExcelTools(enabledTools?: ExcelToolName[]) {
  return Object.entries(excelToolDefinitions)
    .filter(([name]) => !enabledTools || enabledTools.includes(name as ExcelToolName))
    .map(([, def]) => ({
      description: def.description,
      schema: { type: 'object' as const, properties: def.parameters.properties, required: def.parameters.required ?? [] },
      func: def.execute,
    }))
}

export function getExcelTool(name: ExcelToolName): ExcelToolDefinition | undefined {
  return excelToolDefinitions[name]
}

export function getExcelToolDefinitions() {
  return Object.entries(excelToolDefinitions).map(([name, def]) => ({
    name,
    description: def.description,
    parameters: def.parameters,
  }))
}
