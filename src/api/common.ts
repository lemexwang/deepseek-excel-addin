import { Ref } from 'vue'

export function insertResult(result: string, insertType: Ref): void {
  Excel.run(async ctx => {
    const range = ctx.workbook.getSelectedRange()
    range.load(['address', 'rowCount', 'columnCount'])
    await ctx.sync()

    if (insertType.value === 'replace') {
      range.values = [[result]]
    } else if (insertType.value === 'append') {
      // Append to the cell below the selection
      const sheet = ctx.workbook.worksheets.getActiveWorksheet()
      const used = sheet.getUsedRange()
      used.load('rowCount')
      await ctx.sync()
      sheet.getRange(`A${used.rowCount + 1}`).values = [[result]]
    } else {
      // newLine / default: write to selected cell
      range.values = [[result]]
    }
    await ctx.sync()
  }).catch(error => {
    console.error('Excel.run error:', error)
  })
}

export async function insertFormattedResult(result: string, insertType: Ref): Promise<void> {
  // Excel has no markdown formatter — fall back to plain text insertion
  insertResult(result, insertType)
}
