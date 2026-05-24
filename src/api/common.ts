import { Ref } from 'vue'
import { WordFormatter } from '@/utils/wordFormatter'

export function insertResult(result: string, insertType: Ref): void {
  const paragraph = result
    .replace(/(\r\n|\n|\r)/g, '\n')
    .replace(/\n+/g, '\n')
    .split('\n')
  
  Word.run(async context => {
    const range = context.document.getSelection()
    
    if (insertType.value === 'replace') {
      range.insertText(paragraph[0], 'Replace')
      for (let i = 1; i < paragraph.length; i++) {
        range.insertParagraph(paragraph[i], 'After')
      }
    } else if (insertType.value === 'append') {
      range.insertText(paragraph[0], 'End')
      for (let i = 1; i < paragraph.length; i++) {
        range.insertParagraph(paragraph[i], 'After')
      }
    } else if (insertType.value === 'newLine') {
      for (let i = 0; i < paragraph.length; i++) {
        range.insertParagraph(paragraph[i], 'After')
      }
    }
    
    await context.sync()
  }).catch(error => {
    console.error('Word.run error:', error)
  })
}

export async function insertFormattedResult(result: string, insertType: Ref): Promise<void> {
  try {
    await WordFormatter.insertFormattedResult(result, insertType)
  } catch (error) {
    console.warn('Formatted insertion failed, falling back to plain text:', error)
    insertResult(result, insertType)
  }
}
