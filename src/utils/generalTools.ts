import { DynamicStructuredTool } from '@langchain/core/tools'
import { evaluate } from 'mathjs'
import { z } from 'zod'

export type GeneralToolName = 'fetchWebContent' | 'searchWeb' | 'getCurrentDate' | 'calculateMath'

export interface GeneralToolDefinition {
  name: GeneralToolName
  description: string
  tool: DynamicStructuredTool
}

const fetchWebContentTool = new DynamicStructuredTool({
  name: 'fetchWebContent',
  description:
    'Fetches content from a given URL. Useful for gathering reference material, quotes, or information to include in the document. Returns the main text content of the webpage.',
  schema: z.object({
    url: z.string().describe('The URL to fetch content from'),
  }),
  func: async ({ url }) => {
    try {
      const response = await fetch(url, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        },
      })

      if (!response.ok) {
        return `Failed to fetch content: ${response.status} ${response.statusText}`
      }

      const html = await response.text()

      const textContent = html
        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()

      const maxLength = 5000
      const content = textContent.length > maxLength ? textContent.substring(0, maxLength) + '...' : textContent

      return `Content from ${url}:\n\n${content}`
    } catch (error: any) {
      return `Error fetching content: ${error.message}`
    }
  },
})

const SNIPPET_MAX = 300

const fetchWithTimeout = (url: string, timeoutMs: number): Promise<Response> => {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  return fetch(url, { signal: controller.signal }).finally(() => clearTimeout(timer))
}

const formatDDGResults = (results: any[], max: number): string =>
  results
    .slice(0, max)
    .map((r: any, i: number) => `[${i + 1}] ${r.title}\n${r.href}\n${(r.body || '').slice(0, SNIPPET_MAX)}`)
    .join('\n\n')

const searchWebTool = new DynamicStructuredTool({
  name: 'searchWeb',
  description:
    'Searches the web for information. Returns top search results with titles and snippets. ' +
    'Use maxResults=3 for quick facts, maxResults=5 (default) for richer coverage.',
  schema: z.object({
    query: z.string().describe('The search query'),
    maxResults: z.number().optional().default(5).describe('Number of results (default 5, max 10). Use 3 for simple facts.'),
  }),
  func: async ({ query, maxResults = 5 }) => {
    const n = Math.min(Math.max(1, maxResults), 10)
    const q = encodeURIComponent(query)

    // Primary: horosama DDG proxy
    try {
      const res = await fetchWithTimeout(
        `https://ddgs.horosama.com/search/text?query=${q}&max_results=${n}`,
        6000,
      )
      if (res.ok) {
        const data = await res.json()
        if (data.results?.length > 0) return formatDDGResults(data.results, n)
      }
    } catch { /* fall through */ }

    // Fallback: DuckDuckGo instant answer API (no key required)
    try {
      const res = await fetchWithTimeout(
        `https://api.duckduckgo.com/?q=${q}&format=json&no_html=1&skip_disambig=1`,
        6000,
      )
      if (res.ok) {
        const data = await res.json()
        const parts: string[] = []
        if (data.Answer) parts.push(`[Direct Answer] ${data.Answer}`)
        if (data.AbstractText) parts.push(`[1] ${data.Heading}\n${data.AbstractURL}\n${data.AbstractText.slice(0, SNIPPET_MAX)}`)
        data.RelatedTopics?.slice(0, n - 1).forEach((t: any, i: number) => {
          if (t.Text) parts.push(`[${parts.length + 1}] ${t.Text.slice(0, SNIPPET_MAX)}\n${t.FirstURL || ''}`)
        })
        if (parts.length > 0) return parts.join('\n\n')
      }
    } catch { /* fall through */ }

    return 'Web search unavailable (both endpoints failed). Use your knowledge or ask the user to provide the information directly.'
  },
})

const getCurrentDateTool = new DynamicStructuredTool({
  name: 'getCurrentDate',
  description:
    'Returns the current date and time. Useful for adding timestamps, dates to documents, or understanding temporal context.',
  schema: z.object({
    format: z
      .enum(['full', 'date', 'time', 'iso'])
      .optional()
      .default('full')
      .describe('Format: "full" (date and time), "date" (date only), "time" (time only), "iso" (ISO 8601)'),
  }),
  func: async ({ format = 'full' }) => {
    const now = new Date()

    switch (format) {
      case 'date':
        return now.toLocaleDateString('en-US', {
          year: 'numeric',
          month: 'long',
          day: 'numeric',
        })
      case 'time':
        return now.toLocaleTimeString('en-US', {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        })
      case 'iso':
        return now.toISOString()
      case 'full':
      default:
        return now.toLocaleString('en-US', {
          year: 'numeric',
          month: 'long',
          day: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        })
    }
  },
})

const calculateMathTool = new DynamicStructuredTool({
  name: 'calculateMath',
  description:
    'Evaluates mathematical expressions safely. Useful for calculations, statistics, or numerical data in documents. Supports basic arithmetic (+, -, *, /), parentheses, and common math functions.',
  schema: z.object({
    expression: z.string().describe('The mathematical expression to evaluate (e.g., "2 + 2 * 3")'),
  }),
  func: async ({ expression }) => {
    try {
      const result = evaluate(expression)

      if (typeof result !== 'number' && typeof result !== 'bigint') {
        return `Calculation completed, but result is not a simple number: ${result}`
      }

      return `${expression} = ${result}`
    } catch (error: any) {
      return `Error evaluating expression: ${error.message}`
    }
  },
})

export const generalToolDefinitions: GeneralToolDefinition[] = [
  {
    name: 'fetchWebContent',
    description: fetchWebContentTool.description,
    tool: fetchWebContentTool,
  },
  {
    name: 'searchWeb',
    description: searchWebTool.description,
    tool: searchWebTool,
  },
  {
    name: 'getCurrentDate',
    description: getCurrentDateTool.description,
    tool: getCurrentDateTool,
  },
  {
    name: 'calculateMath',
    description: calculateMathTool.description,
    tool: calculateMathTool,
  },
]

export function createGeneralTools(enabledTools?: GeneralToolName[]): DynamicStructuredTool[] {
  if (!enabledTools || enabledTools.length === 0) {
    return generalToolDefinitions.map(def => def.tool)
  }

  return generalToolDefinitions.filter(def => enabledTools.includes(def.name)).map(def => def.tool)
}

export function getGeneralToolDefinitions(): GeneralToolDefinition[] {
  return generalToolDefinitions
}

export function getGeneralTool(name: GeneralToolName): GeneralToolDefinition | undefined {
  return generalToolDefinitions.find(def => def.name === name)
}
