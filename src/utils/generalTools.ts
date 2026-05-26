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
    const tryFetch = async (fetchUrl: string, timeout: number): Promise<string | null> => {
      const controller = new AbortController()
      const timer = setTimeout(() => controller.abort(), timeout)
      try {
        const response = await fetch(fetchUrl, {
          signal: controller.signal,
          headers: { Accept: 'text/html,text/plain,*/*' },
        })
        if (!response.ok) return null
        return await response.text()
      } catch {
        return null
      } finally {
        clearTimeout(timer)
      }
    }

    const truncate = (text: string, max = 5000) =>
      text.length > max ? text.substring(0, max) + '...' : text

    // 1. Local dev-server proxy — Node.js, no CORS, tries Jina then direct fetch
    const proxyText = await tryFetch(`/api/fetch?url=${encodeURIComponent(url)}`, 15000)
    if (proxyText && proxyText.length > 50) {
      return `Content from ${url}:\n\n${truncate(proxyText)}`
    }

    // 2. Jina.ai reader direct (fallback if not in dev mode)
    const jinaText = await tryFetch(`https://r.jina.ai/${url}`, 12000)
    if (jinaText && jinaText.length > 100) {
      return `Content from ${url}:\n\n${truncate(jinaText)}`
    }

    return `Error fetching content from ${url}: unable to retrieve content (CORS restrictions or network unavailable)`
  },
})

const searchWebTool = new DynamicStructuredTool({
  name: 'searchWeb',
  description:
    'Searches the web for information. Returns top search results with titles and snippets. ' +
    'Use maxResults=3 for quick facts, maxResults=5 (default) for richer coverage. ' +
    'Avoid requesting more than needed — fewer results means faster response.',
  schema: z.object({
    query: z.string().describe('The search query'),
    maxResults: z.number().optional().default(5).describe('Number of results to return (default: 5, max: 10). Use 3 for simple factual queries.'),
  }),
  func: async ({ query, maxResults = 5 }) => {
    const clampedMax = Math.min(Math.max(1, maxResults), 10)
    const SNIPPET_MAX = 400
    const serperKey = localStorage.getItem('serperAPIKey') || ''

    // 1. Local dev-server proxy — Node.js side, no CORS restrictions, tries Serper→ddgs→DDG-IA
    try {
      const controller = new AbortController()
      const timer = setTimeout(() => controller.abort(), 20000)
      try {
        const response = await fetch(`/api/search?q=${encodeURIComponent(query)}&max=${clampedMax}`, {
          signal: controller.signal,
          headers: serperKey ? { 'X-Serper-Key': serperKey } : {},
        })
        if (response.ok) {
          const data = await response.json()
          if (data.results?.length > 0) {
            return data.results
              .map((r: any) => `[${r.index}] ${r.title}\n${r.url}\n${r.snippet}`)
              .join('\n\n')
          }
        }
      } finally {
        clearTimeout(timer)
      }
    } catch { /* fall through */ }

    // 2. Direct Serper.dev — has proper CORS headers, works from browser
    if (serperKey) {
      const controller = new AbortController()
      const timer = setTimeout(() => controller.abort(), 8000)
      try {
        const response = await fetch('https://google.serper.dev/search', {
          method: 'POST',
          signal: controller.signal,
          headers: { 'X-API-KEY': serperKey, 'Content-Type': 'application/json' },
          body: JSON.stringify({ q: query, num: clampedMax }),
        })
        if (response.ok) {
          const data = await response.json()
          const items: any[] = data.organic || []
          if (items.length > 0) {
            return items
              .slice(0, clampedMax)
              .map((r: any, i: number) => `[${i + 1}] ${r.title}\n${r.link}\n${(r.snippet || '').slice(0, SNIPPET_MAX)}`)
              .join('\n\n')
          }
        }
      } catch { /* fall through */ } finally {
        clearTimeout(timer)
      }
    }

    return 'Search timed out. The search service may be unavailable. Try again or use your existing knowledge.'
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
