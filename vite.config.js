import fs from 'node:fs'
import os from 'node:os'
import { fileURLToPath, URL } from 'node:url'

import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'
import json5Plugin from 'vite-plugin-json5'

const certDir = `${os.homedir()}/.office-addin-dev-certs`
const httpsConfig = fs.existsSync(`${certDir}/localhost.crt`)
  ? { key: fs.readFileSync(`${certDir}/localhost.key`), cert: fs.readFileSync(`${certDir}/localhost.crt`) }
  : true

const SNIPPET_MAX = 400
const DDG_UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

const decodeDDGHtml = s =>
  s.replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#x27;/g, "'").replace(/&quot;/g, '"').replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ').trim()

const parseDDGHtml = (html, maxResults) => {
  const titleRe = /class="result__a"[^>]+href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/g
  const snippetRe = /class="result__snippet[^"]*"[^>]*>([\s\S]*?)<\/a>/g
  const titles = []
  const snippets = []
  let m
  while ((m = titleRe.exec(html)) !== null) {
    const uddg = m[1].match(/uddg=([^&]+)/)
    titles.push({ url: uddg ? decodeURIComponent(uddg[1]) : m[1], title: decodeDDGHtml(m[2]) })
  }
  while ((m = snippetRe.exec(html)) !== null) {
    const text = decodeDDGHtml(m[1])
    if (text) snippets.push(text)
  }
  const items = []
  for (let i = 0; i < Math.min(titles.length, snippets.length, maxResults); i++) {
    items.push({ index: i + 1, title: titles[i].title, url: titles[i].url, snippet: snippets[i].substring(0, SNIPPET_MAX) })
  }
  return items
}

const stripHtmlNode = html =>
  html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#x27;/g, "'")
    .replace(/\s+/g, ' ').trim()

// Server-side URL fetch proxy — Node.js, no CORS restrictions
const fetchProxyPlugin = {
  name: 'fetch-proxy',
  configureServer(server) {
    server.middlewares.use(async (req, res, next) => {
      if (!req.url?.startsWith('/api/fetch')) {
        next()
        return
      }
      const urlObj = new URL(req.url, 'https://localhost')
      const targetUrl = urlObj.searchParams.get('url') || ''

      res.setHeader('Content-Type', 'text/plain; charset=utf-8')
      res.setHeader('Cache-Control', 'no-store')

      if (!targetUrl) { res.statusCode = 400; res.end('Missing url'); return }

      // 1. Jina.ai reader — returns clean markdown
      try {
        const resp = await fetch(`https://r.jina.ai/${targetUrl}`, {
          headers: { Accept: 'text/plain,*/*' },
          signal: AbortSignal.timeout(12000),
        })
        if (resp.ok) {
          const text = await resp.text()
          if (text && text.length > 100) { res.end(text.substring(0, 8000)); return }
        }
      } catch {}

      // 2. Direct fetch from Node.js (no CORS)
      try {
        const resp = await fetch(targetUrl, {
          headers: { 'User-Agent': DDG_UA, Accept: 'text/html,text/plain,*/*' },
          signal: AbortSignal.timeout(10000),
        })
        if (resp.ok) {
          const html = await resp.text()
          const text = stripHtmlNode(html)
          if (text.length > 50) { res.end(text.substring(0, 8000)); return }
        }
      } catch {}

      res.statusCode = 502
      res.end(`Unable to fetch: ${targetUrl}`)
    })
  },
}

// Server-side search proxy — Node.js, no CORS restrictions
const searchProxyPlugin = {
  name: 'search-proxy',
  configureServer(server) {
    server.middlewares.use(async (req, res, next) => {
      if (!req.url?.startsWith('/api/search')) {
        next()
        return
      }

      const urlObj = new URL(req.url, 'https://localhost')
      const query = urlObj.searchParams.get('q') || ''
      const maxResults = Math.min(parseInt(urlObj.searchParams.get('max') || '5'), 10)
      const serperKey = req.headers['x-serper-key'] || ''

      res.setHeader('Content-Type', 'application/json; charset=utf-8')
      res.setHeader('Cache-Control', 'no-store')

      if (!query.trim()) {
        res.end(JSON.stringify({ results: [] }))
        return
      }

      // 1. Serper.dev — fastest, most reliable (requires API key)
      if (serperKey) {
        try {
          const resp = await fetch('https://google.serper.dev/search', {
            method: 'POST',
            headers: { 'X-API-KEY': serperKey, 'Content-Type': 'application/json' },
            body: JSON.stringify({ q: query, num: maxResults }),
            signal: AbortSignal.timeout(8000),
          })
          if (resp.ok) {
            const data = await resp.json()
            const items = (data.organic || []).slice(0, maxResults).map((r, i) => ({
              index: i + 1, title: r.title || '', url: r.link || '',
              snippet: (r.snippet || '').substring(0, SNIPPET_MAX),
            }))
            if (items.length > 0) {
              res.end(JSON.stringify({ results: items, source: 'serper' }))
              return
            }
          }
        } catch { /* fall through */ }
      }

      // 2. DuckDuckGo HTML search — real web results, no API key needed
      try {
        const resp = await fetch(
          `https://html.duckduckgo.com/html/?q=${encodeURIComponent(query)}`,
          { headers: { 'User-Agent': DDG_UA, 'Accept-Language': 'en-US,en;q=0.9' }, signal: AbortSignal.timeout(10000) },
        )
        if (resp.ok) {
          const html = await resp.text()
          const items = parseDDGHtml(html, maxResults)
          if (items.length > 0) {
            res.end(JSON.stringify({ results: items, source: 'ddg-html' }))
            return
          }
        }
      } catch { /* fall through */ }

      // 3. DuckDuckGo Instant Answer — fast fallback for factual queries
      try {
        const resp = await fetch(
          `https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_html=1&skip_disambig=1`,
          { signal: AbortSignal.timeout(6000) },
        )
        if (resp.ok) {
          const data = await resp.json()
          const items = []
          if (data.AbstractText) {
            items.push({ index: 1, title: data.Heading || query, url: data.AbstractURL || '', snippet: data.AbstractText.substring(0, SNIPPET_MAX) })
          }
          for (const t of (data.RelatedTopics || []).slice(0, maxResults - items.length)) {
            if (t.Text) items.push({ index: items.length + 1, title: '', url: t.FirstURL || '', snippet: t.Text.substring(0, SNIPPET_MAX) })
          }
          if (items.length > 0) {
            res.end(JSON.stringify({ results: items, source: 'ddg-ia' }))
            return
          }
        }
      } catch { /* fall through */ }

      res.end(JSON.stringify({ results: [] }))
    })
  },
}

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [tailwindcss(), vue(), json5Plugin(), fetchProxyPlugin, searchProxyPlugin],
  server: {
    https: httpsConfig,
    port: 3000,
    headers: {
      'Cache-Control': 'no-store, no-cache, must-revalidate',
      'Pragma': 'no-cache',
    },
  },
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
      async_hook: fileURLToPath(new URL('./async_hook.js', import.meta.url)),
      'node:async_hooks': fileURLToPath(new URL('./async_hook.js', import.meta.url)),
    },
  },
})
