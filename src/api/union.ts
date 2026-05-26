import { BaseChatModel } from '@langchain/core/language_models/chat_models'
import { ChatGoogleGenerativeAI } from '@langchain/google-genai'
import { ChatGroq } from '@langchain/groq'
// import { MemorySaver } from '@langchain/langgraph'
import { ChatOllama } from '@langchain/ollama'
import { AzureChatOpenAI, ChatOpenAI } from '@langchain/openai'
import { createAgent } from 'langchain'

import { IndexedDBSaver } from '@/api/checkpoints'

export const providerCapabilities: Record<
  string,
  { supportsFunctionCalling: boolean; supportsVision: boolean; supportsWebSearch: boolean }
> = {
  deepseek: { supportsFunctionCalling: true, supportsVision: false, supportsWebSearch: false },
  official: { supportsFunctionCalling: true, supportsVision: true, supportsWebSearch: false },
  gemini: { supportsFunctionCalling: true, supportsVision: true, supportsWebSearch: true },
  groq: { supportsFunctionCalling: true, supportsVision: false, supportsWebSearch: false },
  ollama: { supportsFunctionCalling: true, supportsVision: false, supportsWebSearch: false },
  azure: { supportsFunctionCalling: true, supportsVision: true, supportsWebSearch: false },
}

import {
  AgentOptions,
  AzureOptions,
  DeepseekOptions,
  GeminiOptions,
  GroqOptions,
  OllamaOptions,
  OpenAIOptions,
  ProviderOptions,
} from './types'

const ModelCreators: Record<string, (opts: any) => BaseChatModel> = {
  deepseek: (opts: DeepseekOptions) => {
    const modelName = opts.model || 'deepseek-chat'
    return new ChatOpenAI({
      modelName,
      configuration: {
        apiKey: opts.config.apiKey,
        baseURL: 'https://api.deepseek.com/v1',
      },
      temperature: opts.temperature ?? 0.7,
      maxTokens: opts.maxTokens ?? 4096,
    })
  },

  official: (opts: OpenAIOptions) => {
    const modelName = opts.model || 'gpt-5'
    return new ChatOpenAI({
      modelName,
      configuration: {
        apiKey: opts.config.apiKey,
        baseURL: opts.config.baseURL || 'https://api.openai.com/v1',
      },
      temperature: opts.temperature ?? 0.7,
      maxTokens: opts.maxTokens ?? 800,
    })
  },

  ollama: (opts: OllamaOptions) => {
    return new ChatOllama({
      model: opts.ollamaModel,
      baseUrl: opts.ollamaEndpoint?.replace(/\/$/, '') || 'http://localhost:11434',
      temperature: opts.temperature,
    })
  },

  groq: (opts: GroqOptions) => {
    return new ChatGroq({
      model: opts.groqModel,
      apiKey: opts.groqAPIKey,
      temperature: opts.temperature ?? 0.5,
      maxTokens: opts.maxTokens ?? 1024,
    })
  },

  gemini: (opts: GeminiOptions) => {
    return new ChatGoogleGenerativeAI({
      model: opts.geminiModel ?? 'gemini-3-pro-preview',
      apiKey: opts.geminiAPIKey,
      temperature: opts.temperature ?? 0.7,
      maxOutputTokens: opts.maxTokens ?? 800,
    })
  },

  azure: (opts: AzureOptions) => {
    return new AzureChatOpenAI({
      model: opts.azureDeploymentName,
      temperature: opts.temperature ?? 0.7,
      maxTokens: opts.maxTokens ?? 800,
      azureOpenAIApiKey: opts.azureAPIKey,
      azureOpenAIEndpoint: opts.azureAPIEndpoint,
      azureOpenAIApiDeploymentName: opts.azureDeploymentName,
      azureOpenAIApiVersion: opts.azureAPIVersion ?? '2024-10-01',
    })
  },
}

// const checkpointer = new MemorySaver()
const checkpointer = new IndexedDBSaver()

async function executeChatFlow(model: BaseChatModel, options: ProviderOptions): Promise<void> {
  try {
    if (!options.threadId) {
      options.threadId = crypto.randomUUID()
      console.log(`[Chat] New thread started: ${options.threadId}`)
    }
    const agent = createAgent({
      model,
      tools: [],
      checkpointer,
    })
    const stream = await agent.stream(
      {
        messages: options.messages,
      },
      {
        signal: options.abortSignal,
        configurable: { thread_id: options.threadId },
        streamMode: 'messages',
      },
    )

    let fullContent = ''
    for await (const chunk of stream) {
      if (options.abortSignal?.aborted) {
        break
      }

      const content = typeof chunk[0].content === 'string' ? chunk[0].content : ''
      fullContent += content
      options.onStream(fullContent)
    }
  } catch (error: any) {
    if (error.name === 'AbortError' || options.abortSignal?.aborted) {
      // Don't mark as error if intentionally aborted
      throw error
    }
    options.errorIssue.value = true
    console.error(error)
  } finally {
    options.loading.value = false
  }
}

async function executeAgentFlow(model: BaseChatModel, options: AgentOptions): Promise<void> {
  try {
    if (!options.threadId) {
      options.threadId = crypto.randomUUID()
      console.log(`[Agent] New thread started: ${options.threadId}`)
    }
    const agent = createAgent({
      model,
      tools: options.tools || [],
      checkpointer,
    })

    const stream = await agent.stream(
      { messages: options.messages },
      {
        recursionLimit: Number(options.recursionLimit),
        signal: options.abortSignal,
        configurable: {
          thread_id: options.threadId,
          checkpoint_id: options.checkpointId,
        },
        streamMode: 'messages',
      },
    )

    const announcedToolIds = new Set<string>()
    let currentContent = ''

    for await (const chunk of stream) {
      if (options.abortSignal?.aborted) break

      const msg = chunk[0] as any
      if (!msg) continue
      const msgType = msg._getType?.()

      // Detect tool calls from streaming AI chunks (fire once per tool call id)
      if (msgType === 'ai' && msg.tool_call_chunks?.length > 0) {
        for (const tc of msg.tool_call_chunks) {
          if (tc.name && tc.id && !announcedToolIds.has(tc.id)) {
            announcedToolIds.add(tc.id)
            currentContent = ''
            if (options.onToolCall) options.onToolCall(tc.name, {})
          }
        }
        continue
      }

      // Complete tool result messages
      if (msgType === 'tool') {
        if (options.onToolResult) {
          options.onToolResult(String(msg.name || ''), String(msg.content || ''))
        }
        currentContent = ''
        continue
      }

      // Stream AI text tokens for the final response
      if (msgType === 'ai' && typeof msg.content === 'string' && msg.content) {
        currentContent += msg.content
        options.onStream(currentContent)
      }
    }

    console.log('[Agent] Flow completed')
  } catch (error: any) {
    if (error.name === 'AbortError' || options.abortSignal?.aborted) {
      throw error
    }
    if (error.name === 'GraphRecursionError') {
      options.errorIssue.value = 'recursionLimitExceeded'
    } else {
      options.errorIssue.value = true
    }
    console.error('[Agent] Error:', error)
  } finally {
    options.loading.value = false
  }
}

export async function getChatResponse(options: ProviderOptions) {
  const creator = ModelCreators[options.provider]
  if (!creator) {
    throw new Error(`Unsupported provider: ${options.provider}`)
  }
  const model = creator(options)
  return executeChatFlow(model, options)
}

export async function getAgentResponse(options: AgentOptions) {
  const creator = ModelCreators[options.provider]
  if (!creator) {
    throw new Error(`Unsupported provider: ${options.provider}`)
  }
  const model = creator(options)
  return executeAgentFlow(model, options)
}
