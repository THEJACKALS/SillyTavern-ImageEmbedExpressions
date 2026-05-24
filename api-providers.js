/**
 * API Providers Library for Image Embed Expressions
 * Supports multiple AI chat completion providers with unified interface
 */
import { getRequestHeaders } from '/script.js';

const HORDE_ANONYMOUS_API_KEY = '0000000000';

export const API_PROVIDERS = {
    // OpenAI and OpenAI-compatible APIs
    openai: {
        name: 'OpenAI',
        baseUrl: 'https://api.openai.com/v1',
        models: ['gpt-4', 'gpt-4-turbo', 'gpt-3.5-turbo'],
        defaultModel: 'gpt-3.5-turbo',
        requiresAuth: true
    },
    anthropic: {
        name: 'Anthropic (Claude)',
        baseUrl: 'https://api.anthropic.com/v1',
        models: ['claude-3-opus', 'claude-3-sonnet', 'claude-3-haiku'],
        defaultModel: 'claude-3-haiku',
        requiresAuth: true
    },
    deepseek: {
        name: 'Deepseek',
        baseUrl: 'https://api.deepseek.com/v1',
        models: ['deepseek-chat', 'deepseek-coder'],
        defaultModel: 'deepseek-chat',
        requiresAuth: true
    },
    xai: {
        name: 'xAI (Grok)',
        baseUrl: 'https://api.x.ai/v1',
        models: ['grok-beta'],
        defaultModel: 'grok-beta',
        requiresAuth: true
    },
    groq: {
        name: 'Groq',
        baseUrl: 'https://api.groq.com/openai/v1',
        models: ['mixtral-8x7b-32768', 'llama2-70b-4096'],
        defaultModel: 'mixtral-8x7b-32768',
        requiresAuth: true
    },
    perplexity: {
        name: 'Perplexity',
        baseUrl: 'https://api.perplexity.ai',
        models: ['pplx-7b-online', 'pplx-70b-online'],
        defaultModel: 'pplx-7b-online',
        requiresAuth: true
    },
    openrouter: {
        name: 'OpenRouter',
        baseUrl: 'https://openrouter.ai/api/v1',
        models: ['auto'],
        defaultModel: 'auto',
        requiresAuth: true,
        editable: true,
        description: 'Supports hundreds of models'
    },
    ollama: {
        name: 'Ollama (Local)',
        baseUrl: 'http://localhost:11434/v1',
        models: ['mistral', 'llama2', 'neural-chat'],
        defaultModel: 'mistral',
        requiresAuth: false,
        editable: true,
        description: 'Local Ollama server'
    },
    llamacpp: {
        name: 'Llama.cpp (Local)',
        baseUrl: 'http://localhost:8000/v1',
        models: ['gpt-3.5-turbo'],
        defaultModel: 'gpt-3.5-turbo',
        requiresAuth: false,
        editable: true,
        description: 'Local Llama.cpp server'
    },
    lmstudio: {
        name: 'LM Studio (Local)',
        baseUrl: 'http://localhost:1234/v1',
        models: ['local-model'],
        defaultModel: 'local-model',
        requiresAuth: false,
        editable: true,
        description: 'Local LM Studio server'
    },
    aphrodite: {
        name: 'Aphrodite Engine (Local)',
        baseUrl: 'http://localhost:5000/v1',
        models: ['gpt-3.5-turbo'],
        defaultModel: 'gpt-3.5-turbo',
        requiresAuth: false,
        editable: true,
        description: 'Local Aphrodite Engine server'
    },
    horde: {
        name: 'Horde AI',
        baseUrl: 'https://oai.aihorde.net/v1',
        models: [],
        defaultModel: '',
        requiresAuth: false,
        optionalAuth: true,
        description: 'Distributed AI. Leave the API key empty to use anonymous key 0000000000.'
    },
    custom: {
        name: 'Custom / OpenAI-compatible',
        baseUrl: '',
        models: [],
        defaultModel: '',
        requiresAuth: false,
        editable: true,
        customModel: true,
        description: 'Use any chat-completions endpoint with your own base URL and model'
    }
};

/**
 * Call API to get expression suggestion
 * @param {string} provider - Provider key (openai, anthropic, etc)
 * @param {string} apiKey - API key for authentication
 * @param {string} prompt - The prompt to send to AI
 * @param {Object} options - Additional options
 * @returns {Promise<string>} - AI response
 */
export async function callAIProvider(provider, apiKey, prompt, options = {}) {
    const providerConfig = API_PROVIDERS[provider];
    
    if (!providerConfig) {
        throw new Error(`Unknown provider: ${provider}`);
    }

    const customBaseUrl = providerConfig.editable ? options.customBaseUrl : '';
    const baseUrl = resolveProviderBaseUrl(provider, customBaseUrl || providerConfig.baseUrl);
    const model = options.model || providerConfig.defaultModel;
    const maxTokens = options.maxTokens || 50;
    const temperature = options.temperature ?? 0.3;
    const resolvedApiKey = resolveProviderApiKey(provider, apiKey);

    try {
        switch (provider) {
            case 'anthropic':
                return await callAnthropicAPI(baseUrl, apiKey, model, prompt, maxTokens, temperature);
            
            case 'ollama':
            case 'llamacpp':
            case 'lmstudio':
            case 'aphrodite':
            case 'horde':
            case 'custom':
                // Use the ST backend as a proxy for local/custom providers to avoid browser CORS/preflight issues.
                return await callOpenAICompatibleProxyAPI(baseUrl, resolvedApiKey, model, prompt, maxTokens, temperature);
            
            case 'openai':
            case 'deepseek':
            case 'xai':
            case 'groq':
            case 'perplexity':
            case 'openrouter':
            default:
                // OpenAI-compatible API
                return await callOpenAICompatibleAPI(baseUrl, resolvedApiKey, model, prompt, maxTokens, temperature);
        }
    } catch (error) {
        console.error(`Error calling ${provider} API:`, error);
        throw error;
    }
}

/**
 * Call OpenAI-compatible API
 */
async function callOpenAICompatibleAPI(baseUrl, apiKey, model, prompt, maxTokens, temperature) {
    const endpoint = `${normalizeBaseUrl(baseUrl)}/chat/completions`;
    const headers = {
        'Content-Type': 'application/json',
    };

    if (apiKey) {
        headers.Authorization = `Bearer ${apiKey}`;
    }
    
    const response = await fetch(endpoint, {
        method: 'POST',
        headers,
        body: JSON.stringify({
            model: model,
            messages: [
                {
                    role: 'system',
                    content: 'You are a helpful assistant that analyzes character emotions and expressions.'
                },
                {
                    role: 'user',
                    content: prompt
                }
            ],
            max_tokens: maxTokens,
            temperature: temperature
        })
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`API error (${response.status}): ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    
    if (data.choices && data.choices[0] && data.choices[0].message) {
        return data.choices[0].message.content.trim();
    }

    throw new Error(getUnexpectedResponseMessage(data, baseUrl));
}

async function callOpenAICompatibleProxyAPI(baseUrl, apiKey, model, prompt, maxTokens, temperature) {
    const normalizedBaseUrl = normalizeBaseUrl(baseUrl);
    const response = await fetch('/api/backends/chat-completions/generate', {
        method: 'POST',
        headers: getRequestHeaders(),
        body: JSON.stringify({
            chat_completion_source: 'custom',
            custom_url: normalizedBaseUrl,
            custom_include_headers: apiKey ? JSON.stringify({ Authorization: `Bearer ${apiKey}` }) : '',
            model,
            temperature,
            max_tokens: maxTokens,
            stream: false,
            messages: [
                {
                    role: 'system',
                    content: 'You are a helpful assistant that analyzes character emotions and expressions.',
                },
                {
                    role: 'user',
                    content: prompt,
                },
            ],
        }),
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok || data?.error) {
        const errorMessage = data?.error?.message || response.statusText || 'Request failed';
        throw new Error(`API error (${response.status}): ${errorMessage}`);
    }

    if (data.choices && data.choices[0] && data.choices[0].message) {
        return data.choices[0].message.content.trim();
    }

    throw new Error(getUnexpectedResponseMessage(data, normalizedBaseUrl));
}

/**
 * Call Anthropic Claude API
 */
async function callAnthropicAPI(baseUrl, apiKey, model, prompt, maxTokens, temperature) {
    const endpoint = `${baseUrl}/messages`;
    
    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
            model: model,
            max_tokens: maxTokens,
            system: 'You are a helpful assistant that analyzes character emotions and expressions.',
            messages: [
                {
                    role: 'user',
                    content: prompt
                }
            ],
            temperature: temperature
        })
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(`API error (${response.status}): ${errorData.error?.message || response.statusText}`);
    }

    const data = await response.json();
    
    if (data.content && data.content[0] && data.content[0].text) {
        return data.content[0].text.trim();
    }

    throw new Error('Invalid Claude API response format');
}

/**
 * Validate API key (basic check, actual validation happens on API call)
 */
export function validateApiKey(provider, apiKey) {
    const providerConfig = API_PROVIDERS[provider];
    if (!providerConfig) {
        return false;
    }

    // Local providers don't require API key
    if (!providerConfig.requiresAuth) {
        return true;
    }

    if (!apiKey || typeof apiKey !== 'string') {
        return false;
    }

    // Most API keys are at least 20 characters
    return apiKey.length >= 10;
}

/**
 * Get provider configuration
 */
export function getProviderConfig(provider) {
    return API_PROVIDERS[provider] || null;
}

/**
 * Get all available providers
 */
export function getAllProviders() {
    return Object.keys(API_PROVIDERS).map(key => ({
        id: key,
        ...API_PROVIDERS[key]
    }));
}

export function providerRequiresApiKey(provider) {
    return !!API_PROVIDERS[provider]?.requiresAuth;
}

export async function fetchProviderModels(provider, apiKey, options = {}) {
    const providerConfig = API_PROVIDERS[provider];
    if (!providerConfig) {
        throw new Error(`Unknown provider: ${provider}`);
    }

    if (provider === 'horde') {
        return await fetchHordeModels();
    }

    const customBaseUrl = providerConfig.editable ? options.customBaseUrl : '';
    const baseUrl = resolveProviderBaseUrl(provider, customBaseUrl || providerConfig.baseUrl);
    const resolvedApiKey = resolveProviderApiKey(provider, apiKey);
    const staticModels = Array.isArray(providerConfig.models) ? [...providerConfig.models] : [];

    if (!shouldFetchModelsThroughProxy(provider, baseUrl)) {
        return staticModels;
    }

    const response = await fetch('/api/backends/chat-completions/status', {
        method: 'POST',
        headers: getRequestHeaders(),
        cache: 'no-cache',
        body: JSON.stringify({
            chat_completion_source: 'custom',
            custom_url: baseUrl,
            custom_include_headers: resolvedApiKey ? JSON.stringify({ Authorization: `Bearer ${resolvedApiKey}` }) : '',
        }),
    });

    if (!response.ok) {
        throw new Error(`Failed to fetch models: ${response.statusText}`);
    }

    const data = await response.json().catch(() => ({}));
    const models = Array.isArray(data?.data)
        ? data.data.map(model => String(model?.id || '').trim()).filter(Boolean)
        : [];

    return models.length ? models : staticModels;
}

async function fetchHordeModels() {
    const response = await fetch('/api/horde/text-models', {
        method: 'POST',
        headers: getRequestHeaders(),
        body: JSON.stringify({ force: false }),
    });

    if (!response.ok) {
        throw new Error(`Failed to fetch Horde models: ${response.statusText}`);
    }

    const data = await response.json().catch(() => ([]));
    const models = Array.isArray(data)
        ? data.map(model => String(model?.name || '').trim()).filter(Boolean)
        : [];

    if (!models.length) {
        throw new Error('Horde did not return any text models');
    }

    return [...new Set(models)];
}

function normalizeBaseUrl(baseUrl) {
    return String(baseUrl ?? '').trim().replace(/\/+$/, '');
}

function resolveProviderApiKey(provider, apiKey) {
    const normalizedApiKey = typeof apiKey === 'string' ? apiKey.trim() : '';
    if (provider === 'horde') {
        return normalizedApiKey || HORDE_ANONYMOUS_API_KEY;
    }
    return normalizedApiKey;
}

function resolveProviderBaseUrl(provider, baseUrl) {
    const normalized = normalizeBaseUrl(baseUrl);
    if (!normalized) {
        return normalized;
    }

    if (provider === 'lmstudio') {
        try {
            const url = new URL(normalized);
            if (!/^\/v\d+$/i.test(url.pathname) && url.pathname !== '/v1') {
                url.pathname = url.pathname && url.pathname !== '/' ? url.pathname.replace(/\/+$/, '') : '';
                url.pathname = `${url.pathname}/v1`.replace(/\/{2,}/g, '/');
                return normalizeBaseUrl(url.toString());
            }
        } catch {
            if (!/\/v\d+$/i.test(normalized)) {
                return `${normalized}/v1`;
            }
        }
    }

    return normalized;
}

function getUnexpectedResponseMessage(data, baseUrl) {
    const apiMessage = data?.error?.message || data?.message || data?.error;
    if (apiMessage) {
        return String(apiMessage);
    }

    const hint = /\/v1$/i.test(normalizeBaseUrl(baseUrl))
        ? 'Unexpected API response format'
        : 'Unexpected API response format. Try using a base URL ending with /v1';

    return hint;
}

function shouldFetchModelsThroughProxy(provider, baseUrl) {
    if (!normalizeBaseUrl(baseUrl)) {
        return false;
    }

    return ['ollama', 'llamacpp', 'lmstudio', 'aphrodite', 'horde', 'custom'].includes(provider);
}

/**
 * Test API connection
 */
export async function testAPIConnection(provider, apiKey, options = {}) {
    const providerConfig = API_PROVIDERS[provider];
    
    if (!providerConfig) {
        throw new Error(`Unknown provider: ${provider}`);
    }

    const testPrompt = 'Respond with just the word "working". Nothing else.';
    
    try {
        let resolvedModel = options.model || providerConfig.defaultModel || '';

        if (!resolvedModel) {
            const availableModels = await fetchProviderModels(provider, apiKey, options);
            resolvedModel = Array.isArray(availableModels)
                ? String(availableModels.find(model => String(model || '').trim()) || '').trim()
                : '';
        }

        if (!resolvedModel) {
            throw new Error('No model available for this provider. Try refreshing models or enter a model manually.');
        }

        const response = await callAIProvider(provider, apiKey, testPrompt, {
            ...options,
            model: resolvedModel,
            maxTokens: 10,
            temperature: 0
        });
        
        return {
            success: true,
            message: 'API connection successful',
            response: response
        };
    } catch (error) {
        return {
            success: false,
            message: error.message,
            error: error
        };
    }
}
