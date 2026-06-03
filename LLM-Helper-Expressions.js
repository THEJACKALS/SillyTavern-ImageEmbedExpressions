import { getStringHash } from '/scripts/utils.js';
import {
    callAIProvider,
    getProviderConfig,
    providerRequiresApiKey,
} from './api-providers.js';

export const EXPRESSION_DETECTION_VERSION = 10;

const AI_CACHE_STORAGE_KEY = `imageEmbedsExpressions_aiCache_v${EXPRESSION_DETECTION_VERSION}`;
const AI_CACHE_MAX_ENTRIES = 500;
const MESSAGE_CONTEXT_TOKEN_LIMIT = 2000;
const APPROX_CHARS_PER_TOKEN = 4;

function loadAiExpressionCache() {
    try {
        const raw = localStorage.getItem(AI_CACHE_STORAGE_KEY);
        if (!raw) return new Map();
        const obj = JSON.parse(raw);
        if (typeof obj !== 'object' || obj === null) return new Map();
        return new Map(Object.entries(obj));
    } catch {
        return new Map();
    }
}

function saveAiExpressionCache(cacheMap) {
    try {
        let entries = [...cacheMap.entries()];
        if (entries.length > AI_CACHE_MAX_ENTRIES) {
            entries = entries.slice(entries.length - AI_CACHE_MAX_ENTRIES);
        }
        localStorage.setItem(AI_CACHE_STORAGE_KEY, JSON.stringify(Object.fromEntries(entries)));
    } catch {
        // localStorage might be full or unavailable.
    }
}

let aiExpressionCache = loadAiExpressionCache();

export function clearExpressionAiCache() {
    aiExpressionCache.clear();
    try { localStorage.removeItem(AI_CACHE_STORAGE_KEY); } catch { /* ignore */ }
}

function trimMessageToTokenLimit(messageText, tokenLimit = MESSAGE_CONTEXT_TOKEN_LIMIT) {
    const text = String(messageText || '').trim();
    const maxChars = tokenLimit * APPROX_CHARS_PER_TOKEN;
    if (text.length <= maxChars) {
        return text;
    }

    return text.slice(text.length - maxChars).trim();
}

function buildAvailableExpressionList(entries, parseEntryName) {
    return entries
        .map(entry => ({
            name: entry.name,
            displayName: parseEntryName(entry.name).expression,
        }))
        .filter(entry => entry.displayName);
}

function buildStatelessExpressionPrompt({ characterName, messageText, availableExpressions }) {
    return `You are the stateless Advanced Expressions selector for SillyTavern.

Task:
Choose ONE existing Character Expression image for "${characterName}".
You are not generating a new image. You are only selecting one item from the provided expression list.

Memory rules:
Use only the single generated message below.
Ignore any prior chat history, prior choices, and external assumptions.
The message is capped to about ${MESSAGE_CONTEXT_TOKEN_LIMIT} tokens.

Available Character Expressions for "${characterName}":
${availableExpressions.map(entry => `- ${entry.displayName}`).join('\n')}

Generated message to inspect:
${messageText}

Selection rules:
- Pick the expression that best matches "${characterName}" in this message.
- Prefer visible facial/body cues, dialogue delivery, and immediate action.
- Do not choose based on another character's emotion.
- If several emotions are possible, choose the strongest visible expression.
- Respond with exactly one expression name from the list above.
- No explanation, no punctuation, no extra text.`;
}

export async function detectExpressionWithAI({
    messageText,
    entries,
    characterName,
    settings,
    parseEntryName,
    normalizeName,
}) {
    if (!entries || entries.length === 0) return null;

    const providerConfig = getProviderConfig(settings.apiProvider);

    if (!settings.apiProvider) {
        console.warn('Advanced Expressions: API provider not configured');
        return null;
    }

    if (!providerConfig) {
        console.warn('Advanced Expressions: invalid API provider configuration');
        return null;
    }

    if (providerRequiresApiKey(settings.apiProvider) && !settings.apiKey) {
        console.warn('Advanced Expressions: API key not configured');
        return null;
    }

    if (providerConfig.editable && !settings.customBaseUrl && !providerConfig.baseUrl) {
        console.warn('Advanced Expressions: custom base URL is required for this provider');
        return null;
    }

    const scopedMessageText = trimMessageToTokenLimit(messageText);
    const entrySignature = entries.map(entry => entry.name).join('|');
    const messageHash = getStringHash(`${EXPRESSION_DETECTION_VERSION}::${characterName}::${scopedMessageText}::${entrySignature}`);
    if (aiExpressionCache.has(messageHash)) {
        return aiExpressionCache.get(messageHash);
    }

    try {
        const availableExpressions = buildAvailableExpressionList(entries, parseEntryName);

        if (availableExpressions.length === 0) return null;
        if (!settings.apiModel) {
            console.warn('Advanced Expressions: API model not configured');
            return null;
        }

        const prompt = buildStatelessExpressionPrompt({
            characterName,
            messageText: scopedMessageText,
            availableExpressions,
        });

        const aiResponse = await callAIProvider(
            settings.apiProvider,
            settings.apiKey,
            prompt,
            {
                model: settings.apiModel || undefined,
                customBaseUrl: settings.customBaseUrl || undefined,
                maxTokens: 50,
                temperature: 0.3,
            },
        );

        if (!aiResponse) {
            console.warn('AI expression detection returned empty response');
            return null;
        }

        const responseText = normalizeName(aiResponse.trim()).replace(/_/g, ' ');
        const matchedExpression = availableExpressions.find(entry =>
            responseText.includes(normalizeName(entry.displayName).replace(/_/g, ' '))
        );
        const selectedExpression = matchedExpression ? matchedExpression.name : null;

        aiExpressionCache.set(messageHash, selectedExpression);
        saveAiExpressionCache(aiExpressionCache);

        return selectedExpression;
    } catch (error) {
        console.error('Error calling AI for expression detection:', error);
        return null;
    }
}
