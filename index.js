import { characters, chat, eventSource, event_types, saveSettingsDebounced, this_chid } from '/script.js';
import { extension_settings, renderExtensionTemplateAsync } from '/scripts/extensions.js';
import { getBase64Async, getFileExtension, getStringHash, saveBase64AsFile } from '/scripts/utils.js';
import {
    fetchProviderModels,
    getAllProviders,
    getProviderConfig,
    providerRequiresApiKey,
    testAPIConnection,
} from './api-providers.js';
import { detectActiveCharacterKeys } from './character-presence.js';
import {
    clearExpressionAiCache,
    detectExpressionWithAI,
    EXPRESSION_DETECTION_VERSION,
} from './LLM-Helper-Expressions.js';
import {
    EXPRESSION_HINTS,
    getExpressionHintKeys,
} from './hint-bank.js';

const EXTENSION_ID = (() => {
    const match = new URL(import.meta.url).pathname.match(/scripts\/extensions\/(.+)\/index\.js$/);
    if (match?.[1]) {
        return match[1];
    }

    // Fallback: prefer third-party path when loaded from user extensions
    const isThirdParty = import.meta.url.includes('/third-party/');
    return isThirdParty ? 'third-party/image-embeds-expressions' : 'image-embeds-expressions';
})();
const SETTINGS_KEY = 'imageEmbedsExpressions';
const STORAGE_FOLDER = 'image-embeds-expressions';
const PLACEHOLDER_REGEX = /\{\{img::(.*?)\}\}/gi;
const CODE_TAGS = new Set(['code', 'pre', 'samp', 'kbd']);
const defaultSettings = {
    characters: {},
    enabled: true,
    doubleEnabled: false,
    showUserMode: true,
    advancedExpressionsEnabled: false,
    apiProvider: '',
    apiKey: '',
    apiModel: '',
    customBaseUrl: '',
    autoConnectLastServer: false
};
const DEFAULT_CHARACTER_GROUP = '__default__';
let lastAssistantMessageId = null;
let currentMode = 'character'; // 'character' or 'user'
let messagePlacementCache = new Map();
const llmEligibleMessageIds = new Set();

const ADVANCED_AI_RECENT_MESSAGE_WINDOW = 2;
const LEADING_CHARACTER_ACTIONS = [
    'said', 'says', 'asked', 'asks', 'replied', 'replies', 'responded', 'responds',
    'whispered', 'whispers', 'muttered', 'mutters', 'shouted', 'shouts', 'yelled', 'yells',
    'laughed', 'laughs', 'chuckled', 'chuckles', 'smiled', 'smiles', 'grinned', 'grins',
    'sighed', 'sighs', 'squeaked', 'squeaks', 'snapped', 'snaps', 'growled', 'growls', 'hissed', 'hisses',
    'looked', 'looks', 'glared', 'glares', 'blinked', 'blinks', 'nodded', 'nods',
    'shook', 'turns', 'turned', 'tapped', 'waves', 'waved', 'gestured', 'gestures',
    'paused', 'pauses', 'leaned', 'leans', 'stepped', 'steps', 'walked', 'walks',
    'sat', 'sits', 'stood', 'stands', 'crossed', 'uncrossed', 'clenched', 'relaxed',
];

let modelFetchToken = 0;
let modelRefreshTimeout = null;
let refreshAllDebounceTimer = null;         // debounce timer for refreshAllMessages
const autoInjectInFlight = new Set();       // guard: prevent concurrent autoInject for same messageId
const processedMessages = new WeakMap();    // track already-processed message roots (stores generation#)
let processedMessageGeneration = 0;         // bump to invalidate all processedMessages entries

function clearAiExpressionCache() {
    clearExpressionAiCache();
    messagePlacementCache.clear();
    // Bump generation so all cached "already processed" roots get re-evaluated.
    processedMessageGeneration++;
}

function setModelSelectOptions(models, preferredModel = '') {
    const modelSelect = $('#image_embeds_api_model');
    const modelCustomInput = $('#image_embeds_api_model_custom');
    const uniqueModels = [...new Set((models || []).map(model => String(model || '').trim()).filter(Boolean))];

    modelSelect.empty();
    modelSelect.append('<option value="">-- Select or type model --</option>');

    if (preferredModel && !uniqueModels.includes(preferredModel)) {
        uniqueModels.unshift(preferredModel);
    }

    for (const model of uniqueModels) {
        modelSelect.append(
            $('<option></option>')
                .attr('value', model)
                .text(model),
        );
    }

    const hasPresetModels = uniqueModels.length > 0;
    modelSelect.toggle(hasPresetModels);
    modelCustomInput.toggle(!hasPresetModels);

    const resolvedModel = preferredModel && uniqueModels.includes(preferredModel)
        ? preferredModel
        : (preferredModel || (uniqueModels[0] || ''));

    modelSelect.val(resolvedModel);
    modelCustomInput.val(resolvedModel);

    const settings = ensureSettings();
    settings.apiModel = resolvedModel;
}

async function refreshProviderModels(provider, { silent = true } = {}) {
    if (!provider) return;

    const token = ++modelFetchToken;
    const settings = ensureSettings();
    const providerConfig = getProviderConfig(provider);
    const statusDiv = $('#image_embeds_connection_status');
    const preferredModel = settings.apiModel || providerConfig?.defaultModel || '';

    if (!silent) {
        statusDiv.html('<span style="color: #ffd93d;">Loading models...</span>');
    }

    try {
        const models = await fetchProviderModels(provider, settings.apiKey, {
            customBaseUrl: settings.customBaseUrl || undefined,
        });

        if (token !== modelFetchToken) return;

        setModelSelectOptions(models, preferredModel);

        if (!silent) {
            const count = Array.isArray(models) ? models.length : 0;
            statusDiv.html(`<span style="color: #6bcf7f;"><i class="fa-solid fa-check"></i> Loaded ${count} model${count === 1 ? '' : 's'}</span>`);
            setTimeout(() => {
                if (token === modelFetchToken) {
                    statusDiv.html('');
                }
            }, 2500);
        }
    } catch (error) {
        if (token !== modelFetchToken) return;

        setModelSelectOptions(providerConfig?.models || [], preferredModel);

        if (!silent) {
            statusDiv.html(`<span style="color: #ff6b6b;"><i class="fa-solid fa-xmark"></i> ${error.message}</span>`);
        }
    }
}

function scheduleProviderModelsRefresh(provider, delay = 450) {
    clearTimeout(modelRefreshTimeout);
    modelRefreshTimeout = setTimeout(() => {
        void refreshProviderModels(provider, { silent: true });
    }, delay);
}

function renderConnectionStatus(state = 'disconnected', message = '') {
    const statusDiv = $('#image_embeds_connection_status');
    if (!statusDiv.length) return;

    const normalizedState = ['connected', 'connecting', 'disconnected'].includes(state)
        ? state
        : 'disconnected';
    const label = message || (
        normalizedState === 'connected'
            ? 'Connected'
            : (normalizedState === 'connecting' ? 'Connecting...' : 'Disconnected')
    );

    statusDiv
        .removeClass('connected connecting disconnected')
        .addClass(normalizedState)
        .empty()
        .append(
            $('<span class="image-embeds-connection-lamps" aria-hidden="true"></span>').append(
                $('<span class="image-embeds-lamp image-embeds-lamp-red"></span>'),
                $('<span class="image-embeds-lamp image-embeds-lamp-green"></span>'),
            ),
            $('<span class="image-embeds-connection-label"></span>').text(label),
        );
}

function formatExpressionKey(key) {
    return String(key || '')
        .replace(/_/g, ' ')
        .replace(/\b\w/g, letter => letter.toUpperCase());
}

function renderExpressionDocumentation() {
    const list = $('#image_embeds_expression_docs_list');
    if (!list.length) return;

    list.empty();

    const entries = Object.entries(EXPRESSION_HINTS)
        .sort(([left], [right]) => left.localeCompare(right));

    for (const [key, hints] of entries) {
        const card = $('<div class="image-embeds-doc-card"></div>');
        const header = $('<div class="image-embeds-doc-card-header"></div>');
        header.append(
            $('<div class="image-embeds-doc-card-title"></div>').text(formatExpressionKey(key)),
            $('<div class="image-embeds-doc-card-count"></div>').text(`${hints.length} support terms`),
        );

        const hintWrap = $('<div class="image-embeds-doc-hints"></div>');
        const previewHints = hints.slice(0, 6);

        for (const hint of previewHints) {
            hintWrap.append($('<span class="image-embeds-doc-hint"></span>').text(hint));
        }

        if (hints.length > previewHints.length) {
            hintWrap.append($('<span class="image-embeds-doc-hint image-embeds-doc-hint-more"></span>').text(`+${hints.length - previewHints.length} more`));
        }

        card.append(header, hintWrap);
        list.append(card);
    }
}

function openExpressionDocumentation() {
    renderExpressionDocumentation();
    $('#image_embeds_expression_docs_popup')
        .appendTo('body')
        .addClass('is-open')
        .attr('aria-hidden', 'false')
        .css('display', 'flex');
}

function closeExpressionDocumentation() {
    $('#image_embeds_expression_docs_popup')
        .removeClass('is-open')
        .attr('aria-hidden', 'true')
        .hide();
}

async function restartProviderConnection({ signal } = {}) {
    const settings = ensureSettings();
    const providerConfig = getProviderConfig(settings.apiProvider);

    if (!settings.apiProvider) {
        renderConnectionStatus('disconnected', 'Select a provider first');
        return false;
    }

    if (providerRequiresApiKey(settings.apiProvider) && !settings.apiKey) {
        renderConnectionStatus('disconnected', 'This provider needs an API key');
        return false;
    }

    if (providerConfig?.editable && !settings.customBaseUrl && !providerConfig.baseUrl) {
        renderConnectionStatus('disconnected', 'Enter a base URL first');
        return false;
    }

    renderConnectionStatus('connecting', 'Restarting...');

    try {
        if (settings.apiProvider === 'horde') {
            const models = await fetchProviderModels(settings.apiProvider, settings.apiKey, {
                customBaseUrl: settings.customBaseUrl || undefined,
            });
            setModelSelectOptions(models, settings.apiModel || providerConfig?.defaultModel || models[0] || '');
        } else {
            const result = await testAPIConnection(
                settings.apiProvider,
                settings.apiKey,
                {
                    model: settings.apiModel || undefined,
                    customBaseUrl: settings.customBaseUrl || undefined,
                    signal,
                },
            );

            if (!result.success) {
                throw new Error(result.message || 'Connection failed');
            }

            await refreshProviderModels(settings.apiProvider, { silent: true });
        }

        if (signal?.aborted) return false;

        renderConnectionStatus('connected', 'Connected');
        return true;
    } catch (error) {
        if (signal?.aborted || error.name === 'AbortError') {
            renderConnectionStatus('disconnected', 'Restart aborted');
        } else {
            renderConnectionStatus('disconnected', error.message || 'Disconnected');
        }
        return false;
    }
}

function escapeRegExp(value) {
    return String(value ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function textIncludesNeedle(text, needle) {
    const value = String(needle || '').trim();
    if (!value) return false;
    if (/\s/.test(value)) {
        return String(text || '').includes(value);
    }
    return new RegExp(`\\b${escapeRegExp(value)}\\b`, 'i').test(String(text || ''));
}

function normalizeHintText(text) {
    return String(text || '')
        .toLowerCase()
        .replace(/[^\w\s-]/g, ' ')
        .replace(/[-_]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function rememberAssistantMessage(messageId) {
    const numericId = Number(messageId);
    if (Number.isNaN(numericId)) return null;

    const message = chat?.[numericId];
    if (message && !message.is_user && !message.is_system) {
        lastAssistantMessageId = numericId;
    }

    return numericId;
}

function getActiveAssistantMessageId() {
    if (typeof lastAssistantMessageId === 'number' && chat?.[lastAssistantMessageId]) {
        return lastAssistantMessageId;
    }

    const fallback = chat.length - 1;
    if (fallback >= 0 && chat?.[fallback]) {
        return fallback;
    }

    return null;
}

function markMessageEligibleForLLM(messageId) {
    const numericId = Number(messageId);
    if (Number.isNaN(numericId)) return;
    const message = chat?.[numericId];
    if (!message || message.is_user || message.is_system) return;
    llmEligibleMessageIds.add(numericId);
}

function canUseLLMForMessage(messageId) {
    const numericId = Number(messageId);
    return !Number.isNaN(numericId) && llmEligibleMessageIds.has(numericId);
}

function consumeLLMEligibility(messageId) {
    const numericId = Number(messageId);
    if (!Number.isNaN(numericId)) {
        llmEligibleMessageIds.delete(numericId);
    }
}

function ensureSettings() {
    if (!extension_settings[SETTINGS_KEY] || typeof extension_settings[SETTINGS_KEY] !== 'object') {
        extension_settings[SETTINGS_KEY] = { ...defaultSettings, characters: {} };
    }

    if (!extension_settings[SETTINGS_KEY].characters || typeof extension_settings[SETTINGS_KEY].characters !== 'object') {
        extension_settings[SETTINGS_KEY].characters = {};
    }

    if (typeof extension_settings[SETTINGS_KEY].enabled !== 'boolean') {
        extension_settings[SETTINGS_KEY].enabled = true;
    }

    if (typeof extension_settings[SETTINGS_KEY].doubleEnabled !== 'boolean') {
        extension_settings[SETTINGS_KEY].doubleEnabled = false;
    }

    if (typeof extension_settings[SETTINGS_KEY].showUserMode !== 'boolean') {
        extension_settings[SETTINGS_KEY].showUserMode = true;
    }

    if (typeof extension_settings[SETTINGS_KEY].advancedExpressionsEnabled !== 'boolean') {
        extension_settings[SETTINGS_KEY].advancedExpressionsEnabled = false;
    }

    if (typeof extension_settings[SETTINGS_KEY].apiProvider !== 'string') {
        extension_settings[SETTINGS_KEY].apiProvider = '';
    }

    if (typeof extension_settings[SETTINGS_KEY].apiKey !== 'string') {
        extension_settings[SETTINGS_KEY].apiKey = '';
    }

    if (typeof extension_settings[SETTINGS_KEY].apiModel !== 'string') {
        extension_settings[SETTINGS_KEY].apiModel = '';
    }

    if (typeof extension_settings[SETTINGS_KEY].customBaseUrl !== 'string') {
        extension_settings[SETTINGS_KEY].customBaseUrl = '';
    }

    if (typeof extension_settings[SETTINGS_KEY].autoConnectLastServer !== 'boolean') {
        extension_settings[SETTINGS_KEY].autoConnectLastServer = false;
    }

    return extension_settings[SETTINGS_KEY];
}

function getCharacterKey() {
    const settings = ensureSettings();
    const avatar = characters?.[this_chid]?.avatar;
    if (!avatar) {
        return null;
    }
    if (!settings.characters[avatar]) {
        settings.characters[avatar] = { entries: [] };
    }
    if (!Array.isArray(settings.characters[avatar].entries)) {
        settings.characters[avatar].entries = [];
    }
    return avatar;
}

function getCharacterEntries() {
    const key = getCharacterKey();
    if (!key) return [];
    const settings = ensureSettings();

    // Migrate legacy global entries to the current character once.
    if (Array.isArray(settings.entries) && settings.entries.length && (!settings.characters[key]?.entries?.length)) {
        settings.characters[key] = { entries: settings.entries };
        delete settings.entries;
        saveSettingsDebounced();
    }

    return settings.characters[key].entries;
}

function getCharacterFolder() {
    const key = getCharacterKey();
    if (!key) return STORAGE_FOLDER;
    return `${STORAGE_FOLDER}/${normalizeName(key)}`;
}

function getUserEntries() {
    const key = getCharacterKey();
    if (!key) return [];
    const settings = ensureSettings();
    if (!settings.characters[key]) {
        settings.characters[key] = { entries: [] };
    }
    if (!settings.characters[key].userEntries) {
        settings.characters[key].userEntries = [];
    }
    if (!Array.isArray(settings.characters[key].userEntries)) {
        settings.characters[key].userEntries = [];
    }
    return settings.characters[key].userEntries;
}

function getUserFolder() {
    const key = getCharacterKey();
    if (!key) return STORAGE_FOLDER;
    return `${STORAGE_FOLDER}/${normalizeName(key)}_user`;
}

function getActiveCharacterExpressionKey() {
    return normalizeName(characters?.[this_chid]?.name || '');
}

function normalizeName(name) {
    return String(name ?? '')
        .trim()
        .toLowerCase()
        .replace(/[\\\/\s]+/g, '_');
}

function findEntryByName(name) {
    const target = normalizeName(name);

    // Search character entries first
    let entry = getCharacterEntries().find(entry => normalizeName(entry.name) === target);
    if (entry) return entry;

    // Then search user entries if available
    if (currentMode === 'user') {
        entry = getUserEntries().find(entry => normalizeName(entry.name) === target);
    }
    return entry || null;
}

function parseEntryName(name) {
    const raw = String(name ?? '').trim();

    // Supported separators (in priority order): /  \  |  -  _
    // Strategy: try each separator explicitly so we don't mis-split names that
    // contain a separator as part of the character/expression name itself.
    //
    // Examples that must all work:
    //   Evelyn/smile        → character=evelyn,  expression=smile
    //   NameChar/pose       → character=namechar, expression=pose
    //   NameChar_pose       → character=namechar, expression=pose
    //   Violet-huhu         → character=violet,  expression=huhu
    //   Miko|focused        → character=miko,    expression=focused
    //   Jean-Luc/smile      → character=jean-luc, expression=smile  (/ wins over -)
    //   just_expression     → character='',      expression=just_expression
    //   {{img::name}}       → treated as raw, no split

    // Priority list: explicit separators tried from highest to lowest priority.
    // `/` and `|` are unambiguous; `\` rare but supported.
    // `-` and `_` are lower priority because they often appear inside names.
    const SEPARATORS = [
        { sep: '/',  re: /^(.+?)\/(.+)$/ },
        { sep: '\\', re: /^(.+?)\\(.+)$/ },
        { sep: '|',  re: /^(.+?)\|(.+)$/ },
        { sep: '-',  re: /^(.+?)-(.+)$/ },
        { sep: '_',  re: /^(.+?)_(.+)$/ },
    ];

    let character = '';
    let expression = raw;

    for (const { re } of SEPARATORS) {
        const m = raw.match(re);
        if (m) {
            const candidateChar = m[1].trim();
            const candidateExpr = m[2].trim();
            // Both sides must be non-empty to count as a valid character/expression split.
            if (candidateChar && candidateExpr) {
                character = normalizeName(candidateChar);
                expression = candidateExpr;
                break;
            }
        }
    }

    return {
        raw,
        character,
        expression,
        normalized: normalizeName(raw),
    };
}

function buildNeedles(entry, includeFirstPerson = false) {
    const parsed = parseEntryName(entry.name);
    const needles = [
        normalizeName(entry.name).replace(/_/g, ' '),
        parsed.character?.replace(/_/g, ' '),
        normalizeName(parsed.expression).replace(/_/g, ' '),
    ].filter(Boolean);

    // Add first person variations for user expressions
    if (includeFirstPerson) {
        const firstPersonVariations = [
            `i'm ${needles[0]}`,
            `i am ${needles[0]}`,
            `im ${needles[0]}`,
            `i'm ${needles[needles.length - 1]}`,
            `i am ${needles[needles.length - 1]}`,
            `im ${needles[needles.length - 1]}`,
        ].filter(Boolean);
        needles.push(...firstPersonVariations);
    }

    return needles;
}

function groupEntriesByCharacter(entries) {
    const groups = new Map();

    for (const entry of entries) {
        const parsed = parseEntryName(entry.name);
        const key = parsed.character || DEFAULT_CHARACTER_GROUP;
        const bucket = groups.get(key) || [];
        bucket.push({ entry, parsed });
        groups.set(key, bucket);
    }

    return groups;
}

function detectCharacterFromText(text, characters) {
    const scored = scoreCharacters(text, characters);
    return scored[0]?.character || null;
}

function getLeadingNarrativeCharacterKey(text) {
    const value = String(text || '')
        .replace(/^[\s>*_"'`~()[\]{}]+/, '')
        .trim();
    if (!value) return '';

    const possessiveMatch = value.match(/^([A-Z][A-Za-z0-9' -]{0,48}?)(?:'s|’s)\b/);
    if (possessiveMatch?.[1]) {
        return normalizeName(possessiveMatch[1]);
    }

    const labelMatch = value.match(/^([A-Z][A-Za-z0-9' -]{0,48}?)(?=\s*(?::|[–—]|\s-\s))/);
    if (labelMatch?.[1]) {
        return normalizeName(labelMatch[1]);
    }

    const actionPattern = LEADING_CHARACTER_ACTIONS.map(escapeRegExp).join('|');
    const actionMatch = value.match(new RegExp(`^([A-Z][A-Za-z0-9' -]{0,48}?)\\s+(${actionPattern})\\b`));
    if (!actionMatch?.[1]) return '';

    const candidate = actionMatch[1]
        .trim()
        .replace(/\s+(?:and|with|to|at|from)$/i, '')
        .trim();
    const normalizedCandidate = normalizeName(candidate);
    const pronouns = new Set(['i', 'you', 'he', 'she', 'they', 'we', 'it']);

    return candidate && !pronouns.has(normalizedCandidate) ? normalizedCandidate : '';
}

function hasStrictCharacterScopedEntries(entries) {
    const parsedEntries = (entries || []).map(entry => parseEntryName(entry.name));
    return parsedEntries.length > 0 && parsedEntries.every(parsed => !!parsed.character);
}

function isUnknownLeadingCharacter(messageText, characterKeys = []) {
    const leadingCharacter = getLeadingNarrativeCharacterKey(messageText);
    if (!leadingCharacter) return false;
    if (/^(?:the|a|an|when|while|as|after|before|for)_/.test(leadingCharacter)) return false;
    return !(characterKeys || []).some(key => characterKeysMatch(key, leadingCharacter));
}

function getMessageSpeakerKey(message, characterKeys = []) {
    const speaker = normalizeName(message?.name || '');
    if (!speaker) return '';
    return characterKeys.includes(speaker) ? speaker : '';
}

function getCompactCharacterKey(characterName) {
    return normalizeName(characterName).replace(/[-|_\s]+/g, '');
}

function characterKeysMatch(a, b) {
    const normalizedA = normalizeName(a);
    const normalizedB = normalizeName(b);
    if (!normalizedA || !normalizedB) return false;
    return normalizedA === normalizedB || getCompactCharacterKey(normalizedA) === getCompactCharacterKey(normalizedB);
}

function getCharacterSearchAliases(characterName) {
    const normalized = normalizeName(characterName).replace(/(?:^|[_-])(chan|san|sama|kun|senpai|sensei)$/i, '');
    const plainName = normalized.replace(/_/g, ' ').trim();
    const parts = plainName.split(/\s+/).filter(Boolean);
    return [...new Set([
        plainName,
        parts.at(-1) || '',
    ].filter(Boolean))];
}

function entryBelongsToCharacter(entry, characterKey) {
    const parsed = parseEntryName(entry?.name);
    return !!parsed.character && characterKeysMatch(parsed.character, characterKey);
}

function filterEntriesForCharacter(entries, characterKey, { includeDefault = true } = {}) {
    const normalizedCharacter = normalizeName(characterKey);
    if (!normalizedCharacter) {
        return entries || [];
    }

    return (entries || []).filter(entry => {
        const parsed = parseEntryName(entry.name);
        if (!parsed.character) {
            return includeDefault;
        }
        return characterKeysMatch(parsed.character, normalizedCharacter);
    });
}

function getCrossPackEntriesForMessage(messageText) {
    const settings = ensureSettings();
    const allEntries = Object.values(settings.characters || {})
        .flatMap(character => Array.isArray(character?.entries) ? character.entries : [])
        .filter(entry => entry?.name);
    const groups = groupEntriesByCharacter(allEntries);
    const characterKeys = Array.from(groups.keys()).filter(key => key && key !== DEFAULT_CHARACTER_GROUP);
    const activeKeys = detectActiveCharacterKeys(String(messageText || ''), characterKeys);
    const scoredKeys = scoreCharacters(String(messageText || ''), characterKeys)
        .filter(result => result.score > 0)
        .map(result => result.character);
    const wantedKeys = [...new Set([...activeKeys, ...scoredKeys])];

    if (!wantedKeys.length) return [];

    return allEntries.filter(entry => {
        const parsed = parseEntryName(entry.name);
        return parsed.character && wantedKeys.some(key => characterKeysMatch(parsed.character, key));
    });
}

function resolveMessageCharacterKey(messageId, characterKeys = []) {
    const message = chat?.[messageId];
    const messageText = String(message?.mes || '');
    const currentScores = scoreCharacters(messageText, characterKeys);
    const topCurrent = currentScores[0];
    const secondCurrent = currentScores[1];
    const speakerKey = getMessageSpeakerKey(message, characterKeys);

    if (speakerKey) {
        return speakerKey;
    }

    if (topCurrent?.score >= 1 && (!secondCurrent || topCurrent.score > secondCurrent.score)) {
        return topCurrent.character;
    }

    const recentContext = buildRecentInteractionContext(messageId, characterKeys);
    if (recentContext.primary) {
        return recentContext.primary;
    }

    return topCurrent?.character || '';
}

function buildRecentInteractionContext(messageId, characterKeys, windowSize = 4) {
    if (!Array.isArray(characterKeys) || characterKeys.length === 0) {
        return { primary: '', scores: new Map() };
    }

    const start = Math.max(0, Number(messageId) - windowSize);
    const end = Math.min(chat.length - 1, Number(messageId));
    const scores = new Map(characterKeys.map(key => [key, 0]));

    for (let index = start; index <= end; index++) {
        const message = chat?.[index];
        if (!message || message.is_system) continue;

        const recencyWeight = end - index + 1;
        const speakerKey = getMessageSpeakerKey(message, characterKeys);
        if (speakerKey) {
            scores.set(speakerKey, (scores.get(speakerKey) || 0) + (message.is_user ? 1 : 6) * recencyWeight);
        }

        const detectedCharacters = scoreCharacters(String(message.mes || ''), characterKeys);
        for (const result of detectedCharacters.slice(0, 2)) {
            const mentionWeight = message.is_user ? 4 : 2;
            scores.set(result.character, (scores.get(result.character) || 0) + result.score * mentionWeight * recencyWeight);
        }
    }

    const sorted = Array.from(scores.entries())
        .filter(([, score]) => score > 0)
        .sort((a, b) => b[1] - a[1]);

    return {
        primary: sorted[0]?.[0] || '',
        scores,
    };
}

function scoreCharacters(text, characters) {
    const lowerText = String(text || '').toLowerCase();
    const cleaned = lowerText.replace(/[^\w\s]/g, ' ');
    const presenceVerbs = ['said', 'says', 'ask', 'asked', 'asks', 'reply', 'replied', 'replies', 'respond', 'responded', 'responds', 'yell', 'yelled', 'yells', 'shout', 'shouted', 'shouts', 'whisper', 'whispered', 'whispers', 'mutter', 'muttered', 'mutters', 'laughed', 'laughs', 'laughing', 'smiled', 'smiles', 'smiling', 'nodded', 'nods', 'grinned', 'grins', 'grinning', 'looked', 'looks', 'looking', 'turned', 'turns', 'walk', 'walking', 'walked', 'walks', 'stood', 'stands', 'standing', 'sat', 'sits', 'sitting', 'hitched', 'shivered', 'giggled', 'smirked', 'melted', 'rested', 'traced', 'dried', 'watched', 'climbed', 'ground', 'grabbed', 'come', 'comes', 'came', 'arrive', 'arrives', 'arrived', 'blushed', 'cried', 'stared'];
    const imaginationHints = ['memory of', 'remembering', 'image of', 'imagination of', 'imagining', 'fantasy of', 'thinking of', 'thought of', 'dream of', 'dreaming of', 'idea of', 'vision of'];
    const results = [];

    for (const character of characters) {
        const aliases = getCharacterSearchAliases(character);
        if (!aliases.length) continue;

        let score = 0;
        let firstIndex = Infinity;
        let mentionCount = 0;
        let match;

        for (const plainName of aliases) {
            const wordPattern = new RegExp(`\\b${escapeRegExp(plainName)}\\b`, 'g');
            const speakingPattern = new RegExp(`(^|\\n)\\s*${escapeRegExp(plainName)}\\s*[:\\-\\u2013\\u2014]`, 'g');

            while ((match = wordPattern.exec(cleaned)) !== null) {
                const aliasWeight = plainName === aliases[0] ? 1 : 0.75;
                score += aliasWeight;
                mentionCount += 1;
                if (match.index < firstIndex) {
                    firstIndex = match.index;
                }

                const window = lowerText.slice(Math.max(0, match.index - 24), match.index + plainName.length + 24);
                if (presenceVerbs.some(v => new RegExp(`\\b${escapeRegExp(v)}\\b`).test(window))) {
                    score += aliasWeight;
                }
                if (imaginationHints.some(h => window.includes(h))) {
                    score -= aliasWeight;
                }
            }

            while ((match = speakingPattern.exec(lowerText)) !== null) {
                score += plainName === aliases[0] ? 2 : 1.5;
                if (match.index < firstIndex) {
                    firstIndex = match.index;
                }
            }
        }

        if (firstIndex === 0) {
            score += 1;
        }

        score = Math.max(score, 0);

        if (score > 0) {
            results.push({ character, score, firstIndex, mentionCount });
        }
    }

    return results.sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        if (a.firstIndex !== b.firstIndex) return a.firstIndex - b.firstIndex;
        return (b.mentionCount || 0) - (a.mentionCount || 0);
    });
}

function scoreExpressionFromText(messageText, availableEntries) {
    const lower = String(messageText || '').toLowerCase();
    const cleaned = normalizeHintText(lower);
    let best = null;

    for (const entry of availableEntries) {
        const parsed = parseEntryName(entry.name);
        const expression = normalizeName(parsed.expression).replace(/_/g, ' ').trim();
        const hintKeys = getExpressionHintKeys(expression);
        let score = 0;

        if (expression && new RegExp(`\\b${escapeRegExp(expression)}\\b`, 'i').test(cleaned)) {
            score += 4;
        }

        for (const key of hintKeys) {
            for (const hint of EXPRESSION_HINTS[key] || []) {
                const normalizedHint = normalizeHintText(hint);
                if (normalizedHint && textIncludesNeedle(cleaned, normalizedHint)) {
                    score += hint.includes(' ') ? 2 : 1;
                }
            }
        }

        if (score > 0 && (!best || score > best.score)) {
            best = { entry, score };
        }
    }

    return best?.score > 0 ? best.entry : null;
}

function selectEntryForCharacter(groups, characterKey, messageText) {
    const entries = groups.get(characterKey) || [];
    if (!entries.length) return null;

    const cleaned = String(messageText || '').toLowerCase().replace(/[^\w\s]/g, ' ');

    for (const { entry, parsed } of entries) {
        const needles = [
            parsed.normalized.replace(/_/g, ' ').trim(),
            normalizeName(parsed.expression).replace(/_/g, ' ').trim(),
        ].filter(Boolean);

        if (needles.some(needle => textIncludesNeedle(cleaned, needle))) {
            return entry;
        }
    }

    const scoredEntry = scoreExpressionFromText(messageText, entries.map(item => item.entry));
    if (scoredEntry) {
        return scoredEntry;
    }

    return entries.length === 1 ? entries[0].entry : null;
}

function getFallbackEntryForCharacter(groups, characterKey) {
    const entries = groups.get(characterKey) || [];
    if (!entries.length) return null;

    const preferred = entries.find(({ parsed }) => {
        const expression = normalizeName(parsed.expression);
        return expression.includes('neutral') || expression.includes('normal') || expression.includes('default');
    });

    return (preferred || entries[0]).entry;
}

function findEntryMatchInText(entries, messageText, preferredCharacter = '') {
    const cleaned = String(messageText || '').toLowerCase().replace(/[^\w\s]/g, ' ');
    const normalizedPreferredCharacter = normalizeName(preferredCharacter);
    const prioritizedEntries = normalizedPreferredCharacter
        ? [
            ...entries.filter(entry => parseEntryName(entry.name).character === normalizedPreferredCharacter),
            ...entries.filter(entry => parseEntryName(entry.name).character !== normalizedPreferredCharacter),
        ]
        : entries;

    for (const entry of prioritizedEntries) {
        const parsed = parseEntryName(entry.name);
        const needles = [
            normalizeName(parsed.raw).replace(/_/g, ' '),
            normalizeName(parsed.expression).replace(/_/g, ' '),
        ].filter(Boolean);

        if (needles.some(needle => textIncludesNeedle(cleaned, needle))) {
            return entry;
        }
    }

    return null;
}

function splitParagraphsWithOffsets(text) {
    const value = String(text || '');
    const paragraphs = [];
    const regex = /[^\r\n]+/g;
    let match;

    while ((match = regex.exec(value)) !== null) {
        const raw = match[0];
        const trimmed = raw.trim();
        if (!trimmed) continue;
        const start = match.index;
        const end = start + raw.length;
        paragraphs.push({ text: trimmed, start, end });
    }

    return paragraphs;
}

function splitTextSegmentsWithOffsets(text, characterKeys = []) {
    const value = String(text || '');
    const segments = [];
    const regex = /[^\r\n.!?]+(?:[.!?]+|$)/g;
    let match;

    while ((match = regex.exec(value)) !== null) {
        const raw = match[0];
        const trimmed = raw.trim();
        if (!trimmed) continue;

        segments.push({
            text: trimmed,
            start: match.index,
            end: match.index + raw.length,
        });
    }

    const baseSegments = segments.length ? segments : splitParagraphsWithOffsets(value);
    const expandedSegments = [];

    for (const segment of baseSegments) {
        const mentionSegments = splitSegmentByCharacterMentions(segment, characterKeys);
        expandedSegments.push(...mentionSegments);
    }

    return expandedSegments.length ? expandedSegments : baseSegments;
}

function splitSegmentByCharacterMentions(segment, characterKeys = []) {
    if (!Array.isArray(characterKeys) || characterKeys.length < 2) {
        return [segment];
    }

    const matches = [];

    for (const character of characterKeys) {
        const plainName = String(character || '').replace(/_/g, ' ').trim();
        if (!plainName) continue;

        const pattern = new RegExp(`\\b${escapeRegExp(plainName)}\\b\\s*[:\\-\\u2013\\u2014]?`, 'gi');
        let match;

        while ((match = pattern.exec(segment.text)) !== null) {
            matches.push({
                character,
                index: match.index,
            });
        }
    }

    matches.sort((a, b) => a.index - b.index);

    const uniqueCharacters = new Set(matches.map(match => match.character));
    if (matches.length < 2 || uniqueCharacters.size < 2) {
        return [segment];
    }

    const result = [];
    for (let index = 0; index < matches.length; index++) {
        const start = matches[index].index;
        const end = matches[index + 1]?.index ?? segment.text.length;
        const text = segment.text.slice(start, end).trim();
        if (!text) continue;

        const leadingWhitespace = segment.text.slice(start, end).search(/\S/);
        const adjustedStart = segment.start + start + Math.max(leadingWhitespace, 0);
        result.push({
            text,
            start: adjustedStart,
            end: segment.start + end,
        });
    }

    return result.length ? result : [segment];
}

function analyzeParagraphDominance(text, characterKeys) {
    const segments = splitTextSegmentsWithOffsets(text, characterKeys);
    const assignments = segments.map(segment => {
        const scores = scoreCharacters(segment.text, characterKeys);
        const top = scores[0];
        return {
            ...segment,
            character: top?.score > 0 ? top.character : null,
            score: top?.score || 0,
        };
    });

    const counts = new Map();
    const blocks = [];
    let currentBlock = null;

    for (const paragraph of assignments) {
        if (!paragraph.character) {
            currentBlock = null;
            continue;
        }

        if (!currentBlock || currentBlock.character !== paragraph.character) {
            currentBlock = {
                character: paragraph.character,
                start: paragraph.start,
                end: paragraph.end,
                count: 1,
                score: paragraph.score,
            };
            blocks.push(currentBlock);
        } else {
            currentBlock.end = paragraph.end;
            currentBlock.count += 1;
            currentBlock.score += paragraph.score;
        }

        counts.set(paragraph.character, (counts.get(paragraph.character) || 0) + 1);
    }

    const firstChar = assignments.find(p => p.character)?.character || null;
    const lastChar = [...assignments].reverse().find(p => p.character)?.character || null;
    const sortedCounts = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
    const primary = sortedCounts[0]?.[0] || null;
    const secondary = sortedCounts[1]?.[0] || null;
    const totalParagraphs = segments.length || 1;
    const primaryCount = primary ? counts.get(primary) || 0 : 0;
    const secondaryCount = secondary ? counts.get(secondary) || 0 : 0;
    const isSingleDominant = !!primary && firstChar === primary && lastChar === primary && primaryCount >= totalParagraphs * 0.6 && primaryCount >= Math.max(secondaryCount * 1.5, 1);
    const blocksByCharacter = new Map();

    for (const block of blocks) {
        const arr = blocksByCharacter.get(block.character) || [];
        arr.push(block);
        blocksByCharacter.set(block.character, arr);
    }

    return {
        assignments,
        blocksByCharacter,
        counts,
        primary,
        secondary,
        isSingleDominant,
    };
}

function getTextForCharacter(text, dominance, characterKey) {
    const key = normalizeName(characterKey || '');
    if (!key || !dominance?.assignments?.length) {
        return String(text || '');
    }

    const segments = dominance.assignments
        .filter(segment => segment.character === key)
        .map(segment => segment.text)
        .filter(Boolean);

    return segments.length ? segments.join('\n') : String(text || '');
}

function pickDominantBlock(blocksByCharacter, character) {
    const blocks = blocksByCharacter.get(character) || [];
    if (!blocks.length) return null;
    const sorted = blocks.slice().sort((a, b) => {
        if (b.count !== a.count) return b.count - a.count;
        const spanA = a.end - a.start;
        const spanB = b.end - b.start;
        if (spanB !== spanA) return spanB - spanA;
        return a.start - b.start;
    });
    return sorted[0];
}

function buildPlaceholder(name) {
    return `{{img::${name}}}`;
}

function createId() {
    return crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function isInsideCode(node) {
    let parent = node?.parentNode;

    while (parent) {
        if (CODE_TAGS.has(parent.nodeName.toLowerCase())) {
            return true;
        }
        parent = parent.parentNode;
    }

    return false;
}

function revertInjectedPlaceholders(root) {
    root.querySelectorAll?.('.image-embed-expression').forEach(node => {
        const placeholder = node.dataset.placeholder;
        if (placeholder) {
            node.replaceWith(document.createTextNode(placeholder));
        }
    });
}

function clearAutoInjectedExpressions(root) {
    if (!root) return;

    root.querySelectorAll?.('.image-embeds-ai-marker').forEach(node => node.remove());
    revertInjectedPlaceholders(root);
    removeDuplicatePlaceholders(root);
}

function getCurrentCardExpressionEntries() {
    return [...getCharacterEntries(), ...getUserEntries()]
        .filter(entry => entry?.name);
}

function removeExpressionPlaceholdersForEntries(root, entries) {
    if (!root || !entries?.length) return;

    const names = new Set();
    for (const entry of entries) {
        names.add(normalizeName(entry.name));
    }

    root.querySelectorAll?.('.image-embed-expression').forEach(node => {
        // Regenerate is intentionally strict for the open chat: remove rendered
        // expression images so old cross-character picks cannot block reinjection.
        node.remove();
    });

    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const updates = [];
    let current;

    while ((current = walker.nextNode())) {
        if (isInsideCode(current)) continue;
        const text = current.nodeValue || '';
        PLACEHOLDER_REGEX.lastIndex = 0;
        if (!PLACEHOLDER_REGEX.test(text)) continue;

        PLACEHOLDER_REGEX.lastIndex = 0;
        let changed = false;
        const rebuilt = text.replace(PLACEHOLDER_REGEX, (fullMatch, rawName) => {
            if (names.has(normalizeName(rawName))) {
                changed = true;
                return '';
            }
            return fullMatch;
        });

        if (changed) {
            updates.push({ node: current, value: rebuilt.replace(/\n{3,}/g, '\n\n') });
        }
    }

    for (const update of updates) {
        update.node.nodeValue = update.value;
    }

    PLACEHOLDER_REGEX.lastIndex = 0;
    removeDuplicatePlaceholders(root);
}

function removeDuplicatePlaceholders(root) {
    const seen = new Set();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const updates = [];
    let current;

    while ((current = walker.nextNode())) {
        if (isInsideCode(current)) continue;
        const text = current.nodeValue || '';
        PLACEHOLDER_REGEX.lastIndex = 0;
        if (!PLACEHOLDER_REGEX.test(text)) {
            continue;
        }

        PLACEHOLDER_REGEX.lastIndex = 0;
        let match;
        let lastIndex = 0;
        let rebuilt = '';
        let changed = false;

        while ((match = PLACEHOLDER_REGEX.exec(text)) !== null) {
            const [fullMatch] = match;
            const key = fullMatch.toLowerCase();
            if (!seen.has(key)) {
                seen.add(key);
                rebuilt += text.slice(lastIndex, match.index) + fullMatch;
            } else {
                rebuilt += text.slice(lastIndex, match.index);
                changed = true;
            }
            lastIndex = PLACEHOLDER_REGEX.lastIndex;
        }

        if (changed) {
            rebuilt += text.slice(lastIndex);
            updates.push({ node: current, value: rebuilt });
        }
    }

    for (const update of updates) {
        update.node.nodeValue = update.value;
    }

    PLACEHOLDER_REGEX.lastIndex = 0;
}

function createImageNode(entry, rawName) {
    const placeholder = buildPlaceholder(rawName.trim() || entry.name || '');
    const wrapper = document.createElement('span');
    wrapper.className = 'image-embed-expression';
    wrapper.dataset.placeholder = placeholder;

    const image = document.createElement('img');
    image.src = entry.url;
    image.alt = entry.name || rawName || 'expression';
    image.loading = 'lazy';
    image.referrerPolicy = 'no-referrer';
    wrapper.appendChild(image);

    const label = document.createElement('span');
    label.className = 'image-embeds-label';
    label.textContent = entry.name || rawName || '';
    wrapper.appendChild(label);

    return wrapper;
}

function replaceTextNode(textNode) {
    const text = textNode.nodeValue;
    PLACEHOLDER_REGEX.lastIndex = 0;
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;
    let match;

    while ((match = PLACEHOLDER_REGEX.exec(text)) !== null) {
        const [fullMatch, rawName] = match;
        if (match.index > lastIndex) {
            fragment.append(document.createTextNode(text.slice(lastIndex, match.index)));
        }

        const entry = findEntryByName(rawName);
        if (entry?.url) {
            fragment.append(createImageNode(entry, rawName));
        } else {
            fragment.append(document.createTextNode(fullMatch));
        }

        lastIndex = PLACEHOLDER_REGEX.lastIndex;
    }

    if (lastIndex < text.length) {
        fragment.append(document.createTextNode(text.slice(lastIndex)));
    }

    textNode.replaceWith(fragment);
}

function buildPlacementForEntry(entry, character, dominance, targetCharacter = character) {
    const block = pickDominantBlock(dominance.blocksByCharacter, targetCharacter);
    const targetOffset = block ? (block.start + block.end) / 2 : null;
    return { entry, character, targetCharacter, targetOffset };
}

function getPlacementCharacterKey(placement) {
    const parsedCharacter = parseEntryName(placement?.entry?.name).character;
    const key = parsedCharacter || placement?.targetCharacter || placement?.character || DEFAULT_CHARACTER_GROUP;
    return normalizeName(key || DEFAULT_CHARACTER_GROUP);
}

async function refinePlacementWithAI(placement, allEntries, messageText, dominance, characterKey) {
    const targetCharacter = normalizeName(characterKey || getPlacementCharacterKey(placement));
    if (!targetCharacter || targetCharacter === DEFAULT_CHARACTER_GROUP) {
        return placement;
    }

    const entries = filterEntriesForCharacter(allEntries, targetCharacter, { includeDefault: true });
    if (!entries.length) {
        return placement;
    }

    const scopedMessageText = getTextForCharacter(messageText, dominance, targetCharacter);
    const displayName = targetCharacter.replace(/_/g, ' ');
    const selectedExpressionName = await detectExpressionWithAI({
        messageText: scopedMessageText,
        entries,
        characterName: displayName,
        settings: ensureSettings(),
        parseEntryName,
        normalizeName,
    });
    if (!selectedExpressionName) {
        return placement;
    }

    const selectedEntry = entries.find(entry => entry.name === selectedExpressionName);
    if (!selectedEntry) {
        return placement;
    }

    return { ...placement, entry: selectedEntry, character: targetCharacter, targetCharacter };
}

async function refinePlacementsWithAI(placements, allEntries, messageText, dominance, characterKeys, allowMultiple) {
    const result = [];
    const usedCharacters = new Set();
    const maxCount = allowMultiple ? 2 : 1;

    for (const placement of placements || []) {
        if (result.length >= maxCount) break;
        const character = getPlacementCharacterKey(placement);
        if (usedCharacters.has(character)) continue;

        const refined = await refinePlacementWithAI(placement, allEntries, messageText, dominance, character);
        result.push(refined);
        usedCharacters.add(character);
    }

    for (const character of characterKeys || []) {
        if (result.length >= maxCount) break;
        const normalizedCharacter = normalizeName(character);
        if (!normalizedCharacter || usedCharacters.has(normalizedCharacter)) continue;

        const fallbackEntry = getFallbackEntryForCharacter(groupEntriesByCharacter(allEntries), normalizedCharacter);
        if (!fallbackEntry) continue;

        const placement = buildPlacementForEntry(fallbackEntry, normalizedCharacter, dominance, normalizedCharacter);
        const refined = await refinePlacementWithAI(placement, allEntries, messageText, dominance, normalizedCharacter);
        result.push(refined);
        usedCharacters.add(normalizedCharacter);
    }

    return result.length ? result.slice(0, maxCount) : placements;
}

function dedupePlacementsByEntryAndCharacter(placements, allowMultiple = false) {
    const result = [];
    const seenEntries = new Set();
    const seenCharacters = new Set();
    const maxCount = allowMultiple ? 2 : 1;

    for (const placement of placements || []) {
        const entry = placement?.entry;
        if (!entry) continue;

        const entryKey = entry.id || entry.name || entry.url;
        const characterKey = getPlacementCharacterKey(placement);
        if (seenEntries.has(entryKey) || seenCharacters.has(characterKey)) {
            continue;
        }

        seenEntries.add(entryKey);
        seenCharacters.add(characterKey);
        result.push(placement);
        if (result.length >= maxCount) break;
    }

    return result;
}

function isRecentAssistantMessage(messageId, windowSize = ADVANCED_AI_RECENT_MESSAGE_WINDOW) {
    const numericId = Number(messageId);
    if (Number.isNaN(numericId) || !Array.isArray(chat) || !chat.length) {
        return false;
    }

    let seenAssistantMessages = 0;
    for (let index = chat.length - 1; index >= 0; index--) {
        const message = chat?.[index];
        if (!message || message.is_system || message.is_user) continue;
        if (index === numericId) {
            return true;
        }
        seenAssistantMessages++;
        if (seenAssistantMessages >= windowSize) {
            return false;
        }
    }

    return false;
}

function pickEntriesForMessageLegacy(messageId, allowMultiple = false) {
    const message = chat?.[messageId];
    if (!message || message.is_system) return [];

    // Handle user messages - use same detection logic as character messages
    if (message.is_user) {
        const userEntries = getUserEntries();
        if (!userEntries.length) return [];

        const messageText = String(message.mes || '').toLowerCase();
        const cleaned = messageText.replace(/[^\w\s]/g, ' ');

        const maxCount = allowMultiple ? 2 : 1;
        const selected = [];
        const seenEntries = new Set();

        // Try to find entries that match keywords in user message (with first person support)
        for (const entry of userEntries) {
            const needles = buildNeedles(entry, true); // true = include first person variations
            if (needles.some(needle => cleaned.includes(needle))) {
                const key = entry.id || entry.name || entry.url;
                if (!seenEntries.has(key)) {
                    selected.push({ entry, character: 'user' });
                    seenEntries.add(key);
                    if (selected.length >= maxCount) break;
                }
            }
        }

        // If no direct keyword match found, try emotion hint scoring before falling back.
        if (selected.length === 0 && userEntries.length > 0) {
            const scoredEntry = scoreExpressionFromText(messageText, userEntries);
            const fallbackEntry = scoredEntry || (userEntries.length === 1 ? userEntries[0] : null);
            if (fallbackEntry) {
                selected.push({ entry: fallbackEntry, character: 'user' });
            }
        }

        return selected;
    }

    // Handle character/assistant messages
    const messageTextRaw = String(message.mes || '');
    const activeCardEntries = getCharacterEntries();
    const usingCrossPackEntries = activeCardEntries.length === 0;
    const entries = activeCardEntries.length
        ? activeCardEntries
        : getCrossPackEntriesForMessage(messageTextRaw);
    if (!entries.length) return [];

    const settings = ensureSettings();
    const messageText = messageTextRaw.toLowerCase();
    const grouped = groupEntriesByCharacter(entries);
    const characterKeys = Array.from(grouped.keys()).filter(key => key && key !== DEFAULT_CHARACTER_GROUP);
    const strictCharacterScopedEntries = hasStrictCharacterScopedEntries(entries);
    let activeMessageCharacterKeys = strictCharacterScopedEntries
        ? detectActiveCharacterKeys(messageTextRaw, characterKeys)
        : [];
    if (usingCrossPackEntries && strictCharacterScopedEntries && !activeMessageCharacterKeys.length) {
        activeMessageCharacterKeys = scoreCharacters(messageTextRaw, characterKeys)
            .filter(result => result.score > 0)
            .map(result => result.character);
    }
    const eligibleCharacterKeys = activeMessageCharacterKeys.length ? activeMessageCharacterKeys : characterKeys;
    const messageCharacterKey = resolveMessageCharacterKey(messageId, eligibleCharacterKeys);
    const activeCharacterKey = getActiveCharacterExpressionKey();
    const recentContext = buildRecentInteractionContext(messageId, eligibleCharacterKeys);
    const characterScores = eligibleCharacterKeys.length ? scoreCharacters(messageText, eligibleCharacterKeys) : [];
    const dominance = analyzeParagraphDominance(messageTextRaw, eligibleCharacterKeys);
    const selected = [];
    let maxCount = allowMultiple ? 2 : 1;

    if (strictCharacterScopedEntries && isUnknownLeadingCharacter(messageTextRaw, characterKeys)) {
        return [];
    }

    if (strictCharacterScopedEntries && characterKeys.length > 1 && !activeMessageCharacterKeys.length) {
        return [];
    }

    if (allowMultiple && activeMessageCharacterKeys.length > 1) {
        maxCount = Math.min(2, activeMessageCharacterKeys.length);
    } else if (allowMultiple && characterScores.length > 1) {
        const primaryScore = characterScores[0];
        const secondaryScore = characterScores[1];
        if (!secondaryScore || secondaryScore.score < 1 || (primaryScore && secondaryScore.score < primaryScore.score * 0.5)) {
            maxCount = 1;
        }
    }

    if (dominance.isSingleDominant && activeMessageCharacterKeys.length <= 1) {
        maxCount = 1;
    }

    const seenEntries = new Set();
    const desiredCharacters = [];

    if (messageCharacterKey && grouped.has(messageCharacterKey)) {
        desiredCharacters.push(messageCharacterKey);
    }

    for (const character of activeMessageCharacterKeys) {
        if (grouped.has(character) && !desiredCharacters.includes(character)) {
            desiredCharacters.push(character);
        }
    }

    if (dominance.primary) {
        desiredCharacters.push(dominance.primary);
    } else if (characterScores[0]) {
        desiredCharacters.push(characterScores[0].character);
    }

    if (!desiredCharacters.length && recentContext.primary && grouped.has(recentContext.primary)) {
        desiredCharacters.push(recentContext.primary);
    }

    if (!desiredCharacters.length && !strictCharacterScopedEntries && activeCharacterKey && grouped.has(activeCharacterKey)) {
        desiredCharacters.push(activeCharacterKey);
    }

    if (allowMultiple && !dominance.isSingleDominant) {
        const secondaryCandidate = dominance.secondary || characterScores[1]?.character;
        if (secondaryCandidate && !desiredCharacters.includes(secondaryCandidate)) {
            desiredCharacters.push(secondaryCandidate);
        }
    }

    for (const character of desiredCharacters) {
        const scopedMessageText = getTextForCharacter(messageTextRaw, dominance, character);
        const fullMessageEntry = selectEntryForCharacter(grouped, character, messageTextRaw);
        const scopedEntry = selectEntryForCharacter(grouped, character, scopedMessageText);
        const entry = (activeMessageCharacterKeys.length <= 1 ? fullMessageEntry : scopedEntry)
            || scopedEntry
            || fullMessageEntry
            || (activeMessageCharacterKeys.includes(character) ? getFallbackEntryForCharacter(grouped, character) : null);
        const key = entry ? (entry.id || entry.url || entry.name) : null;
        if (entry && !seenEntries.has(key)) {
            selected.push(buildPlacementForEntry(entry, character, dominance, character));
            seenEntries.add(key);
        }
        if (selected.length >= maxCount) break;
    }

    if (selected.length >= maxCount) {
        return selected;
    }

    const preferredDirectCharacter = messageCharacterKey || dominance.primary || characterScores[0]?.character || activeCharacterKey;
    const directMessageText = preferredDirectCharacter
        ? getTextForCharacter(messageTextRaw, dominance, preferredDirectCharacter)
        : messageText;
    const directEntries = preferredDirectCharacter
        ? filterEntriesForCharacter(entries, preferredDirectCharacter, { includeDefault: true })
        : entries;
    const directMatch = findEntryMatchInText(directEntries, directMessageText, preferredDirectCharacter);
    const directKey = directMatch ? (directMatch.id || directMatch.url || directMatch.name) : null;
    if (directMatch && !seenEntries.has(directKey)) {
        const parsed = parseEntryName(directMatch.name);
        const targetCharacter = parsed.character || preferredDirectCharacter || activeCharacterKey;
        selected.push(buildPlacementForEntry(directMatch, targetCharacter, dominance, targetCharacter));
        seenEntries.add(directKey);
        if (selected.length >= maxCount) {
            return selected;
        }
    }

    if (selected.length === 0 && grouped.has(DEFAULT_CHARACTER_GROUP)) {
        const defaultEntry = selectEntryForCharacter(grouped, DEFAULT_CHARACTER_GROUP, messageText);
        const defaultKey = defaultEntry ? (defaultEntry.id || defaultEntry.url || defaultEntry.name) : null;
        if (defaultEntry && !seenEntries.has(defaultKey)) {
            const targetCharacter = preferredDirectCharacter || activeCharacterKey || DEFAULT_CHARACTER_GROUP;
            selected.push(buildPlacementForEntry(defaultEntry, DEFAULT_CHARACTER_GROUP, dominance, targetCharacter));
            seenEntries.add(defaultKey);
            if (selected.length >= maxCount) {
                return selected;
            }
        }
    }

    if (selected.length === 0) {
        const scoredEntries = preferredDirectCharacter
            ? filterEntriesForCharacter(entries, preferredDirectCharacter, { includeDefault: true })
            : entries;
        const scoredFallback = scoreExpressionFromText(messageTextRaw, scoredEntries.length ? scoredEntries : entries);
        const scoredKey = scoredFallback ? (scoredFallback.id || scoredFallback.url || scoredFallback.name) : null;

        if (scoredFallback && !seenEntries.has(scoredKey)) {
            const parsed = parseEntryName(scoredFallback.name);
            const targetCharacter = parsed.character || preferredDirectCharacter || activeCharacterKey || DEFAULT_CHARACTER_GROUP;
            selected.push(buildPlacementForEntry(scoredFallback, targetCharacter, dominance, targetCharacter));
            seenEntries.add(scoredKey);
            if (selected.length >= maxCount) {
                return selected;
            }
        }
    }

    const disambiguationNeeded = grouped.size > 1 || entries.length > 1;
    if (disambiguationNeeded) {
        return selected;
    }

    const fallbackCharacter = (activeCharacterKey && grouped.has(activeCharacterKey))
        ? activeCharacterKey
        : (characterKeys[0] || DEFAULT_CHARACTER_GROUP);
    const fallbackMessageText = getTextForCharacter(messageTextRaw, dominance, fallbackCharacter);
    const fallbackEntry = selectEntryForCharacter(grouped, fallbackCharacter, fallbackMessageText);
    const fallbackKey = fallbackEntry ? (fallbackEntry.id || fallbackEntry.url || fallbackEntry.name) : null;
    if (fallbackEntry && !seenEntries.has(fallbackKey)) {
        selected.push(buildPlacementForEntry(fallbackEntry, fallbackCharacter, dominance, fallbackCharacter));
    }

    return selected.slice(0, maxCount);
}

async function pickEntriesForMessage(messageId, allowMultiple = false) {
    const message = chat?.[messageId];
    const settings = ensureSettings();
    const cacheEntries = message?.is_user ? getUserEntries() : getCharacterEntries();
    const groupedEntries = groupEntriesByCharacter(cacheEntries);
    const contextCharacterKeys = Array.from(groupedEntries.keys()).filter(key => key && key !== DEFAULT_CHARACTER_GROUP);
    const strictCharacterScopedEntries = !message?.is_user && hasStrictCharacterScopedEntries(cacheEntries);
    const activeMessageCharacterKeys = strictCharacterScopedEntries
        ? detectActiveCharacterKeys(String(message?.mes || ''), contextCharacterKeys)
        : [];
    const eligibleCharacterKeys = activeMessageCharacterKeys.length ? activeMessageCharacterKeys : contextCharacterKeys;
    const messageCharacterKey = resolveMessageCharacterKey(messageId, eligibleCharacterKeys);
    const recentContext = buildRecentInteractionContext(messageId, eligibleCharacterKeys);
    const speakerCharacterKey = getMessageSpeakerKey(message, eligibleCharacterKeys);
    const activeCharacterKey = getActiveCharacterExpressionKey();
    const targetCharacterKey = speakerCharacterKey
        || messageCharacterKey
        || (!strictCharacterScopedEntries && contextCharacterKeys.some(key => characterKeysMatch(key, activeCharacterKey)) ? activeCharacterKey : '')
        || (eligibleCharacterKeys.length === 1 ? eligibleCharacterKeys[0] : '');
    const aiCharacterName = targetCharacterKey || recentContext.primary || activeCharacterKey || characters?.[this_chid]?.name || 'Character';
    const cacheKey = getStringHash(JSON.stringify({
        messageId,
        messageText: String(message?.mes || ''),
        isUser: !!message?.is_user,
        allowMultiple: !!allowMultiple,
        advancedEnabled: !!settings.advancedExpressionsEnabled,
        detectionVersion: EXPRESSION_DETECTION_VERSION,
        currentMode,
        activeCharacter: activeCharacterKey,
        speakerCharacter: speakerCharacterKey,
        targetCharacter: targetCharacterKey,
        messageCharacter: messageCharacterKey,
        recentContext: recentContext.primary,
        activeMessageCharacters: activeMessageCharacterKeys,
        llmEligible: canUseLLMForMessage(messageId),
        entries: cacheEntries.map(entry => `${entry.id || ''}:${entry.name || ''}:${entry.url || ''}`),
    }));

    if (messagePlacementCache.has(cacheKey)) {
        return messagePlacementCache.get(cacheKey);
    }

    const legacySelections = pickEntriesForMessageLegacy(messageId, allowMultiple);

    const canUseAdvancedLLM = !!settings.advancedExpressionsEnabled && canUseLLMForMessage(messageId);

    if (!message || message.is_system || message.is_user || !canUseAdvancedLLM) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    const allEntries = getCharacterEntries();
    if (!allEntries.length) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    if (hasStrictCharacterScopedEntries(allEntries) && isUnknownLeadingCharacter(String(message.mes || ''), contextCharacterKeys)) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    if (hasStrictCharacterScopedEntries(allEntries) && contextCharacterKeys.length > 1 && !activeMessageCharacterKeys.length) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    if (!targetCharacterKey && contextCharacterKeys.length > 1) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    // Filter entries to only include expressions belonging to the speaking character.
    // If Flo is speaking, Miyuki/Sayaka-prefixed entries must never be candidates.
    const entries = targetCharacterKey
        ? filterEntriesForCharacter(allEntries, targetCharacterKey, { includeDefault: true })
        : allEntries;

    if (!entries.length) {
        messagePlacementCache.set(cacheKey, legacySelections);
        return legacySelections;
    }

    const messageTextRaw = String(message.mes || '');
    const dominance = analyzeParagraphDominance(messageTextRaw, eligibleCharacterKeys);

    if (allowMultiple && settings.doubleEnabled && activeMessageCharacterKeys.length > 1) {
        const refinedSelections = await refinePlacementsWithAI(
            legacySelections,
            allEntries,
            messageTextRaw,
            dominance,
            activeMessageCharacterKeys,
            true,
        );
        messagePlacementCache.set(cacheKey, refinedSelections);
        consumeLLMEligibility(messageId);
        return refinedSelections;
    }

    if (legacySelections.length > 0) {
        const refinedSelections = await refinePlacementsWithAI(
            legacySelections,
            allEntries,
            messageTextRaw,
            dominance,
            activeMessageCharacterKeys,
            allowMultiple && settings.doubleEnabled,
        );
        messagePlacementCache.set(cacheKey, refinedSelections);
        consumeLLMEligibility(messageId);
        return refinedSelections;
    }

    const aiSelectedExpressionName = await detectExpressionWithAI({
        messageText: getTextForCharacter(messageTextRaw, dominance, targetCharacterKey),
        entries,
        characterName: aiCharacterName,
        settings,
        parseEntryName,
        normalizeName,
    });

    if (!aiSelectedExpressionName) {
        messagePlacementCache.set(cacheKey, legacySelections);
        consumeLLMEligibility(messageId);
        return legacySelections;
    }

    const aiSelectedEntry = entries.find(e => e.name === aiSelectedExpressionName);
    if (!aiSelectedEntry) {
        messagePlacementCache.set(cacheKey, legacySelections);
        consumeLLMEligibility(messageId);
        return legacySelections;
    }

    if (allowMultiple && settings.doubleEnabled) {
        const alreadyIncluded = legacySelections.some(item => (item.entry.id || item.entry.name || item.entry.url) === (aiSelectedEntry.id || aiSelectedEntry.name || aiSelectedEntry.url));
        if (alreadyIncluded) {
            messagePlacementCache.set(cacheKey, legacySelections);
            consumeLLMEligibility(messageId);
            return legacySelections;
        }

        const result = [legacySelections[0], { entry: aiSelectedEntry, character: aiCharacterName, targetCharacter: aiCharacterName, targetOffset: null }]
            .filter(Boolean)
            .slice(0, 2);
        messagePlacementCache.set(cacheKey, result);
        consumeLLMEligibility(messageId);
        return result;
    }

    const replacement = legacySelections[0]
        ? { ...legacySelections[0], entry: aiSelectedEntry }
        : { entry: aiSelectedEntry, character: aiCharacterName, targetCharacter: aiCharacterName, targetOffset: null };

    const result = [replacement];
    messagePlacementCache.set(cacheKey, result);
    consumeLLMEligibility(messageId);
    return result;
}

function insertPlaceholderNearMatch(root, entry) {
    const token = buildPlaceholder(entry.name);
    const needles = buildNeedles(entry);
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let targetNode = null;

    while (!targetNode) {
        const node = walker.nextNode();
        if (!node) break;
        const text = (node.nodeValue || '').toLowerCase();
        if (needles.some(needle => text.includes(needle))) {
            targetNode = node;
            break;
        }
    }

    const textNode = document.createTextNode(`\n${token}\n`);
    if (targetNode && targetNode.parentNode) {
        targetNode.parentNode.insertBefore(textNode, targetNode.nextSibling);
    } else {
        root.append(textNode);
    }
}

function collectTextNodesWithOffsets(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes = [];
    let currentOffset = 0;
    let node;

    while ((node = walker.nextNode())) {
        const text = node.nodeValue || '';
        const length = text.length;
        nodes.push({
            node,
            start: currentOffset,
            end: currentOffset + length,
            textLower: text.toLowerCase(),
        });
        currentOffset += length;
    }

    return nodes;
}

function findBestNodeForOffset(nodes, offset, usedNodes) {
    let candidate = null;
    let bestDistance = Infinity;

    for (const meta of nodes) {
        if (usedNodes.has(meta.node)) continue;
        if (offset >= meta.start && offset <= meta.end) {
            return meta.node;
        }
        const distance = offset < meta.start ? meta.start - offset : offset - meta.end;
        if (distance < bestDistance) {
            bestDistance = distance;
            candidate = meta.node;
        }
    }

    return candidate;
}

function findNodeMatchingNeedles(nodes, entry, usedNodes) {
    const needles = buildNeedles(entry);
    for (const meta of nodes) {
        if (usedNodes.has(meta.node)) continue;
        if (needles.some(needle => meta.textLower.includes(needle))) {
            return meta.node;
        }
    }
    return null;
}

function insertPlaceholdersSequential(root, entries) {
    if (!entries?.length) return;
    const nodes = collectTextNodesWithOffsets(root);
    let minOffset = 0;
    const usedNodes = new Set();

    for (const entry of entries) {
        const token = buildPlaceholder(entry.name);
        const needles = buildNeedles(entry);
        let chosen = null;

        for (const meta of nodes) {
            if (meta.end < minOffset) continue;
            if (usedNodes.has(meta.node)) continue;
            if (needles.some(needle => meta.textLower.includes(needle))) {
                chosen = meta.node;
                break;
            }
        }

        if (!chosen) {
            for (const meta of nodes) {
                if (usedNodes.has(meta.node)) continue;
                if (needles.some(needle => meta.textLower.includes(needle))) {
                    chosen = meta.node;
                    break;
                }
            }
        }

        const textNode = document.createTextNode(`\n${token}\n`);
        if (chosen && chosen.parentNode) {
            chosen.parentNode.insertBefore(textNode, chosen.nextSibling);
            usedNodes.add(chosen);
            const meta = nodes.find(n => n.node === chosen);
            if (meta) {
                minOffset = meta.end;
            }
        } else {
            root.append(textNode);
        }
    }
}

function insertPlaceholdersWithTargets(root, placements) {
    if (!placements?.length) return;
    const nodes = collectTextNodesWithOffsets(root);
    const usedNodes = new Set();
    const ordered = [...placements].sort((a, b) => {
        const aOffset = Number.isFinite(a.targetOffset) ? a.targetOffset : Infinity;
        const bOffset = Number.isFinite(b.targetOffset) ? b.targetOffset : Infinity;
        return aOffset - bOffset;
    });

    for (const placement of ordered) {
        const entry = placement.entry;
        const token = buildPlaceholder(entry.name);
        let targetNode = null;

        if (Number.isFinite(placement.targetOffset)) {
            targetNode = findBestNodeForOffset(nodes, placement.targetOffset, usedNodes);
        }

        if (!targetNode) {
            targetNode = findNodeMatchingNeedles(nodes, entry, usedNodes);
        }

        const textNode = document.createTextNode(`\n${token}\n`);
        if (targetNode && targetNode.parentNode) {
            targetNode.parentNode.insertBefore(textNode, targetNode.nextSibling);
            usedNodes.add(targetNode);
        } else {
            root.append(textNode);
        }
    }
}

async function autoInjectAfterGeneration(root, messageId) {
    const settings = ensureSettings();
    if (!settings.enabled) return;
    if (!root) return;
    const message = chat?.[messageId];
    if (message?.is_system) return;
    const textContent = (root.textContent || '').replace(/\u200b/g, '').trim();
    if (!textContent || textContent === '...' || textContent === '…') return;

    // Skip if an embed is already present from previous processing.
    if (root.querySelector('.image-embed-expression')) {
        return;
    }

    PLACEHOLDER_REGEX.lastIndex = 0;
    const hasPlaceholder = PLACEHOLDER_REGEX.test(root.textContent || '');
    if (hasPlaceholder) return;

    // Skip if this root was already fully processed in the current generation
    // (prevents redundant work when refreshAllMessages iterates all messages).
    if (processedMessages.get(root) === processedMessageGeneration) return;

    // Prevent concurrent calls for the same messageId (e.g. rapid scroll/render events).
    if (autoInjectInFlight.has(messageId)) return;
    autoInjectInFlight.add(messageId);

    try {
        clearAutoInjectedExpressions(root);

        const placements = await pickEntriesForMessage(messageId, settings.doubleEnabled);
        if (!placements.length) {
            // Still mark as processed so we don't retry on every scroll tick.
            processedMessages.set(root, processedMessageGeneration);
            return;
        }

        const uniqueEntries = dedupePlacementsByEntryAndCharacter(placements, settings.doubleEnabled);

        insertPlaceholdersWithTargets(root, uniqueEntries);
        renderPlaceholders(root);

        // Mark this root as done for the current generation.
        processedMessages.set(root, processedMessageGeneration);
    } finally {
        autoInjectInFlight.delete(messageId);
    }
}

function renderPlaceholders(root) {
    if (!root) return;
    revertInjectedPlaceholders(root);
    removeDuplicatePlaceholders(root);

    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodesToReplace = [];
    let current;

    while ((current = walker.nextNode())) {
        if (isInsideCode(current)) continue;
        PLACEHOLDER_REGEX.lastIndex = 0;
        if (PLACEHOLDER_REGEX.test(current.nodeValue)) {
            nodesToReplace.push(current);
        }
    }

    for (const node of nodesToReplace) {
        replaceTextNode(node);
    }
}

function renderList() {
    const list = $('#image_embeds_list');
    let entries;
    let storageKey;

    if (!list.length) return;

    list.empty();

    // Check if character is selected
    const charKey = getCharacterKey();

    if (currentMode === 'user') {
        // User mode requires a character to be selected
        if (!charKey) {
            list.append($('<div class="image-embeds-empty">Open a character chat to manage User Expressions.</div>'));
            return;
        }
        entries = getUserEntries();
        storageKey = charKey;

        if (!entries.length) {
            list.append($('<div class="image-embeds-empty">No user expressions added yet. Click + to add one.</div>'));
            return;
        }
    } else {
        // Character mode
        if (!charKey) {
            list.append($('<div class="image-embeds-empty">Open a character chat to manage Image Embeds.</div>'));
            return;
        }
        entries = getCharacterEntries();
        storageKey = charKey;

        if (!entries.length) {
            list.append($('<div class="image-embeds-empty">There are no expressions for this character yet. Click + to add one.</div>'));
            return;
        }
    }

    for (const entry of entries) {
        const row = $('<div class="image-embeds-row"></div>');
        const preview = $('<img class="image-embeds-preview" loading="lazy" alt="Expression preview">').attr('src', entry.url || '');
        const input = $('<input type="text" class="text_pole" autocomplete="off">').val(entry.name || '');
        const placeholder = $('<code class="image-embeds-placeholder"></code>').text(buildPlaceholder(entry.name || ''));
        const remove = $('<button type="button" class="menu_button menu_button_icon" title="Remove expression"><i class="fa-solid fa-trash-can"></i></button>');

        input.on('input', (event) => {
            entry.name = event.target.value;
            placeholder.text(buildPlaceholder(entry.name || ''));
            clearAiExpressionCache();
            saveSettingsDebounced();
            refreshAllMessages();
        });

        remove.on('click', async () => {
            // Delete the physical file from server
            await deleteExpressionFile(entry.url);

            // Remove from settings
            if (currentMode === 'user') {
                extension_settings[SETTINGS_KEY].characters[storageKey].userEntries =
                    getUserEntries().filter(x => x.id !== entry.id);
            } else {
                extension_settings[SETTINGS_KEY].characters[storageKey].entries =
                    getCharacterEntries().filter(x => x.id !== entry.id);
            }
            clearAiExpressionCache();
            saveSettingsDebounced();
            renderList();
            refreshAllMessages();
        });

        row.append(preview, input, placeholder, remove);
        list.append(row);
    }
}

function ensureUniqueName(baseName) {
    const entries = currentMode === 'user' ? getUserEntries() : getCharacterEntries();
    const normBase = normalizeName(baseName) || 'expression';
    let candidate = normBase;
    let counter = 1;

    while (entries.some(e => normalizeName(e.name) === candidate)) {
        candidate = `${normBase}_${counter}`;
        counter += 1;
    }

    return candidate;
}

async function deleteExpressionFile(url) {
    if (!url) return;

    try {
        const response = await fetch('/api/files/delete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: url }),
        });

        if (!response.ok) {
            console.warn('Failed to delete expression file:', url, response.statusText);
        }
    } catch (error) {
        console.warn('Error deleting expression file:', url, error);
    }
}

async function addExpressionFromFile(file) {
    if (!file) return;

    let charKey = getCharacterKey();

    if (!charKey) {
        toastr.warning('Open a character chat to manage Image Embeds.', 'Image Embeds');
        return;
    }

    if (!file.type?.startsWith('image/')) {
        toastr.error('File must be an image.', 'Image Embeds');
        return;
    }

    try {
        let folder;
        if (currentMode === 'user') {
            folder = getUserFolder();
        } else {
            folder = getCharacterFolder();
        }

        const base64 = await getBase64Async(file);
        const base64Data = base64.split(',')[1];
        const extension = getFileExtension(file) || file.type.split('/')[1] || 'png';
        const slug = getStringHash(file.name || base64Data);
        const fileName = `${Date.now()}_${slug}`;
        const url = await saveBase64AsFile(base64Data, folder, fileName, extension);

        if (!url) {
            toastr.error('Failed to save image.', 'Image Embeds');
            return;
        }

        const defaultName = ensureUniqueName((file.name || 'expression').replace(/\.[^.]+$/, ''));
        const settings = ensureSettings();

        if (currentMode === 'user') {
            settings.characters[charKey].userEntries.push({
                id: createId(),
                name: defaultName,
                url,
                originalName: file.name || '',
            });
        } else {
            settings.characters[charKey].entries.push({
                id: createId(),
                name: defaultName,
                url,
                originalName: file.name || '',
            });
        }

        clearAiExpressionCache();
        saveSettingsDebounced();
        renderList();
        refreshAllMessages();
        toastr.success('Image expression added.', 'Image Embeds');
    } catch (error) {
        console.error('Failed to add image embed', error);
        toastr.error('An error occurred while adding the image.', 'Image Embeds');
    }
}

async function addExpressionsFromFiles(files) {
    if (!files?.length) return;
    for (const file of files) {
        // eslint-disable-next-line no-await-in-loop
        await addExpressionFromFile(file);
    }
}

/**
 * IntersectionObserver-based lazy processor.
 * Only calls autoInjectAfterGeneration for messages actually visible in the viewport,
 * which eliminates the main source of scroll lag.
 */
let _lazyObserver = null;

function getLazyObserver() {
    if (_lazyObserver) return _lazyObserver;
    _lazyObserver = new IntersectionObserver((entries) => {
        for (const entry of entries) {
            if (!entry.isIntersecting) continue;
            const mes = entry.target;
            const messageId = Number(mes.getAttribute('mesid'));
            const textNode = mes.querySelector('.mes_text');
            if (!textNode || Number.isNaN(messageId)) continue;
            _lazyObserver.unobserve(mes);
            renderPlaceholders(textNode);
            void autoInjectAfterGeneration(textNode, messageId);
        }
    }, { rootMargin: '200px 0px' }); // 200 px look-ahead so images load just before they scroll into view
    return _lazyObserver;
}

function refreshAllMessages() {
    // Debounce: if called repeatedly (e.g. CHAT_CHANGED then MORE_MESSAGES_LOADED in quick
    // succession) collapse into a single pass 80 ms later.
    clearTimeout(refreshAllDebounceTimer);
    refreshAllDebounceTimer = setTimeout(() => {
        const observer = getLazyObserver();
        document.querySelectorAll('#chat .mes').forEach(mes => {
            const messageId = Number(mes.getAttribute('mesid'));
            const textNode = mes.querySelector('.mes_text');
            if (!textNode || Number.isNaN(messageId)) return;

            // Always immediately render explicit {{img::}} placeholders (fast, sync).
            renderPlaceholders(textNode);

            // For auto-injection (slow, possibly async AI call) use the observer
            // so we only process messages that are in/near the viewport.
            observer.observe(mes);
        });
    }, 80);
}

async function regenerateCurrentCharacterExpressions() {
    const charKey = getCharacterKey();
    if (!charKey) {
        toastr.warning('Open a character chat first.', 'Image Embeds');
        return;
    }

    const entries = getCurrentCardExpressionEntries();
    if (!entries.length) {
        toastr.warning('No expressions found for this character card.', 'Image Embeds');
        return;
    }

    clearAiExpressionCache();
    processedMessageGeneration++;
    autoInjectInFlight.clear();
    llmEligibleMessageIds.clear();

    const messageRoots = Array.from(document.querySelectorAll('#chat .mes'))
        .map(mes => ({
            messageId: Number(mes.getAttribute('mesid')),
            root: mes.querySelector('.mes_text'),
        }))
        .filter(item => item.root && !Number.isNaN(item.messageId));

    for (const { root } of messageRoots) {
        processedMessages.delete(root);
        removeExpressionPlaceholdersForEntries(root, entries);
    }

    toastr.info('Old expressions cleared. Regenerating in 2 seconds...', 'Image Embeds');
    await delay(2000);

    for (const { root, messageId } of messageRoots) {
        renderPlaceholders(root);
        await autoInjectAfterGeneration(root, messageId);
    }

    toastr.success('Expressions regenerated for the current character card.', 'Image Embeds');
}

function onMessageRendered(messageId) {
    const root = document.querySelector(`#chat .mes[mesid="${messageId}"] .mes_text`);
    if (root) {
        // Force re-processing for this specific message (swipe, update, new generation).
        processedMessages.delete(root);
        renderPlaceholders(root);
        void autoInjectAfterGeneration(root, messageId);
    }
}

function scheduleMessageRender(messageId) {
    const targetId = rememberAssistantMessage(messageId);
    if (targetId === null || targetId === undefined || Number.isNaN(targetId)) return;
    requestAnimationFrame(() => onMessageRendered(targetId));
}

/**
 * Populate API providers dropdown
 */
function populateAPIProviders() {
    const select = $('#image_embeds_api_provider');
    select.empty();
    select.append('<option value="">-- Select Provider --</option>');

    const providers = getAllProviders();
    for (const provider of providers) {
        const optionText = provider.requiresAuth ? `${provider.name} (API Key required)` : provider.name;
        const option = $('<option></option>')
            .attr('value', provider.id)
            .text(optionText);
        select.append(option);
    }

    // Load saved provider if exists
    const settings = ensureSettings();
    if (settings.apiProvider) {
        select.val(settings.apiProvider);
    }
}

/**
 * Update API provider UI based on selected provider
 */
async function updateAPIProviderUI(provider) {
    if (!provider) {
        return;
    }

    const providerConfig = getProviderConfig(provider);
    if (!providerConfig) return;

    // Update provider info
    const infoDiv = $('#image_embeds_provider_info');
    let infoText = `<strong>${providerConfig.name}</strong>`;
    if (providerConfig.description) {
        infoText += ` - ${providerConfig.description}`;
    }
    infoDiv.html(infoText);

    // Show/hide custom base URL field
    if (providerConfig.editable) {
        $('#image_embeds_baseurl_container').show();
    } else {
        $('#image_embeds_baseurl_container').hide();
    }

    // Update model dropdown
    setModelSelectOptions(providerConfig.models || [], '');

    // Show/hide model container
    $('#image_embeds_model_container').show();

    // Update API key UI
    const apiKeyRequired = providerRequiresApiKey(provider);
    const apiKeyOptional = !!providerConfig.optionalAuth;
    $('#image_embeds_api_key').prop('disabled', !(apiKeyRequired || apiKeyOptional));
    $('#image_embeds_api_key').attr(
        'placeholder',
        apiKeyRequired
            ? 'Enter your API key'
            : (apiKeyOptional ? 'Optional. Leave empty to use the provider default.' : 'Not required for this provider'),
    );
    $('#image_embeds_api_key_hint').text(
        apiKeyRequired
            ? 'Your API key is stored locally and only sent to the provider'
            : (apiKeyOptional
                ? 'Optional. Leave this empty to use the anonymous/default provider key.'
                : 'This provider can work without an API key. Leave it empty unless your endpoint needs one.')
    );

    // Load saved model if exists
    const settings = ensureSettings();
    if (settings.apiModel) {
        $('#image_embeds_api_model').val(settings.apiModel);
        $('#image_embeds_api_model_custom').val(settings.apiModel);
    } else if (providerConfig.defaultModel) {
        settings.apiModel = providerConfig.defaultModel;
        $('#image_embeds_api_model').val(providerConfig.defaultModel);
        $('#image_embeds_api_model_custom').val(providerConfig.defaultModel);
    }

    if (providerConfig.editable) {
        $('#image_embeds_custom_baseurl').val(settings.customBaseUrl || '');
        $('#image_embeds_custom_baseurl').attr('placeholder', providerConfig.baseUrl || 'e.g., http://localhost:11434/v1');
    } else {
        $('#image_embeds_custom_baseurl').val('');
    }

    await refreshProviderModels(provider, { silent: true });
}

function bindUi() {
    $('#image_embeds_add').on('click', () => {
        $('#image_embeds_file_input').val('').trigger('click');
    });

    $('#image_embeds_file_input').on('change', async (event) => {
        const files = Array.from(event.target.files ?? []);
        event.target.value = '';
        await addExpressionsFromFiles(files);
    });

    $('#image_embeds_enabled').on('change', (event) => {
        ensureSettings().enabled = !!event.target.checked;
        saveSettingsDebounced();
    });

    $('#image_embeds_double_enabled').on('change', (event) => {
        const settings = ensureSettings();
        const previousValue = !!settings.doubleEnabled;
        const nextValue = !!event.target.checked;
        const shouldRestart = confirm('Changing Double Expressions needs a page restart so old injected expressions are cleared. Restart now?');

        if (!shouldRestart) {
            event.target.checked = previousValue;
            return;
        }

        settings.doubleEnabled = nextValue;
        clearAiExpressionCache();
        saveSettingsDebounced();
        setTimeout(() => window.location.reload(), 750);
    });

    $('#image_embeds_advanced_enabled').on('change', (event) => {
        ensureSettings().advancedExpressionsEnabled = !!event.target.checked;
        clearAiExpressionCache();
        saveSettingsDebounced();

        // Show/hide advanced settings section
        if (event.target.checked) {
            $('#image_embeds_advanced_settings').show();
            // Trigger AI detection untuk semua pesan
            refreshAllMessages();
        } else {
            $('#image_embeds_advanced_settings').hide();
            // Ketika disabled, biarkan ekspresi yang sudah ada tetap ada
            console.log('Advanced Expressions disabled - keeping existing expressions');
        }
    });

    // API Provider Selection
    $('#image_embeds_api_provider').on('change', (event) => {
        const provider = event.target.value;
        const providerConfig = getProviderConfig(provider);
        ensureSettings().apiProvider = provider;
        ensureSettings().apiModel = '';
        if (!providerConfig?.editable) {
            ensureSettings().customBaseUrl = '';
            $('#image_embeds_custom_baseurl').val('');
        }
        clearAiExpressionCache();

        renderConnectionStatus('disconnected', provider ? 'Disconnected' : 'Select a provider first');

        if (provider) {
            void updateAPIProviderUI(provider).then(() => {
                if (ensureSettings().autoConnectLastServer) {
                    void restartProviderConnection();
                }
            });
        } else {
            $('#image_embeds_provider_info').empty();
            $('#image_embeds_baseurl_container').hide();
            $('#image_embeds_model_container').hide();
        }
        saveSettingsDebounced();
    });

    // API Model Selection
    $('#image_embeds_api_model').on('change', (event) => {
        const model = event.target.value;
        ensureSettings().apiModel = model;
        $('#image_embeds_api_model_custom').val(model);
        clearAiExpressionCache();
        saveSettingsDebounced();
    });

    // API Model Custom Input
    $('#image_embeds_api_model_custom').on('input', (event) => {
        const model = event.target.value;
        ensureSettings().apiModel = model;
        clearAiExpressionCache();
        saveSettingsDebounced();
    });

    // Custom Base URL Input
    $('#image_embeds_custom_baseurl').on('input', (event) => {
        ensureSettings().customBaseUrl = event.target.value;
        clearAiExpressionCache();
        const provider = ensureSettings().apiProvider;
        if (provider) {
            scheduleProviderModelsRefresh(provider);
        }
        saveSettingsDebounced();
    });

    // API Key Input
    $('#image_embeds_api_key').on('input', (event) => {
        ensureSettings().apiKey = event.target.value;
        clearAiExpressionCache();
        const provider = ensureSettings().apiProvider;
        if (provider) {
            scheduleProviderModelsRefresh(provider);
        }
        saveSettingsDebounced();
    });

    // Restart Connection Button
    let restartConnectionAbortController = null;

    $('#image_embeds_test_connection').on('click', async () => {
        const button = $('#image_embeds_test_connection');
        const abortButton = $('#image_embeds_abort_test');

        restartConnectionAbortController = new AbortController();
        const { signal } = restartConnectionAbortController;

        button.prop('disabled', true);
        abortButton.show();

        try {
            await restartProviderConnection({ signal });
        } finally {
            restartConnectionAbortController = null;
            button.prop('disabled', false);
            abortButton.hide();
        }
    });

    $('#image_embeds_auto_connect_last_server').on('change', (event) => {
        const settings = ensureSettings();
        settings.autoConnectLastServer = !!event.target.checked;
        saveSettingsDebounced();

        if (settings.autoConnectLastServer && settings.apiProvider) {
            void restartProviderConnection();
        } else if (!settings.autoConnectLastServer) {
            renderConnectionStatus('disconnected', 'Disconnected');
        }
    });

    // Abort Restart Button
    $('#image_embeds_abort_test').on('click', () => {
        if (restartConnectionAbortController) {
            restartConnectionAbortController.abort();
        }
    });

    $('#image_embeds_user_enabled').on('change', (event) => {
        ensureSettings().showUserMode = !!event.target.checked;
        saveSettingsDebounced();

        // Show/hide user mode button
        if (event.target.checked) {
            $('#image_embeds_mode_user').css('display', 'inline-block');
        } else {
            $('#image_embeds_mode_user').css('display', 'none');
            // Switch back to character mode if user disables mode switching
            if (currentMode === 'user') {
                currentMode = 'character';
                $('#image_embeds_mode_char').css('font-weight', 'bold');
                $('#image_embeds_mode_user').css('font-weight', 'normal');
                $('#image_embeds_mode_label').text('Character Expressions');
            }
        }
        renderList();
    });

    $('#image_embeds_regenerate_current').on('click', async () => {
        const button = $('#image_embeds_regenerate_current');
        button.prop('disabled', true);
        try {
            await regenerateCurrentCharacterExpressions();
        } finally {
            button.prop('disabled', false);
        }
    });

    $('#image_embeds_expression_docs').on('click', () => {
        const popup = $('#image_embeds_expression_docs_popup');
        if (popup.hasClass('is-open')) {
            closeExpressionDocumentation();
        } else {
            openExpressionDocumentation();
        }
    });

    $('#image_embeds_expression_docs_close').on('click', () => {
        closeExpressionDocumentation();
    });

    $('#image_embeds_expression_docs_popup').on('click', (event) => {
        if (event.target === event.currentTarget) {
            closeExpressionDocumentation();
        }
    });

    $('#image_embeds_mode_char').on('click', () => {
        if (!getCharacterKey()) return;
        currentMode = 'character';
        $('#image_embeds_mode_char').css('font-weight', 'bold');
        $('#image_embeds_mode_user').css('font-weight', 'normal');
        $('#image_embeds_mode_label').text('Character Expressions');
        renderList();
    });

    $('#image_embeds_mode_user').on('click', () => {
        const settings = ensureSettings();
        if (!getCharacterKey() || !settings.showUserMode) return;
        currentMode = 'user';
        $('#image_embeds_mode_char').css('font-weight', 'normal');
        $('#image_embeds_mode_user').css('font-weight', 'bold');
        $('#image_embeds_mode_label').text('User Expressions');
        renderList();
    });
}

async function injectSettingsUi() {
    if ($('#image_embeds_expressions_container').length) return;
    const settingsHtml = $(await renderExtensionTemplateAsync(EXTENSION_ID, 'settings'));
    const container = $('<div class="extension_container" id="image_embeds_expressions_container"></div>');
    container.append(settingsHtml);
    $('#extensions_settings2').append(container);

    // Initialize toggle state
    const settings = ensureSettings();
    $('#image_embeds_enabled').prop('checked', !!settings.enabled);
    $('#image_embeds_advanced_enabled').prop('checked', !!settings.advancedExpressionsEnabled);
    $('#image_embeds_double_enabled').prop('checked', !!settings.doubleEnabled);
    $('#image_embeds_user_enabled').prop('checked', !!settings.showUserMode);
    $('#image_embeds_auto_connect_last_server').prop('checked', !!settings.autoConnectLastServer);
    renderConnectionStatus('disconnected', settings.apiProvider ? 'Disconnected' : 'Select a provider first');

    // Initialize API configuration
    populateAPIProviders();

    // Inject Abort Restart button next to Restart Connection if not already present
    if (!$('#image_embeds_abort_test').length) {
        const abortBtn = $('<button type="button" id="image_embeds_abort_test" class="menu_button" title="Abort the connection restart" style="display:none; margin-left:6px;"><i class="fa-solid fa-ban"></i> Abort Restart</button>');
        $('#image_embeds_test_connection').after(abortBtn);
    }

    // Load saved API settings
    $('#image_embeds_custom_baseurl').val(settings.customBaseUrl || '');
    $('#image_embeds_api_key').val(settings.apiKey || '');
    $('#image_embeds_api_model_custom').val(settings.apiModel || '');

    // Show/hide advanced settings section
    if (settings.advancedExpressionsEnabled) {
        $('#image_embeds_advanced_settings').show();
        if (settings.apiProvider) {
            await updateAPIProviderUI(settings.apiProvider);
            if (settings.autoConnectLastServer) {
                void restartProviderConnection();
            }
        }
    } else {
        $('#image_embeds_advanced_settings').hide();
    }

    // Hide user mode buttons if disabled
    if (!settings.showUserMode) {
        $('#image_embeds_mode_user').css('display', 'none');
        currentMode = 'character';
    }

    renderExpressionDocumentation();
}

function bindEvents() {
    eventSource.makeLast(event_types.USER_MESSAGE_RENDERED, (messageId) => scheduleMessageRender(messageId));
    eventSource.makeLast(event_types.CHARACTER_MESSAGE_RENDERED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MESSAGE_UPDATED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MESSAGE_SWIPED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MORE_MESSAGES_LOADED, () => refreshAllMessages());
    eventSource.on(event_types.CHAT_CHANGED, () => {
        lastAssistantMessageId = null;
        llmEligibleMessageIds.clear();
        renderList();
        refreshAllMessages();
    });
    eventSource.on(event_types.EXTENSIONS_FIRST_LOAD, () => refreshAllMessages());
    const renderActiveAssistantMessage = () => {
        const targetId = getActiveAssistantMessageId();
        if (targetId !== null && targetId !== undefined) {
            markMessageEligibleForLLM(targetId);
            scheduleMessageRender(targetId);
        }
    };
    eventSource.on(event_types.GENERATION_STOPPED, renderActiveAssistantMessage);
    eventSource.on(event_types.GENERATION_ENDED, renderActiveAssistantMessage);
    eventSource.on(event_types.CHARACTER_RENAMED, (oldAvatar, newAvatar) => {
        const settings = ensureSettings();
        if (settings.characters?.[oldAvatar]) {
            settings.characters[newAvatar] = settings.characters[oldAvatar];
            delete settings.characters[oldAvatar];
            saveSettingsDebounced();
        }
    });
    eventSource.on(event_types.CHARACTER_DELETED, (data) => {
        const avatar = data?.character?.avatar;
        if (!avatar) return;
        const settings = ensureSettings();
        if (settings.characters?.[avatar]) {
            delete settings.characters[avatar];
            saveSettingsDebounced();
        }
    });
}

jQuery(async function () {
    ensureSettings();
    await injectSettingsUi();
    bindUi();
    bindEvents();
    renderList();
    refreshAllMessages();
});
