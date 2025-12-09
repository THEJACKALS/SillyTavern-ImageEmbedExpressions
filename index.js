import { characters, chat, eventSource, event_types, saveSettingsDebounced, this_chid } from '/script.js';
import { extension_settings, renderExtensionTemplateAsync } from '/scripts/extensions.js';
import { getBase64Async, getFileExtension, getStringHash, saveBase64AsFile } from '/scripts/utils.js';

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
const defaultSettings = { characters: {}, enabled: true };
const DEFAULT_CHARACTER_GROUP = '__default__';
let lastAssistantMessageId = null;

function escapeRegExp(value) {
    return String(value ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

function normalizeName(name) {
    return String(name ?? '')
        .trim()
        .toLowerCase()
        .replace(/[\\\/\s]+/g, '_');
}

function findEntryByName(name) {
    const target = normalizeName(name);
    return getCharacterEntries().find(entry => normalizeName(entry.name) === target);
}

function parseEntryName(name) {
    const raw = String(name ?? '').trim();
    const splitMatch = raw.match(/^(?<character>[^\/\\|\-_]+)[\/\\|\-_]+(?<expression>.+)$/);
    const character = splitMatch?.groups?.character ? normalizeName(splitMatch.groups.character) : '';
    const expression = splitMatch?.groups?.expression || raw;

    return {
        raw,
        character,
        expression,
        normalized: normalizeName(raw),
    };
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
    const lowerText = String(text || '').toLowerCase();
    const cleaned = lowerText.replace(/[^\w\s]/g, ' ');
    const presenceVerbs = ['said', 'says', 'ask', 'asked', 'asks', 'reply', 'replied', 'replies', 'respond', 'responded', 'responds', 'yell', 'yelled', 'yells', 'shout', 'shouted', 'shouts', 'whisper', 'whispered', 'whispers', 'mutter', 'muttered', 'mutters', 'laughed', 'laughs', 'laughing', 'smiled', 'smiles', 'smiling', 'nodded', 'nods', 'grinned', 'grins', 'grinning', 'looked', 'looks', 'looking', 'turned', 'turns', 'walking', 'walked', 'walks', 'stood', 'stands', 'standing', 'sat', 'sits', 'sitting'];
    const imaginationHints = ['memory of', 'remembering', 'image of', 'imagination of', 'imagining', 'fantasy of', 'thinking of', 'thought of', 'dream of', 'dreaming of', 'idea of', 'vision of'];
    let best = null;

    for (const character of characters) {
        const plainName = character.replace(/_/g, ' ').trim();
        if (!plainName) continue;

        const wordPattern = new RegExp(`\\b${escapeRegExp(plainName)}\\b`, 'g');
        const speakingPattern = new RegExp(`(^|\\n)\\s*${escapeRegExp(plainName)}\\s*[:\\-\\u2013\\u2014]`, 'g');
        let score = 0;
        let firstIndex = Infinity;
        let match;

        while ((match = wordPattern.exec(cleaned)) !== null) {
            score += 1;
            if (match.index < firstIndex) {
                firstIndex = match.index;
            }

            const window = lowerText.slice(Math.max(0, match.index - 24), match.index + plainName.length + 24);
            if (presenceVerbs.some(v => new RegExp(`\\b${escapeRegExp(v)}\\b`).test(window))) {
                score += 1;
            }
            if (imaginationHints.some(h => window.includes(h))) {
                score -= 1;
            }
        }

        while ((match = speakingPattern.exec(lowerText)) !== null) {
            score += 2;
            if (match.index < firstIndex) {
                firstIndex = match.index;
            }
        }

        if (firstIndex === 0) {
            score += 1;
        }

        score = Math.max(score, 0);

        if (score > 0 && (!best || score > best.score || (score === best.score && firstIndex < best.firstIndex))) {
            best = { character, score, firstIndex };
        }
    }

    return best?.character || null;
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

        if (needles.some(needle => cleaned.includes(needle))) {
            return entry;
        }
    }

    return entries[0].entry;
}

function findEntryMatchInText(entries, messageText) {
    const cleaned = String(messageText || '').toLowerCase().replace(/[^\w\s]/g, ' ');

    for (const entry of entries) {
        const parsed = parseEntryName(entry.name);
        const needles = [
            normalizeName(parsed.raw).replace(/_/g, ' '),
            normalizeName(parsed.expression).replace(/_/g, ' '),
            parsed.character?.replace(/_/g, ' '),
        ].filter(Boolean);

        if (needles.some(needle => cleaned.includes(needle))) {
            return entry;
        }
    }

    return null;
}

function buildPlaceholder(name) {
    return `{{img::${name}}}`;
}

function createId() {
    return crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
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

function pickEntryForMessage(messageId) {
    const message = chat?.[messageId];
    if (!message || message.is_user || message.is_system) return null;

    const entries = getCharacterEntries();
    if (!entries.length) return null;

    const messageText = String(message.mes || '').toLowerCase();
    const grouped = groupEntriesByCharacter(entries);
    const characterKeys = Array.from(grouped.keys()).filter(key => key && key !== DEFAULT_CHARACTER_GROUP);
    const detectedCharacter = characterKeys.length ? detectCharacterFromText(messageText, characterKeys) : null;
    const hasMultipleCharacters = characterKeys.length > 1;

    if (detectedCharacter) {
        const entry = selectEntryForCharacter(grouped, detectedCharacter, messageText);
        if (entry) return entry;
    }

    if (hasMultipleCharacters) {
        return null;
    }

    const directMatch = findEntryMatchInText(entries, messageText);
    if (directMatch) return directMatch;

    return selectEntryForCharacter(grouped, characterKeys[0] || DEFAULT_CHARACTER_GROUP, messageText) || null;
}

function insertPlaceholderNearMatch(root, entry) {
    const token = buildPlaceholder(entry.name);
    const parsed = parseEntryName(entry.name);
    const needles = [
        normalizeName(entry.name).replace(/_/g, ' '),
        parsed.character?.replace(/_/g, ' '),
        normalizeName(parsed.expression).replace(/_/g, ' '),
    ].filter(Boolean);
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

function autoInjectAfterGeneration(root, messageId) {
    const settings = ensureSettings();
    if (!settings.enabled) return;
    if (!root) return;
    const message = chat?.[messageId];
    if (message?.is_user || message?.is_system) return;
    const textContent = (root.textContent || '').replace(/\u200b/g, '').trim();
    if (!textContent || textContent === '...' || textContent === '…') return;

    // Skip if an embed is already present from previous processing.
    if (root.querySelector('.image-embed-expression')) {
        return;
    }

    PLACEHOLDER_REGEX.lastIndex = 0;
    const hasPlaceholder = PLACEHOLDER_REGEX.test(root.textContent || '');
    if (hasPlaceholder) return;

    const entry = pickEntryForMessage(messageId);
    if (!entry) return;

    insertPlaceholderNearMatch(root, entry);
    renderPlaceholders(root);
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
    const entries = getCharacterEntries();

    if (!list.length) return;

    list.empty();

    if (!getCharacterKey()) {
        list.append($('<div class="image-embeds-empty">Open a character chat to manage Image Embeds.</div>'));
        return;
    }

    if (!entries.length) {
        list.append($('<div class="There are no expressions for this character yet. Click + to add one.</div>'));
        return;
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
            saveSettingsDebounced();
            refreshAllMessages();
        });

        remove.on('click', () => {
            const key = getCharacterKey();
            if (!key) return;
            extension_settings[SETTINGS_KEY].characters[key].entries = getCharacterEntries().filter(x => x.id !== entry.id);
            saveSettingsDebounced();
            renderList();
            refreshAllMessages();
        });

        row.append(preview, input, placeholder, remove);
        list.append(row);
    }
}

function ensureUniqueName(baseName) {
    const entries = getCharacterEntries();
    const normBase = normalizeName(baseName) || 'expression';
    let candidate = normBase;
    let counter = 1;

    while (entries.some(e => normalizeName(e.name) === candidate)) {
        candidate = `${normBase}_${counter}`;
        counter += 1;
    }

    return candidate;
}

async function addExpressionFromFile(file) {
    if (!file) return;
    const charKey = getCharacterKey();
    if (!charKey) {
        toastr.warning('Open a character chat to manage Image Embeds.', 'Image Embeds');
        return;
    }

    if (!file.type?.startsWith('image/')) {
        toastr.error('File must be an image.', 'Image Embeds');
        return;
    }

    try {
        const base64 = await getBase64Async(file);
        const base64Data = base64.split(',')[1];
        const extension = getFileExtension(file) || file.type.split('/')[1] || 'png';
        const slug = getStringHash(file.name || base64Data);
        const fileName = `${Date.now()}_${slug}`;
        const url = await saveBase64AsFile(base64Data, getCharacterFolder(), fileName, extension);

        if (!url) {
            toastr.error('Failed to save image.', 'Image Embeds');
            return;
        }

        const defaultName = ensureUniqueName((file.name || 'expression').replace(/\.[^.]+$/, ''));
        const settings = ensureSettings();
        settings.characters[charKey].entries.push({
            id: createId(),
            name: defaultName,
            url,
            originalName: file.name || '',
        });

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

function refreshAllMessages() {
    document.querySelectorAll('#chat .mes').forEach(mes => {
        const messageId = Number(mes.getAttribute('mesid'));
        const textNode = mes.querySelector('.mes_text');
        if (!textNode || Number.isNaN(messageId)) return;
        renderPlaceholders(textNode);
        autoInjectAfterGeneration(textNode, messageId);
    });
}

function onMessageRendered(messageId) {
    const root = document.querySelector(`#chat .mes[mesid="${messageId}"] .mes_text`);
    if (root) {
        renderPlaceholders(root);
        autoInjectAfterGeneration(root, messageId);
    }
}

function scheduleMessageRender(messageId) {
    const targetId = rememberAssistantMessage(messageId);
    if (targetId === null || targetId === undefined || Number.isNaN(targetId)) return;
    requestAnimationFrame(() => onMessageRendered(targetId));
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
}

function bindEvents() {
    eventSource.makeLast(event_types.USER_MESSAGE_RENDERED, (messageId) => scheduleMessageRender(messageId));
    eventSource.makeLast(event_types.CHARACTER_MESSAGE_RENDERED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MESSAGE_UPDATED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MESSAGE_SWIPED, (messageId) => scheduleMessageRender(messageId));
    eventSource.on(event_types.MORE_MESSAGES_LOADED, () => refreshAllMessages());
    eventSource.on(event_types.CHAT_CHANGED, () => {
        lastAssistantMessageId = null;
        renderList();
        refreshAllMessages();
    });
    eventSource.on(event_types.EXTENSIONS_FIRST_LOAD, () => refreshAllMessages());
    const renderActiveAssistantMessage = () => {
        const targetId = getActiveAssistantMessageId();
        if (targetId !== null && targetId !== undefined) {
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
