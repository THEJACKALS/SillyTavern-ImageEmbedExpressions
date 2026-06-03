const ACTIVE_ACTIONS = [
    'said', 'says', 'asked', 'asks', 'replied', 'replies', 'responded', 'responds',
    'whispered', 'whispers', 'muttered', 'mutters', 'shouted', 'shouts', 'yelled', 'yells',
    'laughed', 'laughs', 'chuckled', 'chuckles', 'smiled', 'smiles', 'grinned', 'grins',
    'gasped', 'gasps', 'sighed', 'sighs', 'squeaked', 'squeaks', 'snapped', 'snaps', 'growled', 'growls',
    'looked', 'looks', 'glared', 'glares', 'blinked', 'blinks', 'nodded', 'nods',
    'trembled', 'trembles', 'protested', 'protests', 'noticed', 'notices',
    'approached', 'approaches', 'stopped', 'stops', 'leaned', 'leans', 'wrapped', 'wraps',
    'reached', 'reaches', 'trailed', 'trails', 'cupped', 'cups', 'traced', 'traces',
    'shifted', 'shifts', 'softened', 'softens', 'remained', 'remains',
    'continued', 'continues', 'worked', 'works', 'pulled', 'pulls', 'licked', 'licks',
    'moaned', 'moans', 'hummed', 'hums', 'flicked', 'flicks', 'bobbed', 'bobs',
    'commanded', 'commands', 'glanced', 'glances', 'complied', 'complies',
    'pressed', 'presses', 'serviced', 'services', 'stroked', 'strokes', 'angled', 'angles',
    'tightened', 'tightens', 'purred', 'purrs',
    'helped', 'helps', 'helping', 'filled', 'fills', 'touched', 'touches', 'touching',
    'patted', 'patting', 'pat', 'pats', 'pats',
    'hitched', 'hitches', 'shivered', 'shivers', 'giggled', 'giggles', 'smirked', 'smirks',
    'melted', 'melts', 'rested', 'rests', 'traced', 'traces', 'dried', 'dries',
    'watched', 'watches', 'climbed', 'climbs', 'ground', 'grinds', 'grabbed', 'grabs',
    'walk', 'come', 'comes', 'came', 'arrive', 'arrives', 'arrived',
    'blushed', 'blushes', 'cried', 'cries', 'stared', 'stares',
];

const ACTIVE_NOUNS = [
    'expression', 'face', 'eyes', 'gaze', 'smile', 'voice', 'tone', 'hands', 'hand',
    'arms', 'shoulders', 'body', 'breathing', 'ears', 'tail', 'posture',
    'mouth', 'head', 'lips', 'tongue', 'grip', 'pace', 'brain', 'mind', 'thoughts',
    'breath', 'chest', 'skin', 'fingers', 'forehead', 'jaw', 'cheeks',
];

function normalizeName(name) {
    return String(name ?? '').trim().toLowerCase().replace(/[\\/\s]+/g, '_');
}

function compactName(name) {
    return normalizeName(name).replace(/[-|_\s]+/g, '');
}

function stripHonorifics(name) {
    return normalizeName(name).replace(/(?:^|[_-])(chan|san|sama|kun|senpai|sensei)$/i, '');
}

function getCharacterAliases(name) {
    const normalized = stripHonorifics(name);
    const parts = normalized.split(/[_-]+/).filter(Boolean);
    return [...new Set([
        normalized,
        parts.at(-1) || '',
    ].filter(Boolean))];
}

function keysMatch(a, b) {
    const normalizedA = normalizeName(a);
    const normalizedB = stripHonorifics(b);
    if (!normalizedA || !normalizedB) return false;

    return getCharacterAliases(normalizedA).some(alias => (
        alias === normalizedB
        || compactName(alias) === compactName(normalizedB)
    ));
}

function escapeRegExp(value) {
    return String(value ?? '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findKnownCharacter(candidate, characterKeys) {
    return (characterKeys || []).find(key => keysMatch(key, candidate)) || '';
}

function stripLeadingDecorators(text) {
    return String(text || '').replace(/^[\s>*_"'`~()[\]{}]+/, '').trim();
}

function getNarrativeText(segment) {
    return String(segment || '')
        .replace(/"[^"]*"/g, ' ')
        .replace(/“[^”]*”/g, ' ')
        .replace(/'[^']*'/g, ' ');
}

function addActiveCharacter(active, seen, character) {
    if (!character || seen.has(character)) return;
    seen.add(character);
    active.push(character);
}

function detectActiveCharactersFromSegment(segment, characterKeys) {
    const text = stripLeadingDecorators(segment);
    const narrativeText = getNarrativeText(text);
    const searchText = `${narrativeText} ${text}`;
    const active = [];
    const seen = new Set();
    if (!text) return active;

    const labelMatch = text.match(/^([A-Z][A-Za-z0-9' -]{0,48}?)(?=\s*(?::|[–—]|\s-\s))/);
    const labelCharacter = labelMatch?.[1] ? findKnownCharacter(labelMatch[1], characterKeys) : '';
    addActiveCharacter(active, seen, labelCharacter);

    const actionPattern = ACTIVE_ACTIONS.map(escapeRegExp).join('|');
    const actionMatch = narrativeText.match(new RegExp(`^([A-Z][A-Za-z0-9' -]{0,48}?)\\s+(${actionPattern})\\b`));
    const actionCharacter = actionMatch?.[1] ? findKnownCharacter(actionMatch[1], characterKeys) : '';
    addActiveCharacter(active, seen, actionCharacter);

    const possessivePattern = ACTIVE_NOUNS.map(escapeRegExp).join('|');
    const possessiveMatch = narrativeText.match(new RegExp(`^([A-Z][A-Za-z0-9' -]{0,48}?)(?:'s|’s)\\s+(${possessivePattern})\\b`));
    const possessiveCharacter = possessiveMatch?.[1] ? findKnownCharacter(possessiveMatch[1], characterKeys) : '';
    addActiveCharacter(active, seen, possessiveCharacter);

    for (const characterKey of characterKeys || []) {
        for (const plainName of getCharacterAliases(characterKey).map(alias => alias.replace(/_/g, ' '))) {
            if (!plainName) continue;

            const namePattern = escapeRegExp(plainName);
            const actionAnywhere = new RegExp(`\\b${namePattern}(?:'s|’s)?(?:[-\\s]+(?:chan|san|sama|kun|senpai|sensei))?\\b[^.!?\\n]{0,80}\\b(${actionPattern})\\b`, 'i');
            const possessiveAnywhere = new RegExp(`\\b${namePattern}(?:'s|’s)\\s+(${possessivePattern})\\b`, 'i');
            const actionBeforeName = new RegExp(`\\b(${actionPattern})\\b[^.!?\\n]{0,40}\\b${namePattern}\\b`, 'i');

            if (
                actionAnywhere.test(searchText)
                || possessiveAnywhere.test(searchText)
                || actionBeforeName.test(searchText)
            ) {
                addActiveCharacter(active, seen, characterKey);
                break;
            }
        }
    }

    return active;
}

export function detectActiveCharacterKeys(text, characterKeys = []) {
    if (!Array.isArray(characterKeys) || characterKeys.length === 0) return [];

    const active = [];
    const seen = new Set();
    const segments = String(text || '')
        .split(/(?:\r?\n)+|(?<=\*)\s+(?=\*)/)
        .map(segment => segment.trim())
        .filter(Boolean);

    for (const segment of segments) {
        const characters = detectActiveCharactersFromSegment(segment, characterKeys);
        for (const character of characters) {
            addActiveCharacter(active, seen, character);
        }
    }

    return active;
}
