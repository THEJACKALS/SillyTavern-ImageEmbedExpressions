function normalizeExpressionName(name) {
    return String(name ?? '')
        .trim()
        .toLowerCase()
        .replace(/[\\\/\s]+/g, '_')
        .replace(/_/g, ' ')
        .trim();
}

export const EXPRESSION_HINTS = {
    admiration: [
        'admire', 'admired', 'admiration', 'impressed', 'impressive', 'amazed by you',
        'looked up to', 'respect', 'respected', 'in awe', 'awe', 'wonder in her eyes',
        'wonder in his eyes', 'starry eyed', 'starry-eyed', 'perfect', 'dangerously perfect',
        'heart-meltingly perfect', 'being considerate', 'genuine', 'present',
    ],
    amusement: [
        'amused', 'amusement', 'chuckled', 'chuckling', 'laughed softly', 'laughing softly',
        'snickered', 'snickering', 'giggle', 'giggled', 'giggling', 'playful smile',
        'teasing smile', 'smirked', 'smirking', 'grinned', 'threw her head back laughing',
        'threw his head back laughing', 'huhu', 'huhu~', 'complained playfully',
        'mischievous', 'mischievously', 'grinned mischievously', 'greedy i like it',
    ],
    anger: [
        'angry', 'anger', 'furious', 'rage', 'glared', 'glaring', 'snapped at', 'snarled',
        'growled', 'hissed', 'clenched fist', 'clenched fists', 'clenched jaw',
        'jaw tightened', 'scowled', 'scowl', 'seething',
    ],
    annoyance: [
        'annoyed', 'annoyance', 'irritated', 'irritation', 'exasperated', 'rolled her eyes',
        'rolled his eyes', 'sighed sharply', 'tch', 'clicked her tongue', 'clicked his tongue',
        'flat look', 'deadpan stare', 'brow twitched',
    ],
    approval: [
        'approved', 'approval', 'approving', 'nodded approvingly', 'nodded', 'good job',
        'well done', 'proud smile', 'satisfied smile', 'gave a thumbs up', 'thumbs up',
        'pleased', 'pleased smile',
    ],
    caring: [
        'caring', 'concerned', 'concern', 'worried for you', 'are you okay', 'gentle concern',
        'softened her voice', 'softened his voice', 'comforting', 'comforted', 'soothing',
        'tended to', 'checked on you', 'protective', 'care about', 'do you care about',
        'hurt her feelings', 'hurt his feelings', 'hurt mixed with concern',
    ],
    confusion: [
        'confused', 'confusion', 'puzzled', 'bewildered', 'tilted her head', 'tilted his head',
        'furrowed her brow', 'furrowed his brow', 'blinked in confusion', 'what do you mean',
        'did not understand', 'does not understand', 'lost', 'processed this information carefully',
        'are you serious', 'do you think', 'what do you think', 'genuine uncertainty',
        'expression showed genuine uncertainty', 'unless you are saying', "unless you're saying",
        'when did this happen', 'gaze flicked between them', 'flicked between them',
        'confusion flickering', 'something else flickering',
    ],
    curiosity: [
        'curious', 'curiosity', 'intrigued', 'interested', 'leaned in', 'leans in',
        'raised an eyebrow', 'arched an eyebrow', 'tell me more', 'questioning look',
        'inquisitive', 'studied you', 'ara', 'ara~',
    ],
    desire: [
        'desire', 'desired', 'want', 'wanted', 'longing', 'yearning', 'hungry gaze',
        'looked at your lips', 'half lidded eyes', 'half-lidded eyes', 'leaned closer',
        'breath hitched', 'bit her lip', 'bit his lip', 'licked her lips', 'licked his lips',
        'kiss', 'kissed', 'melted into the kiss', 'against his lips', 'against her lips',
        'pressed closer', 'traced patterns', 'hungry eyes', 'appealing', 'tease me',
        'tease her', 'tease him', 'ground against', 'straddling', 'straddled', 'moan'
    ],
    arousal: [
        'arousal', 'aroused', 'heated gaze', 'heat pooled', 'breath hitched',
        'breathing grew heavier', 'heavy breathing', 'pulse quickened', 'skin flushed',
        'body reacted', 'shivered with need', 'desire stirred', 'cum', 'orgasm', 'climax',
        'climax rippling', 'convulsions subsided', 'body tensed', 'screamed',
        'lapped up every drop', 'needs attention now', 'my turn',
        'body responding instantly', 'body responded instantly', 'hardness against her',
        'hardness against him', 'joining us', 'bare breasts', 'pressed against his bare chest',
        'pressed against her bare chest', 'ground against him', 'ground against her',
        'cried out', 'gasped', 'walls clenched', 'clenched around him', 'clenched around her',
        'fill me', 'filled me', 'fill me so completely', 'thrust deeper', 'harder please',
        'living room echoed with their moans', 'couch creaked',
    ],
    disappointment: [
        'disappointed', 'disappointment', 'let down', 'deflated', 'shoulders slumped',
        'sank', 'sighed sadly', 'lowered her gaze', 'lowered his gaze', 'not what i hoped',
        'expected better', 'almost hurt', 'voice was soft almost hurt', 'did not think to mention',
        "didn't think to mention", 'mention this earlier',
    ],
    disapproval: [
        'disapproved', 'disapproval', 'disapproving', 'stern look', 'frowned', 'frowning',
        'shook her head', 'shook his head', 'not okay', 'unacceptable', 'tsk', 'reproachful',
        'judging stare',
    ],
    disgust: [
        'disgust', 'disgusted', 'gross', 'revolted', 'repulsed', 'nauseated', 'wrinkled her nose',
        'wrinkled his nose', 'grimaced', 'grimacing', 'sickened', 'ew', 'ugh',
    ],
    embarrassment: [
        'blush', 'blushed', 'blushing', 'flush', 'flushed', 'flushing', 'face red', 'red face',
        'bright red', 'turned bright red', 'turning bright red', 'cheeks red', 'red cheeks',
        'went red', 'turns red', 'turned red', 'ears burning',
        'looked away', 'looks away', 'averted', 'avoids eye contact', 'avoided eye contact',
        'not quite meeting your eyes', 'could not meet your eyes', 'mortified', 'embarrassed',
        'embarrassment', 'sheepish', 'shy smile', 'cheeks burning', 'face burning', 'immediately regretted',
        'wanted to disappear', 'turned away to hide', 'face burned so hot',
        'might actually combust', 'wide and mortified', 'scandalized',
        'mortifying beyond belief', 'die from embarrassment',
        'actually die from embarrassment', 'please tell me you did not',
        "please tell me you didn't", 'you looked through my', 'privacy for a reason',
        'brain officially went offline', 'brain went offline', 'mind short-circuited',
        'short-circuited', 'could not form words', 'could not breathe properly',
        'gasp and a squeak', 'tiny involuntary sound', 'felt her face heat up',
        'felt his face heat up', 'wide eyes and embarrassment', 'shocked and embarrassed',
        'embarrassingly breathless', 'brain was too fried', 'face heat up to dangerous temperatures',
        'state of undress', 'towels barely covering', 'sitting intimately close',
        'take a shower together', 'took a shower together',
        'blushed deeply', 'cheeks flushed', 'stammered', 'r ryo', 'cut herself off',
        'cut himself off', 'crossed lines', 'sputtered', 'taking a step back',
        'took a step back', 'n-no comparison', 'oversized sweater',
    ],
    excitement: [
        'excited', 'excitement', 'thrilled', 'eager', 'bounced', 'bouncing', 'eyes sparkling',
        'lit up', 'brightened', 'can not wait', "can't wait", 'energetic', 'enthusiastic',
        'thrilling', 'hurry up', 'want to see what happens',
    ],
    fear: [
        'afraid', 'fear', 'fearful', 'scared', 'terrified', 'frightened', 'pale',
        'wide eyed', 'wide-eyed', 'froze', 'frozen', 'trembled', 'backed away',
        'panic', 'panicked', 'horrified', 'turned pale', 'turn pale', 'turnpale',
    ],
    gratitude: [
        'grateful', 'gratitude', 'thank you', 'thanks', 'thankful', 'appreciate it',
        'appreciated', 'relieved smile', 'soft thankful smile', 'bowed her head',
        'bowed his head', 'thank you for this', 'for trusting me',
    ],
    grief: [
        'grief', 'grieving', 'mourning', 'mournful', 'sobbed', 'sobbing', 'wept',
        'heartbroken', 'loss', 'bereft', 'voice broke', 'voice cracked', 'tears fell',
    ],
    joy: [
        'happy', 'happiness', 'joy', 'joyful', 'delighted', 'beamed', 'beaming',
        'grinned', 'grinning', 'laughed', 'laughing', 'cheerful', 'radiant smile',
        'smiled brightly', 'squealed', 'squeaked with surprise',
    ],
    love: [
        'love', 'loving', 'affection', 'affectionate', 'soft smile', 'warm smile',
        'smiled warmly', 'smiles warmly', 'tenderly', 'fondly', 'loving gaze',
        'gentle touch', 'caressed', 'hugged', 'embraced', 'warmth in her eyes',
        'warmth in his eyes', 'warm and fuzzy', 'heart fluttered', 'heart-melting',
        'heart melting', 'fall for', 'falling for', 'fell for', 'gotten under her skin',
        'under her skin', 'butterflies', 'stomach did gymnastics',
        'stomach do complicated gymnastics', 'dangerously perfect', 'casual intimacy',
        'patted her head', 'patted his head', 'loved it more than was probably healthy',
        'overwhelming feelings', 'casual touch', 'melted into the kiss',
        'rested her forehead against his', 'rested his forehead against her',
        'other people\'s hearts', 'other peoples hearts',
    ],
    anxious: [
        'anxious', 'anxiety', 'anxiously', 'uneasy', 'uneasiness', 'worried',
        'worry', 'worrying', 'tense', 'tension', 'apprehensive', 'apprehension',
        'on edge', 'restless', 'fidgeted', 'fidgeting', 'hesitated', 'hesitating',
        'swallowed hard', 'rapid breathing', 'breathing was rapid',
        'higher pitched than usual', 'voice was higher pitched',
        'mind was racing', 'inner panic', 'heart did something alarming',
        'hands clenched anxiously', 'fingers twisting', 'trying not to be trouble',
        'too much trouble', 'afraid of being a burden', 'breathing rapidly',
        'clearly distressed', 'what did you think', 'should i throw it away',
        'am i a terrible person',
    ],
    nervousness: [
        'nervous', 'nervousness', 'anxious', 'uneasy', 'tremor in her voice',
        'tremor in his voice', 'voice trembled', 'voice shaking', 'shaking', 'shaky',
        'fidgeted', 'fidgeting', 'swallowed hard', 'sweaty palms', 'hesitated',
        'hesitating', 'uncertain', 'worried', 'rapid breathing', 'breathing was rapid',
        'higher pitched than usual', 'voice was higher pitched', 'shaky breath',
        'could not breathe properly', 'mind was racing', 'nervous energy',
    ],
    jealous: [
        'jealous', 'jealousy', 'envy', 'envious', 'looked jealous', 'possessive glare',
        'who was that', 'why were you with', 'attention on someone else', 'saw you with',
        'bitter smile', 'forced smile', 'green-eyed', 'someone else needs attention',
        'her gaze lingered', 'his gaze lingered', 'where she sat on', 'where he sat on',
        'hard seeing you two together', 'seeing you two together', 'force you to choose',
        'would not be fair', "wouldn't be fair", 'hope you want me too', 'want me too',
    ],
    neutral: [
        'neutral', 'calm', 'blank expression', 'blank face', 'expressionless', 'deadpan',
        'flat tone', 'even tone', 'matter of fact', 'matter-of-fact', 'composed',
        'unreadable', 'stoic', 'waited listening', 'listening to footsteps',
    ],
    optimism: [
        'optimistic', 'optimism', 'hopeful', 'hope', 'confident smile', 'encouraging smile',
        'it will be okay', 'we can do this', 'bright side', 'positive', 'reassuring',
    ],
    pride: [
        'proud', 'pride', 'puffed her chest', 'puffed his chest',
        'chin lifted', 'boasted',
        'boasting', 'satisfied', 'absolutely perfect', 'make it count',
    ],
    smug: [
        'smug', 'smugly', 'smug smile', 'smug grin', 'smirked knowingly',
        'knowing smirk', 'self satisfied', 'self-satisfied', 'satisfied grin',
        'satisfied smirk', 'gave a smug look', 'looked smug',
        'pleased with herself', 'pleased with himself',
    ],
    realization: [
        'realized', 'realization', 'dawned on her', 'dawned on him', 'it clicked',
        'understood', 'understanding dawned', 'eyes widened in realization',
        'suddenly understood', 'oh', 'aha',
    ],
    relief: [
        'relieved', 'relief', 'sighed in relief', 'let out a breath', 'breathed out',
        'thank goodness', 'shoulders relaxed', 'tension left', 'safe now',
        'catch her breath', 'catch his breath',
    ],
    remorse: [
        'remorse', 'remorseful', 'sorry', 'apologized', 'apologetic', 'guilt', 'guilty',
        'regret', 'regretful', 'ashamed', 'looked guilty', 'lowered her head',
        'lowered his head',
    ],
    sadness: [
        'sad', 'sadness', 'sorrow', 'tearful', 'tears', 'crying', 'cried', 'downcast',
        'looked down', 'melancholy', 'lonely', 'hurt', 'pained smile', 'almost hurt',
        'hurt mixed with concern', 'regret it', 'unless you regret it',
    ],
    surprise: [
        'surprised', 'surprise', 'startled', 'blinked', 'eyes widened', 'wide eyes',
        'gasped', 'taken aback', 'caught by surprise', 'shocked', 'stunned',
        'wide eyes', 'eyes widened to', 'comical degree', 'freeze', 'froze up',
        'frozen up', 'brain officially went offline', 'short-circuited', 'could not do anything except stare',
        'jaw dropped', 'two whole bottles', 'no wonder you do not remember',
        "no wonder you don't remember", 'appeared in the doorway', 'what the',
        'froze in the doorway', 'purple eyes widening',
    ],
    wink: [
        'wink', 'winked', 'winking', 'gave a wink', 'playful wink',
        'teasing wink', 'mischievous wink', 'one eye closed',
        'closed one eye', 'winked teasingly', 'winked playfully',
    ],
    agitation: [
        'agitation', 'agitated', 'restless', 'paced', 'pacing', 'irritated',
        'tail lashed', 'tail lashing', 'lashed her tail', 'lashed his tail',
        'tapped her foot', 'tapped his foot', 'drummed her fingers', 'drummed his fingers',
        'bristled', 'bristling',
    ],
    dominant: [
        'dominant', 'commanding', 'commanded', 'ordered', 'authoritative', 'authority',
        'took control', 'in control', 'firm voice', 'stern command', 'held your chin',
        'pinned you', 'towered over', 'uncompromising stare',
    ],
    awkward: [
        'awkward', 'awkwardly', 'awkward silence', 'awkward pause',
        'uncomfortable silence', 'uncomfortable pause', 'shifted awkwardly',
        'laughed awkwardly', 'awkward laugh', 'awkward smile', 'forced awkward smile',
        'followed willingly', 'hesitated awkwardly', 'not sure what to say',
        'did not know what to say', 'unsure what to say',
    ],
    flustered: [
        'flustered', 'stammered', 'stammering', 'stuttered', 'stuttering', 'spluttered',
        'protested', 'protesting', 'waved her hands', 'waved his hands', 'hands flailed',
        'fumbled', 'fumbling', 'caught off guard', 'off guard', 'panicked denial',
        'no no', 'wait wait', 'that is not', "that's not", 'i mean', 'not like that', 'rambling', 'blurting out',
        'blurted out', 'trying to act normal', 'words tumbled out',
        'tumbled out in a rush', 'i will put it away', "i'll put it away",
        'right now immediately', 'snatched the manga', 'clutched the book tighter',
        'hid it against her chest', 'clutching the book protectively',
        'flustered state', 'trying desperately to regain control',
        'using them as a barrier', 'using it as a barrier', 'definitely shutting up',
        'breathless', 'breathless voice', 'shut up now', 'definitely shutting up',
        'shaky breath', 'nervous energy', 'talking faster now', 'rambling again',
        'h-ah', 'mid-reach', 'protectively against her chest',
        'protectively against his chest', 'holding them protectively',
        'held them protectively',
    ],
    frustration: [
        'frustration', 'frustrated', 'frustrations', 'frustations', 'frustrating',
        'growled in frustration', 'sighed in frustration', 'pinched the bridge of her nose',
        'pinched the bridge of his nose', 'threw up her hands', 'threw up his hands',
        'gritted her teeth', 'gritted his teeth',
    ],
    horny: [
        'horny', 'lustful', 'lust', 'needy', 'needily', 'turned on', 'want you',
        'wanted you', 'craving', 'craved', 'aching need', 'desperate need',
        'bedroom eyes', 'seductive smile', 'climax building', 'ready to go', 'wet', 'drenched', 'soaked',
        'straddle your lap', 'straddled your lap', 'straddling your lap',
        'dripping wet', 'still-hard', 'still hard', 'pressing against your',
        'pressed against your', 'my turn', 'hardness', 'joining us', 'naked',
        'bare breasts', 'ground against', 'pressed closer', 'hungry eyes',
        'ready for round', 'lost count', 'bathroom fun',
    ],
    possessive: [
        'possessive', 'posessive', 'possessiveness', 'mine', 'you are mine',
        "you're mine", 'belongs to me', 'belong to me', 'claimed you', 'claiming you',
        'pulled you closer', 'kept you close', 'protective grip',
    ],
    suspicious: [
        'suspicious', 'suspicion', 'distrustful', 'distrust', 'narrowed her eyes',
        'narrowed his eyes', 'eyes narrowed', 'skeptical', 'skepticism', 'wary',
        'looked unconvinced', 'raised an eyebrow suspiciously', 'are you hiding something',
        'tone unreadable', 'how interesting', 'appeared in the doorway',
        'immediately noticed', 'noticed their state of undress',
        'towels barely covering', 'sitting intimately close',
        'did you two', 'take a shower together', 'took a shower together',
    ],
    vulnerable: [
        'vulnerable', 'vurnerable', 'vulnerability', 'unguarded', 'fragile',
        'voice softened', 'voice small', 'small voice', 'looked away helplessly',
        'opened up', 'let her guard down', 'let his guard down', 'teary smile',
        'uncertain smile', 'voice came out smaller', 'smaller than intended',
        'barely audible', 'quietly vulnerable', 'caught off guard by kindness',
        'not used to kindness', 'felt too exposed', 'asking for permission to exist',
        'almost vulnerable', 'someone you barely know', 'scary moving situation',
        'less stressful', 'please tell me', 'try not to be too much trouble',
        'i can be intense', 'afraid of being too much', 'shaky breath',
        'barely above a whisper', 'using them as a barrier', 'protectively against her chest',
        'admitted quietly', 'fingers fidgeted', 'fidgeted', 'met his gaze',
        'met her gaze', 'vulnerability was clear', 'because i do very much',
        'soft almost hurt', 'voice was soft', 'genuine uncertainty',
        'stepped closer', 'just us', 'trusting me',
    ],
};

export const EXPRESSION_ALIASES = {
    admiration: 'admiration',
    admire: 'admiration',
    impressed: 'admiration',
    amused: 'amusement',
    funny: 'amusement',
    playful: 'amusement',
    embarrassed: 'embarrassment',
    shy: 'embarrassment',
    blush: 'embarrassment',
    blushing: 'embarrassment',
    fluster: 'flustered',
    flustered: 'flustered',
    excitement: 'excitement',
    excited: 'excitement',
    panic: 'fear',
    panicked: 'fear',
    nervous: 'nervousness',
    anxiety: 'anxious',
    anxious: 'anxious',
    angry: 'anger',
    mad: 'anger',
    annoyed: 'annoyance',
    irritating: 'annoyance',
    irritated: 'agitation',
    affection: 'love',
    affectionate: 'love',
    love: 'love',
    loving: 'love',
    happy: 'joy',
    smile: 'joy',
    smiling: 'joy',
    grateful: 'gratitude',
    thankful: 'gratitude',
    sorry: 'remorse',
    apologetic: 'remorse',
    guilty: 'remorse',
    grief: 'grief',
    grieving: 'grief',
    sad: 'sadness',
    hurt: 'sadness',
    crying: 'sadness',
    scared: 'fear',
    afraid: 'fear',
    shocked: 'surprise',
    startled: 'surprise',
    confused: 'confusion',
    uncertainty: 'confusion',
    uncertain: 'confusion',
    curious: 'curiosity',
    ara: 'curiosity',
    desire: 'desire',
    wanting: 'desire',
    aroused: 'arousal',
    arousal: 'arousal',
    moaning: 'arousal',
    gasping: 'arousal',
    disappointed: 'disappointment',
    disapproving: 'disapproval',
    disgusted: 'disgust',
    neutral: 'neutral',
    optimistic: 'optimism',
    hopeful: 'optimism',
    proud: 'pride',
    smug: 'smug',
    satisfied: 'pride',
    realized: 'realization',
    relieved: 'relief',
    dominant: 'dominant',
    commanding: 'dominant',
    frustrated: 'frustration',
    frustration: 'frustration',
    frustrations: 'frustration',
    frustations: 'frustration',
    horny: 'horny',
    lust: 'horny',
    lustful: 'horny',
    lewd: 'horny',
    jealous: 'jealous',
    jealousy: 'jealous',
    possessive: 'possessive',
    posessive: 'possessive',
    suspicious: 'suspicious',
    suspicion: 'suspicious',
    vulnerable: 'vulnerable',
    vurnerable: 'vulnerable',
    awkward: 'awkward',
    worry: 'anxious',
    worried: 'anxious',
    huhu: 'amusement',
    wink: 'wink',
    turnpale: 'fear',
    pale: 'fear',
};

export function getExpressionHintKeys(expressionName) {
    const normalized = normalizeExpressionName(expressionName);
    const tokens = normalized.split(/\s+/).filter(Boolean);
    const keys = new Set();

    if (EXPRESSION_HINTS[normalized]) {
        keys.add(normalized);
    }

    for (const token of tokens) {
        if (EXPRESSION_HINTS[token]) {
            keys.add(token);
        }
        if (EXPRESSION_ALIASES[token]) {
            keys.add(EXPRESSION_ALIASES[token]);
        }
    }

    if (normalized.includes('fluster')) keys.add('flustered');
    if (normalized.includes('admir') || normalized.includes('impress')) keys.add('admiration');
    if (normalized.includes('amus') || normalized.includes('playful') || normalized.includes('teas') || normalized.includes('huhu')) keys.add('amusement');
    if (normalized.includes('embarrass') || normalized.includes('blush') || normalized.includes('shy')) keys.add('embarrassment');
    if (normalized.includes('excit') || normalized.includes('thrill') || normalized.includes('eager')) keys.add('excitement');
    if (normalized.includes('anx') || normalized.includes('worry')) keys.add('anxious');
    if (normalized.includes('nerv') || normalized.includes('anx')) keys.add('nervousness');
    if (normalized.includes('ang') || normalized.includes('mad') || normalized.includes('furious')) keys.add('anger');
    if (normalized.includes('annoy') || normalized.includes('exasperat')) keys.add('annoyance');
    if (normalized.includes('approv') || normalized.includes('pleased')) keys.add('approval');
    if (normalized.includes('caring') || normalized.includes('concern') || normalized.includes('comfort')) keys.add('caring');
    if (normalized.includes('confus') || normalized.includes('puzzl') || normalized.includes('uncertain')) keys.add('confusion');
    if (normalized.includes('curio') || normalized.includes('intrigu') || normalized.includes('ara')) keys.add('curiosity');
    if (normalized.includes('desire') || normalized.includes('want') || normalized.includes('longing')) keys.add('desire');
    if (normalized.includes('arous') || normalized.includes('heated') || normalized.includes('moan') || normalized.includes('gasp')) keys.add('arousal');
    if (normalized.includes('disappoint')) keys.add('disappointment');
    if (normalized.includes('disapprov')) keys.add('disapproval');
    if (normalized.includes('disgust') || normalized.includes('repuls') || normalized.includes('gross')) keys.add('disgust');
    if (normalized.includes('grat') || normalized.includes('thank')) keys.add('gratitude');
    if (normalized.includes('grief') || normalized.includes('mourn')) keys.add('grief');
    if (normalized.includes('affection') || normalized.includes('love') || normalized.includes('tender')) keys.add('love');
    if (normalized.includes('agitat') || normalized.includes('irritat') || normalized.includes('restless')) keys.add('agitation');
    if (normalized.includes('neutral') || normalized.includes('blank') || normalized.includes('deadpan')) keys.add('neutral');
    if (normalized.includes('optim') || normalized.includes('hope')) keys.add('optimism');
    if (normalized.includes('pride') || normalized.includes('proud') || normalized.includes('satisfied')) keys.add('pride');
    if (normalized.includes('smug')) keys.add('smug');
    if (normalized.includes('realiz') || normalized.includes('realis') || normalized.includes('understand')) keys.add('realization');
    if (normalized.includes('relief') || normalized.includes('reliev')) keys.add('relief');
    if (normalized.includes('remorse') || normalized.includes('sorry') || normalized.includes('guilt') || normalized.includes('regret')) keys.add('remorse');
    if (normalized.includes('sad') || normalized.includes('cry') || normalized.includes('tear') || normalized.includes('hurt')) keys.add('sadness');
    if (normalized.includes('fear') || normalized.includes('scared') || normalized.includes('afraid') || normalized.includes('pale') || normalized.includes('turnpale')) keys.add('fear');
    if (normalized.includes('happy') || normalized.includes('joy') || normalized.includes('smile') || normalized.includes('laugh')) keys.add('joy');
    if (normalized.includes('surpris') || normalized.includes('shock') || normalized.includes('startl')) keys.add('surprise');
    if (normalized.includes('wink')) keys.add('wink');
    if (normalized.includes('jealous') || normalized.includes('envy') || normalized.includes('envious')) keys.add('jealous');
    if (normalized.includes('domin') || normalized.includes('command')) keys.add('dominant');
    if (normalized.includes('frustrat') || normalized.includes('frustat')) keys.add('frustration');
    if (normalized.includes('horny') || normalized.includes('lust') || normalized.includes('lewd')) keys.add('horny');
    if (normalized.includes('possess') || normalized.includes('posess')) keys.add('possessive');
    if (normalized.includes('suspic') || normalized.includes('distrust') || normalized.includes('skeptic') || normalized.includes('wary')) keys.add('suspicious');
    if (normalized.includes('awkward')) keys.add('awkward');
    if (normalized.includes('vulner') || normalized.includes('vurner') || normalized.includes('unguarded') || normalized.includes('uncertain')) keys.add('vulnerable');

    return [...keys];
}
