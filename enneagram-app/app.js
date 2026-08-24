'use strict';

const STORAGE_KEY = 'enneagramApp:v12:state';
const LEGACY_STORAGE_KEYS = [
  'enneagramApp:v1:state', 'enneagramApp:v2:state', 'enneagramApp:v3:state',
  'enneagramApp:v4:state', 'enneagramApp:v5:state', 'enneagramApp:v6:state',
  'enneagramApp:v7:state', 'enneagramApp:v8:state', 'enneagramApp:v9:state',
  'enneagramApp:v10:state',
  'enneagramApp:v11:state'
];
const MODE_CONFIG = {
  standard: { name:'精度優先・標準版', count:78, typeItemsPerType:8, maxTypeScore:32 },
  short: { name:'短縮版', count:42, typeItemsPerType:4, maxTypeScore:16 }
};
const RATING_OPTIONS = [
  { value:'exactly', score:4, label:'まさに自分に当てはまる' },
  { value:'mostly', score:3, label:'かなり当てはまる' },
  { value:'somewhat', score:2, label:'少し当てはまる' },
  { value:'hardly', score:1, label:'ほとんど当てはまらない' },
  { value:'notAtAll', score:0, label:'まったく当てはまらない' }
];
const RATING_SCORE_MAP = Object.fromEntries(RATING_OPTIONS.map(function (option) {
  return [option.value, option.score];
}));
const FACET_GROUPS = {
  core: { label:'核（恐れ・欲求）' },
  defense: { label:'防衛（反応・回避）' },
  acquisition: { label:'獲得（満足・戦略）' },
  automatic: { label:'自動性（注意・自動反応）' }
};
const LIFE_DOMAIN_LABELS = {
  A:'一人の時間', B:'親しい人', C:'自分のための選択', D:'小さな予定外'
};
const TYPE_CORES = {
  1:{ fear:'自分が悪く、欠陥のある存在であること', desire:'高潔でありたい' },
  2:{ fear:'愛されるにふさわしくないこと', desire:'愛されたい' },
  3:{ fear:'自分に本来の価値がないこと', desire:'価値ある存在でありたい' },
  4:{ fear:'自分固有の存在意義がないこと', desire:'自分自身でありたい' },
  5:{ fear:'無力で無能であること', desire:'有能でありたい' },
  6:{ fear:'支えや導きがないこと', desire:'安全でありたい' },
  7:{ fear:'必要なものを奪われ、苦痛から逃れられないこと', desire:'幸福で満たされたい' },
  8:{ fear:'人に傷つけられ、コントロールされること', desire:'自分自身を守りたい' },
  9:{ fear:'つながりの喪失や分裂', desire:'平和でありたい' }
};
const CROSS_TYPE_MAP = {
  'autonomy|effort':1, 'attention|effort':2, 'attention|demand':3,
  'attention|withdraw':4, 'security|withdraw':5, 'security|effort':6,
  'security|demand':7, 'autonomy|demand':8, 'autonomy|withdraw':9
};
const TYPE_CHARACTER_FILES = {
  1:'type-01-reformer-v1.png', 2:'type-02-helper-v1.png', 3:'type-03-achiever-v1.png',
  4:'type-04-individualist-v1.png', 5:'type-05-investigator-v1.png', 6:'type-06-loyalist-v1.png',
  7:'type-07-enthusiast-v1.png', 8:'type-08-challenger-v1.png', 9:'type-09-peacemaker-v1.png'
};

const APP_DATA = JSON.parse(document.getElementById('app-data').textContent);
const TYPE_QUESTIONS = APP_DATA.typeQuestions;
const CROSS_QUESTIONS = APP_DATA.crossCheckQuestions;
const ALL_QUESTIONS = new Map(TYPE_QUESTIONS.concat(CROSS_QUESTIONS).map(function (question) {
  return [question.id, question];
}));
const TYPE_RESULTS = APP_DATA.typeResults;
const app = document.getElementById('app');
let state = emptyState();

function clearApp() { app.replaceChildren(); }

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function button(text, className, handler) {
  const node = el('button', className, text);
  node.type = 'button';
  node.addEventListener('click', handler);
  return node;
}

function svgEl(tag, attributes, text) {
  const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
  Object.entries(attributes || {}).forEach(function (entry) {
    node.setAttribute(entry[0], String(entry[1]));
  });
  if (text !== undefined) node.textContent = text;
  return node;
}

function showError(message) {
  const node = el('div', 'error', message);
  node.setAttribute('role', 'alert');
  return node;
}

function scrollPageTop() {
  requestAnimationFrame(function () {
    window.scrollTo({ top:0, left:0, behavior:'auto' });
  });
}

function emptyState() {
  return { mode:null, currentIndex:0, answers:{} };
}

function selectedOrder(mode) {
  return mode && MODE_CONFIG[mode] ? APP_DATA.orders[mode] : [];
}

function validateData() {
  const expectedIds = new Set();
  TYPE_QUESTIONS.forEach(function (question) {
    expectedIds.add(question.id);
  });
  CROSS_QUESTIONS.forEach(function (question) { expectedIds.add(question.id); });
  const typeDataValid = Array.isArray(TYPE_QUESTIONS) && TYPE_QUESTIONS.length === 108 &&
    TYPE_QUESTIONS.every(function (question) {
      return /^(ST[1-9]-0[1-8]|SH[1-9]-0[1-4])$/.test(question.id) &&
        MODE_CONFIG[question.version] && question.typeId >= 1 && question.typeId <= 9 &&
        FACET_GROUPS[question.facetGroup] && LIFE_DOMAIN_LABELS[question.lifeDomain] &&
        question.facet && question.text;
    });
  const crossValid = Array.isArray(CROSS_QUESTIONS) && CROSS_QUESTIONS.length === 12 &&
    CROSS_QUESTIONS.every(function (question) {
      return /^(SC|SS|HC|HS)-0[1-3]$/.test(question.id) && MODE_CONFIG[question.version] &&
        question.axis && question.key && question.label && LIFE_DOMAIN_LABELS[question.lifeDomain] &&
        question.text;
    });
  const modeDataValid = Object.keys(MODE_CONFIG).every(function (mode) {
    const order = selectedOrder(mode);
    const modeQuestions = TYPE_QUESTIONS.filter(function (question) { return question.version === mode; });
    const modeCross = CROSS_QUESTIONS.filter(function (question) { return question.version === mode; });
    const expectedModeIds = new Set(modeQuestions.concat(modeCross).map(function (question) { return question.id; }));
    const typeCounts = {};
    modeQuestions.forEach(function (question) {
      typeCounts[question.typeId] = (typeCounts[question.typeId] || 0) + 1;
    });
    return modeQuestions.length === MODE_CONFIG[mode].typeItemsPerType * 9 && modeCross.length === 6 &&
      Array.from({ length:9 }, function (_, index) {
        return typeCounts[index + 1] === MODE_CONFIG[mode].typeItemsPerType;
      }).every(Boolean) && Array.isArray(order) && order.length === MODE_CONFIG[mode].count &&
      new Set(order).size === order.length && order.every(function (id) {
        return expectedModeIds.has(id) && ALL_QUESTIONS.get(id).version === mode;
      }) && order.every(function (id) { return expectedIds.has(id); }) &&
      expectedModeIds.size === order.length;
  });
  const supportingValid = APP_DATA.reflectionQuestions.length === 3 &&
    Object.keys(TYPE_RESULTS).length === 9 && Object.keys(TYPE_CHARACTER_FILES).length === 9 &&
    APP_DATA.closingMessage;
  if (expectedIds.size !== 120 || !typeDataValid || !crossValid || !modeDataValid || !supportingValid) {
    throw new Error('診断データに不備があります。管理者へ連絡してください。');
  }
}

function isRatingValue(value) {
  return Object.prototype.hasOwnProperty.call(RATING_SCORE_MAP, value);
}

function normalizeState(candidate) {
  if (!candidate || candidate.schemaVersion !== 12 || !MODE_CONFIG[candidate.mode] ||
      !Number.isInteger(candidate.currentIndex) || candidate.currentIndex < 0 ||
      candidate.currentIndex >= selectedOrder(candidate.mode).length ||
      !candidate.answers || typeof candidate.answers !== 'object' || Array.isArray(candidate.answers)) return null;
  const allowed = new Set(selectedOrder(candidate.mode));
  const normalized = { mode:candidate.mode, currentIndex:candidate.currentIndex, answers:{} };
  for (const entry of Object.entries(candidate.answers)) {
    if (!allowed.has(entry[0]) || !isRatingValue(entry[1])) return null;
    normalized.answers[entry[0]] = entry[1];
  }
  return normalized;
}

function loadState() {
  try {
    LEGACY_STORAGE_KEYS.forEach(function (key) { sessionStorage.removeItem(key); });
    return normalizeState(JSON.parse(sessionStorage.getItem(STORAGE_KEY))) || emptyState();
  } catch (_) {
    return emptyState();
  }
}

function saveState() {
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify({
      schemaVersion:12,
      mode:state.mode,
      currentIndex:state.currentIndex,
      answers:state.answers
    }));
  } catch (_) {
    // 一時保存できない環境でも診断は継続します。
  }
}

function answeredCount() {
  return selectedOrder(state.mode).reduce(function (total, id) {
    return total + (isRatingValue(state.answers[id]) ? 1 : 0);
  }, 0);
}

function beginMode(mode) {
  if (!MODE_CONFIG[mode]) return;
  const continuing = state.mode === mode && answeredCount() > 0;
  if (!continuing && answeredCount() > 0 && !window.confirm('進行中の回答を消して、別の版を始めますか？')) return;
  if (!continuing) state = { mode:mode, currentIndex:0, answers:{} };
  saveState();
  if (continuing) showQuestion(true);
  else showRoleOffWarmup(true);
}

function showRoleOffWarmup(shouldScrollTop) {
  clearApp();
  app.append(el('p', 'scenario-number', MODE_CONFIG[state.mode].name));
  app.append(el('h1', '', '仕事の自分を、いったん置いていく'));
  app.append(el('p', 'lead', 'ここからは、肩書きや責任から離れた私生活の自分を思い出します。'));
  const card = el('section', 'card role-off-card');
  card.append(el('h2', '', '始める前に、3つの時間を思い出してください'));
  const scenes = el('div', 'role-off-scenes');
  scenes.append(el('p', '', '家で一人で過ごした時間'));
  scenes.append(el('p', '', '親しい人と気楽に過ごした時間'));
  scenes.append(el('p', '', '自分のためだけに何かを選んだ時間'));
  card.append(scenes);
  card.append(el('p', 'scenario-prompt', '立場上そうすべき自分ではなく、誰にも指示されなくても繰り返す心の動きで答えます。'));
  app.append(card);
  const actions = el('div', 'actions');
  actions.append(button('版の選択へ戻る', 'secondary', function () { showStart(true); }));
  actions.append(button('私生活の自分を思い出して始める', 'primary', function () { showQuestion(true); }));
  app.append(actions);
  if (shouldScrollTop) scrollPageTop();
}

function renderModeCard(mode, description) {
  const config = MODE_CONFIG[mode];
  const card = el('section', 'card mode-card');
  card.append(el('p', 'scenario-number', config.count + ' QUESTIONS'));
  card.append(el('h2', '', config.name));
  card.append(el('p', '', description));
  const continuing = state.mode === mode && answeredCount() > 0;
  if (continuing) {
    card.append(el('p', 'answer-summary', '回答済み ' + answeredCount() + ' / ' + config.count));
  }
  card.append(button(continuing ? 'この版を続ける' : 'この版を始める', 'primary', function () {
    beginMode(mode);
  }));
  return card;
}

function showStart(shouldScrollTop) {
  clearApp();
  app.append(el('h1', '', 'エニアグラム・タイプ診断'));
  app.append(el('p', 'lead', '同じ行動でも、心の奥にある「なぜそうするのか」を一つずつ確かめ、タイプ候補とウイングを探ります。'));
  const guide = el('section', 'card');
  guide.append(el('h2', '', '答えるときの基準'));
  guide.append(el('p', '', '診断を始める前に、家で一人の時間、親しい人との時間、自分のための選択を思い出します。'));
  guide.append(el('p', '', '仕事上そうすべき自分ではなく、誰にも指示されなくても繰り返す心の動きで答えてください。'));
  app.append(guide);
  const modes = el('div', 'mode-grid');
  modes.append(renderModeCard('standard', '9タイプを各8側面から確認する72問と、判定を別角度から確かめる6問で、候補を詳しく比較します。'));
  modes.append(renderModeCard('short', '日常の具体的な場面で、各タイプの恐れ・欲求・戦略・自動反応を確認する独自36問と、6問のクロスチェックで傾向を見ます。'));
  app.append(modes);
  app.append(el('p', 'privacy', '回答内容はこの端末のセッション内だけで一時保存され、外部へ送信されません。結果はタイプを確定するものではなく、自己観察の候補です。'));
  if (shouldScrollTop) scrollPageTop();
}

function renderRatingActions(question, selectedValue) {
  const actions = el('div', 'rating-actions');
  const ratingButtons = [];
  RATING_OPTIONS.forEach(function (option) {
    const node = button(option.label, 'rating-button', function () {
      state.answers[question.id] = option.value;
      ratingButtons.forEach(function (ratingButton) {
        ratingButton.setAttribute('aria-pressed', String(ratingButton === node));
      });
      const optionCard = actions.closest('.motive-option');
      if (optionCard) optionCard.classList.add('answered');
      const progressBar = document.querySelector('.progress > div');
      if (progressBar) {
        progressBar.style.width = ((answeredCount() / selectedOrder(state.mode).length) * 100) + '%';
      }
      const error = document.querySelector('.scenario-card .error');
      if (error) error.remove();
      saveState();
    });
    node.setAttribute('aria-pressed', String(selectedValue === option.value));
    node.setAttribute('aria-label', question.text + '：' + option.label);
    ratingButtons.push(node);
    actions.append(node);
  });
  return actions;
}

function showQuestion(shouldScrollTop, errorMessage) {
  const order = selectedOrder(state.mode);
  if (!order.length) {
    showStart(true);
    return;
  }
  const question = ALL_QUESTIONS.get(order[state.currentIndex]);
  clearApp();
  const progressHead = el('div', 'progress-head');
  progressHead.append(el('span', '', MODE_CONFIG[state.mode].name));
  progressHead.append(el('span', '', '質問 ' + (state.currentIndex + 1) + ' / ' + order.length));
  app.append(progressHead);
  const progress = el('div', 'progress');
  const progressBar = el('div');
  progressBar.style.width = ((answeredCount() / order.length) * 100) + '%';
  progress.append(progressBar);
  app.append(progress);
  const card = el('section', 'card scenario-card');
  card.append(el('p', 'scenario-number', 'QUESTION ' + (state.currentIndex + 1)));
  card.append(el('p', 'life-domain', '生活場面：' + LIFE_DOMAIN_LABELS[question.lifeDomain]));
  card.append(el('h1', '', question.text));
  card.append(el('p', 'scenario-prompt', 'この心の動きは、普段の自分にどの程度当てはまりますか？'));
  if (errorMessage) card.append(showError(errorMessage));
  const optionCard = el('article', 'motive-option');
  if (isRatingValue(state.answers[question.id])) optionCard.classList.add('answered');
  optionCard.append(renderRatingActions(question, state.answers[question.id]));
  card.append(optionCard);
  app.append(card);
  const actions = el('div', 'actions');
  if (state.currentIndex > 0) {
    actions.append(button('戻る', 'secondary', function () {
      state.currentIndex -= 1;
      saveState();
      showQuestion(true);
    }));
  }
  actions.append(button(state.currentIndex === order.length - 1 ? '結果を見る' : '次へ', 'primary', proceedQuestion));
  app.append(actions);
  if (shouldScrollTop) scrollPageTop();
}

function proceedQuestion() {
  const order = selectedOrder(state.mode);
  const questionId = order[state.currentIndex];
  if (!isRatingValue(state.answers[questionId])) {
    showQuestion(false, 'この質問への回答を選んでください。');
    return;
  }
  if (state.currentIndex < order.length - 1) {
    state.currentIndex += 1;
    saveState();
    showQuestion(true);
    return;
  }
  if (answeredCount() !== order.length) {
    const firstMissing = order.findIndex(function (id) { return !isRatingValue(state.answers[id]); });
    state.currentIndex = Math.max(firstMissing, 0);
    saveState();
    showQuestion(true, '未回答の質問があります。');
    return;
  }
  showResults(calculateDiagnosis());
}

function facetGroupFor(question) {
  return question.facetGroup;
}

function scoreRanking(scores) {
  return Object.keys(scores).map(Number).sort(function (a, b) {
    return scores[b] - scores[a] || a - b;
  }).map(function (typeId, index, sorted) {
    const rank = index > 0 && scores[typeId] === scores[sorted[index - 1]] ?
      null : index + 1;
    return { typeId:typeId, score:scores[typeId], rank:rank };
  }).reduce(function (items, item, index) {
    if (item.rank === null) item.rank = items[index - 1].rank;
    items.push(item);
    return items;
  }, []);
}

function highestKeys(scoreObject) {
  const maxScore = Math.max.apply(null, Object.values(scoreObject));
  return Object.keys(scoreObject).filter(function (key) { return scoreObject[key] === maxScore; });
}

function determineWing(coreType, scores) {
  const leftType = coreType === 1 ? 9 : coreType - 1;
  const rightType = coreType === 9 ? 1 : coreType + 1;
  const leftScore = scores[leftType];
  const rightScore = scores[rightType];
  const balanced = leftScore === rightScore;
  const wingType = balanced ? null : (leftScore > rightScore ? leftType : rightType);
  return {
    coreType:coreType,
    wingType:wingType,
    wingLabel:balanced ? coreType + '（左右同点）' : coreType + 'w' + wingType,
    wingAType:leftType,
    wingBType:rightType,
    wingAScore:leftScore,
    wingBScore:rightScore,
    balanced:balanced
  };
}

function calculateDiagnosis() {
  const scores = { 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0, 9:0 };
  const facetScores = {};
  const crossScores = {
    center:{ autonomy:0, attention:0, security:0 },
    strategy:{ demand:0, effort:0, withdraw:0 }
  };
  Object.keys(scores).forEach(function (typeId) {
    facetScores[typeId] = {};
    Object.keys(FACET_GROUPS).forEach(function (group) {
      facetScores[typeId][group] = { score:0, max:0, count:0 };
    });
  });
  selectedOrder(state.mode).forEach(function (id) {
    const question = ALL_QUESTIONS.get(id);
    const answer = state.answers[id];
    if (!isRatingValue(answer)) throw new Error('未回答の質問があるため診断できません。');
    const score = RATING_SCORE_MAP[answer];
    if (question.typeId) {
      scores[question.typeId] += score;
      const facetGroup = facetGroupFor(question);
      facetScores[question.typeId][facetGroup].score += score;
      facetScores[question.typeId][facetGroup].max += 4;
      facetScores[question.typeId][facetGroup].count += 1;
    } else {
      crossScores[question.axis][question.key] = score;
    }
  });
  const ranking = scoreRanking(scores);
  const topCandidates = ranking.slice(0, 3);
  const topScore = ranking[0].score;
  const exactTopTypes = ranking.filter(function (item) { return item.score === topScore; }).map(function (item) { return item.typeId; });
  const centerKeys = highestKeys(crossScores.center);
  const strategyKeys = highestKeys(crossScores.strategy);
  const crossCandidates = [];
  centerKeys.forEach(function (center) {
    strategyKeys.forEach(function (strategy) {
      const typeId = CROSS_TYPE_MAP[center + '|' + strategy];
      if (typeId && !crossCandidates.includes(typeId)) crossCandidates.push(typeId);
    });
  });
  return {
    mode:state.mode,
    answers:Object.assign({}, state.answers),
    scores:scores,
    facetScores:facetScores,
    ranking:ranking,
    topCandidates:topCandidates,
    exactTopTypes:exactTopTypes,
    crossScores:crossScores,
    centerKeys:centerKeys,
    strategyKeys:strategyKeys,
    crossCandidates:crossCandidates,
    uniformAnswers:new Set(Object.values(state.answers)).size === 1
  };
}

function renderWingReference(wingResult) {
  const section = el('div', 'section-card wing-reference');
  section.append(el('h3', '', 'ウイング参考：' + wingResult.wingLabel));
  section.append(el('p', '', '隣接するタイプ' + wingResult.wingAType + 'は' + wingResult.wingAScore + '点、タイプ' + wingResult.wingBType + 'は' + wingResult.wingBScore + '点です。'));
  section.append(el('p', 'chart-help', 'ウイングはコア候補を確定する根拠には加えず、左右タイプの得点を参考表示しています。'));
  return section;
}

function renderFacetSummary(result, typeId) {
  const list = el('ul', 'facet-list');
  Object.keys(FACET_GROUPS).forEach(function (group) {
    const value = result.facetScores[typeId][group];
    if (!value.count) return;
    list.append(el('li', '', FACET_GROUPS[group].label + '：' + value.score + ' / ' + value.max + '点'));
  });
  return list;
}

function renderCandidate(result, item, position) {
  const type = TYPE_RESULTS[item.typeId];
  const config = MODE_CONFIG[result.mode];
  const card = el('section', 'card candidate-card');
  card.append(el('p', 'scenario-number', '候補 ' + position));
  card.append(el('h2', '', 'タイプ' + item.typeId + '：' + type.nameJa));
  card.append(el('p', 'lead', item.score + ' / ' + config.maxTypeScore + '点（平均 ' + (item.score / config.typeItemsPerType).toFixed(2) + ' / 4）'));
  card.append(el('p', '', '根元的恐れ：' + TYPE_CORES[item.typeId].fear));
  card.append(el('p', '', '根元的欲求：' + TYPE_CORES[item.typeId].desire));
  card.append(renderFacetSummary(result, item.typeId));
  card.append(renderWingReference(determineWing(item.typeId, result.scores)));
  return card;
}

function crossLabel(axis, key) {
  const question = CROSS_QUESTIONS.find(function (item) {
    return item.axis === axis && item.key === key;
  });
  return question ? question.label : key;
}

function renderCrossCheck(result) {
  const section = el('section', 'card');
  section.append(el('h2', '', '構造クロスチェック'));
  section.append(el('p', '', 'タイプ点とは別に、「何を求めるか」と「どう得ようとするか」の組み合わせを確認します。'));
  section.append(el('p', 'lead', '中心：' + result.centerKeys.map(function (key) { return crossLabel('center', key); }).join('・') + ' ／ 方法：' + result.strategyKeys.map(function (key) { return crossLabel('strategy', key); }).join('・')));
  const candidateText = result.crossCandidates.map(function (typeId) {
    return 'タイプ' + typeId + '「' + TYPE_RESULTS[typeId].nameJa + '」';
  }).join('・');
  section.append(el('p', '', 'この組み合わせが示す候補：' + candidateText));
  const topIds = result.topCandidates.map(function (item) { return item.typeId; });
  const matches = result.crossCandidates.filter(function (typeId) { return topIds.includes(typeId); });
  if (matches.length) {
    section.append(el('div', 'answer-summary', '上位3候補との一致：タイプ' + matches.join('・タイプ')));
  } else {
    section.append(el('div', 'notice', '上位3候補とは一致しませんでした。回答の誤りとは決めず、クロスチェック側の候補も読み比べてください。'));
  }
  return section;
}

function polarPoint(typeId, factor, radius) {
  const angle = -Math.PI / 2 + ((typeId - 1) * Math.PI * 2 / 9);
  return { x:260 + Math.cos(angle) * radius * factor, y:260 + Math.sin(angle) * radius * factor };
}

function pointsAttribute(points) {
  return points.map(function (point) {
    return point.x.toFixed(2) + ',' + point.y.toFixed(2);
  }).join(' ');
}

function renderRadarChart(result) {
  const maxScore = MODE_CONFIG[result.mode].maxTypeScore;
  const svg = svgEl('svg', { class:'radar-svg', viewBox:'0 0 520 520', role:'img', 'aria-labelledby':'radar-title radar-description' });
  svg.append(svgEl('title', { id:'radar-title' }, '9タイプの得点レーダーチャート'));
  svg.append(svgEl('desc', { id:'radar-description' }, '9タイプの得点を同じ上限で比較します。'));
  [0.25, 0.5, 0.75, 1].forEach(function (factor) {
    const points = [];
    for (let typeId = 1; typeId <= 9; typeId += 1) points.push(polarPoint(typeId, factor, 168));
    svg.append(svgEl('polygon', { points:pointsAttribute(points), class:factor === 1 ? 'radar-grid radar-grid-outer' : 'radar-grid' }));
  });
  for (let typeId = 1; typeId <= 9; typeId += 1) {
    const end = polarPoint(typeId, 1, 168);
    svg.append(svgEl('line', { x1:260, y1:260, x2:end.x, y2:end.y, class:'radar-axis' }));
  }
  const scorePoints = [];
  for (let typeId = 1; typeId <= 9; typeId += 1) scorePoints.push(polarPoint(typeId, result.scores[typeId] / maxScore, 168));
  svg.append(svgEl('polygon', { points:pointsAttribute(scorePoints), class:'radar-area' }));
  const topIds = result.topCandidates.map(function (item) { return item.typeId; });
  scorePoints.forEach(function (point, index) {
    svg.append(svgEl('circle', { cx:point.x, cy:point.y, r:5, class:topIds.includes(index + 1) ? 'radar-dot radar-dot-top' : 'radar-dot' }));
  });
  for (let typeId = 1; typeId <= 9; typeId += 1) {
    const labelPoint = polarPoint(typeId, 1, 205);
    const label = svgEl('text', {
      x:labelPoint.x, y:labelPoint.y,
      class:topIds.includes(typeId) ? 'radar-label radar-label-top' : 'radar-label',
      tabindex:'0', role:'link',
      'aria-label':'タイプ' + typeId + '、' + TYPE_RESULTS[typeId].nameJa + '、' + result.scores[typeId] + '点。解説へ移動'
    }, 'T' + typeId + ' ' + result.scores[typeId]);
    label.addEventListener('click', function () { jumpToType(typeId); });
    label.addEventListener('keydown', function (event) {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        jumpToType(typeId);
      }
    });
    svg.append(label);
  }
  return svg;
}

function renderRadarSection(result) {
  const section = el('section', 'card result-overview');
  const maxScore = MODE_CONFIG[result.mode].maxTypeScore;
  section.append(el('h2', '', '9タイプの得点バランス'));
  section.append(el('p', 'chart-help', '外側ほど得点が高く、この版の上限は' + maxScore + '点です。得点は質問数が同じタイプ同士で比較しています。'));
  const wrap = el('div', 'radar-wrap');
  wrap.append(renderRadarChart(result));
  section.append(wrap);
  const nav = el('div', 'score-nav');
  result.ranking.forEach(function (item) {
    const type = TYPE_RESULTS[item.typeId];
    const scoreButton = button(item.rank + '位　タイプ' + item.typeId + ' ' + type.nameJa + '　' + item.score + '点', result.topCandidates.some(function (candidate) { return candidate.typeId === item.typeId; }) ? 'score-button top' : 'score-button', function () {
      jumpToType(item.typeId);
    });
    nav.append(scoreButton);
  });
  section.append(nav);
  return section;
}

function renderReflection() {
  const section = el('section', 'card');
  section.append(el('h2', '', '結果を読んだ後の振り返り'));
  section.append(el('p', '', 'ここは採点も保存もしません。上位候補を読み比べながら、自分の言葉で考えてみてください。'));
  const list = el('ol');
  APP_DATA.reflectionQuestions.forEach(function (question) { list.append(el('li', '', question)); });
  section.append(list);
  return section;
}

function ratingOption(value) {
  return RATING_OPTIONS.find(function (option) { return option.value === value; });
}

function renderAnswerReview(result) {
  const details = el('details', 'type-detail answer-review');
  details.append(el('summary', '', '回答内容とタイプ対応を確認する'));
  const body = el('div', 'type-detail-body');
  body.append(el('p', 'chart-help', '診断中は伏せていたタイプ対応と点数を、ここで確認できます。クロスチェック項目はタイプ点に加えていません。'));
  const list = el('ol');
  selectedOrder(result.mode).forEach(function (id) {
    const question = ALL_QUESTIONS.get(id);
    const rating = ratingOption(result.answers[id]);
    const target = question.typeId ? 'タイプ' + question.typeId : 'クロスチェック';
    list.append(el('li', '', target + '・' + rating.label + '（' + rating.score + '点）— ' + question.text));
  });
  body.append(list);
  details.append(body);
  return details;
}

function renderTypeIllustration(type) {
  const illustrations = {
    1:'より良く整える、公正なまなざし', 2:'思いやりで人を支える、温かな心',
    3:'目標を形にし、前へ進む力', 4:'深い感情を、意味ある表現へ変える',
    5:'静かに観察し、本質を見抜く', 6:'先を読み、備えと信頼で守る',
    7:'可能性を見つけ、喜びへ飛び込む', 8:'力強く道を切り開き、弱い者を守る',
    9:'違いを包み、穏やかな調和をつくる'
  };
  const figure = el('figure', 'type-illustration');
  const image = el('img', 'illustration-img');
  image.src = 'assets/characters-v1/' + TYPE_CHARACTER_FILES[type.typeId];
  image.alt = 'タイプ' + type.typeId + '「' + type.nameJa + '」を表す人物イラスト';
  image.loading = 'lazy';
  image.decoding = 'async';
  figure.append(image);
  figure.append(el('figcaption', 'illustration-caption', illustrations[type.typeId]));
  return figure;
}

function paragraphSection(title, text, className) {
  const section = el('section', 'section-card ' + (className || ''));
  section.append(el('h3', '', title));
  section.append(el('p', '', text));
  return section;
}

function renderTypeDetail(type, open) {
  const details = el('details', 'type-detail');
  details.id = 'type-detail-' + type.typeId;
  details.open = Boolean(open);
  details.append(el('summary', '', 'タイプ' + type.typeId + '：' + type.nameJa + '（' + type.nameEn + '）'));
  const body = el('div', 'type-detail-body');
  body.append(el('p', 'subtitle', '〜' + type.subtitle + '〜'));
  body.append(renderTypeIllustration(type));
  body.append(paragraphSection('根元的恐れ', TYPE_CORES[type.typeId].fear));
  body.append(paragraphSection('根元的欲求', TYPE_CORES[type.typeId].desire));
  body.append(paragraphSection('唯一無二の至高の才能', type.essence));
  body.append(paragraphSection('自動操縦（トランス状態）に陥るサイン', type.warningSignal));
  body.append(paragraphSection('内なる「超自我の行進命令」', type.innerCommand));
  body.append(paragraphSection('強みが裏目に出る悲劇（自己成就的予言）', type.tragedy, 'runaway'));
  body.append(paragraphSection('心の鎧を脱ぐための「癒しの態度」', type.healing, 'healing'));
  details.append(body);
  return details;
}

function renderAllTypeDetails(result) {
  const section = el('section', 'type-details');
  section.append(el('h2', '', '各タイプの特徴'));
  section.append(el('p', 'chart-help', '上位候補だけでなく、クロスチェックが示した候補も読み比べられます。'));
  const openTypes = new Set(result.topCandidates.map(function (item) { return item.typeId; }).concat(result.crossCandidates));
  for (let typeId = 1; typeId <= 9; typeId += 1) section.append(renderTypeDetail(TYPE_RESULTS[typeId], openTypes.has(typeId)));
  return section;
}

function jumpToType(typeId) {
  const target = document.getElementById('type-detail-' + typeId);
  if (!target) return;
  target.open = true;
  target.scrollIntoView({ behavior:'smooth', block:'start' });
  requestAnimationFrame(function () {
    const summary = target.querySelector('summary');
    if (summary) summary.focus({ preventScroll:true });
  });
}

function renderClosingMessage() {
  const section = el('section', 'card closing-message');
  section.append(el('h2', '', APP_DATA.closingMessage.title));
  APP_DATA.closingMessage.paragraphs.forEach(function (paragraph) { section.append(el('p', '', paragraph)); });
  return section;
}

function showResults(result) {
  clearApp();
  app.append(el('p', 'scenario-number', MODE_CONFIG[result.mode].name));
  app.append(el('h1', '', '最も強く表れたタイプ候補'));
  app.append(el('p', 'lead', '行動そのものではなく、その奥にある恐れ・欲求・反応の得点から、上位3候補を表示しています。'));
  if (result.exactTopTypes.length > 1) {
    app.append(el('div', 'notice', '最高得点が同点です：タイプ' + result.exactTopTypes.join('・タイプ') + '。1つへ絞らず、動機を読み比べてください。'));
  }
  if (result.uniformAnswers) {
    app.append(el('div', 'notice', 'すべての質問に同じ段階で回答しています。タイプ間の違いが出にくいため、結果は判定保留として読み、必要なら回答を見直してください。'));
  }
  result.topCandidates.forEach(function (item, index) { app.append(renderCandidate(result, item, index + 1)); });
  app.append(renderCrossCheck(result));
  app.append(renderRadarSection(result));
  app.append(el('p', 'notice', '点差だけでタイプを確定しません。根元的恐れと根元的欲求が、繰り返す反応を最もよく説明する候補を本人が確認してください。短縮版は標準版より確認する側面が少ないため、より暫定的な結果です。'));
  app.append(renderReflection());
  app.append(renderAnswerReview(result));
  app.append(renderAllTypeDetails(result));
  app.append(renderClosingMessage());
  app.append(button('版の選択へ戻って診断し直す', 'primary', restart));
  scrollPageTop();
}

function restart() {
  state = emptyState();
  try {
    sessionStorage.removeItem(STORAGE_KEY);
    LEGACY_STORAGE_KEYS.forEach(function (key) { sessionStorage.removeItem(key); });
  } catch (_) {
    // 保存不可の場合も開始画面へ戻します。
  }
  showStart(true);
}

try {
  validateData();
  state = loadState();
  document.getElementById('app-version').textContent = APP_VERSION;
  showStart();
} catch (error) {
  clearApp();
  app.append(showError(error.message));
}
