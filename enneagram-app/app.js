'use strict';

const STORAGE_KEY = 'enneagramApp:v14:state';
const LEGACY_STORAGE_KEYS = [
  'enneagramApp:v1:state', 'enneagramApp:v2:state', 'enneagramApp:v3:state',
  'enneagramApp:v4:state', 'enneagramApp:v5:state', 'enneagramApp:v6:state',
  'enneagramApp:v7:state', 'enneagramApp:v8:state', 'enneagramApp:v9:state',
  'enneagramApp:v10:state',
  'enneagramApp:v11:state', 'enneagramApp:v12:state', 'enneagramApp:v13:state'
];
const MODE_CONFIG = {
  standard: { name:'標準版', count:78, typeItemsPerType:8, maxTypeScore:32 },
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
let previousQuestionsCleared = false;

function clearApp() {
  app.replaceChildren();
  app.classList.remove('question-screen', 'start-screen');
  document.body.classList.remove('question-view');
}

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

function isMode(mode) {
  return Object.prototype.hasOwnProperty.call(MODE_CONFIG, mode);
}

function selectedOrder(mode) {
  return isMode(mode) ? APP_DATA.orders[mode] : [];
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
  if (!candidate || candidate.schemaVersion !== 14 || !isMode(candidate.mode) ||
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
    LEGACY_STORAGE_KEYS.forEach(function (key) {
      if (sessionStorage.getItem(key) !== null) previousQuestionsCleared = true;
      sessionStorage.removeItem(key);
    });
    return normalizeState(JSON.parse(sessionStorage.getItem(STORAGE_KEY))) || emptyState();
  } catch (_) {
    return emptyState();
  }
}

function saveState() {
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify({
      schemaVersion:14,
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
  if (!isMode(mode)) return;
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
  card.append(el('p', 'scenario-prompt', '最近の気分だけでなく、以前から繰り返してきた傾向を思い出してください。理想や役割上の振る舞いではなく、自然に優先する方向で答えます。'));
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
  app.classList.add('start-screen');
  if (previousQuestionsCleared) app.append(el('p', 'notice', '質問を改訂したため、旧版の途中回答をリセットしました。新しい質問で最初から回答してください。'));
  app.append(el('h1', '', 'エニアグラム・タイプ診断'));
  app.append(el('p', 'lead', '同じ行動でも、心の奥にある「なぜそうするのか」を一つずつ確かめ、タイプ候補とウイングを探ります。'));
  const guide = el('section', 'card');
  guide.append(el('h2', '', '答えるときの基準'));
  guide.append(el('p', '', '診断を始める前に、家で一人の時間、親しい人との時間、自分のための選択を思い出します。'));
  guide.append(el('p', '', '最近の出来事だけでなく、長年、場面が変わっても繰り返してきた傾向を振り返ります。迷った質問は「保留して次の未回答へ」で後から答えられます。'));
  guide.append(el('p', '', 'どの文も良し悪しを決めるものではありません。理想の自分ではなく、私生活で自然に繰り返す心の動きで答えてください。'));
  app.append(guide);
  const modes = el('div', 'mode-grid');
  modes.append(renderModeCard('standard', '9タイプを各8側面から確認する72問と、回答傾向を振り返る補助6問で、候補を詳しく比較します。'));
  modes.append(renderModeCard('short', '日常の具体的な場面で、各タイプの恐れ・欲求・戦略・自動反応を確認する独自36問と、補助6問で傾向を振り返ります。'));
  app.append(modes);
  app.append(el('p', 'chart-help', '標準版は確認する質問数が多い版です。両版の診断精度の差は実測していません。'));
  app.append(el('p', 'privacy', '回答内容はこの端末のセッション内だけで一時保存され、外部へ送信されません。結果はタイプを確定するものではなく、自己観察の候補です。'));
  if (shouldScrollTop) scrollPageTop();
}

// 繰り返し表示していた説明は、回答を保持したまま開けるヒントにまとめます。
function showAnswerHelp() {
  const dialog = el('dialog', 'answer-help');
  dialog.setAttribute('aria-labelledby', 'answer-help-title');
  const title = el('h2', '', '回答のヒント');
  title.id = 'answer-help-title';
  dialog.append(title);
  dialog.append(el('p', '', '理想ではなく、以前から繰り返してきた自分に当てはまる度合いを選びます。'));
  dialog.append(el('p', '', '迷うときは「保留」で次の未回答へ進めます。「少し当てはまる」は、わからないという意味ではありません。'));
  dialog.append(el('p', '', '回答を選んだら「次へ」を押してください。「戻る」で前の回答を変更できます。'));
  dialog.append(button('閉じる', 'primary', function () { dialog.close(); }));
  dialog.addEventListener('close', function () { dialog.remove(); });
  document.body.append(dialog);
  dialog.showModal();
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
      const countLabel = document.querySelector('.answer-count');
      if (countLabel) {
        countLabel.textContent = '回答済み ' + answeredCount() + ' / ' + selectedOrder(state.mode).length;
        countLabel.classList.remove('error');
        countLabel.removeAttribute('role');
        countLabel.removeAttribute('aria-label');
      }
      const optionCard = actions.closest('.motive-option');
      if (optionCard) optionCard.classList.add('answered');
      const progressBar = document.querySelector('.progress > div');
      if (progressBar) {
        progressBar.style.width = ((answeredCount() / selectedOrder(state.mode).length) * 100) + '%';
        progressBar.parentElement.setAttribute('aria-valuenow', String(answeredCount()));
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
  if (!order.length) { showStart(true); return; }
  const question = ALL_QUESTIONS.get(order[state.currentIndex]);
  clearApp();
  document.body.classList.add('question-view');
  app.classList.add('question-screen');
  const status = el('div', 'question-status');
  const progressHead = el('div', 'progress-head');
  progressHead.append(el('strong', '', '質問 ' + (state.currentIndex + 1) + ' / ' + order.length));
  const countLabel = el('span', 'answer-count', '回答済み ' + answeredCount() + ' / ' + order.length);
  if (errorMessage) {
    // エラーは既存の進捗行に表示し、回答欄や「次へ」を押し下げません。
    countLabel.classList.add('error');
    countLabel.textContent = errorMessage.includes('だけ') ? '残りはこの1問です' : '回答を選んでください';
    countLabel.setAttribute('role', 'alert');
    countLabel.setAttribute('aria-label', errorMessage);
  }
  progressHead.append(countLabel);
  const help = button('?', 'help-button', showAnswerHelp);
  help.setAttribute('aria-label', '回答のヒント');
  progressHead.append(help);
  status.append(progressHead);
  const progress = el('div', 'progress');
  progress.setAttribute('role', 'progressbar');
  progress.setAttribute('aria-label', MODE_CONFIG[state.mode].name + 'の回答進捗');
  progress.setAttribute('aria-valuemin', '0');
  progress.setAttribute('aria-valuemax', String(order.length));
  progress.setAttribute('aria-valuenow', String(answeredCount()));
  const progressBar = el('div');
  progressBar.style.width = ((answeredCount() / order.length) * 100) + '%';
  progress.append(progressBar);
  status.append(progress);
  app.append(status);
  const card = el('section', 'card scenario-card');
  const copy = el('div', 'question-copy');
  copy.append(el('p', 'life-domain', '生活場面：' + LIFE_DOMAIN_LABELS[question.lifeDomain]));
  const title = el('h1', '', question.text);
  title.tabIndex = -1;
  copy.append(title);
  card.append(copy);
  const optionCard = el('article', 'motive-option');
  if (isRatingValue(state.answers[question.id])) optionCard.classList.add('answered');
  optionCard.append(renderRatingActions(question, state.answers[question.id]));
  card.append(optionCard);
  app.append(card);
  // 操作は一つの行に収め、初問でも「次へ」の位置を変えません。
  const actions = el('div', 'actions');
  if (state.currentIndex > 0) {
    actions.append(button('戻る', 'secondary', function () {
      state.currentIndex -= 1;
      saveState();
      showQuestion(true);
    }));
  } else actions.append(el('span', 'back-placeholder'));
  const defer = button('保留', 'secondary', deferQuestion);
  defer.setAttribute('aria-label', '保留して次の未回答へ');
  actions.append(defer);
  const complete = answeredCount() === order.length;
  const next = button(complete ? '結果を更新' : (state.currentIndex === order.length - 1 ? '結果を見る' : '次へ'), 'primary', function () {
    if (answeredCount() === order.length) showResults(calculateDiagnosis());
    else proceedQuestion();
  });
  if (complete) next.setAttribute('aria-label', '修正した回答で結果を更新');
  actions.append(next);
  app.append(actions);
  if (shouldScrollTop) {
    scrollPageTop();
    requestAnimationFrame(function () { title.focus({ preventScroll:true }); });
  }
}

// 保留は0点にせず、未回答のまま循環して再確認します。
function deferQuestion() {
  const order = selectedOrder(state.mode);
  for (let step = 1; step < order.length; step += 1) {
    const index = (state.currentIndex + step) % order.length;
    if (!isRatingValue(state.answers[order[index]])) {
      state.currentIndex = index;
      saveState();
      showQuestion(true);
      return;
    }
  }
  if (answeredCount() === order.length) showResults(calculateDiagnosis());
  else showQuestion(false, '未回答はこの質問だけです。思い浮かぶ具体例を振り返ってから選んでください。');
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

function determineWing(coreType, scores) {
  const leftType = coreType === 1 ? 9 : coreType - 1;
  const rightType = coreType === 9 ? 1 : coreType + 1;
  const leftScore = scores[leftType];
  const rightScore = scores[rightType];
  const balanced = leftScore === rightScore;
  const inconclusive = Math.abs(leftScore - rightScore) <= 1;
  const wingType = inconclusive ? null : (leftScore > rightScore ? leftType : rightType);
  return {
    coreType:coreType,
    wingType:wingType,
    wingLabel:inconclusive ? coreType + (balanced ? '（左右同点・保留）' : '（左右1点差・保留）') : coreType + 'w' + wingType,
    wingAType:leftType,
    wingBType:rightType,
    wingAScore:leftScore,
    wingBScore:rightScore,
    balanced:balanced
  };
}

function calculateDiagnosis() {
  if (!isMode(state.mode)) throw new Error('診断版を選択してください。');
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
  // 3番目の得点に並ぶタイプをすべて含め、番号順による候補の脱落を防ぎます。
  const flatProfile = new Set(Object.values(scores)).size === 1;
  const cutoff = ranking[2].score;
  const topCandidates = flatProfile ? [] : ranking.filter(function (item) { return item.score > 0 && item.score >= cutoff; });
  const topScore = ranking[0].score;
  const exactTopTypes = ranking.filter(function (item) { return item.score === topScore; }).map(function (item) { return item.typeId; });
  return {
    mode:state.mode,
    answers:Object.assign({}, state.answers),
    scores:scores,
    facetScores:facetScores,
    ranking:ranking,
    topCandidates:topCandidates,
    exactTopTypes:exactTopTypes,
    crossScores:crossScores,
    domainStability:calculateDomainStability(scores),
    flatProfile:flatProfile,
    sensitivity:calculateSensitivity(scores),
    uniformAnswers:new Set(TYPE_QUESTIONS.filter(function (q) { return q.version === state.mode; }).map(function (q) { return state.answers[q.id]; })).size === 1
  };
}

// 各タイプで同数の質問を外し、残る3領域の最高点集合を比較します。
// 生活場面と質問内容の影響は分離できず、正答確率を表すものではありません。
function calculateDomainStability(scores) {
  const leaders = function (values) {
    if (new Set(Object.values(values)).size === 1) return [];
    const maximum = Math.max.apply(null, Object.values(values));
    return Object.keys(values).map(Number).filter(function (id) { return values[id] === maximum; });
  };
  const original = leaders(scores);
  const checks = Object.keys(LIFE_DOMAIN_LABELS).map(function (domain) {
    const remaining = Object.assign({}, scores);
    TYPE_QUESTIONS.filter(function (q) { return q.version === state.mode && q.lifeDomain === domain; }).forEach(function (q) {
      remaining[q.typeId] -= RATING_SCORE_MAP[state.answers[q.id]];
    });
    const topTypes = leaders(remaining);
    return { excludedDomain:domain, scores:remaining, topTypes:topTypes,
      sameTop:original.length > 0 && original.join(',') === topTypes.join(','),
      maxScore:MODE_CONFIG[state.mode].maxTypeScore * 3 / 4 };
  });
  return { checks:checks, stable:checks.every(function (check) { return check.sameTop; }) };
}

function renderDomainStability(result) {
  const section = el('details', 'card domain-stability');
  section.append(el('summary', '', '生活場面による違いを確認する'));
  section.append(el('p', '', '一つの生活場面の回答を外し、残る3場面で最高点の候補を比べます。質問内容も同時に変わるため、違いの原因を場面だけに特定することはできません。'));
  result.domainStability.checks.forEach(function (check) {
    section.append(el('p', '', LIFE_DOMAIN_LABELS[check.excludedDomain] + 'を除く：' +
      (check.topTypes.length ? 'タイプ' + check.topTypes.join('・タイプ') : '全タイプ同点・保留') +
      '（各タイプ' + check.maxScore + '点満点）'));
  });
  section.append(el('p', 'chart-help', '違いがあっても誤回答とは限りません。具体的な場面を振り返る手がかりです。'));
  return section;
}

// 1つの回答を1段階だけ変えた場合の最高点集合を実際に再計算します。
// これは採点の感度であり、統計的な信頼度や正答確率ではありません。
function calculateSensitivity(scores) {
  const topSet = function (values) {
    const max = Math.max.apply(null, Object.values(values));
    return Object.keys(values).map(Number).filter(function (id) { return values[id] === max; });
  };
  const original = topSet(scores).join(',');
  const alternatives = new Set();
  let changeable = false;
  TYPE_QUESTIONS.filter(function (q) { return q.version === state.mode; }).forEach(function (q) {
    const value = RATING_SCORE_MAP[state.answers[q.id]];
    [-1, 1].forEach(function (delta) {
      if (value + delta < 0 || value + delta > 4) return;
      const changed = Object.assign({}, scores);
      changed[q.typeId] += delta;
      const leaders = topSet(changed);
      if (leaders.join(',') !== original) {
        changeable = true;
        leaders.forEach(function (id) { alternatives.add(id); });
      }
    });
  });
  return { changeable:changeable, types:Array.from(alternatives).sort(function (a, b) { return a - b; }) };
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

function renderCandidate(result, item) {
  const type = TYPE_RESULTS[item.typeId];
  const config = MODE_CONFIG[result.mode];
  const card = el('section', 'card candidate-card');
  card.append(el('p', 'scenario-number', '得点 ' + item.rank + '位' + (result.ranking.filter(function (other) { return other.score === item.score; }).length > 1 ? '（同点）' : '')));
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
  const section = el('details', 'card auxiliary-reflection');
  section.append(el('summary', '', '補助6問の振り返り'));
  section.append(el('p', '', 'この6問は、気になりやすいことや行動の取り方を振り返る項目です。組み合わせからタイプを推定せず、タイプ点にも加えません。'));
  CROSS_QUESTIONS.filter(function (q) { return q.version === result.mode; }).forEach(function (q) {
    section.append(el('h3', '', q.label));
    section.append(el('p', '', q.text));
    const answer = RATING_OPTIONS.find(function (option) { return option.value === result.answers[q.id]; });
    section.append(el('p', 'chart-help', answer.label + '（' + answer.score + '/4）'));
  });
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
  body.append(el('p', 'chart-help', '診断中は伏せていたタイプ対応と点数を、ここで確認できます。補助項目はタイプ点に加えていません。'));
  const list = el('ol');
  selectedOrder(result.mode).forEach(function (id) {
    const question = ALL_QUESTIONS.get(id);
    const rating = ratingOption(result.answers[id]);
    const target = question.typeId ? 'タイプ' + question.typeId : '補助項目';
    const row = el('li', '', target + '・' + rating.label + '（' + rating.score + '点）— ' + question.text);
    row.append(button('この回答を見直す', 'secondary', function () {
      state.currentIndex = selectedOrder(state.mode).indexOf(id);
      saveState();
      showQuestion(true);
    }));
    list.append(row);
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
  body.append(el('p', 'chart-help', '以下はタイプ理論上の傾向の例です。すべてが本人に当てはまると決めず、具体例と照合してください。'));
  body.append(paragraphSection('注意が向きやすいところ', type.warningSignal));
  body.append(paragraphSection('強みとして生かせるところ', type.essence));
  body.append(paragraphSection('行き過ぎたときに起こりうること', type.tragedy, 'runaway'));
  body.append(paragraphSection('助けになる関わり方', type.healing, 'healing'));
  details.append(body);
  return details;
}

function renderAllTypeDetails(result) {
  const section = el('section', 'type-details');
  section.append(el('h2', '', '各タイプの特徴'));
  section.append(el('p', 'chart-help', '上位候補に限らず、気になるタイプを読み比べられます。'));
  const openTypes = new Set(result.topCandidates.map(function (item) { return item.typeId; }));
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

// 原票の文章・採点表は移植せず、既存の動機定義から独自に作成した比較用の原稿です。
// 各列は「満たしたいこと」「予定外への初動」「手放しにくいこと」をそろえています。
const COMPARISON_PROMPTS = [
  '長い目で見ると、どちらを満たしたときに安心しやすいですか？',
  '思いどおりに進まないとき、どちらへ心が向きやすいですか？',
  '自分の選択を振り返ると、どちらを手放しにくいですか？'
];
const COMPARISON_MOTIVES = {
  1:['自分で納得できる筋道に沿って選べていること。', 'どこを直せば納得できる状態に戻るかを考える。', '気になる点を残したまま終えず、自分の納得する形まで整えること。'],
  2:['自分の働きかけが身近な人に届き、必要とされていること。', '相手が何を望んでいるかを探り、こちらから働きかける。', '自分から関わることで、大切な人との結びつきを保つこと。'],
  3:['自分の工夫や努力が形になり、価値を認められること。', '今できる方法に切り替え、うまくできる自分を取り戻す。', '自分にはできると感じられ、人にも伝わる形を残すこと。'],
  4:['自分の感じ方を置き去りにせず、自分らしくいられること。', '自分にとって何が失われたように感じるのかを見つめる。', '周りに合わせるだけで終えず、自分にしかない意味を大切にすること。'],
  5:['自分で理解でき、必要な力を蓄えられていること。', 'いったん距離と時間を取り、状況を理解してから関わる。', '十分に理解するための時間と、自分で使える余力を確保すること。'],
  6:['不安な点を確かめ、頼れる支えがあると感じられること。', '見落としや別の可能性を確認し、信頼できる手掛かりを探す。', '自分だけで決めつけず、確かめられる根拠や信頼関係を持つこと。'],
  7:['この先に楽しみや選べる道があり、気持ちが開けること。', 'ほかに試せる道や、気持ちを前に向けられる可能性を探す。', '一つのつらい状態に閉じ込められず、次の可能性を残すこと。'],
  8:['自分の大切な領域を守り、決める力を持てていること。', '自分で状況に働きかけ、決める力を取り戻そうとする。', '大切な選択を人任せにせず、自分の意思で引き受けること。'],
  9:['周りとのつながりが穏やかで、自分のペースも保てること。', 'まず波立ちを鎮め、無理なく続けられる落ち着きどころを探す。', '自分の主張で関係を揺らすより、つながりと心の落ち着きを保つこと。']
};

function renderResultGuidance(result) {
  const section = el('section', 'card');
  section.append(el('h2', '', '結果の読み方と、次の確認'));
  const status = result.flatProfile ? '全タイプ同点のため保留' :
    result.exactTopTypes.length > 1 ? '最高点が同点の候補を比較' :
    result.sensitivity.changeable ? '僅差のため候補を比較' :
    !result.domainStability.stable ? '生活場面・質問による違いを確認' : '上位候補の動機を比較';
  section.append(el('p', 'notice result-status', status));
  if (!result.flatProfile) section.append(el('p', 'chart-help', result.domainStability.stable ?
    'どの生活場面を一つ外しても最高点の候補は同じでした。正しさの保証ではありません。' :
    '生活場面を一つ外すと最高点の候補が変わる場合があります。下の「生活場面による違い」で確認できます。'));
  if (!result.flatProfile) {
    const gap = result.ranking[0].score - result.ranking[1].score;
    section.append(el('p', '', '最高点と次の得点の差：' + gap + '点。点数は当てはまりの合計で、確率や能力の高さではありません。'));
    section.append(el('p', 'chart-help', result.sensitivity.changeable ?
      'タイプ質問の回答を1つ、1段階だけ変えると最高点の候補が変わるか同点になります。タイプ' + result.sensitivity.types.join('・タイプ') + 'の違いを丁寧に確認してください。' :
      'タイプ質問の回答を1つ、1段階だけ変えても最高点の候補は変わりません。ただし、診断が正しいことを保証する検証ではありません。'));
  }
  section.append(el('p', '', '得点と合わない候補も比較できます。具体的な出来事を一つ思い出し、「その行動で何を守りたかったか」を確かめてください。決めきれなければ保留で構いません。'));
  section.append(el('p', 'privacy', '自己理解のための独自質問です。心理測定上の精度・再検査信頼性は未検証です。採用・配置・人事評価の判断には使いません。'));
  return section;
}

function summarizeComparison(typeA, typeB, choices) {
  if (!Number.isInteger(typeA) || !Number.isInteger(typeB) || typeA === typeB ||
      !COMPARISON_MOTIVES[typeA] || !COMPARISON_MOTIVES[typeB]) throw new Error('異なる2タイプを選んでください。');
  if (!Array.isArray(choices) || choices.length !== 3 || choices.some(function (value) {
    return ![null, 'a', 'b', 'hold'].includes(value);
  })) throw new Error('比較回答が不正です。');
  const count = function (value) { return choices.filter(function (choice) { return choice === value; }).length; };
  return { complete:!choices.includes(null), a:count('a'), b:count('b'), hold:count('hold') };
}

function renderPairComparison(result) {
  const section = el('section', 'card pair-comparison');
  section.append(el('h2', '', '2タイプの動機を比べる（任意）'));
  section.append(el('p', '', '9タイプから気になる2つを選び、3つの観点で「長年の自分により近い方」を比べます。両方に当てはまるときも優先しやすい方を考え、選べなければ保留にしてください。'));
  section.append(el('p', 'chart-help', 'TK式の二者比較・同点確認の考え方を参考にした独自の振り返りです。正式なTK式診断ではありません。元の得点や順位は変更しません。比較の回答はこの画面を離れると消えます。'));
  const controls = el('div', 'pair-select');
  const selectors = ['A', 'B'].map(function (side) {
    const label = el('label', '', '比較するタイプ ' + side);
    const select = el('select');
    select.setAttribute('aria-label', '比較するタイプ ' + side);
    const placeholder = el('option', '', 'タイプを選択');
    placeholder.value = '';
    select.append(placeholder);
    for (let id = 1; id <= 9; id += 1) {
      const option = el('option', '', 'タイプ' + id + '：' + TYPE_RESULTS[id].nameJa);
      option.value = String(id);
      select.append(option);
    }
    label.append(select);
    controls.append(label);
    return select;
  });
  section.append(controls);
  const body = el('div');
  section.append(body);
  function updatePair() {
    body.replaceChildren();
    const a = Number(selectors[0].value);
    const b = Number(selectors[1].value);
    if (!a || !b) return;
    if (a === b) {
      body.append(showError('異なる2タイプを選んでください。'));
      return;
    }
    const choices = [null, null, null];
    const output = el('p', 'comparison-output', '3つの観点に答えると比較内容をまとめます。');
    output.setAttribute('role', 'status');
    COMPARISON_PROMPTS.forEach(function (prompt, index) {
      const field = el('fieldset', 'pair-question');
      field.append(el('legend', '', (index + 1) + '. ' + prompt));
      const actions = el('div', 'pair-options');
      const nodes = [];
      [['a', 'A：' + COMPARISON_MOTIVES[a][index]], ['b', 'B：' + COMPARISON_MOTIVES[b][index]], ['hold', '保留（選べない・どちらも違う）']].forEach(function (entry) {
        const node = button(entry[1], 'secondary', function () {
          choices[index] = entry[0];
          nodes.forEach(function (item) { item.setAttribute('aria-pressed', String(item === node)); });
          const summary = summarizeComparison(a, b, choices);
          output.textContent = summary.complete ?
            '今回選んだ動機：タイプ' + a + 'が' + summary.a + '件、タイプ' + b + 'が' + summary.b + '件、保留が' + summary.hold + '件。これは3観点の振り返りで、診断の確定や順位の決着ではありません。選んだ理由が説明できる実体験と、当てはまらない例も思い出してください。' :
            '確認済み ' + choices.filter(function (value) { return value !== null; }).length + ' / 3。すべて確認してから読み比べます。';
        });
        node.setAttribute('aria-pressed', 'false');
        nodes.push(node);
        actions.append(node);
      });
      field.append(actions);
      body.append(field);
    });
    body.append(output);
  }
  selectors.forEach(function (select) { select.addEventListener('change', updatePair); });
  return section;
}

function showResults(result) {
  clearApp();
  app.append(el('p', 'scenario-number', MODE_CONFIG[result.mode].name));
  app.append(el('h1', '', result.flatProfile ? '今回はタイプを絞り込めません' : '比較して確かめるタイプ候補'));
  app.append(el('p', 'lead', result.flatProfile ? '9タイプすべてが同点です。番号順で候補を選ばず、判定を保留します。' : '得点があるタイプの上位3件を目安に、3件目と同点のタイプもすべて表示します。同点の表示順はタイプ番号順です。'));
  if (!result.flatProfile && result.exactTopTypes.length > 1) {
    app.append(el('div', 'notice', '最高得点が同点です：タイプ' + result.exactTopTypes.join('・タイプ') + '。1つへ絞らず、動機を読み比べてください。'));
  }
  if (result.uniformAnswers) {
    app.append(el('div', 'notice', 'タイプ質問すべてに同じ段階で回答しています。タイプ間の違いが出にくいため、結果は判定保留として読み、必要なら回答を見直してください。'));
  }
  app.append(renderResultGuidance(result));
  app.append(renderPairComparison(result));
  result.topCandidates.forEach(function (item, index) { app.append(renderCandidate(result, item, index + 1)); });
  app.append(renderDomainStability(result));
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
    LEGACY_STORAGE_KEYS.forEach(function (key) {
      if (sessionStorage.getItem(key) !== null) previousQuestionsCleared = true;
      sessionStorage.removeItem(key);
    });
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

