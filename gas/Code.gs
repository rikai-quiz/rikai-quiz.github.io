// りかいクイズ用のサーバー(Googleスプレッドシートに紐づけて使う)
// 初回だけ setup() を実行すると、必要なシートと最初の20問が作られる

const SHEET = {
  questions: "questions",
  quizzes: "quizzes",
  results: "results",
  suggestions: "suggestions",
};

const HEADERS = {
  questions: ["id", "question", "a", "b", "c", "active", "source"],
  quizzes: ["id", "name", "questions", "answers", "created_at", "key"],
  results: ["quiz_id", "name", "score", "total", "created_at", "picks"],
  suggestions: ["created_at", "question", "a", "b", "c", "by", "status"],
};

// ---- 入口 ----

function doGet(e) {
  return respond(e.parameter || {});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return respond(body);
}

function respond(p) {
  let out;
  try {
    switch (p.action) {
      case "questions": out = { questions: activeQuestions() }; break;
      case "create": out = createQuiz(p); break;
      case "quiz": out = getQuiz(p.id); break;
      case "submit": out = submit(p); break;
      case "ranking": out = rankingFor(p); break;
      case "suggest": out = suggest(p); break;
      case "updateAnswers": out = updateAnswers(p); break;
      default: throw new Error("unknown action");
    }
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---- 処理 ----

function activeQuestions() {
  return rows(SHEET.questions)
    .filter(r => r.active === true && r.question)
    .map(r => ({ id: String(r.id), q: r.question, a: [r.a, r.b, r.c] }));
}

function createQuiz(p) {
  const name = text(p.name, 20);
  if (!name) throw new Error("名前を入力してください");
  const byId = new Map(activeQuestions().map(q => [q.id, q]));
  const qids = Array.isArray(p.qids) ? p.qids.map(String) : [];
  const answers = Array.isArray(p.answers) ? p.answers : [];
  if (!qids.length || qids.length !== answers.length) throw new Error("回答の数が合いません");
  const snapshot = qids.map(id => {
    const q = byId.get(id);
    if (!q) throw new Error("質問が更新されました。最初からやり直してください");
    return q;
  });
  if (answers.some(a => ![0, 1, 2].includes(a))) throw new Error("回答が不正です");

  // 質問はクイズ作成時点の内容を保存する(あとでシートの質問を変えても、作成済みのクイズは崩れない)
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const id = newId();
    // key は作成者だけが持つ合言葉。これがあれば挑戦しなくても全員の回答を見られる
    const key = Utilities.getUuid().replace(/-/g, "");
    sheet(SHEET.quizzes).appendRow([id, name, JSON.stringify(snapshot), "'" + answers.join(""), new Date(), key]);
    return { id, key };
  } finally {
    lock.releaseLock();
  }
}

function getQuiz(id) {
  const quiz = findQuiz(id);
  // 答えは返さない(採点はサーバー側で行う)
  return { id: quiz.id, name: quiz.name, questions: quiz.questions };
}

function submit(p) {
  const quiz = findQuiz(p.id);
  const name = text(p.name, 20);
  if (!name) throw new Error("名前を入力してください");
  const picks = Array.isArray(p.picks) ? p.picks : [];
  if (picks.length !== quiz.answers.length) throw new Error("回答の数が合いません");
  const score = picks.filter((v, i) => v === quiz.answers[i]).length;
  sheet(SHEET.results).appendRow([quiz.id, name, score, quiz.answers.length, new Date(), "'" + picks.join("")]);
  return { score, total: quiz.answers.length, answers: quiz.answers, ranking: ranking(quiz.id), people: people(quiz.id) };
}

// 作成者が自分の答えを直す。すでに挑戦した人のスコアも新しい答えで計算し直す
function updateAnswers(p) {
  const quiz = findQuiz(p.id);
  if (!p.key || String(p.key) !== quiz.key) throw new Error("作成者だけが答えを直せます");
  const answers = Array.isArray(p.answers) ? p.answers : [];
  if (answers.length !== quiz.answers.length || answers.some(a => ![0, 1, 2].includes(a))) throw new Error("回答が不正です");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const qs = sheet(SHEET.quizzes);
    const qv = qs.getDataRange().getValues();
    const qrow = qv.findIndex((r, i) => i > 0 && String(r[0]) === quiz.id);
    qs.getRange(qrow + 1, HEADERS.quizzes.indexOf("answers") + 1).setValue("'" + answers.join(""));

    const rs = sheet(SHEET.results);
    const rv = rs.getDataRange().getValues();
    const col = name => rv[0].indexOf(name);
    rv.forEach((r, i) => {
      if (i === 0 || String(r[col("quiz_id")]) !== quiz.id) return;
      const picks = String(r[col("picks")]).split("").map(Number);
      const score = picks.filter((v, k) => v === answers[k]).length;
      rs.getRange(i + 1, col("score") + 1).setValue(score);
    });
  } finally {
    lock.releaseLock();
  }
  return rankingFor(p);
}

// 挑戦前の人にはスコアだけ。作成者(key一致)には答えと全員の回答も返す
function rankingFor(p) {
  const quiz = findQuiz(p.id);
  const out = { name: quiz.name, ranking: ranking(quiz.id) };
  if (p.key && String(p.key) === quiz.key) {
    Object.assign(out, { owner: true, questions: quiz.questions, answers: quiz.answers, people: people(quiz.id) });
  }
  return out;
}

function quizResults(id) {
  return rows(SHEET.results)
    .filter(r => String(r.quiz_id) === String(id))
    .sort((x, y) => y.score - x.score || new Date(x.created_at) - new Date(y.created_at));
}

function ranking(id) {
  return quizResults(id).slice(0, 30).map(r => ({ name: r.name, score: r.score, total: r.total }));
}

// 一人ひとりが各問題でどれを選んだか
function people(id) {
  return quizResults(id).map(r => ({ name: r.name, picks: String(r.picks).split("").map(Number) }));
}

function suggest(p) {
  const question = text(p.question, 60);
  const choices = (Array.isArray(p.choices) ? p.choices : []).map(c => text(c, 30));
  if (!question || choices.length !== 3 || choices.some(c => !c)) throw new Error("質問と3つの選択肢を入力してください");
  sheet(SHEET.suggestions).appendRow([new Date(), question, ...choices, text(p.by, 20), "未確認"]);
  return { ok: true };
}

// ---- 補助 ----
// answers / picks は "0210…" のような数字の並び。先頭の0が消えないよう ' を付けて文字列として保存している

function findQuiz(id) {
  const r = rows(SHEET.quizzes).find(r => String(r.id) === String(id || ""));
  if (!r) throw new Error("クイズが見つかりません");
  return { id: String(r.id), name: r.name, questions: JSON.parse(r.questions), answers: String(r.answers).split("").map(Number), key: String(r.key) };
}

function newId() {
  const chars = "abcdefghjkmnpqrstuvwxyz23456789";
  const used = new Set(rows(SHEET.quizzes).map(r => String(r.id)));
  let id;
  do {
    id = Array.from({ length: 6 }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
  } while (used.has(id));
  return id;
}

// 文字数を制限し、先頭が = + - @ の入力は数式として実行されないよう文字列扱いにする
function text(v, max) {
  const s = String(v == null ? "" : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function sheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rows(name) {
  const values = sheet(name).getDataRange().getValues();
  const head = values.shift();
  return values.map(r => Object.fromEntries(head.map((k, i) => [k, r[i]])));
}

// ---- 初期設定(最初に1回だけ実行) ----

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET).forEach(key => {
    let sh = ss.getSheetByName(SHEET[key]);
    if (!sh) sh = ss.insertSheet(SHEET[key]);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[key]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[key].length).setFontWeight("bold");
    }
  });
  const qs = ss.getSheetByName(SHEET.questions);
  if (qs.getLastRow() === 1) {
    SEED.forEach((s, i) => qs.appendRow([i + 1, s[0], s[1], s[2], s[3], true, s[4]]));
    qs.getRange(2, 6, SEED.length, 1).insertCheckboxes();
  }
}

const SEED = [
  ["1億円もらえるが、一生「好きな食べ方」を1つ禁止される。どうする？", "迷わず受け取る", "禁止される内容しだい", "好きなものは譲れないので断る", "https://x.com/4309qwerty/status/2104822498147876894"],
  ["恋愛では？", "追いかけたい", "追われたい", "対等がいい", "https://x.com/fumiya0307/status/2036068688068317201"],
  ["身内でデスゲームが始まったら、自分はどうなる？", "最後まで生き残る", "序盤で脱落する", "警戒されて最初に狙われる", "https://x.com/yuukamiya68/status/1059270960241500160"],
  ["恋人との理想の連絡頻度は？", "毎日おはよう・おやすみ", "数日に1回くらい", "用事があるときだけ", "https://x.com/kadena_love0/status/1976047560990347367"],
  ["友だちが作ったものを見せられた。正直いいと思えなかったら？", "正直に感想を言う", "とりあえず褒める", "良いところだけ探して伝える", "https://x.com/Rower1216/status/2101507160245199096"],
  ["どこからが浮気？", "二人きりで食事に行ったら", "気持ちが傾いた時点で", "体の関係を持ったら", "https://x.com/Happy_Pokonyan1/status/1981846592635503075"],
  ["「明日、私がコップ一杯の水になったら、その水をどうする？」と言われたら？", "飲む", "ずっと飾っておく", "花にあげる", "https://x.com/hukaiyoitaiyo/status/1542460319968407558"],
  ["男女の友情は成立する？", "成立する", "場合による", "成立しない", "https://x.com/piyo_hiyokoman/status/1571712684416176129"],
  ["「5年後どうなっていたい？」と聞かれたら？", "具体的な目標を答えられる", "ぼんやりとならある", "今で精一杯で考えていない", "https://x.com/Shioko5/status/1780013438007435572"],
  ["夕方に送ったLINEの返信が翌日だったら？", "全然気にしない", "少しモヤっとする", "かなり無理", "https://x.com/___tou_latte_/status/2096803000564215955"],
  ["恋人に「なんで怒ってるかわかる？」と言われたら？", "正直に「教えて」と聞く", "思い当たることを謝る", "落ち着くまで待つ", "https://x.com/Motosuzukisan/status/964484802039574528"],
  ["恋人がいても、1人の時間はどれくらい必要？", "ほとんどいらない", "週に何日かは欲しい", "毎日必ず欲しい", "https://x.com/wlbisimportant/status/2045744404833718498"],
  ["恋人と価値観が違ったときは？", "納得するまで話し合う", "相手の考えを受け入れる", "深入りせずそっとしておく", "https://x.com/radran10/status/2103108125624840289"],
  ["失恋した異性の親友に「今日だけ泊めて」と頼まれたら？", "泊める", "話は聞くけど泊めない", "断る", "https://x.com/Ringoxlove/status/2111668157652046069"],
  ["恋人と会う理想の頻度は？", "週2回以上", "週1回くらい", "月に数回でも平気", "https://x.com/wlbisimportant/status/2045744404833718498"],
  ["倦怠期で、好きかわからなくなったら？", "乗り越える努力をする", "少し距離を置いて様子を見る", "気持ちに正直に別れる", "https://x.com/___tou_latte_/status/2112744956318401007"],
  ["恋人の過去の恋愛について知りたい？", "全部知りたい", "別れた理由だけ知りたい", "知りたくない", "https://x.com/wlbisimportant/status/2045744404833718498"],
  ["理想の出会い方は？", "自然に出会って少しずつ", "友だちの紹介", "マッチングアプリでも全然いい", "https://x.com/yagi_mt/status/1793326596981772601"],
  ["記念日はどう過ごしたい？", "しっかりお祝いしたい", "ささやかでいい", "特に気にしない", "https://x.com/kadena_love0/status/1976047560990347367"],
  ["恋人があまり連絡をくれないタイプだったら？", "愛情が足りないと感じる", "そういう性格だと割り切る", "むしろ楽", "https://x.com/Tatarianaster_7/status/2012484681560916104"],
];
