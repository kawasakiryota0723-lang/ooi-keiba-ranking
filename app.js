const fileInput = document.getElementById("fileInput");
const raceList = document.getElementById("raceList");
const message = document.getElementById("message");
const summary = document.getElementById("summary");
const filters = document.getElementById("filters");
const dateChip = document.getElementById("dateChip");
const candidateCount = document.getElementById("candidateCount");
const raceCount = document.getElementById("raceCount");
const template = document.getElementById("raceTemplate");

let races = [];
let activeFilter = "all";
let currentTargetDate = "";

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  setMessage("読み込み中…", `${file.name}を確認しています。`);

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: false });
    const venue = text(workbook.Sheets["自動計算"]?.B3?.v || workbook.Sheets["当日ランキング"]?.B3?.v);
    if (venue !== "大井") throw new Error("大井用の集計Excelを選んでください。競馬場が大井になっているか確認してください。");
    const sheet = workbook.Sheets["当日全レース"];

    if (!sheet) {
      throw new Error("「当日全レース」シートが見つかりません。Excelでシート名を確認してください。");
    }

    const values = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      range: "A1:G16",
      raw: true,
      defval: "",
      blankrows: true,
    });

    const parsed = [];
    for (let index = 4; index <= 15; index += 1) {
      const row = values[index] || [];
      const raceNumber = normalizeRace(row[0], index - 3);
      const horseName = text(row[2]);
      const decision = text(row[6]) || (horseName ? "判定不可" : "対象レースなし");

      parsed.push({
        raceNumber,
        horseNumber: text(row[1]),
        horseName,
        popularity: displayNumber(row[3], "—"),
        score: displayNumber(row[4], "—", 1),
        gap: displayNumber(row[5], "—", 1),
        decision,
        ...readRaceDetails(workbook, Number(sheet.B2?.v), raceNumber, row),
      });
    }

    races = parsed;
    const targetDate = formatTargetDate(sheet.B2?.v);
    try { localStorage.setItem("ooi-keiba-mobile-data-v2", JSON.stringify({ targetDate, races })); } catch {}
    showData(targetDate);
  } catch (error) {
    races = [];
    try { localStorage.removeItem("ooi-keiba-mobile-data-v2"); } catch {}
    raceList.replaceChildren();
    summary.hidden = true;
    filters.hidden = true;
    dateChip.textContent = "読込エラー";
    setMessage("読み込めませんでした", error.message || "Excelファイルを確認してください。", true);
  } finally {
    fileInput.value = "";
  }
});

filters.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-filter]");
  if (!button) return;
  activeFilter = button.dataset.filter;
  document.querySelectorAll(".filter").forEach((item) => item.classList.toggle("active", item === button));
  renderRaces();
});

function showData(targetDate) {
  currentTargetDate = targetDate || "";
  dateChip.textContent = targetDate || "日付不明";
  candidateCount.textContent = races.filter((race) => race.decision === "購入候補").length;
  raceCount.textContent = races.filter((race) => race.horseName).length;
  summary.hidden = false;
  filters.hidden = false;
  message.hidden = true;
  activeFilter = "all";
  document.querySelectorAll(".filter").forEach((item) => item.classList.toggle("active", item.dataset.filter === "all"));
  renderRaces();
}

function renderRaces() {
  const filtered = races.filter((race) => {
    if (activeFilter === "candidate") return race.decision === "購入候補";
    if (activeFilter === "other") return race.decision !== "購入候補";
    return true;
  });

  raceList.replaceChildren();
  if (!filtered.length) {
    setMessage("該当レースはありません", "表示条件を変更してください。");
    return;
  }

  message.hidden = true;
  for (const race of filtered) {
    const node = template.content.cloneNode(true);
    const card = node.querySelector(".race-card");
    const isCandidate = race.decision === "購入候補";
    const isUnavailable = race.decision === "判定不可" || race.decision === "対象レースなし";
    card.classList.toggle("candidate", isCandidate);
    card.classList.toggle("unavailable", isUnavailable);
    node.querySelector(".race-number").textContent = `${race.raceNumber}R`;
    node.querySelector(".horse-name").textContent = race.horseName || "データなし";
    node.querySelector(".horse-meta").textContent = race.horseNumber ? `馬番 ${race.horseNumber}` : "";
    node.querySelector(".decision").textContent = race.decision;
    node.querySelector(".score").textContent = race.score;
    node.querySelector(".gap").textContent = race.gap;
    node.querySelector(".popularity").textContent = race.popularity === "—" ? "—" : `${race.popularity}番`;
    node.querySelector(".race-toggle").setAttribute("aria-label", `${race.raceNumber}Rの全馬ランキング`);
    renderDetails(node.querySelector(".race-details"), race);
    const hint = node.querySelector(".detail-hint");
    card.addEventListener("toggle", () => { hint.textContent = card.open ? "ランキングを閉じる" : "全馬のランキングを見る"; });
    raceList.appendChild(node);
  }
}

function readRaceDetails(workbook, date, raceNumber, overview) {
  const detail = OoiRanking.getDetails(workbook, date, raceNumber);
  if (detail.horses.length) {
    const top = detail.horses[0];
    if (text(top.horseNumber) !== text(overview[1]) || text(top.horseName) !== text(overview[2]) || Math.abs(top.score - Number(overview[4])) > 0.01) {
      return { ...detail, warning: "Excel一覧と再計算点に差があります。全馬順位はスマホ側の再計算結果です。" };
    }
  }
  return detail;
}

function renderDetails(container, race) {
  const horses = race.horses || [];
  if (!horses.length) {
    const note = document.createElement("p");
    note.className = "detail-note";
    note.textContent = race.error || "全馬のランキングを表示するには、Excelをもう一度選択してください。";
    container.appendChild(note);
    return;
  }
  if (race.warning) {
    const warning = document.createElement("p");
    warning.className = "detail-note";
    warning.textContent = race.warning;
    container.appendChild(warning);
  }
  const title = document.createElement("h2");
  title.textContent = `${race.raceNumber}R 全馬ランキング（${horses.length}頭）`;
  container.appendChild(title);
  const table = document.createElement("table");
  table.className = "horse-table";
  const head = document.createElement("thead"), header = document.createElement("tr");
  for (const label of ["順位", "馬番", "馬名・騎手", "能力点"]) {
    const th = document.createElement("th"); th.scope = "col"; th.textContent = label; header.appendChild(th);
  }
  head.appendChild(header); table.appendChild(head);
  const body = document.createElement("tbody");
  for (const horse of horses) {
    const row = document.createElement("tr");
    if (horse.rank <= 3) row.className = `rank-${horse.rank}`;
    const rank = document.createElement("td"); rank.textContent = horse.rank; row.appendChild(rank);
    const number = document.createElement("td"); number.textContent = horse.horseNumber; row.appendChild(number);
    const info = document.createElement("td");
    const name = document.createElement("strong"); name.textContent = `${horse.rating ? horse.rating + " " : ""}${horse.horseName}`; info.appendChild(name);
    const jockey = document.createElement("span"); jockey.className = "jockey"; jockey.textContent = horse.jockey || "騎手未登録"; info.appendChild(jockey); row.appendChild(info);
    const score = document.createElement("td"); score.textContent = displayNumber(horse.score, "—", 1); row.appendChild(score);
    body.appendChild(row);
  }
  table.appendChild(body); container.appendChild(table);
}

function setMessage(title, detail, isError = false) {
  message.hidden = false;
  message.classList.toggle("error", isError);
  message.innerHTML = `<div class="empty-mark">${isError ? "!" : "…"}</div><h2></h2><p></p>`;
  message.querySelector("h2").textContent = title;
  message.querySelector("p").textContent = detail;
}

function normalizeRace(value, fallback) {
  const match = String(value || "").match(/\d+/);
  return match ? Number(match[0]) : fallback;
}

function displayNumber(value, fallback, digits = 0) {
  if (value === "" || value === null || value === undefined) return fallback;
  const number = Number(value);
  if (!Number.isFinite(number)) return text(value) || fallback;
  return digits ? number.toFixed(digits).replace(/\.0$/, "") : String(number);
}

function formatTargetDate(value) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (digits.length === 8) {
    return `${digits.slice(0, 4)}/${digits.slice(4, 6)}/${digits.slice(6, 8)}`;
  }
  return text(value);
}

function text(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

// History is independent of the current-day cache. Excel imports never write it.
const HISTORY_PREFIX = "ooi-keiba-prediction-history-v1:";

function historyDateKey(value) {
  const match = String(value || "").trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (!match) throw new Error("対象日を確認できません。日付の入ったExcelを読み込んでください。");
  const [, year, month, day] = match;
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  if (date.getUTCFullYear() !== Number(year) || date.getUTCMonth() + 1 !== Number(month) || date.getUTCDate() !== Number(day)) {
    throw new Error("対象日が正しくありません。Excelの日付を確認してください。");
  }
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function savePredictionHistory() {
  const dateKey = historyDateKey(currentTargetDate);
  const populated = races.filter(race => race.horseName || race.horses?.length);
  if (!populated.length) throw new Error("保存する予想がありません。先にExcelを読み込んでください。");
  if (populated.some(race => !Array.isArray(race.horses) || !race.horses.length)) {
    throw new Error("全馬ランキングが不足しているレースがあります。全馬データを含むExcelを読み込み直してください。");
  }
  const key = HISTORY_PREFIX + dateKey;
  if (localStorage.getItem(key) !== null) {
    throw new Error(`${dateKey}は保存済みです。最初の予想を保持するため上書きしません。`);
  }
  const savedAt = new Date().toISOString();
  // Serialize all original fields, including overview gap/decision and every horse's
  // rank, number, name, jockey, score, rating and popularity. No recalculation.
  const snapshot = { schemaVersion: 1, venue: "大井", dateKey,
    targetDate: currentTargetDate, savedAt, races: JSON.parse(JSON.stringify(races)) };
  localStorage.setItem(key, JSON.stringify(snapshot));
  return snapshot;
}

function createHistoryUI() {
  const section = document.createElement("section");
  section.id = "ooi-prediction-history";
  section.style.cssText = "margin:16px 0;padding:16px;border:1px solid #aaa;border-radius:12px;";
  const controls = document.createElement("div");
  controls.style.cssText = "display:flex;gap:10px;flex-wrap:wrap";
  const save = document.createElement("button");
  save.type = "button";
  save.textContent = "今日の予想を保存";
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.textContent = "予想履歴";
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute("aria-controls", "ooi-history-panel");
  for (const button of [save, toggle]) {
    button.style.cssText = "padding:10px 14px;min-height:44px;cursor:pointer";
    controls.appendChild(button);
  }
  const note = document.createElement("p");
  note.textContent = "読込中の対象日の予想を固定保存します。同じ日は上書きしません。履歴はこのブラウザに保存され、ブラウザのデータ削除で消えます。";
  const status = document.createElement("p");
  status.setAttribute("role", "status");
  status.setAttribute("aria-live", "polite");
  const panel = document.createElement("div");
  panel.id = "ooi-history-panel";
  panel.hidden = true;
  section.append(controls, note, status, panel);
  raceList.before(section);

  function displaySnapshot(snapshot) {
    panel.replaceChildren();
    const back = document.createElement("button");
    back.type = "button";
    back.textContent = "履歴一覧に戻る";
    back.addEventListener("click", displayList);
    const title = document.createElement("h2");
    title.textContent = `${snapshot.targetDate} 大井・保存済み予想`;
    const stamp = document.createElement("p");
    stamp.textContent = `保存日時：${new Date(snapshot.savedAt).toLocaleString("ja-JP")}`;
    panel.append(back, title, stamp);
    for (const race of snapshot.races) {
      const details = document.createElement("details");
      details.style.cssText = "margin:12px 0;padding:10px;border:1px solid #aaa;border-radius:8px";
      const heading = document.createElement("summary");
      heading.style.cursor = "pointer";
      heading.textContent = `${race.raceNumber}R ${race.horseName || "データなし"} ／ ${race.decision || "—"}`;
      const overview = document.createElement("p");
      overview.textContent = `能力点：${race.score ?? "—"} ／ 能力差：${race.gap ?? "—"} ／ 人気：${race.popularity ?? "—"}`;
      const ranking = document.createElement("div");
      ranking.style.overflowX = "auto";
      renderDetails(ranking, race);
      // The original ranking table omits popularity; show the saved values here.
      const table = ranking.querySelector("table");
      if (table) {
        const th = document.createElement("th");
        th.scope = "col";
        th.textContent = "人気";
        table.querySelector("thead tr").appendChild(th);
        table.querySelectorAll("tbody tr").forEach((row, index) => {
          const td = document.createElement("td");
          td.textContent = race.horses[index].popularity ?? "—";
          row.appendChild(td);
        });
      }
      details.append(heading, overview, ranking);
      panel.appendChild(details);
    }
  }

  function displayList() {
    panel.replaceChildren();
    const title = document.createElement("h2");
    title.textContent = "予想履歴";
    panel.appendChild(title);
    try {
      const keys = [];
      for (let i = 0; i < localStorage.length; i += 1) {
        const key = localStorage.key(i);
        if (key?.startsWith(HISTORY_PREFIX)) keys.push(key);
      }
      keys.sort().reverse();
      if (!keys.length) {
        const empty = document.createElement("p");
        empty.textContent = "保存済みの予想はありません。";
        panel.appendChild(empty);
      }
      for (const key of keys) {
        const item = document.createElement("p");
        try {
          const snapshot = JSON.parse(localStorage.getItem(key));
          if (snapshot?.schemaVersion !== 1 || !Array.isArray(snapshot.races) || snapshot.races.some(race => !race || (race.horses != null && (!Array.isArray(race.horses) || race.horses.some(horse => !horse))))) {
            throw new Error("履歴形式が不正です");
          }
          const button = document.createElement("button");
          button.type = "button";
          button.style.cssText = "padding:10px;min-height:44px;cursor:pointer";
          const count = snapshot.races.filter(race => race.horseName || race.horses?.length).length;
          button.textContent = `${snapshot.targetDate} 大井 ${count}R分 — ${new Date(snapshot.savedAt).toLocaleString("ja-JP")} 保存`;
          button.addEventListener("click", () => displaySnapshot(snapshot));
          item.appendChild(button);
        } catch {
          item.textContent = `${key.slice(HISTORY_PREFIX.length)}：履歴を読み込めません。保存データは保持しています。`;
        }
        panel.appendChild(item);
      }
    } catch {
      status.textContent = "履歴を読み込めません。ブラウザの保存設定を確認してください。";
    }
  }

  save.addEventListener("click", async () => {
    save.disabled = true;
    try {
      // Serialize saves across tabs on browsers with Web Locks support.
      const snapshot = navigator.locks?.request
        ? await navigator.locks.request(HISTORY_PREFIX, savePredictionHistory)
        : savePredictionHistory();
      status.textContent = `${snapshot.targetDate}の予想を保存しました。`;
      if (!panel.hidden) displayList();
    } catch (error) {
      status.textContent = error.name === "QuotaExceededError"
        ? "保存容量が不足しています。履歴は保存できませんでした。既存の履歴は保持しています。"
        : error.name === "SecurityError"
          ? "ブラウザで保存が許可されていません。保存設定を確認してください。"
          : error.message || "予想を保存できませんでした。";
    } finally {
      save.disabled = false;
    }
  });
  toggle.addEventListener("click", () => {
    panel.hidden = !panel.hidden;
    toggle.setAttribute("aria-expanded", String(!panel.hidden));
    toggle.textContent = panel.hidden ? "予想履歴" : "予想履歴を閉じる";
    if (!panel.hidden) displayList();
  });
}

createHistoryUI();

try {
  const saved = JSON.parse(localStorage.getItem("ooi-keiba-mobile-data-v2"));
  if (Array.isArray(saved?.races) && saved.races.length) {
    races = saved.races;
    showData(saved.targetDate);
  }
} catch {
  try { localStorage.removeItem("ooi-keiba-mobile-data-v2"); } catch {}
}
