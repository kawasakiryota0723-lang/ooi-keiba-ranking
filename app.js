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
      return { horses: [], error: "全馬の計算結果と一覧が一致しません。Excelで全レース更新を実行・保存して、読み込み直してください。" };
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

try {
  const saved = JSON.parse(localStorage.getItem("ooi-keiba-mobile-data-v2"));
  if (Array.isArray(saved?.races) && saved.races.length) {
    races = saved.races;
    showData(saved.targetDate);
  }
} catch {
  try { localStorage.removeItem("ooi-keiba-mobile-data-v2"); } catch {}
}
