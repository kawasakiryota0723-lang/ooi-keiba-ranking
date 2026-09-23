(function () {
  const number = v => v === '' || v == null || !Number.isFinite(Number(v)) ? null : Number(v);
  function decision({ raceName, popularity, gap, error, hasHorses = true, finished = false }) {
    if (!hasHorses || error) return '判定不可';
    if (!String(raceName || '').trim()) return 'レース名未取得';
    if (/新馬/.test(raceName)) return '新馬・参考';
    if (finished) return '結果取込済み';
    if (!(number(popularity) >= 1) || number(gap) === null) return '人気・点差未取得';
    return number(popularity) === 1 && number(gap) >= 10 ? '購入候補' : '見送り';
  }
  // The declared field can change after sales open. Never infer a settled bet
  // from the current row count: use the official place payout combinations.
  function placeResult(horse, payouts) {
    if (!horse) return { status: '未確認', hit: null, payout: null };
    if (/取消|除外/.test(String(horse['着順'] || ''))) return { status: '取消・除外（集計対象外）', hit: null, payout: null };
    const entries = (payouts || []).flatMap(row => [1, 2, 3].map(i => ({
      horse: number(row['複勝組番' + i]), amount: number(row['複勝払戻金' + i + '（円）'])
    }))).filter(p => p.horse > 0 && p.amount > 0);
    if (!entries.length) return { status: '払戻未取得', hit: null, payout: null };
    const paid = entries.find(p => p.horse === number(horse['馬番']));
    if (paid) return { status: '複勝的中', hit: 1, payout: paid.amount };
    if (!(number(horse['着順']) > 0) && !/中止|失格/.test(String(horse['着順'] || ''))) return { status: '着順未取得', hit: null, payout: null };
    return { status: '複勝不的中', hit: 0, payout: 0 };
  }
  const rules = { decision, placeResult };
  globalThis.OoiRules = rules;
  if (typeof module !== 'undefined') module.exports = rules;
})();
