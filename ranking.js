// Uses the same scoring formula as the Ooi Excel template. No network calls.
(function () {
const num=v=>v===null||v===undefined||v===''?null:Number.isFinite(Number(v))?Number(v):null;
const rate=v=>{const a=String(v??'').split('-').map(Number);return a.length===4&&a.every(Number.isFinite)?(a[0]+a[1]+a[2]+1.25)/(a.reduce((x,y)=>x+y,0)+5):.25};
const time=v=>{if(!v)return null;const m=String(v).match(/^(\d+):(\d+(?:\.\d+)?)$/);return m?60*+m[1]+ +m[2]:num(v)};
function calc(all,date,race,venue){
 const current=all.filter(r=>+r['競走年月日']===date&&+r['レース番号']===race&&r['競馬場']===venue);
 const dist=current[0]?.['距離'];const prior=all.filter(r=>+r['競走年月日']<date&&+r['着順']>0);
 const out=current.map((r,i)=>{
  const h=prior.filter(a=>a['馬名']===r['馬名']).sort((a,b)=>b['競走年月日']-a['競走年月日']);
  const hd=h.filter(a=>+a['距離']===+dist).slice(0,3);
  const past=h.slice(0,3), weights=[.5,.3,.2];
  const weighted=past.length?past.reduce((s,a,j)=>s+(+a['着順'])*weights[j],0)/weights.slice(0,past.length).reduce((a,b)=>a+b,0):null;
  const avg=hd.length?hd.reduce((s,a)=>s+(num(a['上がり3F数値'])??0),0)/hd.length:null;
  const rr=key=>{const a=prior.filter(q=>q[key]===r[key]&&q['競馬場']===venue&&+q['距離']===+dist);return (a.filter(q=>+q['着順']<=3).length+1.25)/(a.length+5)};
  return {r,i,past,hd,weighted,avg,m:rate(r['当競馬場成績']),n:rate(r['うち当距離成績']),o:rr('騎手名'),p:rr('父馬名'),sec:time(r['最高タイム'])};
 });
 const scaled=(v,field,max)=>{const a=out.map(x=>x[field]).filter(x=>x!==null);return v===null||a.length<2?max/2:(a.length-1-a.filter(x=>x<v).length)/(a.length-1)*max};
 for(const h of out){h.cv=scaled(h.sec,'sec',15);const change=num(h.r['馬体重増減']);h.score=Math.round((h.weighted===null?10:Math.max(0,20-(h.weighted-1)*2))*10+(scaled(h.avg,'avg',20)+h.m*12+h.n*10+h.o*15+h.p*5+(change===null?1.5:Math.abs(change)<=10?3:Math.abs(change)<=20?2:1)+h.cv)*10)/10;}
 const sorted=[...out].sort((a,b)=>b.score-a.score||a.i-b.i);sorted.forEach((h,i)=>h.rank=i+1);
 return {out,sorted,dist,gap:sorted.length>=2?Math.round((sorted[0].score-sorted[1].score)*10)/10:null};
}

const parsedBooks=new WeakMap();
function getDetails(workbook, date, raceNumber) {
  const sheet=workbook.Sheets['NAR統合データ'];
  if(!sheet) return {horses:[],error:'全馬の表示には「NAR統合データ」が入った大井用Excelを選んでください。'};
  let all=parsedBooks.get(workbook);
  if(!all){all=XLSX.utils.sheet_to_json(sheet,{defval:null});parsedBooks.set(workbook,all);}
  const current=all.filter(r=>+r['競走年月日']===date&&+r['レース番号']===raceNumber&&r['競馬場']==='大井');
  if(!current.length) return {horses:[],error:'このレースの出走データがありません。Excelを更新して読み込み直してください。'};
  if(current.some(r=>!num(r['距離']))) return {horses:[],error:'距離データが不足しています。Excelで「すべて更新」を実行してください。'};
  const result=calc(all,date,raceNumber,'大井');
  return {horses:result.sorted.map(h=>({rank:h.rank,horseNumber:h.r['馬番'],horseName:h.r['馬名'],jockey:h.r['騎手名'],score:h.score,popularity:num(h.r['人気']),rating:['◎','○','▲','△','△'][h.rank-1]||''})),error:''};
}
globalThis.OoiRanking={getDetails};
})();
