(function () {
const keys=['競馬場','競走年月日','レース番号'];
const key=r=>keys.map(k=>String(r[k]??'').trim()).join('|');
const positive=v=>Number.isFinite(Number(v))&&Number(v)>0;
const cancelled=r=>/取消|除外|取止/.test(String(r['着差']||'')+String(r['着順']||''));
const today=()=>new Intl.DateTimeFormat('sv-SE',{timeZone:'Asia/Tokyo',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()).replaceAll('-','');
function parse(csv,required,label) {
 const raw=XLSX.read(csv.replace(/^\uFEFF/,''),{type:'string',raw:true});
 const rows=XLSX.utils.sheet_to_json(raw.Sheets[raw.SheetNames[0]],{header:1,defval:''});
 const headers=(rows.shift()||[]).map(v=>String(v).trim());
 if(required.some(k=>!headers.includes(k))) throw new Error(label+'の列が不足しています。NARのCSVを選択してください。');
 return rows.filter(r=>r.some(v=>v!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
}
function unique(rows,fn,label){const seen=new Set();for(const r of rows){const k=fn(r);if(seen.has(k))throw new Error(label+'に重複があります。');seen.add(k);}}
async function merge(book,files,day=today()) {
 const found={};
 for(const f of files){const m=f.name.match(/^(\d{8})_(horselist|racelist|payback|odds)\.csv$/i);
  if(!m||m[1]!==day)throw new Error('日本時間の本日（'+day+'）のNAR CSVを選んでください。ZIPは先に展開してください。');
  if(found[m[2].toLowerCase()])throw new Error('同じ種類のCSVが複数選ばれています。');found[m[2].toLowerCase()]=f;
 }
 if(!['horselist','racelist','payback'].every(k=>found[k]))throw new Error('horselist・racelist・paybackの3ファイルをまとめて選んでください。oddsもあれば一緒に選択します。');
 const need={horselist:[...keys,'馬番','馬名','騎手名','父馬名','着順','着差','人気','最高タイム','当競馬場成績','うち当距離成績','馬体重増減','上がり3F'],racelist:[...keys,'距離','レース名'],payback:[...keys,...[1,2,3].flatMap(i=>['複勝組番'+i,'複勝払戻金'+i+'（円）'])],odds:[...keys,'賭式','番号1','オッズ','人気']};
 const data={odds:[]};
 for(const [kind,f] of Object.entries(found)){
  const rows=parse(await f.text(),need[kind],f.name).filter(r=>String(r['競馬場']).trim()==='大井');
  if(rows.some(r=>String(r['競走年月日']).trim()!==day))throw new Error('CSV内の日付がファイル名と一致しません。');
  data[kind]=rows.map(r=>({...r,'競馬場':'大井','競走年月日':Number(day),'レース番号':Number(r['レース番号'])}));
 }
 if(!data.horselist.length)throw new Error('このCSVに本日の大井の出走馬はありません。非開催またはデータ未取得です。');
 unique(data.racelist,key,'レース一覧');unique(data.horselist,r=>key(r)+'|'+Number(r['馬番']),'出馬表');
 const races=new Map(data.racelist.map(r=>[key(r),r]));
 if(data.horselist.some(r=>!races.has(key(r))||!positive(races.get(key(r))['距離'])||!positive(r['馬番'])))throw new Error('出馬表とレース一覧が一致しないか、距離・馬番が不足しています。');
 const odds=data.odds.filter(r=>r['賭式']==='単勝'&&Number(r['オッズ'])>=1&&Number.isInteger(Number(r['人気']))&&positive(r['人気'])&&positive(r['番号1']));
 unique(odds,r=>key(r)+'|'+Number(r['番号1']),'単勝オッズ');
 const om=new Map(odds.map(r=>[key(r)+'|'+Number(r['番号1']),Number(r['人気'])]));
 const paid=new Map();
 for(const r of data.payback)for(let i=1;i<=3;i++){
  const horse=Number(r['複勝組番'+i]),amount=Number(r['複勝払戻金'+i+'（円）']);if(!(horse>0&&amount>0))continue;
  const k=key(r)+'|'+horse;if(paid.has(k)&&paid.get(k)['複勝払戻金1（円）']!==amount)throw new Error('同じ馬の複勝払戻金が一致しません。');
  paid.set(k,{...Object.fromEntries(keys.map(k=>[k,r[k]])),'複勝組番1':horse,'複勝払戻金1（円）':amount});
 }
 const finished=new Set([...data.horselist.filter(r=>positive(r['着順'])),...paid.values()].map(key));
 const horses=data.horselist.map(r=>{
  const race=races.get(key(r));
  return {...r,'距離':Number(race['距離']),'レース名':race['レース名'],'天候':race['天候'],'馬場':race['馬場'],
   '上がり3F数値':positive(r['上がり3F'])?Number(r['上がり3F']):null,
   '人気':cancelled(r)?null:finished.has(key(r))?(positive(r['人気'])?Number(r['人気']):null):(om.get(key(r)+'|'+Number(r['馬番']))??null)};
 });
 const old=XLSX.utils.sheet_to_json(book.Sheets['NAR統合データ'],{defval:null});
 const oldPay=XLSX.utils.sheet_to_json(book.Sheets['NAR払戻データ']||{},{defval:null});
 const result={...book,Sheets:{...book.Sheets}};
 result.Sheets['NAR統合データ']=XLSX.utils.json_to_sheet([...old.filter(r=>!races.has(key(r))),...horses]);
 result.Sheets['NAR払戻データ']=XLSX.utils.json_to_sheet([...oldPay.filter(r=>!races.has(key(r))),...paid.values()]);
 return {workbook:result,date:Number(day),count:horses.length,odds:odds.length};
}
globalThis.OoiDaily={merge,today};
})();
