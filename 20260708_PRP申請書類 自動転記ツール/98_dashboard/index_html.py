# -*- coding: utf-8 -*-
"""ダッシュボードのフロントエンド（1ページ・依存ゼロ）。app.py から読み込む。"""

INDEX_HTML = r"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PRP 自動転記ツール ダッシュボード</title>
<style>
  :root{
    --bg:#f4f6fb; --panel:#ffffff; --ink:#1a2230; --sub:#5b6b82;
    --line:#e3e8f0; --brand:#2f6df6; --brand-ink:#fff;
    --ok:#1f9d63; --okbg:#e6f6ee; --warn:#c47d09; --warnbg:#fff4d9;
    --danger:#d64545; --dangerbg:#fdeaea; --shadow:0 1px 3px rgba(20,30,60,.08),0 6px 24px rgba(20,30,60,.06);
  }
  @media (prefers-color-scheme:dark){
    :root{ --bg:#0f141c; --panel:#171e29; --ink:#e8edf5; --sub:#9aa8bd;
      --line:#28313f; --okbg:#123024; --warnbg:#332a12; --dangerbg:#3a1e1e; --shadow:0 1px 3px rgba(0,0,0,.4);}
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
    font-family:-apple-system,"Hiragino Kaku Gothic ProN","Yu Gothic UI",Meiryo,system-ui,sans-serif;
    line-height:1.6;-webkit-font-smoothing:antialiased}
  header{background:linear-gradient(135deg,#2f6df6,#4b8bff);color:#fff;padding:20px 24px;box-shadow:var(--shadow)}
  header h1{margin:0;font-size:19px;font-weight:700;letter-spacing:.02em}
  header p{margin:4px 0 0;font-size:12.5px;opacity:.85}
  .wrap{max-width:1080px;margin:0 auto;padding:22px 20px 60px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
  @media (max-width:820px){.grid{grid-template-columns:1fr}}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;box-shadow:var(--shadow)}
  .card h2{margin:0 0 12px;font-size:14px;display:flex;align-items:center;gap:8px}
  .card h2 .n{background:var(--brand);color:#fff;width:22px;height:22px;border-radius:50%;
    display:grid;place-items:center;font-size:12px;font-weight:700;flex:none}
  .muted{color:var(--sub);font-size:12.5px}
  /* dropzone */
  #drop{border:2px dashed var(--line);border-radius:12px;padding:26px 16px;text-align:center;cursor:pointer;
    transition:.15s;background:transparent}
  #drop:hover,#drop.hot{border-color:var(--brand);background:rgba(47,109,246,.06)}
  #drop .big{font-size:30px}
  #drop b{color:var(--brand)}
  /* file list */
  .files{list-style:none;margin:12px 0 0;padding:0;display:flex;flex-direction:column;gap:6px}
  .files li{display:flex;align-items:center;gap:10px;padding:9px 11px;border:1px solid var(--line);
    border-radius:10px;font-size:13px;cursor:pointer;background:transparent}
  .files li:hover{border-color:var(--brand)}
  .files li.sel{border-color:var(--brand);background:rgba(47,109,246,.08)}
  .files li .nm{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .files li .meta{color:var(--sub);font-size:11px;flex:none}
  .dot{width:9px;height:9px;border-radius:50%;background:var(--line);flex:none}
  .files li.sel .dot{background:var(--brand)}
  /* button */
  button{font:inherit;cursor:pointer;border-radius:10px;border:1px solid transparent;padding:10px 16px;font-weight:600}
  .btn{background:var(--brand);color:var(--brand-ink)}
  .btn:disabled{opacity:.5;cursor:not-allowed}
  .btn.ghost{background:transparent;color:var(--brand);border-color:var(--line)}
  .btn.sm{padding:6px 11px;font-size:12px}
  .run-row{display:flex;align-items:center;gap:12px;margin-top:14px}
  /* log console */
  #console{margin-top:14px;background:#0d1420;color:#cdd8ea;border-radius:10px;padding:12px 14px;
    font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:12px;line-height:1.55;
    max-height:260px;overflow:auto;white-space:pre-wrap;display:none}
  #console .l{opacity:.92}
  #console .warn{color:#ffd479}
  #console .err{color:#ff9a9a}
  /* stat cards */
  .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
  @media (max-width:560px){.stats{grid-template-columns:repeat(2,1fr)}}
  .stat{border:1px solid var(--line);border-radius:11px;padding:11px 12px;text-align:center}
  .stat .v{font-size:22px;font-weight:800;line-height:1.2}
  .stat .k{font-size:11px;color:var(--sub);margin-top:2px}
  .stat.ok .v{color:var(--ok)} .stat.warn .v{color:var(--warn)} .stat.danger .v{color:var(--danger)}
  /* table */
  .tblwrap{overflow-x:auto;border:1px solid var(--line);border-radius:10px}
  table{border-collapse:collapse;width:100%;font-size:12px;min-width:560px}
  th,td{padding:7px 9px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
  th{background:rgba(47,109,246,.06);position:sticky;top:0;font-weight:700;font-size:11px;color:var(--sub)}
  tr:last-child td{border-bottom:none}
  .pill{display:inline-block;padding:1px 8px;border-radius:99px;font-size:11px;font-weight:700}
  .pill.warn{background:var(--warnbg);color:var(--warn)}
  .pill.danger{background:var(--dangerbg);color:var(--danger)}
  .empty{padding:24px;text-align:center;color:var(--sub);font-size:13px}
  /* output/history rows */
  .row{display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid var(--line);
    border-radius:10px;margin-bottom:8px;font-size:13px}
  .row .nm{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .row .meta{color:var(--sub);font-size:11px;flex:none}
  .section-t{margin:26px 0 12px;font-size:14px;font-weight:700}
  .toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:var(--ink);color:var(--bg);
    padding:10px 18px;border-radius:99px;font-size:13px;box-shadow:var(--shadow);opacity:0;transition:.25s;pointer-events:none}
  .toast.show{opacity:1}
  .spin{width:16px;height:16px;border:2.5px solid rgba(255,255,255,.4);border-top-color:#fff;border-radius:50%;
    display:inline-block;animation:sp .7s linear infinite;vertical-align:-3px}
  @keyframes sp{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<header>
  <h1>🩺 PRP申請書類 自動転記ツール</h1>
  <p>ヒアリングシートを入れて「転記実行」を押すだけ。書類一式が生成されます。</p>
</header>

<div class="wrap">
  <div class="grid">
    <!-- 左: 入力 & 実行 -->
    <div class="card">
      <h2><span class="n">1</span>ヒアリングシートを選ぶ</h2>
      <div id="drop">
        <div class="big">📄</div>
        <div>ここに <b>PRPヒアリングシート.xlsx</b> をドラッグ&ドロップ<br>
          <span class="muted">またはクリックしてファイルを選択</span></div>
        <input type="file" id="file" accept=".xlsx" hidden>
      </div>
      <ul class="files" id="inputList"></ul>

      <h2 style="margin-top:20px"><span class="n">2</span>転記を実行する</h2>
      <div class="muted" id="selInfo">上の一覧からファイルを選択してください。</div>
      <div class="run-row">
        <button class="btn" id="runBtn" disabled>▶ 転記実行</button>
        <span class="muted" id="runState"></span>
      </div>
      <div id="console"></div>
    </div>

    <!-- 右: 結果 -->
    <div class="card">
      <h2><span class="n">3</span>結果 ・ 要確認項目</h2>
      <div id="result">
        <div class="empty">まだ実行されていません。<br>転記を実行すると、ここに結果と「要確認（黄色）」項目が表示されます。</div>
      </div>
    </div>
  </div>

  <div class="section-t">📦 生成された書類フォルダ</div>
  <div id="outputs"><div class="empty">まだありません。</div></div>

  <div class="section-t">🗂 実行履歴（ログ）</div>
  <div id="history"><div class="empty">まだありません。</div></div>
</div>

<div class="toast" id="toast"></div>

<script>
const $ = s => document.querySelector(s);
let selected = null;      // 選択中の入力ファイル名
let running = false;

function toast(msg){ const t=$("#toast"); t.textContent=msg; t.classList.add("show");
  clearTimeout(t._t); t._t=setTimeout(()=>t.classList.remove("show"),2200); }

function esc(s){ return (s||"").replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

// ---- 一覧の描画 ----
function renderInputs(list){
  const ul=$("#inputList"); ul.innerHTML="";
  if(!list.length){ ul.innerHTML='<li style="cursor:default"><span class="muted">01_input にファイルがありません</span></li>'; return; }
  list.forEach(f=>{
    const li=document.createElement("li");
    if(f.name===selected) li.classList.add("sel");
    li.innerHTML=`<span class="dot"></span><span class="nm">${esc(f.name)}</span>
      <span class="meta">${f.mtime} ・ ${f.size_kb}KB</span>`;
    li.onclick=()=>{ selected=f.name; renderInputs(list); updateSel(); };
    ul.appendChild(li);
  });
}
function updateSel(){
  $("#selInfo").innerHTML = selected
    ? `選択中: <b>${esc(selected)}</b>` : "上の一覧からファイルを選択してください。";
  $("#runBtn").disabled = !selected || running;
}

function renderOutputs(list){
  const box=$("#outputs");
  if(!list||!list.length){ box.innerHTML='<div class="empty">まだありません。</div>'; return; }
  box.innerHTML="";
  list.forEach(o=>{
    const row=document.createElement("div"); row.className="row";
    row.innerHTML=`<span>📁</span><span class="nm">${esc(o.name)}</span>
      <span class="meta">${o.files}ファイル ・ ${o.mtime}</span>
      <button class="btn sm">⬇ ZIP</button>`;
    row.querySelector("button").onclick=()=>{ window.location="/api/download?folder="+encodeURIComponent(o.name); };
    box.appendChild(row);
  });
}

function renderHistory(list){
  const box=$("#history");
  if(!list||!list.length){ box.innerHTML='<div class="empty">まだありません。</div>'; return; }
  box.innerHTML="";
  list.forEach(l=>{
    const row=document.createElement("div"); row.className="row";
    row.innerHTML=`<span>🗒</span><span class="nm">${esc(l.name)}</span>
      <span class="meta">${l.mtime}</span>
      <button class="btn ghost sm">結果を表示</button>
      <button class="btn sm">⬇ Excel</button>`;
    const [showBtn,dlBtn]=row.querySelectorAll("button");
    showBtn.onclick=async()=>{ const r=await fetch("/api/log?name="+encodeURIComponent(l.name));
      if(r.ok) renderResult(await r.json()); else toast("ログを読めませんでした"); window.scrollTo({top:0,behavior:"smooth"}); };
    dlBtn.onclick=()=>{ window.location="/api/download-log?name="+encodeURIComponent(l.name); };
    box.appendChild(row);
  });
}

// ---- 結果パネル ----
function renderResult(res){
  const box=$("#result");
  if(!res||!res.total){ box.innerHTML='<div class="empty">解析できる結果がありませんでした。</div>'; return; }
  const s=res.stats||{};
  const done=s["転記済み"]||0, check=s["確認対象"]||0, unmap=s["未割当(要マッピング)"]||0,
        empty=(s["空欄のため未転記"]||0)+(s["入力元なし"]||0);
  let html=`<div class="stats">
    <div class="stat ok"><div class="v">${done}</div><div class="k">転記済み</div></div>
    <div class="stat warn"><div class="v">${check}</div><div class="k">確認対象</div></div>
    <div class="stat danger"><div class="v">${unmap}</div><div class="k">未割当</div></div>
    <div class="stat"><div class="v">${empty}</div><div class="k">空欄/入力元なし</div></div>
  </div>`;
  if(res.info && res.info["入力"]){
    html+=`<div class="muted" style="margin-bottom:12px">入力: ${esc(res.info["入力"].split("/").pop())}</div>`;
  }
  const checks=res.checks||[];
  html+=`<div style="font-weight:700;font-size:13px;margin-bottom:8px">⚠ 要確認（目視チェック対象）: ${checks.length}件</div>`;
  if(!checks.length){
    html+='<div class="empty">要確認の項目はありません 🎉</div>';
  }else{
    html+='<div class="tblwrap"><table><thead><tr>'+
      '<th>文書</th><th>シート</th><th>セル</th><th>項目</th><th>内容</th><th>状態</th><th>備考</th>'+
      '</tr></thead><tbody>';
    checks.forEach(c=>{
      const danger = c.status.indexOf("未割当")>=0;
      html+=`<tr>
        <td>${esc(c.doc)}</td><td>${esc(c.sheet)}</td><td>${esc(c.cell)}</td>
        <td>${esc(c.var)}</td><td>${esc(c.content)}</td>
        <td><span class="pill ${danger?'danger':'warn'}">${esc(c.status)}</span></td>
        <td>${esc(c.note)}</td></tr>`;
    });
    html+='</tbody></table></div>';
  }
  box.innerHTML=html;
}

// ---- コンソール ----
function logLine(text,cls){
  const c=$("#console"); c.style.display="block";
  const d=document.createElement("div"); d.className="l "+(cls||"");
  d.textContent=text; c.appendChild(d); c.scrollTop=c.scrollHeight;
}

// ---- アップロード ----
async function uploadFile(file){
  if(!file.name.toLowerCase().endsWith(".xlsx")){ toast("拡張子 .xlsx のファイルを選んでください"); return; }
  const fd=new FormData(); fd.append("file",file);
  toast("アップロード中…");
  const r=await fetch("/api/upload",{method:"POST",body:fd});
  const j=await r.json();
  if(!r.ok){ toast(j.error||"アップロード失敗"); return; }
  selected=j.name; toast("アップロード完了: "+j.name);
  await refresh(); updateSel();
}

// ---- 実行(SSE) ----
function run(){
  if(!selected||running) return;
  running=true; updateSel();
  $("#runBtn").innerHTML='<span class="spin"></span> 実行中…';
  $("#runState").textContent="処理しています。書類の数によって数十秒〜数分かかります。";
  $("#console").innerHTML=""; $("#console").style.display="block";

  const es=new EventSource("/api/run?file="+encodeURIComponent(selected));
  es.addEventListener("log", e=>{
    const line=JSON.parse(e.data).line;
    let cls=""; if(/エラー|異常|失敗/.test(line)) cls="err";
    else if(/確認|未転記|未割当|不足/.test(line)) cls="warn";
    logLine(line,cls);
  });
  es.addEventListener("error", e=>{
    try{ logLine("✖ "+JSON.parse(e.data).message,"err"); }catch(_){}
    // EventSource は接続断でも error を投げるので、doneが来ていなければ終了扱い
    finish(false);
  });
  es.addEventListener("done", e=>{
    const d=JSON.parse(e.data);
    logLine("✔ 完了しました。","");
    renderResult(d.result);
    renderOutputs(d.outputs);
    refresh();   // 履歴更新
    es.close(); finish(true);
    toast("転記が完了しました");
  });
  function finish(ok){
    if(!running) return;
    running=false; es.close();
    $("#runBtn").innerHTML="▶ 転記実行";
    $("#runState").textContent = ok? "完了" : "終了（ログを確認してください）";
    updateSel();
  }
}

// ---- 初期化 ----
async function refresh(){
  const r=await fetch("/api/state"); const s=await r.json();
  renderInputs(s.inputs);
  renderOutputs(s.outputs);
  renderHistory(s.logs);
  if(!selected && s.inputs.length) selected=s.inputs[0].name;
  updateSel();
}

const drop=$("#drop"), fileInput=$("#file");
drop.onclick=()=>fileInput.click();
fileInput.onchange=()=>{ if(fileInput.files[0]) uploadFile(fileInput.files[0]); };
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("hot");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("hot");}));
drop.addEventListener("drop",e=>{ if(e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0]); });
$("#runBtn").onclick=run;

refresh();
</script>
</body>
</html>
"""
