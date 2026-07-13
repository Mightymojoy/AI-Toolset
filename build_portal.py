# -*- coding: utf-8 -*-
"""
AI 工具集门户生成器
读取 E:\AI工具集\AI 工具集.xlsx 的 Sheet1，自动生成 index.html。
表格结构（Sheet1）：
  - 分组行：A列有值、B/C为空  -> 作为来源分组（如「自建网页」「飞书文档」）
  - 表头行：AI智能体名称 / 对应链接 / 分类
  - 数据行：三个单元格都有值 -> {name, url, cat, group}
新增工具：直接在表格里加一行即可，重跑本脚本即可刷新门户。
"""
import openpyxl
import json
import os
import datetime

SRC = r"E:\AI工具集\AI 工具集.xlsx"
OUT = r"E:\AI工具集\index.html"

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["Sheet1"]

entries = []
group = None
for row in ws.iter_rows(values_only=True):
    cells = list(row) + [None, None, None]
    a, b, c = cells[0], cells[1], cells[2]
    # 空行跳过
    if a is None and b is None and c is None:
        continue
    a_s = str(a).strip() if a is not None else ""
    b_s = str(b).strip() if b is not None else ""
    c_s = str(c).strip() if c is not None else ""
    # 分组行：仅 A 有值
    if a_s and not b_s and not c_s:
        group = a_s
        continue
    # 表头行
    if a_s == "AI智能体名称":
        continue
    # 数据行：三项齐全
    if a_s and b_s and c_s:
        entries.append({
            "group": group or "",
            "name": a_s,
            "url": b_s,
            "cat": c_s,
        })

if not entries:
    raise SystemExit("未从表格中解析到任何工具条目，请检查 Sheet1 结构。")

# 分类顺序（按出现顺序去重）
cats = []
for e in entries:
    if e["cat"] not in cats:
        cats.append(e["cat"])

data_json = json.dumps(entries, ensure_ascii=False)
cats_json = json.dumps(cats, ensure_ascii=False)
gen_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>ITO · AI 工具集</title>
<script src="https://unpkg.com/lucide@latest"></script>
<style>
  :root{
    --gold:#c9a962; --gold-soft:#e7d9b4; --gold-deep:#a98a45;
    --bg:#faf6ef; --card:#ffffff; --ink:#2c2620; --muted:#8a8175;
    --line:#ece4d6; --shadow:0 6px 22px rgba(120,96,52,.08);
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
    background:var(--bg);color:var(--ink);line-height:1.6;-webkit-font-smoothing:antialiased;
    min-height:100vh;position:relative;overflow-x:hidden}
  /* 金质流光背景 */
  .bg-aura{position:fixed;inset:0;z-index:0;overflow:hidden;pointer-events:none}
  .bg-aura span{position:absolute;border-radius:50%;filter:blur(72px);opacity:.62;
    background:radial-gradient(circle at 30% 30%,var(--gold-soft),transparent 70%);will-change:transform}
  .bg-aura .a1{width:560px;height:560px;top:-160px;left:-140px;
    background:radial-gradient(circle at 30% 30%,rgba(201,169,98,.58),transparent 70%);
    animation:float1 22s ease-in-out infinite}
  .bg-aura .a2{width:500px;height:500px;top:18%;right:-180px;
    background:radial-gradient(circle at 30% 30%,rgba(231,217,180,.68),transparent 70%);
    animation:float2 27s ease-in-out infinite}
  .bg-aura .a3{width:640px;height:640px;bottom:-240px;left:32%;
    background:radial-gradient(circle at 30% 30%,rgba(169,138,69,.45),transparent 70%);
    animation:float3 32s ease-in-out infinite}
  @keyframes float1{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(110px,80px) scale(1.12)}}
  @keyframes float2{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(-100px,70px) scale(1.08)}}
  @keyframes float3{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(80px,-90px) scale(1.15)}}
  /* 粒子互动层 */
  #particles{position:fixed;inset:0;z-index:1;pointer-events:none}
  @media (prefers-reduced-motion:reduce){
    .bg-aura span{animation:none!important}
  }
  .wrap{max-width:1180px;margin:0 auto;padding:40px 24px 64px;position:relative;z-index:2}
  header.top{display:flex;align-items:center;gap:14px;margin-bottom:6px}
  .logo{width:42px;height:42px;border-radius:11px;background:linear-gradient(135deg,var(--gold),var(--gold-deep));
    display:flex;align-items:center;justify-content:center;color:#fff;flex:none}
  .logo svg{width:22px;height:22px}
  h1{font-size:25px;font-weight:700;letter-spacing:.5px}
  .sub{color:var(--muted);font-size:14px;margin:2px 0 22px}
  .bar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;margin-bottom:18px}
  .search{flex:1;min-width:220px;display:flex;align-items:center;gap:8px;background:var(--card);
    border:1px solid var(--line);border-radius:12px;padding:10px 14px;box-shadow:var(--shadow)}
  .search svg{width:17px;height:17px;color:var(--muted);flex:none}
  .search input{border:none;outline:none;background:transparent;font-size:14px;width:100%;color:var(--ink)}
  .chips{display:flex;flex-wrap:wrap;gap:8px}
  .chip{border:1px solid var(--line);background:var(--card);color:var(--muted);
    padding:7px 15px;border-radius:20px;font-size:13px;cursor:pointer;transition:.15s;user-select:none}
  .chip:hover{border-color:var(--gold-soft)}
  .chip.active{background:var(--gold);border-color:var(--gold);color:#fff;font-weight:600}
  .cat-block{margin-top:30px}
  .cat-title{display:flex;align-items:center;gap:9px;font-size:17px;font-weight:700;
    color:var(--gold-deep);margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid var(--line)}
  .cat-title svg{width:18px;height:18px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(264px,1fr));gap:16px}
  .card{background:rgba(255,255,255,.82);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
    border:1px solid var(--line);border-radius:16px;padding:18px 18px 16px;
    box-shadow:var(--shadow);text-decoration:none;color:inherit;display:flex;flex-direction:column;gap:10px;
    transition:.18s;position:relative;overflow:hidden}
  .card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--gold);opacity:0;transition:.18s}
  .card:hover{transform:translateY(-3px);border-color:var(--gold-soft);box-shadow:0 12px 28px rgba(120,96,52,.14)}
  .card:hover::before{opacity:1}
  .card h3{font-size:15.5px;font-weight:650;line-height:1.35}
  .card .url{font-size:12px;color:var(--muted);word-break:break-all;display:flex;align-items:center;gap:6px}
  .card .url svg{width:13px;height:13px;flex:none}
  .tags{display:flex;gap:7px;flex-wrap:wrap;margin-top:auto}
  .tag{font-size:11px;padding:3px 9px;border-radius:8px;font-weight:600}
  .tag.cat{background:var(--gold-soft);color:var(--gold-deep)}
  .tag.src{background:#f1ede6;color:#7a7264}
  .tag.login{background:#f7e9e9;color:#b45454}
  .empty{color:var(--muted);text-align:center;padding:50px 0;font-size:14px}
</style>
</head>
<body>
<div class="bg-aura"><span class="a1"></span><span class="a2"></span><span class="a3"></span></div>
<canvas id="particles"></canvas>
<div class="wrap">
  <header class="top">
    <div class="logo"><i data-lucide="boxes"></i></div>
    <h1>ITO · AI 工具集</h1>
  </header>
  <p class="sub">团队内部 AI 工具与文档统一入口 · 点击卡片跳转对应网页</p>

  <div class="bar">
    <label class="search">
      <i data-lucide="search"></i>
      <input id="q" type="text" placeholder="搜索工具名称…" />
    </label>
    <div class="chips" id="chips"></div>
  </div>

  <div id="content"></div>
</div>

<script>
const DATA = __DATA__;
const CATS = __CATS__;
const ICONS = { "数据看板":"bar-chart-3", "知识库与文档":"book-open", "内容创作工具":"pen-tool", "数据表":"table" };

const content = document.getElementById('content');
const chips = document.getElementById('chips');
const q = document.getElementById('q');
let activeCat = "全部";

function renderChips(){
  chips.innerHTML = "";
  const all = ["全部", ...CATS];
  all.forEach(c=>{
    const el = document.createElement('span');
    el.className = 'chip' + (c===activeCat?' active':'');
    el.textContent = c;
    el.onclick = ()=>{ activeCat = c; renderChips(); render(); };
    chips.appendChild(el);
  });
}

function esc(s){ return String(s).replace(/[&<>"]/g, m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m])); }

function card(e){
  const icon = ICONS[e.cat] || "box";
  const tags = `<div class="tags">
      <span class="tag cat">${esc(e.cat)}</span>
    </div>`;
  return `<a class="card" href="${esc(e.url)}" target="_blank" rel="noopener">
      <h3>${esc(e.name)}</h3>
      ${tags}
    </a>`;
}

function render(){
  const kw = q.value.trim().toLowerCase();
  const filtered = DATA.filter(e=>{
    const okCat = activeCat==="全部" || e.cat===activeCat;
    const okKw = !kw || e.name.toLowerCase().includes(kw) || e.url.toLowerCase().includes(kw);
    return okCat && okKw;
  });

  if(!filtered.length){ content.innerHTML = '<div class="empty">没有匹配的工具</div>'; lucide.createIcons(); return; }

  // 按分类分块
  let html = "";
  CATS.forEach(cat=>{
    const list = filtered.filter(e=>e.cat===cat);
    if(!list.length) return;
    html += `<section class="cat-block">
        <div class="cat-title"><i data-lucide="${ICONS[cat]||'box'}"></i>${esc(cat)}<span style="color:var(--muted);font-weight:400;font-size:13px"> · ${list.length}</span></div>
        <div class="grid">${list.map(card).join('')}</div>
      </section>`;
  });
  content.innerHTML = html;
  lucide.createIcons();
}

q.addEventListener('input', render);
renderChips();
render();
lucide.createIcons();

// 粒子互动背景：金色微粒漂浮 + 鼠标拨开 + 邻近淡金连线；尊重系统「减少动态」
(function(){
  if(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const cv = document.getElementById('particles');
  if(!cv) return;
  const ctx = cv.getContext('2d');
  let W, H, dpr;
  function resize(){
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = cv.width  = Math.floor(innerWidth  * dpr);
    H = cv.height = Math.floor(innerHeight * dpr);
    cv.style.width  = innerWidth  + 'px';
    cv.style.height = innerHeight + 'px';
    ctx.setTransform(dpr,0,0,dpr,0,0);
  }
  resize();
  window.addEventListener('resize', resize);

  const N = Math.max(28, Math.min(56, Math.floor(innerWidth / 24)));
  const pts = [];
  for(let i=0;i<N;i++){
    pts.push({
      x: Math.random()*innerWidth, y: Math.random()*innerHeight,
      vx: (Math.random()-.5)*.25, vy: (Math.random()-.5)*.25,
      r: Math.random()*1.5 + .6
    });
  }
  const mouse = {x:-9999, y:-9999};
  window.addEventListener('mousemove', e=>{ mouse.x=e.clientX; mouse.y=e.clientY; });
  window.addEventListener('mouseout',  ()=>{ mouse.x=-9999; mouse.y=-9999; });
  window.addEventListener('touchmove', e=>{ if(e.touches[0]){ mouse.x=e.touches[0].clientX; mouse.y=e.touches[0].clientY; } }, {passive:true});
  window.addEventListener('touchend',  ()=>{ mouse.x=-9999; mouse.y=-9999; });

  const LINK = 118, PUSH = 110;
  function tick(){
    ctx.clearRect(0,0,innerWidth,innerHeight);
    for(let i=0;i<pts.length;i++){
      const p = pts[i];
      p.x += p.vx; p.y += p.vy;
      if(p.x < 0 || p.x > innerWidth)  p.vx *= -1;
      if(p.y < 0 || p.y > innerHeight) p.vy *= -1;
      const dx = p.x - mouse.x, dy = p.y - mouse.y, d = Math.hypot(dx,dy);
      if(d < PUSH && d > 0){ const f = (PUSH - d) / PUSH * 1.0; p.x += dx/d*f; p.y += dy/d*f; }
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, 6.283);
      ctx.fillStyle = 'rgba(201,169,98,.5)';
      ctx.fill();
      for(let j=i+1;j<pts.length;j++){
        const q = pts[j], dd = Math.hypot(p.x-q.x, p.y-q.y);
        if(dd < LINK){
          ctx.beginPath();
          ctx.moveTo(p.x, p.y); ctx.lineTo(q.x, q.y);
          ctx.strokeStyle = 'rgba(201,169,98,' + (0.12*(1 - dd/LINK)).toFixed(3) + ')';
          ctx.lineWidth = .6;
          ctx.stroke();
        }
      }
    }
    requestAnimationFrame(tick);
  }
  tick();
})();
</script>
</body>
</html>
"""

HTML = (HTML
        .replace("__DATA__", data_json)
        .replace("__CATS__", cats_json)
        .replace("__GEN__", gen_time))

with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML)

print(f"OK 解析到 {len(entries)} 个工具，分类 {cats}")
print(f"已生成 -> {OUT}")
