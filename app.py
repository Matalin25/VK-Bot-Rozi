# -*- coding: utf-8 -*-
"""
Интерактивный поиск заказов на фриланс-биржах.

Локальное приложение: Python-бэкенд (стандартная библиотека) отдаёт
HTML-интерфейс и JSON-API /api/search. Поиск выполняется по заданным
параметрам (категории, источники, текстовый запрос, бюджет, свежесть).

Запуск:
    python app.py            # открой http://localhost:8000

Остановка: Ctrl+C.
"""
import json
import os
import re
import io
import csv
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

import parser_freelance as P

# На хостингах (Render/Railway) порт задаётся переменной PORT,
# слушать нужно на 0.0.0.0. Локально — 127.0.0.1:8000.
IS_HOSTED = bool(os.environ.get("PORT"))
HOST = "0.0.0.0" if IS_HOSTED else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8000"))

# Данные последнего успешного поиска (кэш) + защита от параллельных запусков
_CACHE = {"rows": [], "meta": {}}
_LOCK = threading.Lock()
_RUNNING = {"flag": False}

_SOURCE_FUNCS = {
    "fl": ("FL.ru", P.parse_flru),
    "freelance": ("freelance.ru", P.parse_freelance_ru),
    "kwork": ("Kwork.ru", P.parse_kwork),
    "freelancer": ("Freelancer.com", P.parse_freelancer),
    "telegram": ("Telegram", P.parse_telegram),
}


def _parse_price(price):
    """Пытается вытащить число из строки бюджета (в рублях)."""
    if not price:
        return None
    s = str(price).replace("\u00a0", " ").replace("\u202f", " ")
    if "$" in s or "usd" in s.lower():
        m = re.search(r"(\d[\d\s,\.]*)", s)
        if m:
            try:
                return int(float(m.group(1).replace(" ", "").replace(",", "")) * 90)
            except ValueError:
                return None
    m = re.search(r"(\d[\d\s]{2,})", s)
    if m:
        try:
            return int(m.group(1).replace(" ", ""))
        except ValueError:
            return None
    return None


def _collect(sources, pages, query, categories):
    """Парсит выбранные источники и возвращает список заказов."""
    rows = []
    old_pages = P.PAGES
    old_queries = list(P.SEARCH_QUERIES)
    try:
        P.PAGES = max(1, int(pages) if pages else 3)
        if query:
            P.SEARCH_QUERIES = [query]
        else:
            P.SEARCH_QUERIES = old_queries

        for key in sources:
            if key not in _SOURCE_FUNCS:
                continue
            label, func = _SOURCE_FUNCS[key]
            if key == "telegram" and not P.TG_PROXY:
                # t.me часто блокируется; без прокси пробуем, но не ждём долго
                pass
            try:
                found = func()
            except Exception as e:
                print(f"  [!] {label}: {e}")
                found = []
            for r in found:
                r["categories"] = P.classify(r["name"], r["desc"])
                r["price_num"] = _parse_price(r.get("price", ""))
                rows.append(r)
    finally:
        P.PAGES = old_pages
        P.SEARCH_QUERIES = old_queries

    # дедупликация
    seen = set()
    uniq = []
    for r in rows:
        key = r.get("link") or (r["source"], r["name"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)

    # фильтрация по категориям (если выбраны конкретные, а не "все")
    if categories:
        cats = set(categories)
        uniq = [r for r in uniq
                if cats & set(r.get("categories", []))]

    # фильтрация по текстовому запросу
    if query:
        uniq = [r for r in uniq if P.matches_query(r["name"], r["desc"], query)]

    uniq.sort(key=lambda r: (r["source"], r["name"]))
    return uniq


def _export_csv(rows):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["№", "Источник", "Категории", "Название", "Описание",
                "Бюджет", "Ссылка"])
    for i, r in enumerate(rows, 1):
        cats = ", ".join(P.CATEGORIES.get(c, c)
                         for c in r.get("categories", []))
        w.writerow([i, r.get("source", ""), cats, r.get("name", ""),
                    r.get("desc", ""), r.get("price", ""), r.get("link", "")])
    return buf.getvalue().encode("utf-8-sig")


def _export_xlsx(rows):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Заказы"

    columns = ["№", "Источник", "Категории", "Название заказа", "Описание",
               "Бюджет", "Ссылка"]
    header_fill = PatternFill("solid", fgColor="2B5A92")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(columns)
    for c in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = border

    for i, r in enumerate(rows, 1):
        cats = ", ".join(P.CATEGORIES.get(c, c)
                         for c in r.get("categories", []))
        ws.append([i, r.get("source", ""), cats, r.get("name", ""),
                   r.get("desc", ""), r.get("price", ""), r.get("link", "")])

    widths = [5, 16, 26, 48, 80, 18, 55]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in range(2, ws.max_row + 1):
        for c in range(1, len(columns) + 1):
            cell = ws.cell(row=row, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(c in (3, 4, 5, 6, 7)))
        link = ws.cell(row=row, column=7)
        if link.value:
            link.hyperlink = link.value
            link.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[row].height = 45
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8",
              disposition=None):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        if disposition:
            self.send_header("Content-Disposition", disposition)
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path in ("/", "/index.html"):
            self._send(200, INDEX_HTML, "text/html; charset=utf-8")
            return
        if path == "/api/search":
            self._handle_search(parse_qs(parsed.query))
            return
        if path == "/api/status":
            self._send(200, json.dumps({
                "running": _RUNNING["flag"],
                "count": len(_CACHE["rows"]),
                "meta": _CACHE["meta"],
            }))
            return
        self._send(404, json.dumps({"error": "not found"}))

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/export":
            self._send(404, json.dumps({"error": "not found"}))
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length) or b"{}")
            rows = payload.get("rows", [])
            fmt = payload.get("format", "xlsx")
            if fmt == "csv":
                data = _export_csv(rows)
                self._send(200, data, "text/csv; charset=utf-8",
                           f"attachment; filename=orders.csv")
            else:
                data = _export_xlsx(rows)
                self._send(200, data,
                           "application/vnd.openxmlformats-officedocument."
                           "spreadsheetml.sheet",
                           "attachment; filename=orders.xlsx")
        except Exception as e:
            self._send(500, json.dumps({"error": str(e)}))

    def _handle_search(self, qs):
        if _RUNNING["flag"]:
            self._send(409, json.dumps(
                {"error": "Поиск уже выполняется, подождите..."}))
            return

        sources = qs.get("sources", [""])[0].split(",")
        sources = [s for s in sources if s] or list(_SOURCE_FUNCS.keys())
        pages = qs.get("pages", ["3"])[0]
        query = qs.get("q", [""])[0].strip()
        categories = qs.get("categories", [""])[0]
        categories = [c for c in categories.split(",") if c]

        with _LOCK:
            _RUNNING["flag"] = True
        try:
            rows = _collect(sources, pages, query, categories)
            _CACHE["rows"] = rows
            _CACHE["meta"] = {
                "sources": sources,
                "pages": pages,
                "query": query,
                "categories": categories,
                "count": len(rows),
            }
            self._send(200, json.dumps(
                {"rows": rows, "meta": _CACHE["meta"]},
                ensure_ascii=False))
        except Exception as e:
            self._send(500, json.dumps({"error": str(e)}))
        finally:
            with _LOCK:
                _RUNNING["flag"] = False


INDEX_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Поиск заказов - фриланс</title>
<style>
  :root{--bg:#0f1419;--panel:#1a2230;--panel2:#222d3d;--line:#2d3a4d;
        --txt:#e6edf5;--muted:#8b9bb4;--accent:#3d8bfd;--green:#2ecc71;
        --red:#e74c3c;--yellow:#f1c40f}
  *{box-sizing:border-box}
  body{margin:0;font:14px/1.5 "Segoe UI",Roboto,Arial,sans-serif;
       background:var(--bg);color:var(--txt)}
  header{background:var(--panel);padding:14px 20px;border-bottom:1px solid var(--line);
         position:sticky;top:0;z-index:10}
  h1{margin:0;font-size:18px}
  h1 span{color:var(--accent)}
  .panel{background:var(--panel);margin:14px 20px;padding:16px;border-radius:10px;
         border:1px solid var(--line)}
  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
  label{display:block;font-size:12px;color:var(--muted);margin-bottom:6px;
        text-transform:uppercase;letter-spacing:.4px}
  input,select{width:100%;padding:9px 11px;background:var(--panel2);color:var(--txt);
        border:1px solid var(--line);border-radius:7px;font-size:14px}
  .checks{display:flex;flex-wrap:wrap;gap:8px}
  .chk{display:flex;align-items:center;gap:7px;background:var(--panel2);
       padding:7px 11px;border-radius:7px;border:1px solid var(--line);cursor:pointer;
       font-size:13px;user-select:none}
  .chk input{width:auto}
  .chk.on{border-color:var(--accent);background:#25344d}
  .row{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:14px}
  button{background:var(--accent);color:#fff;border:0;padding:10px 22px;
         border-radius:7px;font-size:14px;font-weight:600;cursor:pointer}
  button:hover{filter:brightness(1.1)}
  button:disabled{opacity:.5;cursor:wait}
  .status{color:var(--muted);font-size:13px}
  .cards{margin:0 20px 40px;display:grid;gap:12px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:10px;
        padding:14px 16px;position:relative;border-left:4px solid var(--accent)}
  .card.declined{opacity:.45;border-left-color:var(--red)}
  .card.applied{border-left-color:var(--green)}
  .card h3{margin:0 0 6px;font-size:15px;padding-right:130px}
  .card h3 a{color:var(--txt);text-decoration:none}
  .card h3 a:hover{color:var(--accent)}
  .desc{color:var(--muted);font-size:13px;margin:6px 0;max-height:80px;overflow:hidden}
  .meta{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-top:8px;font-size:12px}
  .tag{background:var(--panel2);padding:3px 9px;border-radius:20px;border:1px solid var(--line);
       color:var(--muted)}
  .tag.cat{color:var(--accent);border-color:#2a4a7a}
  .price{color:var(--yellow);font-weight:600}
  .actions{position:absolute;top:12px;right:12px;display:flex;gap:6px}
  .actions button{padding:5px 10px;font-size:12px;background:var(--panel2);
       border:1px solid var(--line);color:var(--muted)}
  .actions button.on{background:var(--green);color:#fff;border-color:var(--green)}
  .actions button.hide.on{background:var(--red);border-color:var(--red)}
  .empty{text-align:center;color:var(--muted);padding:50px 20px}
  .spinner{display:inline-block;width:15px;height:15px;border:2px solid var(--line);
        border-top-color:var(--accent);border-radius:50%;animation:spin 1s linear infinite;
        vertical-align:-3px;margin-right:7px}
  @keyframes spin{to{transform:rotate(360deg)}}
  .stats{margin:0 20px;display:flex;gap:20px;flex-wrap:wrap;color:var(--muted);font-size:13px}
  .stats b{color:var(--txt)}
</style>
</head>
<body>
<header><h1>Поиск заказов <span>&#9889;</span> фриланс-биржи + Telegram</h1></header>

<div class="panel">
  <div class="grid">
    <div style="grid-column:1/-1">
      <label>Категории заказов</label>
      <div class="checks" id="cats">
        <label class="chk on"><input type="checkbox" value="tracker" checked>Трекеры / канбан</label>
        <label class="chk"><input type="checkbox" value="automation">Автоматизация / боты</label>
        <label class="chk"><input type="checkbox" value="development">Разработка приложений</label>
        <label class="chk"><input type="checkbox" value="crm">CRM / ERP / 1С</label>
        <label class="chk"><input type="checkbox" id="allcats">Искать всё</label>
      </div>
    </div>
    <div style="grid-column:1/-1">
      <label>Источники</label>
      <div class="checks" id="srcs">
        <label class="chk on"><input type="checkbox" value="fl" checked>FL.ru</label>
        <label class="chk on"><input type="checkbox" value="freelance" checked>freelance.ru</label>
        <label class="chk on"><input type="checkbox" value="kwork" checked>Kwork.ru</label>
        <label class="chk on"><input type="checkbox" value="freelancer" checked>Freelancer.com</label>
        <label class="chk"><input type="checkbox" value="telegram">Telegram</label>
      </div>
    </div>
    <div>
      <label>Поисковый запрос (слова через пробел)</label>
      <input id="q" placeholder="например: jira интеграция crm">
    </div>
    <div>
      <label>Страниц ленты на источник</label>
      <input id="pages" type="number" min="1" max="20" value="3">
    </div>
  </div>
  <div class="row">
    <button id="go">&#128269; Искать</button>
    <button id="exportXlsx" style="background:#217346">&#128202; Экспорт Excel</button>
    <button id="exportCsv" style="background:#5a6b82">&#128196; Экспорт CSV</button>
    <span class="status" id="status">Готов к поиску</span>
  </div>
</div>

<div class="panel" style="margin-top:0">
  <div class="grid">
    <div>
      <label>Бюджет от (&#8381;)</label>
      <input id="minPrice" type="number" min="0" placeholder="0">
    </div>
    <div>
      <label>Бюджет до (&#8381;)</label>
      <input id="maxPrice" type="number" min="0" placeholder="&infin;">
    </div>
    <div>
      <label>Поиск по тексту</label>
      <input id="filterText" placeholder="фильтр по названию/описанию">
    </div>
    <div>
      <label>Показывать</label>
      <select id="viewFilter">
        <option value="all">Все заказы</option>
        <option value="applied">Только с откликом</option>
        <option value="fav">Только избранное</option>
        <option value="new">Только непросмотренные</option>
      </select>
    </div>
    <div>
      <label>Сортировка</label>
      <select id="sortBy">
        <option value="source">По источнику</option>
        <option value="budget_desc">Бюджет: дороже</option>
        <option value="budget_asc">Бюджет: дешевле</option>
        <option value="name">По названию</option>
      </select>
    </div>
  </div>
</div>

<div class="stats" id="stats"></div>
<div class="cards" id="cards">
  <div class="empty">Нажмите «Искать», чтобы собрать заказы.<br>
  <span style="font-size:12px">Telegram работает только если настроен TG_PROXY.</span></div>
</div>

<script>
const LS = {
  get(k, d){ try{ return JSON.parse(localStorage.getItem("oz_"+k)) ?? d; }catch(e){ return d; } },
  set(k, v){ localStorage.setItem("oz_"+k, JSON.stringify(v)); }
};
let ALL = [];   // все полученные заказы
let FILTERED = [];  // текущий отфильтрованный список
let META = {};

function bindChk(containerId){
  document.querySelectorAll("#"+containerId+" .chk").forEach(chk=>{
    const inp = chk.querySelector("input");
    chk.classList.toggle("on", inp.checked);
    inp.addEventListener("change", ()=>{
      chk.classList.toggle("on", inp.checked);
      if(containerId==="cats" && inp.checked) document.getElementById("allcats").checked=false;
      if(containerId==="cats" && inp.id==="allcats" && inp.checked){
        document.querySelectorAll("#cats input").forEach(x=>{
          if(x.id!=="allcats"){ x.checked=false; x.closest(".chk").classList.remove("on"); }
        });
      }
    });
  });
}
bindChk("cats"); bindChk("srcs");

function selected(containerId){
  return [...document.querySelectorAll("#"+containerId+" input:checked")]
    .map(x=>x.value).filter(v=>v!=="on" && v!=="allcats");
}

async function search(){
  const go = document.getElementById("go");
  const status = document.getElementById("status");
  const cats = selected("cats");
  const allcats = document.getElementById("allcats").checked;
  const srcs = selected("srcs");
  const q = document.getElementById("q").value.trim();
  const pages = document.getElementById("pages").value || "3";

  const params = new URLSearchParams();
  params.set("sources", srcs.join(","));
  params.set("pages", pages);
  params.set("q", q);
  params.set("categories", allcats ? "" : cats.join(","));

  go.disabled = true;
  status.innerHTML = '<span class="spinner"></span>Идёт сбор заказов... это может занять 1-3 минуты';
  try{
    const r = await fetch("/api/search?"+params.toString());
    const data = await r.json();
    if(data.error){ status.textContent = "Ошибка: "+data.error; go.disabled=false; return; }
    ALL = data.rows || [];
    META = data.meta || {};
    status.textContent = "Найдено заказов: "+ALL.length;
    render();
  }catch(e){
    status.textContent = "Ошибка соединения: "+e;
  }
  go.disabled = false;
}

function catLabel(c){
  return {tracker:"Трекер/канбан", automation:"Автоматизация", development:"Разработка", crm:"CRM/1С"}[c] || c;
}

function render(){
  const minP = parseFloat(document.getElementById("minPrice").value) || 0;
  const maxP = parseFloat(document.getElementById("maxPrice").value) || Infinity;
  const txt = document.getElementById("filterText").value.trim().toLowerCase();
  const view = document.getElementById("viewFilter").value;
  const sort = document.getElementById("sortBy").value;

  const applied = new Set(LS.get("applied", []));
  const fav = new Set(LS.get("fav", []));
  const hidden = new Set(LS.get("hidden", []));

  let list = ALL.filter(r=>{
    const pn = r.price_num;
    if(pn !== null && pn !== undefined){
      if(pn < minP || pn > maxP) return false;
    } else if(minP > 0){ return false; }  // без цены скрываем, если задан минимум
    if(txt && !((r.name+" "+r.desc).toLowerCase().includes(txt))) return false;
    const key = r.link || (r.source+r.name);
    if(view==="applied" && !applied.has(key)) return false;
    if(view==="fav" && !fav.has(key)) return false;
    if(view==="new" && (applied.has(key) || hidden.has(key))) return false;
    return true;
  });

  if(sort==="budget_desc") list.sort((a,b)=>(b.price_num||0)-(a.price_num||0));
  else if(sort==="budget_asc") list.sort((a,b)=>(a.price_num||0)-(b.price_num||0));
  else if(sort==="name") list.sort((a,b)=>a.name.localeCompare(b.name,"ru"));
  else list.sort((a,b)=>(a.source||"").localeCompare(b.source||"")); 

  const cards = document.getElementById("cards");
  FILTERED = list;
  document.getElementById("stats").innerHTML =
    "Показано: <b>"+list.length+"</b> из <b>"+ALL.length+"</b> &nbsp;|&nbsp; "+
    "С откликом: <b>"+applied.size+"</b> &nbsp;|&nbsp;  В избранном: <b>"+fav.size+"</b>";

  if(!list.length){ cards.innerHTML='<div class="empty">Ничего не найдено по текущим фильтрам.</div>'; return; }

  cards.innerHTML = list.map(r=>{
    const key = r.link || (r.source+r.name);
    const isApplied = applied.has(key), isFav = fav.has(key), isHidden = hidden.has(key);
    const cls = isHidden ? "declined" : (isApplied ? "applied" : "");
    const cats = (r.categories||[]).map(c=>'<span class="tag cat">'+catLabel(c)+'</span>').join("");
    const price = r.price ? '<span class="tag price">'+esc(r.price)+'</span>' : "";
    return `<div class="card ${cls}" data-key="${escAttr(key)}">
      <div class="actions">
        <button class="fav ${isFav?'on':''}" title="В избранное">${isFav?'&#9733;':'&#9734;'}</button>
        <button class="apply ${isApplied?'on':''}" title="Откликнулся">${isApplied?'&#10003; Отклик':'Отклик'}</button>
        <button class="hide ${isHidden?'on':''}" title="Скрыть">${isHidden?'&#128065; Вкл':'Скрыть'}</button>
      </div>
      <h3><a href="${escAttr(r.link)}" target="_blank" rel="noopener">${esc(r.name||"(без названия)")}</a></h3>
      <div class="desc">${esc((r.desc||"").slice(0,300))}</div>
      <div class="meta">
        <span class="tag">${esc(r.source)}</span>${cats}${price}
      </div>
    </div>`;
  }).join("");

  cards.querySelectorAll(".card").forEach(card=>{
    const key = card.getAttribute("data-key");
    card.querySelector(".fav").onclick = ()=>{ toggle("fav", key); render(); };
    card.querySelector(".apply").onclick = ()=>{ toggle("applied", key); render(); };
    card.querySelector(".hide").onclick = ()=>{ toggle("hidden", key); render(); };
  });
}

function toggle(name, key){
  const set = new Set(LS.get(name, []));
  set.has(key) ? set.delete(key) : set.add(key);
  LS.set(name, [...set]);
}

async function doExport(format){
  if(!FILTERED.length){ alert("Нет заказов для экспорта. Сначала выполните поиск."); return; }
  const status = document.getElementById("status");
  status.innerHTML = '<span class="spinner"></span>Готовлю файл...';
  try{
    const r = await fetch("/api/export", {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify({format, rows: FILTERED})
    });
    if(!r.ok){ const e = await r.json(); throw new Error(e.error||r.status); }
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const ts = new Date().toISOString().slice(0,16).replace(/[:T]/g,"-");
    a.href = url;
    a.download = "orders_"+ts+(format==="csv"?".csv":".xlsx");
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    status.textContent = "Файл экспортирован ("+FILTERED.length+" заказов)";
  }catch(e){
    status.textContent = "Ошибка экспорта: "+e;
  }
}

function esc(s){ return (s||"").replace(/[&<>]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])); }
function escAttr(s){ return (s||"").replace(/"/g,"&quot;").replace(/[<>]/g,c=>({"<":"&lt;",">":"&gt;"}[c])); }

document.getElementById("go").onclick = search;
document.getElementById("exportXlsx").onclick = ()=>doExport("xlsx");
document.getElementById("exportCsv").onclick = ()=>doExport("csv");
["minPrice","maxPrice","filterText","viewFilter","sortBy"].forEach(id=>
  document.getElementById(id).addEventListener("input", render));
</script>
</body>
</html>
"""


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Сервер запущен: http://{HOST}:{PORT}")
    print("Остановка: Ctrl+C")
    if not IS_HOSTED:
        try:
            webbrowser.open(f"http://127.0.0.1:{PORT}")
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nОстановлено.")
        server.shutdown()


if __name__ == "__main__":
    main()
