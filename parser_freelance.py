# -*- coding: utf-8 -*-
"""
Парсер фриланс-бирж: поиск заказов на разработку трекера задач
(системы управления задачами / канбан / task tracker).

Источники:
  * FL.ru            - https://www.fl.ru/projects/
  * freelance.ru     - https://freelance.ru/task
  * Kwork.ru         - https://kwork.ru/projects
  * Freelancer.com   - https://www.freelancer.com/jobs/
  * Telegram         - публичные каналы через https://t.me/s/<канал>
                       (если t.me блокируется — задайте TG_PROXY)

Логика: биржи не поддерживают фильтр по ключевым словам через URL,
поэтому собираем свежие ленты заказов (несколько страниц) и фильтруем
клиентски по ключевым словам, относящимся к трекерам задач.

Результат: Excel-файл "Заказы_трекер_задач.xlsx" в папке проекта.
"""
import os
import re
import json
import time
import html as htmllib
from datetime import datetime

import requests
from bs4 import BeautifulSoup

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("Установите openpyxl:  python -m pip install openpyxl")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Сколько страниц ленты тянуть с каждой биржи
PAGES = 8
TIMEOUT = (10, 25)
SLEEP = 1.0

# Прокси для Telegram (t.me блокируется у части провайдеров).
# Задайте переменную окружения TG_PROXY, например:
#   set TG_PROXY=socks5h://127.0.0.1:1080
#   set TG_PROXY=http://user:pass@127.0.0.1:8080
TG_PROXY = os.environ.get("TG_PROXY", "").strip()

# Публичные Telegram-каналы с вакансиями/заказами для IT и фриланса.
# Список легко расширить: добавляйте ники без @.
TG_CHANNELS = [
    "ru_freelance",
    "frilans_rabota_vakansii",
    "udalenka_vacansii",
    "forfreelance",
    "freelance_ru",
    "zarabotok_dla_programmistov",
    "PythonJob",
    "job_python",
    "devops_rabota",
    "javascript_jobs",
    "WebDevJobs",
    "prog_job",
]

# Поисковые запросы для бирж с рабочим поиском (freelance.ru, Freelancer)
SEARCH_QUERIES = [
    "трекер", "трекер задач", "таск-трекер", "задачи",
    "управление задачами", "канбан", "приложение для задач",
    "task tracker", "task management", "kanban", "jira", "notion",
    "crm", "автоматизация", "разработка приложения",
]

# --------------------------------------------------------------------------
# Ключевые слова под задачу "разработка трекера задач"
# --------------------------------------------------------------------------
# ПРЯМЫЕ фразы: заказ точно про трекер/управление задачами
STRONG_PHRASES = [
    "трекер задач", "таск-трекер", "таск трекер", "тасктрекер", "таск трекер",
    "task tracker", "tasktracker", "task-tracker",
    "система управления задачами", "системы управления задачами",
    "управление задачами", "управления задачами", "управлению задач",
    "менеджер задач", "менеджера задач", "приложение для задач",
    "список задач", "учет задач", "учёт задач", "трекер проектов",
    "управление проектами и задачами", "постановка задач",
    "kanban", "канбан", "доска задач", "kanban board",
    "task management", "issue tracker", "bug tracker", "баг-трекер",
    "project tracker", "todo list", "todolist", "to-do list",
]
# НАЗВАНИЯ трекеров (нужен ещё контекст разработки/внедрения)
TOOL_TERMS = [
    "jira", "джира", "trello", "трелло", "asana", "асана", "notion",
    "yougile", "youtrack", "wrike", "clickup", "monday.com", "basecamp",
    "redmine", "openproject", "taiga", "kaiten", "weeek", "todoist",
    "ticktick", "битрикс24 задачи",
]
# КОНТЕКСТ разработки/внедрения (нужен для TOOL_TERMS и ядра)
DEV_CONTEXT = [
    "разработ", "приложен", "систем", "софт", "программ", "сервис", "создан",
    "напис", "внедрен", "интеграц", "мобильн", "saas", "фронтенд", "backend",
    "бэкенд", "стартап", "автоматизац", "платформ", "скрипт", "сайт",
    "веб-приложен", "web application", "разработать", "сделать",
]
# ЯДРО: слова, которые сами требуют трекерную связку
CORE_WORDS = [
    "трекер", "tracker", "таск", "канбан", "kanban", "sprint",
]
NEGATIVE_TERMS = [
    "фитнес-трекер", "фитнес трекер", "браслет", "часы трекер", "gps-трекер",
    "gps трекер", "gps-трекинг", "трекинг авто", "трекер для авто",
    "трекер для машин", "трекер для животн", "трекер для собак",
    "трекер для кош", "трекер для детей", "трекер привычек",
    "трекер расходов", "трекер калорий", "трекер сна", "трекер посылок",
    "трекер цен", "трекер акций", "трекер крипт", "трекер звонков",
    "трекер продаж", "трекер времени", "time tracker", "тренажер",
    "тренажёр", "фарм трекер", "игровой трекер", "трекер для игры",
    "торрент трекер", "торрент-трекер", "torrent tracker", "торрент",
    "ticketmaster", "билет", "концерт", "покупка билет",
]


def _phrase_hit(text, phrases):
    return any(p in text for p in phrases)


# для коротких/неоднозначных слов — совпадение по границам слова
def _word_hit(text, words):
    for w in words:
        if re.search(r"(?<![а-яёa-z])" + re.escape(w) + r"(?![а-яёa-z])", text):
            return True
    return False


# совпадение по НАЧАЛУ слова (учёт русских окончаний: портал->портала)
def _stem_hit(text, stems):
    for w in stems:
        if re.search(r"(?<![а-яёa-z])" + re.escape(w), text):
            return True
    return False


def matches_tracker(name, desc):
    title = (name or "").lower()
    # убираем URL, чтобы названия сервисов из ссылок не давали ложных срабатываний
    desc_no_url = re.sub(r"https?://\S+", " ", desc or "")
    full = f"{name} {desc_no_url}".lower()
    if any(n in full for n in NEGATIVE_TERMS):
        return False
    # 1) прямые фразы в заголовке или описании
    if _phrase_hit(full, STRONG_PHRASES):
        return True
    # 2) конкретный инструмент-трекер в ЗАГОЛОВКЕ
    if _word_hit(title, TOOL_TERMS):
        return True
    # 3) ядро трекера в ЗАГОЛОВКЕ + контекст разработки в любом месте
    if _word_hit(title, CORE_WORDS) and _phrase_hit(full, DEV_CONTEXT):
        return True
    return False


# Смежные заказы: разработка веб/мобильных приложений, автоматизация,
# CRM/ERP, дашборды, системы учёта — потенциально близко к "трекеру".
RELATED_TITLE = [
    "приложение", "приложени", "систем", "сервис", "автоматизац", "crm", "erp",
    "dashboard", "дашборд", "панель", "личный кабинет", "личного кабинета",
    "веб-прилож", "web-прилож", "платформ", "portal", "портал", "софт",
    "программ", "учёт", "учет", "база данных", "dasboard", "админ",
    "управлен", "трекер", "tracker", "kanban", "канбан", "задач", "task",
    "мобильн", "saas", "стат", "отчет", "отчёт", "api",
]
RELATED_DEV = [
    "разработ", "создан", "напис", "внедрен", "внедрить", "автоматизац",
    "интеграц", "доработ", "настро", "api", "backend", "бэкенд", "фронтенд",
    "python", "django", "react", "vue", "node", "laravel", "php", "flutter",
]


def matches_related(name, desc):
    """Смежные заказы: разработка приложений/систем/автоматизации.

    Отбор по ЗАГОЛОВКУ: должен присутствовать конкретный
    программный артефакт (приложение, система, бот, API, CRM, ...)
    и признак разработки/технологии.
    """
    title = (name or "").lower()
    if any(n in title for n in NEGATIVE_TERMS):
        return False
    artifact_stems = [
        "приложени", "систем", "crm", "erp", "дашборд", "dashboard",
        "платформ", "портал", "парсер", "parser", "api", "трекер", "tracker",
        "kanban", "канбан", "задач", "task", "бот", "софт", "программ",
        "автоматизац", "интеграц", "saas", "микросервис", "backend",
        "бэкенд", "фронтенд", "сервис", "кабинет", "сайт", "лендинг",
    ]
    artifact_phrases = [
        "web-прилож", "веб-прилож", "telegram-бот", "телеграм-бот",
        "чат-бот", "база данных", "1с", "битрикс",
    ]
    dev_verbs = [
        "разработ", "разработать", "создать", "написать", "внедрить",
        "внедрен", "автоматизировать", "интегрировать", "доработать",
        "доработ", "развернуть", "развертыв", "запрограммировать",
        "спроектировать", "интеграц", "версия", "функционал", "модул",
        "баг", "ошибк", "перенести", "перенос", "подключ", "настро",
    ]
    tech = [
        "python", "django", "flask", "react", "vue", "node", "laravel", "php",
        "flutter", "java", "kotlin", "swift", "typescript", "golang",
        "openai", "gpt", "нейросет", "1с", "битрикс", "make", "zapier",
        "mysql", "postgres", "mongodb", "firebase",
    ]
    full = full_lower(desc, title)
    has_artifact = (_stem_hit(title, artifact_stems)
                    or _phrase_hit(title, artifact_phrases))
    dev_in_title = _phrase_hit(title, dev_verbs)
    has_signal = (dev_in_title or _word_hit(full, tech)
                  or _phrase_hit(full, dev_verbs))
    if has_artifact and has_signal:
        return True
    return False


def full_lower(desc, title):
    return f"{title} {(desc or '').lower()}"


# --------------------------------------------------------------------------
# Категории заказов (для интерактивной таблицы)
# --------------------------------------------------------------------------
CATEGORIES = {
    "tracker": "Трекеры / канбан",
    "automation": "Автоматизация / боты",
    "development": "Разработка приложений / сайтов",
    "crm": "CRM / ERP / 1С",
}

_CAT_KEYWORDS = {
    "tracker": [
        "трекер", "tracker", "канбан", "kanban", "jira", "джира", "trello",
        "трелло", "youtrack", "redmine", "clickup", "kaiten", "weeek",
        "asana", "wrike", "управление задачами", "управления задачами",
        "task management", "доска задач", "бэклог", "backlog", "sprint",
        "спринт", "scrum", "скрам", "notion",
    ],
    "automation": [
        "автоматизац", "бот", "парсер", "parser", "интеграц", "api",
        "make", "zapier", "n8n", "webhook", "вебхук", "скрипт", "cron",
        "выгрузк", "синхронизац", "rpa", "ai-агент", "нейросет", "gpt",
    ],
    "crm": [
        "crm", "erp", "amo crm", "amocrm", "амо", "битрикс24", "битрикс",
        "1с", "1c", "срм", "salesforce", "hubspot", "воронк", "личный кабинет",
        "система лояльности", "учёт", "учет",
    ],
    "development": [
        "разработ", "разработать", "создать", "написать", "приложени",
        "сайт", "лендинг", "веб", "web", "мобильн", "android", "ios",
        "flutter", "react", "vue", "django", "laravel", "python", "php",
        "node", "backend", "бэкенд", "фронтенд", "платформ", "сервис",
        "софт", "программ", "портал", "верстк",
    ],
}


def classify(name, desc):
    """Возвращает список категорий, к которым относится заказ."""
    text = full_lower(desc, name)
    title = (name or "").lower()
    cats = []
    for cat, words in _CAT_KEYWORDS.items():
        if cat == "tracker":
            if matches_tracker(name, desc):
                cats.append(cat)
                continue
        if _stem_hit(text, words) or _stem_hit(title, words):
            cats.append(cat)
    # если ничего не подошло, но это явно заказ разработки — development
    if not cats and matches_related(name, desc):
        cats.append("development")
    return cats


def matches_query(name, desc, query):
    """Проверяет, попадает ли заказ под произвольный поисковый запрос."""
    if not query:
        return True
    text = full_lower(desc, name)
    terms = [t.strip().lower() for t in re.split(r"[\s,]+", query) if t.strip()]
    return all(t in text for t in terms)


def _proxies(url):
    """Прокси применяем только к Telegram, остальное ходит напрямую."""
    if TG_PROXY and "t.me" in url:
        return {"http": TG_PROXY, "https": TG_PROXY}
    return None


def get(url, **kw):
    last = None
    proxies = _proxies(url)
    if proxies:
        kw.setdefault("proxies", proxies)
    for _ in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, **kw)
            if r.status_code == 200:
                return r
            last = r
        except Exception as e:
            last = e
        time.sleep(2)
    if isinstance(last, Exception):
        print(f"    [!] ошибка {url}: {str(last)[:80]}")
    elif last is not None:
        print(f"    [!] статус {last.status_code} {url[:80]}")
    return None


def clean(text):
    if not text:
        return ""
    text = htmllib.unescape(str(text))
    text = re.sub(r"\s+", " ", text).strip()
    return text


# --------------------------------------------------------------------------
# FL.ru  (лента /projects/?kind=1&page=N)
# --------------------------------------------------------------------------
def parse_flru():
    rows = []
    for p in range(1, PAGES + 1):
        url = "https://www.fl.ru/projects/?kind=1" + (f"&page={p}" if p > 1 else "")
        r = get(url)
        if not r:
            break
        soup = BeautifulSoup(r.text, "html.parser")
        grids = soup.select("div.b-post__grid")
        if not grids:
            break
        for g in grids:
            title_a = g.select_one("h2 a")
            if not title_a:
                continue
            name = clean(title_a.get_text())
            href = title_a.get("href", "")
            link = "https://www.fl.ru" + href if href.startswith("/") else href
            desc_el = g.select_one("div.b-post__txt.text-5")
            desc = clean(desc_el.get_text()) if desc_el else ""
            price_el = g.select_one("div.b-post__price span")
            price = clean(price_el.get_text()) if price_el else ""
            rows.append(dict(source="FL.ru", name=name, desc=desc,
                             price=price, link=link))
        time.sleep(SLEEP)
    return rows


# --------------------------------------------------------------------------
# freelance.ru  (лента /task + поиск ?q= — поиск работает)
# --------------------------------------------------------------------------
def _freelance_parse_page(html_text, rows, seen):
    soup = BeautifulSoup(html_text, "html.parser")
    cards = soup.select("article.task-card")
    for card in cards:
        title_a = card.select_one("a.task-card__title-link")
        if not title_a:
            continue
        name = clean(title_a.get_text())
        href = title_a.get("href", "")
        link = "https://freelance.ru" + href if href.startswith("/") else href
        if link in seen:
            continue
        seen.add(link)
        desc_el = card.select_one("p.task-card__desc")
        desc = clean(desc_el.get_text()) if desc_el else ""
        budget_el = card.select_one("div.task-card__budget")
        price = clean(budget_el.get_text()) if budget_el else ""
        cat_el = card.select_one("span.task-chip--cat")
        category = clean(cat_el.get_text()) if cat_el else ""
        rows.append(dict(source="freelance.ru", name=name,
                         desc=(desc + " " + category).strip(),
                         price=price, link=link))
    return len(cards)


def parse_freelance_ru():
    rows = []
    seen = set()
    # 1) поиск по запросам
    for q in SEARCH_QUERIES:
        url = "https://freelance.ru/task?q=" + requests.utils.quote(q)
        r = get(url)
        if r:
            _freelance_parse_page(r.text, rows, seen)
        time.sleep(SLEEP)
    # 2) общая лента
    for p in range(1, PAGES + 1):
        url = "https://freelance.ru/task" + (f"?page={p}" if p > 1 else "")
        r = get(url)
        if not r:
            break
        if _freelance_parse_page(r.text, rows, seen) == 0:
            break
        time.sleep(SLEEP)
    return rows


# --------------------------------------------------------------------------
# Kwork.ru  (лента /projects?c=all&page=N, данные в window.stateData)
# --------------------------------------------------------------------------
def _extract_state_data(html_text):
    m = re.search(r"window\.stateData\s*=\s*", html_text)
    if not m:
        return None
    start = m.end()
    depth = 0
    i = start
    instr = False
    esc = False
    while i < len(html_text):
        c = html_text[i]
        if instr:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                instr = False
        else:
            if c == '"':
                instr = True
            elif c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    break
        i += 1
    try:
        return json.loads(html_text[start:i + 1])
    except json.JSONDecodeError:
        return None


def parse_kwork():
    rows = []
    for p in range(1, PAGES + 1):
        url = "https://kwork.ru/projects?c=all" + (f"&page={p}" if p > 1 else "")
        r = get(url)
        if not r:
            break
        data = _extract_state_data(r.text)
        if not data:
            break
        items = (data.get("wantsListData", {}).get("pagination", {}) or {}).get("data") or []
        if not items:
            break
        for it in items:
            name = clean(it.get("name", ""))
            desc = clean(it.get("description", ""))
            price = it.get("priceLimit") or it.get("possiblePriceLimit") or ""
            if price:
                try:
                    price = f"{float(price):,.0f} ₽".replace(",", " ")
                except (ValueError, TypeError):
                    price = str(price)
            wid = it.get("id")
            link = (f"https://kwork.ru/projects/{wid}"
                    if wid else "https://kwork.ru/projects")
            rows.append(dict(source="Kwork.ru", name=name, desc=desc,
                             price=price, link=link))
        time.sleep(SLEEP)
    return rows


# --------------------------------------------------------------------------
# Freelancer.com  (лента /jobs/, фильтр клиентский)
# --------------------------------------------------------------------------
def parse_freelancer():
    rows = []
    seen = set()
    for p in range(1, PAGES + 1):
        url = "https://www.freelancer.com/jobs/" + (str(p) if p > 1 else "")
        r = get(url)
        if not r:
            break
        soup = BeautifulSoup(r.text, "html.parser")
        cards = soup.select("div.JobSearchCard-item")
        if not cards:
            break
        new = 0
        for card in cards:
            title_a = card.select_one("a.JobSearchCard-primary-heading-link")
            if not title_a:
                continue
            name = clean(title_a.get_text())
            href = title_a.get("href", "")
            if href in seen:
                continue
            seen.add(href)
            new += 1
            link = "https://www.freelancer.com" + href if href.startswith("/") else href
            desc_el = card.select_one("p.JobSearchCard-primary-description")
            desc = clean(desc_el.get_text()) if desc_el else ""
            price_el = card.select_one("div.JobSearchCard-secondary-price")
            price = clean(price_el.get_text()) if price_el else ""
            rows.append(dict(source="Freelancer.com", name=name, desc=desc,
                             price=price, link=link))
        if new == 0:
            break
        time.sleep(SLEEP)
    return rows


# --------------------------------------------------------------------------
# Telegram  (публичные каналы через веб-превью t.me/s/<канал>)
# --------------------------------------------------------------------------
# Примечание: t.me может быть заблокирован у провайдера — тогда нужен
# прокси в переменной окружения TG_PROXY (см. выше).
def _tg_post_title(text):
    """Первая непустая строка поста (обычно заголовок вакансии)."""
    for line in text.splitlines():
        line = clean(line)
        if len(line) >= 5:
            return line[:120]
    return clean(text)[:120]


def parse_telegram():
    rows = []
    seen = set()
    for nick in TG_CHANNELS:
        url = f"https://t.me/s/{nick}"
        r = get(url)
        if not r:
            continue
        soup = BeautifulSoup(r.text, "html.parser")
        posts = soup.select("div.tgme_widget_message")
        for post in posts:
            text_el = post.select_one("div.tgme_widget_message_text")
            if not text_el:
                continue
            text = clean(text_el.get_text(" "))
            if len(text) < 15:
                continue
            post_id = None
            holder = post
            while holder is not None:
                if holder.get("data-post"):
                    post_id = holder["data-post"].split("/")[-1]
                    break
                holder = holder.parent
            if not post_id:
                a = post.select_one("a.tgme_widget_message_date")
                if a and a.get("href"):
                    post_id = a["href"].rstrip("/").split("/")[-1]
            link = f"https://t.me/{nick}/{post_id}" if post_id else url
            if link in seen:
                continue
            seen.add(link)
            name = _tg_post_title(text)
            price = _tg_extract_price(text)
            rows.append(dict(source=f"TG:@{nick}", name=name, desc=text,
                             price=price, link=link))
        time.sleep(SLEEP)
    return rows


def _tg_extract_price(text):
    """Пытаемся вытащить бюджет из текста поста."""
    m = re.search(r"(\d[\d\s]{2,})\s*(?:руб|₽|р\.|rub)", text, re.IGNORECASE)
    if m:
        return f"{clean(m.group(1))} ₽"
    m = re.search(r"\$\s?(\d[\d\s,\.]{1,})", text)
    if m:
        return f"$ {clean(m.group(1))}"
    if re.search(r"бюджет|ставка|оплат|гонорар", text, re.IGNORECASE):
        return "уточняется"
    return ""


SOURCES = [
    ("FL.ru", parse_flru),
    ("freelance.ru", parse_freelance_ru),
    ("Kwork.ru", parse_kwork),
    ("Freelancer.com", parse_freelancer),
    ("Telegram", parse_telegram),
]


def _write_sheet(ws, rows, title, header_fill, header_font, border,
                 collected, columns):
    ws.title = title
    ws.append(columns)
    for c in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = border
    for idx, row in enumerate(rows, start=1):
        ws.append([idx, row["source"], row["name"], row["desc"],
                   row["price"], row["link"], collected])
    widths = [5, 16, 48, 80, 18, 55, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(columns) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(c in (3, 4, 5, 6)))
        link = ws.cell(row=r, column=6)
        if link.value:
            link.hyperlink = link.value
            link.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[r].height = 45
    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"


def save_excel(tracker_rows, related_rows, path):
    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="2B5A92")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    collected = f"{datetime.now():%d.%m.%Y %H:%M}"
    columns = ["№", "Источник", "Название заказа", "Описание", "Бюджет",
               "Ссылка", "Дата сбора"]

    ws = wb.active
    _write_sheet(ws, tracker_rows, "Трекер задач", header_fill, header_font,
                 border, collected, columns)
    ws2 = wb.create_sheet("Смежные заказы")
    _write_sheet(ws2, related_rows, "Смежные заказы", header_fill,
                 header_font, border, collected, columns)

    ws3 = wb.create_sheet("Сводка")
    ws3.append(["Категория", "Найдено заказов"])
    for c in range(1, 3):
        cell = ws3.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    ws3.append(["Трекеры задач (точное совпадение)", len(tracker_rows)])
    ws3.append(["Смежные заказы (разработка/автоматизация)", len(related_rows)])
    ws3.append(["ИТОГО", len(tracker_rows) + len(related_rows)])
    for r in range(2, ws3.max_row + 1):
        ws3.cell(row=r, column=2).alignment = Alignment(horizontal="center")
    ws3.cell(row=ws3.max_row, column=1).font = Font(bold=True)
    ws3.cell(row=ws3.max_row, column=2).font = Font(bold=True)

    ws3.append([])
    ws3.append(["Источники:", "FL.ru, freelance.ru, Kwork.ru, Freelancer.com"])
    ws3.append(["Дата сбора:", collected])
    ws3.append(["Точный критерий:",
                "трекер задач / система управления задачами / канбан / "
                "task tracker / Jira / Notion / кнопки и т.п."])
    ws3.append(["Смежный критерий:",
                "разработка веб/мобильных приложений, систем, автоматизация, "
                "CRM/ERP, личные кабинеты"])
    ws3.append(["Примечание:",
                "биржи не дают фильтр по слову через URL — собирается лента "
                "заказов, затем фильтрация по ключевым словам на стороне скрипта"])
    for r in range(ws3.max_row - 5, ws3.max_row + 1):
        ws3.cell(row=r, column=1).font = Font(bold=True)
        ws3.cell(row=r, column=2).alignment = Alignment(wrap_text=True,
                                                        vertical="top")
    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 70

    wb.save(path)


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    all_rows = []
    for site_name, parser in SOURCES:
        print(f"\n=== {site_name} ===")
        try:
            found = parser()
        except Exception as e:
            print(f"  [!] сбой парсера: {e}")
            found = []
        print(f"  собрано заказов из ленты: {len(found)}")
        all_rows.extend(found)

    # дедупликация по ссылке
    seen = set()
    uniq = []
    for row in all_rows:
        key = row["link"] or (row["source"], row["name"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(row)

    # классификация
    tracker_rows = [r for r in uniq if matches_tracker(r["name"], r["desc"])]
    tracker_links = {r["link"] for r in tracker_rows}
    related_rows = [r for r in uniq
                    if r["link"] not in tracker_links
                    and matches_related(r["name"], r["desc"])]

    tracker_rows.sort(key=lambda r: (r["source"], r["name"]))
    related_rows.sort(key=lambda r: (r["source"], r["name"]))

    xlsx_path = os.path.join(out_dir, "Заказы_трекер_задач.xlsx")
    save_excel(tracker_rows, related_rows, xlsx_path)

    print(f"\nЗаказов про трекеры задач: {len(tracker_rows)}")
    print(f"Смежных заказов: {len(related_rows)}")
    print(f"Файл сохранён: {xlsx_path}")
    return xlsx_path


if __name__ == "__main__":
    main()
