"""
GGPL新宿 ランキングダッシュボード 生成スクリプト
GitHub Actionsで実行し、index.htmlを生成してGitHub Pagesで公開する。
"""

import os
import sys
import json
import re
import base64
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.service_account import Credentials

JST = timezone(timedelta(hours=9))

RING_SHEET_ID = os.environ.get("RING_SHEET_ID", "")
TONAME_SHEET_ID = os.environ.get("TONAME_SHEET_ID", "")
SA_JSON_STR = os.environ.get("SERVICE_ACCOUNT_JSON", "")

SEASON_TAB_PATTERN = re.compile(r"^\d{1,2}/\d{1,2}-\d{1,2}/\d{1,2}$")
MONTH_TAB_PATTERN = re.compile(r"^\d{4}年\d{1,2}月$|^\d{1,2}月$|^\d{4}/\d{1,2}$")
DAILY_TAB_NAME = "デイリー運用"


def get_gc():
    sa_json = SA_JSON_STR.strip()
    try:
        sa_info = json.loads(sa_json)
    except (json.JSONDecodeError, ValueError):
        try:
            fixed = re.sub(
                r'("private_key"\s*:\s*")(.*?)(")',
                lambda m: m.group(1) + m.group(2).replace('\n', '\\n').replace('\r', '') + m.group(3),
                sa_json,
                flags=re.DOTALL
            )
            sa_info = json.loads(fixed)
        except (json.JSONDecodeError, ValueError):
            sa_info = json.loads(base64.b64decode(sa_json.encode()).decode('utf-8'))
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def _detect_rank(cell):
    """Returns rank number 1-10 or None. Handles 1st/2nd/3rd/4th/1位/1 etc."""
    r = cell.strip().replace(" ", "")
    for i in range(1, 11):
        suffix = "st" if i == 1 else "nd" if i == 2 else "rd" if i == 3 else "th"
        if r in (f"{i}{suffix}", f"{i}位", str(i)):
            return i
    return None


def get_ranking_from_sheet(ws):
    rows = ws.get_all_values()
    print(f"    {ws.title}: {len(rows)}行読み込み")
    ranking = []
    for row in rows:
        if len(row) < 2:
            continue
        rank_num = _detect_rank(str(row[0]))
        if rank_num is None:
            continue

        # シートレイアウト: A=順位, B=ポイント, C=名前
        # (以前 B=名前, C=ポイントと逆に読んでいたため修正)
        points = str(row[1]).strip() if len(row) > 1 else ""
        name   = str(row[2]).strip() if len(row) > 2 else ""

        # B列が数字でなくC列が数字なら逆レイアウト (フォールバック)
        pts_b = re.sub(r"[^\d]", "", points)
        pts_c = re.sub(r"[^\d]", "", name)
        if not pts_b and pts_c:
            name, points = points, name
            pts_b = pts_c

        if not name or name in ("#N/A", ""):
            continue
        if not pts_b or int(pts_b) == 0:
            continue

        suffix = "st" if rank_num == 1 else "nd" if rank_num == 2 else "rd" if rank_num == 3 else "th"
        ranking.append({"rank": f"{rank_num}{suffix}", "name": name, "points": points})

    print(f"    → {len(ranking)}件のランキングデータ")
    return ranking


def fetch_daily_ranking(gc, sheet_id):
    try:
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(DAILY_TAB_NAME)
    except Exception as e:
        print(f"  デイリータブ読み込みエラー: {e}")
        return []

    rows = ws.get_all_values()
    by_date = {}
    for row in rows:
        if len(row) < 10:
            continue
        name = str(row[2]).strip()
        points_str = str(row[8]).strip().replace(",", "").replace("-", "").replace("−", "")
        date_str = str(row[9]).strip()
        minus = str(row[8]).strip().startswith("-") or str(row[8]).strip().startswith("−") or str(row[8]).strip().startswith("\\-")

        if not name or not date_str or not re.match(r"\d{4}/\d{2}/\d{2}", date_str):
            continue
        try:
            points = int(float(points_str))
        except ValueError:
            continue
        if minus:
            points = -points
        if points <= 0:
            continue

        if date_str not in by_date:
            by_date[date_str] = {}
        by_date[date_str][name] = by_date[date_str].get(name, 0) + points

    seasons = []
    for date_str in sorted(by_date.keys(), reverse=True):
        players = sorted(by_date[date_str].items(), key=lambda x: x[1], reverse=True)
        ranking = []
        for i, (name, pts) in enumerate(players):
            suffix = ["st", "nd", "rd"][i] if i < 3 else "th"
            ranking.append({"rank": f"{i+1}{suffix}", "name": name, "points": f"{pts:,}pt"})
        m = int(date_str[5:7])
        d = int(date_str[8:10])
        seasons.append({"title": f"{m}/{d}", "ranking": ranking})

    print(f"  デイリー: {len(seasons)}日分のデータ読み込み完了")
    return seasons


def fetch_seasons(gc, sheet_id, tab_pattern):
    try:
        sh = gc.open_by_key(sheet_id)
    except Exception as e:
        print(f"  シート読み込みエラー: {e}")
        return []

    seasons = []
    for ws in sh.worksheets():
        title = ws.title
        if tab_pattern.match(title):
            print(f"  タブ読み込み中: {title}")
            ranking = get_ranking_from_sheet(ws)
            seasons.append({"title": title, "ranking": ranking})
    return seasons


# ── HTML generation ──────────────────────────────────────────────────────────

def _podium_html(ranking, color_key):
    """Top-3 podium + rest table."""
    if not ranking:
        return '<p class="no-data">データなし</p>'

    def card_suit(i):
        return ["♠", "♥", "♦"][i] if i < 3 else ""

    top3 = ranking[:3]
    while len(top3) < 3:
        top3.append(None)
    first, second, third = top3[0], top3[1], top3[2]

    def place_html(player, place, suit, block_class, block_num, crown=False):
        h = f'<div class="podium-place {place}">'
        if player:
            if crown:
                h += '<div class="crown-icon">♛</div>'
            h += f'<div class="podium-suit {color_key}-suit">{suit}</div>'
            h += f'<div class="podium-name">{player["name"]}</div>'
            h += f'<div class="podium-pts {color_key}-pts">{player["points"]}</div>'
        else:
            h += '<div class="podium-empty"></div>'
        h += f'<div class="podium-block {block_class} {color_key}-block">{block_num}</div>'
        h += '</div>'
        return h

    html = '<div class="podium-wrap">'
    html += place_html(second, "second", "♥", "block-silver", "2")
    html += place_html(first,  "first",  "♠", "block-gold",   "1", crown=True)
    html += place_html(third,  "third",  "♦", "block-bronze",  "3")
    html += '</div>'

    rest = ranking[3:]
    if rest:
        html += '<table class="rest-table"><tbody>'
        for r in rest:
            html += (f'<tr>'
                     f'<td class="rest-rank">{r["rank"]}</td>'
                     f'<td class="rest-name">{r["name"]}</td>'
                     f'<td class="rest-pts">{r["points"]}</td>'
                     f'</tr>')
        html += '</tbody></table>'

    return html


def _season_tabs_html(seasons, color_key, no_data_msg):
    if not seasons:
        return f'<p class="no-data">{no_data_msg}</p>'

    tabs = '<div class="season-tabs">'
    for i, s in enumerate(seasons):
        active = "active" if i == 0 else ""
        tabs += (f'<button class="season-btn {active} {color_key}-tab-btn" '
                 f'onclick="switchSeason(this,\'{color_key}-sp-{i}\')">{s["title"]}</button>')
    tabs += '</div>'

    panels = ''
    for i, s in enumerate(seasons):
        active = "active" if i == 0 else ""
        panels += f'<div class="season-panel {active}" id="{color_key}-sp-{i}">'
        panels += _podium_html(s["ranking"], color_key)
        panels += '</div>'

    return tabs + panels


def generate_html(ring_seasons, daily_seasons, toname_seasons, updated_at):
    ring_html   = _season_tabs_html(ring_seasons,   "ring",   "シーズンデータなし")
    daily_html  = _season_tabs_html(daily_seasons,  "daily",  "デイリーデータなし")
    toname_html = _season_tabs_html(toname_seasons, "toname", "シーズンデータなし")

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GGP LIVE SHINJUKU - RANKING</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@700;900&family=Noto+Sans+JP:wght@400;700;900&display=swap" rel="stylesheet">
<style>
/* ── Base ── */
:root {{
  --gold:      #d4af37;
  --gold-lt:   #f5d76e;
  --gold-dk:   #8b6914;
  --ring-hi:   #e84545;
  --ring-dk:   #6b0000;
  --daily-hi:  #52c97a;
  --daily-dk:  #1a4a2e;
  --toname-hi: #5b9bd5;
  --toname-dk: #0d2b5e;
  --bg:        #07070f;
  --surface:   #10101c;
  --border:    rgba(212,175,55,0.25);
  --text:      #e8e0d0;
  --muted:     #888;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{
  font-family: 'Noto Sans JP', sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  overflow-x: hidden;
}}

/* ── Header ── */
header {{
  position: relative;
  text-align: center;
  padding: 36px 20px 28px;
  background: linear-gradient(180deg, #000 0%, #0a0a18 60%, #07070f 100%);
  border-bottom: 1px solid var(--gold);
  overflow: hidden;
}}
header::before {{
  content: '';
  position: absolute; inset: 0;
  background:
    radial-gradient(ellipse 60% 40% at 50% 0%, rgba(212,175,55,0.12) 0%, transparent 70%);
  pointer-events: none;
}}
.logo-suit-left, .logo-suit-right {{
  position: absolute;
  top: 50%; transform: translateY(-50%);
  font-size: 4rem;
  opacity: 0.07;
  color: var(--gold);
  font-family: serif;
}}
.logo-suit-left {{ left: 24px; }}
.logo-suit-right {{ right: 24px; }}

.logo-ggp {{
  font-family: 'Cinzel', serif;
  font-size: clamp(1.6rem, 5vw, 2.8rem);
  font-weight: 900;
  letter-spacing: 0.18em;
  color: var(--gold);
  text-shadow: 0 0 40px rgba(212,175,55,0.6), 0 2px 4px #000;
  line-height: 1;
}}
.logo-live {{
  font-family: 'Cinzel', serif;
  font-size: clamp(0.55rem, 1.8vw, 0.85rem);
  letter-spacing: 0.5em;
  color: rgba(212,175,55,0.7);
  margin-top: 4px;
  text-transform: uppercase;
}}
.logo-divider {{
  width: 120px; height: 1px;
  background: linear-gradient(90deg, transparent, var(--gold), transparent);
  margin: 12px auto;
}}
.logo-ranking {{
  font-family: 'Cinzel', serif;
  font-size: clamp(0.6rem, 2vw, 0.9rem);
  letter-spacing: 0.4em;
  color: var(--text);
  opacity: 0.8;
}}
.updated {{
  font-size: 0.7rem;
  color: var(--muted);
  margin-top: 10px;
  letter-spacing: 0.05em;
}}

/* ── Category nav ── */
.cat-nav {{
  display: flex;
  justify-content: center;
  gap: 10px;
  padding: 24px 16px 0;
  flex-wrap: wrap;
}}
.cat-btn {{
  position: relative;
  padding: 12px 28px;
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 2px;
  font-family: 'Noto Sans JP', sans-serif;
  font-size: 0.82rem;
  font-weight: 700;
  letter-spacing: 0.08em;
  cursor: pointer;
  background: var(--surface);
  color: var(--muted);
  transition: all 0.25s;
  clip-path: polygon(8px 0%, 100% 0%, calc(100% - 8px) 100%, 0% 100%);
}}
.cat-btn::after {{
  content: '';
  position: absolute; bottom: 0; left: 0; right: 0; height: 2px;
  background: transparent;
  transition: background 0.25s;
}}
.cat-btn.ring-cat.active   {{ background: rgba(232,69,69,0.15);  color: var(--ring-hi);   border-color: var(--ring-hi); }}
.cat-btn.daily-cat.active  {{ background: rgba(82,201,122,0.15); color: var(--daily-hi);  border-color: var(--daily-hi); }}
.cat-btn.toname-cat.active {{ background: rgba(91,155,213,0.15); color: var(--toname-hi); border-color: var(--toname-hi); }}
.cat-btn:hover {{ opacity: 0.85; }}

/* ── Category panels ── */
.category-panel {{ display: none; }}
.category-panel.active {{ display: block; }}

/* ── Hero banner ── */
.hero {{
  position: relative;
  padding: 32px 20px 24px;
  text-align: center;
  overflow: hidden;
}}
.ring-hero   {{ background: linear-gradient(160deg, #1a0000 0%, #3d0808 40%, #6b0000 70%, #1a0000 100%); }}
.daily-hero  {{ background: linear-gradient(160deg, #010f06 0%, #0d3320 40%, #1a5c36 70%, #010f06 100%); }}
.toname-hero {{ background: linear-gradient(160deg, #00040f 0%, #0a1f4e 40%, #0d2b5e 70%, #00040f 100%); }}

.hero::before {{
  content: '';
  position: absolute; inset: 0;
  background:
    radial-gradient(ellipse 70% 60% at 50% 50%, rgba(212,175,55,0.08) 0%, transparent 70%);
}}
.hero-suits {{
  position: absolute; inset: 0;
  display: flex; align-items: center; justify-content: space-around;
  font-size: 5rem; opacity: 0.04; color: var(--gold);
  pointer-events: none; font-family: serif;
  letter-spacing: 0.5em;
}}
.hero-label {{
  font-family: 'Cinzel', serif;
  font-size: clamp(0.55rem, 2vw, 0.75rem);
  letter-spacing: 0.5em;
  color: var(--gold-lt);
  opacity: 0.8;
  margin-bottom: 6px;
}}
.hero-title {{
  font-family: 'Cinzel', serif;
  font-size: clamp(1rem, 3.5vw, 1.6rem);
  font-weight: 900;
  letter-spacing: 0.12em;
  color: var(--gold);
  text-shadow: 0 0 30px rgba(212,175,55,0.5);
  margin-bottom: 16px;
  line-height: 1.3;
}}
.prize-row {{
  display: flex;
  justify-content: center;
  gap: 10px;
  flex-wrap: wrap;
  margin-bottom: 10px;
}}
.prize-item {{
  background: rgba(0,0,0,0.4);
  border: 1px solid rgba(212,175,55,0.3);
  border-radius: 2px;
  padding: 5px 14px;
  font-size: 0.75rem;
  font-weight: 700;
  color: var(--gold-lt);
  letter-spacing: 0.06em;
  clip-path: polygon(6px 0%, 100% 0%, calc(100% - 6px) 100%, 0% 100%);
}}
.badge-row {{
  display: flex;
  justify-content: center;
  gap: 8px;
  flex-wrap: wrap;
  margin-top: 8px;
}}
.badge {{
  display: inline-block;
  background: linear-gradient(135deg, rgba(212,175,55,0.2), rgba(212,175,55,0.05));
  border: 1px solid rgba(212,175,55,0.5);
  border-radius: 20px;
  padding: 4px 14px;
  font-size: 0.7rem;
  font-weight: 700;
  color: var(--gold-lt);
  letter-spacing: 0.05em;
}}

/* ── Season tabs ── */
.season-wrap {{
  max-width: 680px;
  margin: 0 auto;
  padding: 20px 16px 48px;
}}
.season-tabs {{
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 20px;
}}
.season-btn {{
  padding: 5px 13px;
  border-radius: 20px;
  border: 1px solid #333;
  font-size: 0.75rem;
  font-weight: 700;
  cursor: pointer;
  background: #111;
  color: var(--muted);
  transition: all 0.2s;
  letter-spacing: 0.04em;
}}
.ring-tab-btn.active   {{ border-color: var(--ring-hi);   color: var(--ring-hi);   background: rgba(232,69,69,0.1); }}
.daily-tab-btn.active  {{ border-color: var(--daily-hi);  color: var(--daily-hi);  background: rgba(82,201,122,0.1); }}
.toname-tab-btn.active {{ border-color: var(--toname-hi); color: var(--toname-hi); background: rgba(91,155,213,0.1); }}

.season-panel {{ display: none; }}
.season-panel.active {{ display: block; }}

/* ── Podium ── */
.podium-wrap {{
  display: flex;
  align-items: flex-end;
  justify-content: center;
  gap: 8px;
  margin: 0 0 24px;
}}
.podium-place {{
  display: flex;
  flex-direction: column;
  align-items: center;
  flex: 1;
  max-width: 180px;
}}
.podium-place.first  {{ order: 2; }}
.podium-place.second {{ order: 1; }}
.podium-place.third  {{ order: 3; }}

.crown-icon {{
  font-size: 1.3rem;
  margin-bottom: 2px;
  color: var(--gold);
  text-shadow: 0 0 12px var(--gold);
  font-family: serif;
}}
.podium-suit {{
  font-size: 1.8rem;
  font-family: serif;
  line-height: 1;
  margin-bottom: 4px;
}}
.ring-suit   {{ color: #e84545; text-shadow: 0 0 12px rgba(232,69,69,0.6); }}
.daily-suit  {{ color: #52c97a; text-shadow: 0 0 12px rgba(82,201,122,0.6); }}
.toname-suit {{ color: #5b9bd5; text-shadow: 0 0 12px rgba(91,155,213,0.6); }}

.podium-name {{
  font-size: 0.82rem;
  font-weight: 900;
  text-align: center;
  margin: 4px 4px 2px;
  color: var(--text);
  line-height: 1.2;
  word-break: break-all;
}}
.podium-pts {{
  font-size: 0.72rem;
  font-weight: 700;
  margin-bottom: 8px;
  letter-spacing: 0.04em;
}}
.ring-pts   {{ color: var(--ring-hi); }}
.daily-pts  {{ color: var(--daily-hi); }}
.toname-pts {{ color: var(--toname-hi); }}

.podium-empty {{ flex: 1; }}

.podium-block {{
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: 'Cinzel', serif;
  font-size: 1.4rem;
  font-weight: 900;
  border-radius: 3px 3px 0 0;
  border-top: 2px solid transparent;
}}
.block-gold   {{ height: 80px; background: linear-gradient(180deg, #c9a227 0%, #7a5700 100%); border-color: var(--gold-lt); }}
.block-silver {{ height: 58px; background: linear-gradient(180deg, #9a9a9a 0%, #555 100%); border-color: #bbb; }}
.block-bronze {{ height: 42px; background: linear-gradient(180deg, #b87333 0%, #6b3a1f 100%); border-color: #cd7f32; }}

/* color-specific block glow */
.ring-block   {{ box-shadow: inset 0 1px 0 rgba(255,255,255,0.2); }}
.daily-block  {{ box-shadow: inset 0 1px 0 rgba(255,255,255,0.2); }}
.toname-block {{ box-shadow: inset 0 1px 0 rgba(255,255,255,0.2); }}

/* ── Rest table (4th+) ── */
.rest-table {{
  width: 100%;
  border-collapse: collapse;
  margin-top: 4px;
  font-size: 0.88rem;
}}
.rest-table tr {{
  border-bottom: 1px solid rgba(255,255,255,0.05);
  transition: background 0.15s;
}}
.rest-table tr:hover {{ background: rgba(255,255,255,0.03); }}
.rest-rank {{
  padding: 10px 8px;
  width: 52px;
  color: var(--muted);
  font-size: 0.78rem;
  font-weight: 700;
  letter-spacing: 0.04em;
}}
.rest-name {{
  padding: 10px 8px;
  font-weight: 700;
}}
.rest-pts {{
  padding: 10px 8px;
  text-align: right;
  font-weight: 700;
  font-variant-numeric: tabular-nums;
  color: var(--gold);
  white-space: nowrap;
}}

.no-data {{ color: var(--muted); font-size: 0.88rem; padding: 24px 0; text-align: center; }}

@media (max-width: 420px) {{
  .hero-title {{ font-size: 0.95rem; }}
  .podium-block {{ font-size: 1rem; }}
  .podium-name {{ font-size: 0.75rem; }}
}}
</style>
</head>
<body>

<!-- ── Header ── -->
<header>
  <div class="logo-suit-left">♠♥</div>
  <div class="logo-suit-right">♦♣</div>
  <div class="logo-ggp">GGP LIVE SHINJUKU</div>
  <div class="logo-live">GRAND POKER LOUNGE</div>
  <div class="logo-divider"></div>
  <div class="logo-ranking">R A N K I N G</div>
  <p class="updated">最終更新: {updated_at}</p>
</header>

<!-- ── Category nav ── -->
<div class="cat-nav">
  <button class="cat-btn ring-cat active"   onclick="switchCat(this,'ring-panel')">♠ リングポイント</button>
  <button class="cat-btn daily-cat"         onclick="switchCat(this,'daily-panel')">♥ デイリーリング</button>
  <button class="cat-btn toname-cat"        onclick="switchCat(this,'toname-panel')">♦ トナメポイント</button>
</div>

<!-- ── Ring panel ── -->
<div class="category-panel ring-panel active" id="ring-panel">
  <div class="hero ring-hero">
    <div class="hero-suits">♠ ♥ ♦ ♣</div>
    <div class="hero-label">SEASON RANKING</div>
    <div class="hero-title">リングゲームポイントランキング</div>
    <div class="prize-row">
      <span class="prize-item">🥇 1位 &nbsp;30,000コイン</span>
      <span class="prize-item">🥈 2位 &nbsp;15,000コイン</span>
      <span class="prize-item">🥉 3位 &nbsp;10,000コイン</span>
    </div>
    <div class="badge-row">
      <span class="badge">☔ 雨の日ポイント 2倍</span>
    </div>
  </div>
  <div class="season-wrap">
    {ring_html}
  </div>
</div>

<!-- ── Daily panel ── -->
<div class="category-panel daily-panel" id="daily-panel">
  <div class="hero daily-hero">
    <div class="hero-suits">♠ ♥ ♦ ♣</div>
    <div class="hero-label">DAILY RANKING</div>
    <div class="hero-title">デイリーリングポイントランキング</div>
    <div class="prize-row">
      <span class="prize-item">🥇 1位 &nbsp;5,000コイン</span>
      <span class="prize-item">🥈 2位 &nbsp;3,000コイン</span>
    </div>
  </div>
  <div class="season-wrap">
    {daily_html}
  </div>
</div>

<!-- ── Tournament panel ── -->
<div class="category-panel toname-panel" id="toname-panel">
  <div class="hero toname-hero">
    <div class="hero-suits">♠ ♥ ♦ ♣</div>
    <div class="hero-label">MONTHLY TOURNAMENT</div>
    <div class="hero-title">トナメポイントランキング</div>
    <div class="prize-row">
      <span class="prize-item">🥇 1位 &nbsp;30,000コイン</span>
      <span class="prize-item">🥈 2位 &nbsp;15,000コイン</span>
      <span class="prize-item">🥉 3位 &nbsp;10,000コイン</span>
    </div>
  </div>
  <div class="season-wrap">
    {toname_html}
  </div>
</div>

<script>
function switchCat(btn, id) {{
  document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.category-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  var el = document.getElementById(id);
  if (el) el.classList.add('active');
}}
function switchSeason(btn, id) {{
  var wrap = btn.closest('.season-wrap');
  wrap.querySelectorAll('.season-btn').forEach(b => b.classList.remove('active'));
  wrap.querySelectorAll('.season-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  var el = document.getElementById(id);
  if (el) el.classList.add('active');
}}
</script>
</body>
</html>"""


def main():
    if not SA_JSON_STR:
        print("❌ SERVICE_ACCOUNT_JSON が設定されていません")
        sys.exit(1)

    gc = get_gc()
    updated_at = datetime.now(JST).strftime("%Y/%m/%d %H:%M JST")

    print("リングポイントランキング読み込み中...")
    ring_seasons = fetch_seasons(gc, RING_SHEET_ID, SEASON_TAB_PATTERN) if RING_SHEET_ID else []

    print("デイリーリングランキング読み込み中...")
    daily_seasons = fetch_daily_ranking(gc, RING_SHEET_ID) if RING_SHEET_ID else []

    print("トナメポイントランキング読み込み中...")
    toname_seasons = fetch_seasons(gc, TONAME_SHEET_ID, MONTH_TAB_PATTERN) if TONAME_SHEET_ID else []

    print("HTML生成中...")
    html = generate_html(ring_seasons, daily_seasons, toname_seasons, updated_at)

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ index.html 生成完了！")


if __name__ == "__main__":
    main()
