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
        # private_key内の実際の改行を\nエスケープに変換して再試行
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


def get_ranking_from_sheet(ws):
    rows = ws.get_all_values()
    ranking = []
    for row in rows:
        if len(row) < 2:
            continue
        rank_cell = str(row[0]).strip()
        name = str(row[1]).strip() if len(row) > 1 else ""
        points = str(row[2]).strip() if len(row) > 2 else ""
        if rank_cell in ("1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th") or \
           rank_cell in ("1位", "2位", "3位", "4位", "5位", "6位", "7位", "8位", "9位", "10位"):
            if not name or name in ("#N/A", ""):
                continue
            pts_num = re.sub(r"[^\d]", "", points)
            if not pts_num or int(pts_num) == 0:
                continue
            ranking.append({
                "rank": rank_cell,
                "name": name,
                "points": points
            })
    return ranking


def fetch_daily_ranking(gc, sheet_id):
    try:
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(DAILY_TAB_NAME)
    except Exception as e:
        print(f"  デイリータブ読み込みエラー: {e}")
        return []

    rows = ws.get_all_values()
    # 日付ごとにデータを集計
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
        # 同日に同じプレイヤーが複数行ある場合は合算
        by_date[date_str][name] = by_date[date_str].get(name, 0) + points

    # 日付降順でシーズンリストを作成
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


def generate_html(ring_seasons, daily_seasons, toname_seasons, updated_at):
    def season_tabs_html(seasons, color_class, no_data_msg):
        if not seasons:
            return f'<p class="no-data">{no_data_msg}</p>'

        tabs_html = '<div class="season-tabs">'
        for i, s in enumerate(seasons):
            active = "active" if i == 0 else ""
            tabs_html += f'<button class="season-btn {active}" onclick="switchSeason(this, \'{color_class}-{i}\')">{s["title"]}</button>'
        tabs_html += "</div>"

        panels_html = ""
        for i, s in enumerate(seasons):
            active = "active" if i == 0 else ""
            panels_html += f'<div class="season-panel {active}" id="{color_class}-{i}">'
            if not s["ranking"]:
                panels_html += '<p class="no-data">データなし</p>'
            else:
                panels_html += '<table class="ranking-table"><thead><tr><th>順位</th><th>名前</th><th>ポイント</th></tr></thead><tbody>'
                for j, r in enumerate(s["ranking"]):
                    medal = ""
                    if j == 0: medal = '<span class="medal gold">🥇</span>'
                    elif j == 1: medal = '<span class="medal silver">🥈</span>'
                    elif j == 2: medal = '<span class="medal bronze">🥉</span>'
                    row_class = f"rank-{j+1}" if j < 3 else ""
                    panels_html += f'<tr class="{row_class}"><td>{medal}{r["rank"]}</td><td>{r["name"]}</td><td>{r["points"]}</td></tr>'
                panels_html += "</tbody></table>"
            panels_html += "</div>"

        return tabs_html + panels_html

    ring_html = season_tabs_html(ring_seasons, "ring", "シーズンデータなし")
    daily_html = season_tabs_html(daily_seasons, "daily", "デイリーデータなし")
    toname_html = season_tabs_html(toname_seasons, "toname", "シーズンデータなし")

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GG新宿 ランキング</title>
<style>
  :root {{
    --red: #e63946;
    --red-light: #ff6b6b;
    --green: #2d6a4f;
    --green-light: #52b788;
    --blue: #1d3557;
    --blue-light: #457b9d;
    --white: #ffffff;
    --gray: #f8f9fa;
    --dark: #1a1a2e;
  }}

  * {{ margin: 0; padding: 0; box-sizing: border-box; }}

  body {{
    font-family: 'Helvetica Neue', Arial, sans-serif;
    background: var(--dark);
    color: var(--white);
    min-height: 100vh;
  }}

  header {{
    text-align: center;
    padding: 32px 16px 16px;
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border-bottom: 2px solid var(--red);
  }}

  header h1 {{
    font-size: 2rem;
    font-weight: 800;
    letter-spacing: 0.1em;
    color: var(--white);
  }}

  header h1 span {{ color: var(--red); }}

  .updated {{ font-size: 0.75rem; color: #aaa; margin-top: 6px; }}

  .category-tabs {{
    display: flex;
    justify-content: center;
    gap: 8px;
    padding: 20px 16px 0;
    flex-wrap: wrap;
  }}

  .cat-btn {{
    padding: 10px 24px;
    border: none;
    border-radius: 24px;
    font-size: 0.9rem;
    font-weight: 700;
    cursor: pointer;
    transition: all 0.2s;
    opacity: 0.5;
  }}

  .cat-btn.ring-cat {{ background: var(--red); color: white; }}
  .cat-btn.daily-cat {{ background: var(--green); color: white; }}
  .cat-btn.toname-cat {{ background: var(--blue-light); color: white; }}
  .cat-btn.active {{ opacity: 1; transform: scale(1.05); box-shadow: 0 4px 16px rgba(0,0,0,0.4); }}

  .category-panel {{ display: none; padding: 20px 16px 40px; max-width: 720px; margin: 0 auto; }}
  .category-panel.active {{ display: block; }}

  .panel-title {{
    font-size: 1.2rem;
    font-weight: 800;
    margin-bottom: 16px;
    padding-bottom: 8px;
  }}

  .ring-panel .panel-title {{ color: var(--red-light); border-bottom: 2px solid var(--red); }}
  .daily-panel .panel-title {{ color: var(--green-light); border-bottom: 2px solid var(--green-light); }}
  .toname-panel .panel-title {{ color: #90e0ef; border-bottom: 2px solid var(--blue-light); }}

  .season-tabs {{
    display: flex;
    gap: 6px;
    flex-wrap: wrap;
    margin-bottom: 16px;
  }}

  .season-btn {{
    padding: 6px 14px;
    border-radius: 16px;
    border: 2px solid transparent;
    font-size: 0.8rem;
    font-weight: 600;
    cursor: pointer;
    background: #2a2a3e;
    color: #aaa;
    transition: all 0.2s;
  }}

  .ring-panel .season-btn.active {{ border-color: var(--red); color: var(--red-light); background: rgba(230,57,70,0.15); }}
  .daily-panel .season-btn.active {{ border-color: var(--green-light); color: var(--green-light); background: rgba(82,183,136,0.15); }}
  .toname-panel .season-btn.active {{ border-color: var(--blue-light); color: #90e0ef; background: rgba(69,123,157,0.15); }}

  .season-panel {{ display: none; }}
  .season-panel.active {{ display: block; }}

  .ranking-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.95rem;
  }}

  .ranking-table th {{
    padding: 10px 12px;
    text-align: left;
    font-size: 0.75rem;
    letter-spacing: 0.1em;
    color: #aaa;
    border-bottom: 1px solid #333;
  }}

  .ranking-table td {{
    padding: 12px;
    border-bottom: 1px solid #222;
  }}

  .rank-1 td {{ background: rgba(255, 215, 0, 0.08); }}
  .rank-2 td {{ background: rgba(192, 192, 192, 0.06); }}
  .rank-3 td {{ background: rgba(205, 127, 50, 0.06); }}

  .medal {{ margin-right: 4px; }}

  .no-data {{ color: #666; font-size: 0.9rem; padding: 20px 0; }}

  @media (max-width: 480px) {{
    header h1 {{ font-size: 1.4rem; }}
    .cat-btn {{ padding: 8px 16px; font-size: 0.8rem; }}
  }}
</style>
</head>
<body>

<header>
  <h1>GG<span>新宿</span> RANKING</h1>
  <p class="updated">最終更新: {updated_at}</p>
</header>

<div class="category-tabs">
  <button class="cat-btn ring-cat active" onclick="switchCategory(this, 'ring-panel')">🃏 リングポイント</button>
  <button class="cat-btn daily-cat" onclick="switchCategory(this, 'daily-panel')">📅 デイリーリング</button>
  <button class="cat-btn toname-cat" onclick="switchCategory(this, 'toname-panel')">🏆 トナメ</button>
</div>

<div class="category-panel ring-panel active">
  <p class="panel-title">リングポイントランキング</p>
  {ring_html}
</div>

<div class="category-panel daily-panel">
  <p class="panel-title">デイリーリングランキング</p>
  {daily_html}
</div>

<div class="category-panel toname-panel">
  <p class="panel-title">トナメポイントランキング</p>
  {toname_html}
</div>

<script>
function switchCategory(btn, panelId) {{
  document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.category-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(panelId) && document.getElementById(panelId).classList.add('active');
  document.querySelector('.' + panelId) && document.querySelector('.' + panelId).classList.add('active');
}}

function switchSeason(btn, panelId) {{
  const parent = btn.closest('.category-panel');
  parent.querySelectorAll('.season-btn').forEach(b => b.classList.remove('active'));
  parent.querySelectorAll('.season-panel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(panelId) && document.getElementById(panelId).classList.add('active');
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
