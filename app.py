#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
土地造成リスク診断Webアプリ
"""

# ============================================================
# 1. Imports & Configuration
# ============================================================
import os
import json
import math
import io
import base64
import sqlite3
import requests
import traceback
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime
from flask import (Flask, render_template, request, jsonify,
                   send_file, redirect, url_for, session, abort)
import stripe
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (Mail, Attachment, FileContent,
                                   FileName, FileType, Disposition)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image, PageBreak,
                                HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# Stripe
stripe.api_key = os.environ.get('STRIPE_SECRET_KEY', '')
STRIPE_PUBLISHABLE_KEY = os.environ.get('STRIPE_PUBLISHABLE_KEY', '')
STRIPE_WEBHOOK_SECRET = os.environ.get('STRIPE_WEBHOOK_SECRET', '')
PRICE_JPY = 30000

# SendGrid
SENDGRID_API_KEY = os.environ.get('SENDGRID_API_KEY', '')
FROM_EMAIL = os.environ.get('FROM_EMAIL', 'noreply@example.com')

# Admin
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
DB_PATH = os.path.join(BASE_DIR, 'land_risk.db')
os.makedirs(REPORTS_DIR, exist_ok=True)


# ============================================================
# 2. Database
# ============================================================
def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def init_db():
    with get_db() as db:
        db.execute('''
            CREATE TABLE IF NOT EXISTS orders (
                id                          INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at                  TEXT    DEFAULT (datetime('now', 'localtime')),
                requester_name              TEXT    NOT NULL,
                email                       TEXT    NOT NULL,
                address                     TEXT    NOT NULL,
                land_use                    TEXT    NOT NULL,
                payment_status              TEXT    DEFAULT 'pending',
                stripe_session_id           TEXT,
                report_status               TEXT    DEFAULT 'pending',
                sent_at                     TEXT,
                latitude                    REAL,
                longitude                   REAL,
                elevation                   REAL,
                elevation_diff              REAL,
                soil_amplification          REAL,
                flood_depth                 REAL,
                landslide_risk              INTEGER,
                overall_rank                TEXT,
                total_score                 INTEGER,
                score_terrain               INTEGER,
                score_soil                  INTEGER,
                score_disaster              INTEGER,
                score_regulation            INTEGER,
                score_cost                  INTEGER,
                grading_cost_per_sqm        INTEGER,
                soil_improvement_cost_per_sqm INTEGER,
                total_cost_per_sqm          INTEGER,
                api_data                    TEXT,
                pdf_path                    TEXT
            )
        ''')
        db.commit()


# gunicorn（Render等）でも起動時にテーブルを作成する
init_db()


# ============================================================
# 3. Japanese Font Registration
# ============================================================
FONT_NAME = 'Helvetica'
_font_registered = False


def _try_register(path, idx=None):
    """フォントファイルを登録する。成功したらTrueを返す。"""
    try:
        if idx is not None:
            pdfmetrics.registerFont(TTFont('Japanese', path, subfontIndex=idx))
        else:
            pdfmetrics.registerFont(TTFont('Japanese', path))
        return True
    except Exception as e:
        print(f'フォント登録失敗 {path}: {e}')
        return False


def register_japanese_font():
    global FONT_NAME, _font_registered
    if _font_registered:
        return

    fonts_dir = os.path.join(BASE_DIR, 'fonts')
    local_ttf  = os.path.join(fonts_dir, 'NotoSansJP-Regular.ttf')

    # (path, subfontIndex or None)  ― ttf は None、ttc は 0
    candidates = [
        (local_ttf,                                                    None),  # リポジトリ同梱 NotoSansJP
        ('C:/Windows/Fonts/meiryo.ttc',                               0),     # Windows
        ('C:/Windows/Fonts/msgothic.ttc',                             0),     # Windows
        ('C:/Windows/Fonts/msmincho.ttc',                             0),     # Windows
        ('C:/Windows/Fonts/YuGothR.ttc',                              0),     # Windows
        ('/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc',           0),     # macOS
        ('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',    0),     # Linux (apt)
        ('/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',    0),     # Linux (alt)
        ('/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc',         0),     # Linux (alt2)
    ]

    for path, idx in candidates:
        if os.path.exists(path):
            if _try_register(path, idx):
                FONT_NAME = 'Japanese'
                _font_registered = True
                print(f'日本語フォント登録成功: {path}')
                return

    # フォールバック: NotoSansJP をダウンロード（静的フォントのURLを使用）
    # Google Fonts は variable font に移行したため static/ サブディレクトリを参照
    download_urls = [
        'https://github.com/google/fonts/raw/main/ofl/notosansjp/static/NotoSansJP-Regular.ttf',
        'https://github.com/notofonts/noto-cjk/raw/main/Sans/Variable/TTF/Subset/NotoSansJP-Regular.ttf',
    ]
    try:
        import urllib.request
        os.makedirs(fonts_dir, exist_ok=True)
        downloaded = False
        for url in download_urls:
            try:
                print(f'NotoSansJP をダウンロード中: {url}')
                urllib.request.urlretrieve(url, local_ttf)
                # ダウンロードしたファイルが最低限のサイズか確認（HTML エラーページでないか）
                if os.path.getsize(local_ttf) > 50_000:
                    downloaded = True
                    break
                else:
                    print(f'ダウンロードファイルが小さすぎます（HTMLエラーページの可能性）: {os.path.getsize(local_ttf)} bytes')
                    os.remove(local_ttf)
            except Exception as e:
                print(f'ダウンロード失敗 {url}: {e}')
        if downloaded and _try_register(local_ttf):
            FONT_NAME = 'Japanese'
            _font_registered = True
            print(f'日本語フォント登録成功（ダウンロード）: {local_ttf}')
            return
    except Exception as e:
        print(f'フォントダウンロード処理エラー: {e}')

    print('警告: 日本語フォントが見つかりません。PDFのテキストが文字化けする可能性があります。')
    print(f'  対処法: {local_ttf} にNotoSansJP-Regular.ttfを配置してください。')
    _font_registered = True


# ============================================================
# 4. External API Functions
# ============================================================
def geocode_address(address: str) -> dict:
    """国土地理院APIで住所→緯度経度変換"""
    try:
        resp = requests.get(
            'https://msearch.gsi.go.jp/address-search/AddressSearch',
            params={'q': address}, timeout=10
        )
        data = resp.json()
        if data:
            coords = data[0]['geometry']['coordinates']  # [lon, lat]
            return {
                'lon': coords[0],
                'lat': coords[1],
                'title': data[0]['properties'].get('title', address)
            }
    except Exception as e:
        print(f'ジオコーディングエラー: {e}')
    return {'lon': None, 'lat': None, 'title': address}


def get_elevation(lon: float, lat: float) -> float:
    """国土地理院APIで標高取得"""
    try:
        resp = requests.get(
            'https://cyberjapandata2.gsi.go.jp/general/dem/scripts/getelevation.php',
            params={'lon': lon, 'lat': lat, 'outtype': 'JSON'},
            timeout=10
        )
        data = resp.json()
        elev = data.get('elevation')
        if elev is not None and elev != 'e':
            return float(elev)
    except Exception as e:
        print(f'標高取得エラー: {e}')
    return 0.0


def get_elevation_diff(lon: float, lat: float, radius_deg: float = 0.005) -> float:
    """周辺4点+中心の標高差（最大-最小）を返す"""
    try:
        points = [
            (lon + radius_deg, lat),
            (lon - radius_deg, lat),
            (lon, lat + radius_deg),
            (lon, lat - radius_deg),
            (lon, lat),
        ]
        elevations = [get_elevation(p[0], p[1]) for p in points]
        valid = [e for e in elevations if e > -9999]
        if len(valid) >= 2:
            return round(max(valid) - min(valid), 2)
    except Exception as e:
        print(f'標高差計算エラー: {e}')
    return 1.0


def get_jshis_data(lon: float, lat: float) -> dict:
    """J-SHISから地盤増幅率（AVS30）を取得"""
    try:
        url = (
            f'https://www.j-shis.bosai.go.jp/map/api/pshm/Y2020/DM/'
            f'meshAvs30/5/{lon:.4f}/{lat:.4f}.json'
        )
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            amp = data.get('value', 1.5)
            if amp and float(amp) > 0:
                return {'amplification': float(amp), 'raw': data}
    except Exception as e:
        print(f'J-SHIS APIエラー: {e}')
    # フォールバック: 標高から簡易推定（低地=軟弱地盤傾向）
    try:
        elev = get_elevation(lon, lat)
        if elev < 2:
            return {'amplification': 2.2, 'raw': {}}
        elif elev < 5:
            return {'amplification': 1.8, 'raw': {}}
        elif elev < 20:
            return {'amplification': 1.5, 'raw': {}}
        else:
            return {'amplification': 1.2, 'raw': {}}
    except Exception:
        pass
    return {'amplification': 1.5, 'raw': {}}


def get_hazard_data(lon: float, lat: float) -> dict:
    """標高をもとに浸水リスクを推定（ハザードマップAPIの代替簡易判定）"""
    result = {'flood_depth': 0.0, 'landslide': 0}
    try:
        elev = get_elevation(lon, lat)
        if elev <= 0:
            result['flood_depth'] = 3.0
        elif elev <= 2:
            result['flood_depth'] = 2.0
        elif elev <= 5:
            result['flood_depth'] = 1.0
        elif elev <= 10:
            result['flood_depth'] = 0.5
        else:
            result['flood_depth'] = 0.0

        # 傾斜地判定（標高差が大きければ土砂リスクあり）
        elevation_diff = get_elevation_diff(lon, lat)
        if elevation_diff > 5:
            result['landslide'] = 1
    except Exception as e:
        print(f'ハザードデータ取得エラー: {e}')
    return result


# ============================================================
# 5. Risk Assessment Logic
# ============================================================
RANK_TABLE = [
    (85, 100, 'A'),
    (70,  84, 'B'),
    (50,  69, 'C'),
    ( 0,  49, 'D'),
]

RANK_LABELS = {
    'A': '優良 — 造成適性が高く、リスクは低水準です',
    'B': '良好 — 一部対策で造成可能です',
    'C': '要注意 — 専門的検討と対策が必要です',
    'D': '高リスク — 慎重な判断と専門家相談を強く推奨します',
}

RANK_COLORS_HEX = {
    'A': '#1B5E20', 'B': '#1565C0', 'C': '#E65100', 'D': '#B71C1C',
}
RANK_BG_HEX = {
    'A': '#E8F5E9', 'B': '#E3F2FD', 'C': '#FFF3E0', 'D': '#FFEBEE',
}


def calc_score_terrain(elevation_diff: float) -> int:
    if elevation_diff <= 1:   return 20
    if elevation_diff <= 3:   return 12
    if elevation_diff <= 5:   return 6
    return 2


def calc_score_soil(amplification: float) -> int:
    if amplification <= 1.2:  return 20
    if amplification <= 1.5:  return 15
    if amplification <= 2.0:  return 8
    return 3


def calc_score_disaster(flood_depth: float, landslide: int) -> int:
    score = 20
    if flood_depth > 2:       score -= 15
    elif flood_depth > 0.5:   score -= 8
    elif flood_depth > 0:     score -= 4
    if landslide > 0:         score -= 8
    return max(0, score)


def calc_score_regulation(land_use: str) -> int:
    return {'住宅': 16, '商業': 14, '太陽光': 18, 'その他': 12}.get(land_use, 12)


def calc_score_cost(total_cost_per_sqm: int) -> int:
    if total_cost_per_sqm <= 3000:   return 20
    if total_cost_per_sqm <= 6000:   return 15
    if total_cost_per_sqm <= 10000:  return 8
    if total_cost_per_sqm <= 15000:  return 4
    return 1


def calc_grading_cost(elevation_diff: float) -> int:
    if elevation_diff <= 1:   return 3000
    if elevation_diff <= 3:   return 6000
    return 12000


def calc_soil_improvement_cost(amplification: float) -> int:
    if amplification > 2.0:   return 8000
    if amplification > 1.5:   return 4000
    return 0


def get_rank(total_score: int) -> str:
    for low, high, rank in RANK_TABLE:
        if low <= total_score <= high:
            return rank
    return 'D'


def assess_risk(elevation_diff: float, amplification: float,
                flood_depth: float, landslide: int, land_use: str) -> dict:
    grading      = calc_grading_cost(elevation_diff)
    soil_imp     = calc_soil_improvement_cost(amplification)
    total_cost   = grading + soil_imp

    s_terrain    = calc_score_terrain(elevation_diff)
    s_soil       = calc_score_soil(amplification)
    s_disaster   = calc_score_disaster(flood_depth, landslide)
    s_regulation = calc_score_regulation(land_use)
    s_cost       = calc_score_cost(total_cost)
    total        = s_terrain + s_soil + s_disaster + s_regulation + s_cost

    return {
        'overall_rank':                   get_rank(total),
        'total_score':                    total,
        'score_terrain':                  s_terrain,
        'score_soil':                     s_soil,
        'score_disaster':                 s_disaster,
        'score_regulation':               s_regulation,
        'score_cost':                     s_cost,
        'grading_cost_per_sqm':           grading,
        'soil_improvement_cost_per_sqm':  soil_imp,
        'total_cost_per_sqm':             total_cost,
    }


# ============================================================
# 6. Radar Chart (matplotlib → PNG bytes)
# ============================================================
def generate_radar_chart(order: dict) -> bytes:
    categories  = ['地形', '地盤', '災害\nリスク', '法規制', 'コスト']
    raw_scores  = [
        order.get('score_terrain', 0),
        order.get('score_soil', 0),
        order.get('score_disaster', 0),
        order.get('score_regulation', 0),
        order.get('score_cost', 0),
    ]
    N       = len(categories)
    angles  = [n / N * 2 * math.pi for n in range(N)]
    angles += angles[:1]
    values  = [v / 20 for v in raw_scores]
    values += values[:1]

    # フォント設定（NotoSansJP）
    try:
        import matplotlib
        import matplotlib.font_manager as fm
        font_path = os.path.join(os.path.dirname(__file__), 'fonts', 'NotoSansJP-Regular.ttf')
        fm.fontManager.addfont(font_path)
        matplotlib.rcParams['font.family'] = 'Noto Sans JP'
    except Exception:
        pass

    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, size=11)
    ax.set_ylim(0, 1)
    ax.set_yticks([0.25, 0.5, 0.75, 1.0])
    ax.set_yticklabels(['5', '10', '15', '20'], size=8, color='grey')
    ax.plot(angles, values, 'o-', linewidth=2, color='#2196F3')
    ax.fill(angles, values, alpha=0.25, color='#2196F3')

    for angle, val, raw in zip(angles[:-1], values[:-1], raw_scores):
        ax.annotate(str(raw),
                    xy=(angle, val + 0.08),
                    fontsize=11, ha='center', va='bottom',
                    color='#1565C0', fontweight='bold')

    ax.set_title('リスク評価レーダーチャート', size=13, pad=20)
    ax.grid(color='grey', linestyle='--', linewidth=0.5, alpha=0.7)
    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ============================================================
# 7. PDF Score Helper Functions
# ============================================================
def _score_label(score: int, max_score: int) -> str:
    r = score / max_score
    if r >= 0.85: return '優良'
    if r >= 0.70: return '良好'
    if r >= 0.50: return '要注意'
    return '要対策'


def _score_comment(category: str, score: int) -> str:
    comments = {
        'terrain': {
            (18, 20): '平坦で造成難易度は低い',
            (12, 17): '中程度の造成工事が必要',
            ( 6, 11): '大規模造成工事が見込まれる',
            ( 0,  5): '急峻地形、造成難易度が非常に高い',
        },
        'soil': {
            (18, 20): '良好地盤、標準基礎で対応可',
            (12, 17): 'やや軟弱、地盤調査を推奨',
            ( 6, 11): '軟弱地盤、地盤改良工事が必要',
            ( 0,  5): '非常に軟弱、専門的調査が必須',
        },
        'disaster': {
            (18, 20): '災害リスクは低い',
            (12, 17): '軽微なリスク、対策で対応可',
            ( 6, 11): '複数リスクあり、慎重な設計が必要',
            ( 0,  5): '高リスク、総合的対策が必須',
        },
        'regulation': {
            (18, 20): '規制は標準的',
            (12, 17): '一般的な法規制に注意',
            ( 6, 11): '複数法規制の確認が必要',
            ( 0,  5): '複雑な法規制、専門家相談必須',
        },
        'cost': {
            (18, 20): '低コスト、経済的に有利',
            (12, 17): '標準的なコスト水準',
            ( 6, 11): '高コスト、資金計画に注意',
            ( 0,  5): '非常に高コスト、事業性の再検討を推奨',
        },
    }
    for (lo, hi), text in comments.get(category, {}).items():
        if lo <= score <= hi:
            return text
    return ''


def _get_recommended_actions(order: dict) -> list:
    """優先度・タイトル・内容のリストを返す"""
    ed   = order.get('elevation_diff', 0) or 0
    amp  = order.get('soil_amplification', 1.5) or 1.5
    fd   = order.get('flood_depth', 0) or 0
    ls   = order.get('landslide_risk', 0) or 0
    use  = order.get('land_use', '住宅')

    actions = []

    if ed > 3:
        actions.append(('高', '大規模造成の詳細設計が必要',
                        '標高差が大きく、擁壁・盛土設計を専門家に依頼してください。地盤安定性の検証も必須です。'))
    elif ed > 1:
        actions.append(('中', '造成計画の事前検討',
                        '中程度の造成工事が予想されます。複数施工業者への見積依頼と工法比較を推奨します。'))

    if amp > 1.5:
        actions.append(('高', '地盤調査の実施',
                        'スウェーデン式サウンディングまたはボーリング調査を実施し、基礎・地盤改良工法を決定してください。'))

    if fd > 0.5:
        actions.append(('高', '浸水対策の設計',
                        f'想定浸水深 {fd:.1f}m に対応した嵩上げ・防水設計が必要です。1階床高さを適切に設定してください。'))

    if ls > 0:
        actions.append(('高', '土砂災害警戒区域の確認',
                        '都道府県の土砂災害警戒区域指定状況を確認し、必要に応じて開発許可を取得してください。'))

    if use == '太陽光':
        actions.append(('中', 'FIT認定・農地転用の確認',
                        '農地法・森林法・FIT制度の要件を事前確認し、農業委員会への届出・許可取得を行ってください。'))

    actions.append(('中', '自治体窓口での法規制確認',
                    '都市計画課にて用途地域・開発許可要件・地区計画・条例を確認してください。1,000㎡以上の開発は許可申請が必要です。'))

    actions.append(('低', '複数業者による見積取得',
                    '造成・地盤改良について3社以上から見積を取得し、工法・価格・実績を比較検討することを推奨します。'))

    return actions[:6]


# ============================================================
# 8. PDF Generation (A4・8ページ)
# ============================================================
def build_pdf(order: dict) -> bytes:
    register_japanese_font()
    F   = FONT_NAME
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=2 * cm, leftMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2 * cm,
        title='土地造成リスク診断レポート'
    )

    rank        = order.get('overall_rank', 'C') or 'C'
    score       = order.get('total_score', 0) or 0
    rank_color  = colors.HexColor(RANK_COLORS_HEX.get(rank, '#E65100'))
    rank_bg     = colors.HexColor(RANK_BG_HEX.get(rank, '#FFF3E0'))

    # ── スタイル ────────────────────────────────────────────────
    def S(name, **kw):
        return ParagraphStyle(name, fontName=F, **kw)

    s_body    = S('Body',    fontSize=10, leading=17, spaceAfter=4)
    s_h1      = S('H1',     fontSize=16, leading=22, spaceAfter=6,
                  spaceBefore=14, textColor=colors.HexColor('#1A237E'))
    s_h2      = S('H2',     fontSize=12, leading=17, spaceAfter=4,
                  spaceBefore=10, textColor=colors.HexColor('#283593'))
    s_small   = S('Small',  fontSize=8,  leading=13, textColor=colors.HexColor('#555555'))
    s_center  = S('Center', fontSize=10, leading=14, alignment=TA_CENTER)

    story = []

    def section_header(title):
        story.append(Spacer(1, 0.4 * cm))
        story.append(HRFlowable(width='100%', thickness=2,
                                color=colors.HexColor('#1A237E')))
        story.append(Paragraph(title, s_h1))

    def kv_table(data, col_widths=None):
        if col_widths is None:
            col_widths = [5 * cm, 12 * cm]
        tbl = Table(data, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ('FONTNAME',       (0, 0), (-1, -1), F),
            ('FONTSIZE',       (0, 0), (-1, -1), 10),
            ('BACKGROUND',     (0, 0), (0, -1),  colors.HexColor('#E8EAF6')),
            ('GRID',           (0, 0), (-1, -1), 0.5, colors.HexColor('#9E9E9E')),
            ('VALIGN',         (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 0), (-1, -1),
             [colors.white, colors.HexColor('#F5F5F5')]),
            ('TOPPADDING',     (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING',  (0, 0), (-1, -1), 5),
            ('LEFTPADDING',    (0, 0), (-1, -1), 8),
        ]))
        return tbl

    def score_table_style():
        return TableStyle([
            ('FONTNAME',       (0, 0), (-1, -1), F),
            ('FONTSIZE',       (0, 0), (-1, -1), 10),
            ('BACKGROUND',     (0, 0), (-1,  0), colors.HexColor('#283593')),
            ('TEXTCOLOR',      (0, 0), (-1,  0), colors.white),
            ('GRID',           (0, 0), (-1, -1), 0.5, colors.HexColor('#9E9E9E')),
            ('ALIGN',          (1, 0), (3, -1),  'CENTER'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [colors.white, colors.HexColor('#F5F5F5')]),
            ('TOPPADDING',     (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING',  (0, 0), (-1, -1), 5),
            ('LEFTPADDING',    (0, 0), (-1, -1), 8),
        ])

    ed   = order.get('elevation_diff', 1.0) or 1.0
    amp  = order.get('soil_amplification', 1.5) or 1.5
    fd   = order.get('flood_depth', 0) or 0
    ls   = order.get('landslide_risk', 0) or 0
    use  = order.get('land_use', '住宅') or '住宅'

    # ============================================================
    # PAGE 1: 表紙
    # ============================================================
    story.append(Spacer(1, 0.5 * cm))

    # バナー
    banner = Table([['土地造成リスク診断レポート']], colWidths=[17 * cm])
    banner.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('FONTSIZE',      (0, 0), (-1, -1), 20),
        ('BACKGROUND',    (0, 0), (-1, -1), colors.HexColor('#1A237E')),
        ('TEXTCOLOR',     (0, 0), (-1, -1), colors.white),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING',    (0, 0), (-1, -1), 14),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 14),
    ]))
    story.append(banner)
    story.append(Spacer(1, 0.8 * cm))

    # 依頼者情報
    story.append(Paragraph('■ 依頼者情報', s_h2))
    story.append(kv_table([
        ['依頼者名',  order.get('requester_name', '')],
        ['メール',    order.get('email', '')],
        ['土地住所',  order.get('address', '')],
        ['利用用途',  use],
        ['診断日時',  order.get('created_at', '')],
    ]))
    story.append(Spacer(1, 0.8 * cm))

    # 総合ランク
    story.append(Paragraph('■ 総合評価', s_h2))
    rank_tbl = Table([[rank]], colWidths=[17 * cm])
    rank_tbl.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('FONTSIZE',      (0, 0), (-1, -1), 96),
        ('TEXTCOLOR',     (0, 0), (-1, -1), rank_color),
        ('BACKGROUND',    (0, 0), (-1, -1), rank_bg),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING',    (0, 0), (-1, -1), 18),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 18),
        ('BOX',           (0, 0), (-1, -1), 2, rank_color),
    ]))
    story.append(rank_tbl)
    story.append(Spacer(1, 0.2 * cm))

    label_tbl = Table(
        [[Paragraph(f'総合スコア: {score}/100点　　{RANK_LABELS.get(rank, "")}',
                    ParagraphStyle('RankLabel', fontName=F, fontSize=9, leading=13,
                                   textColor=rank_color, alignment=TA_CENTER))]],
        colWidths=[17 * cm])
    label_tbl.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('BACKGROUND',    (0, 0), (-1, -1), rank_bg),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING',    (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
    ]))
    story.append(label_tbl)
    story.append(Spacer(1, 0.6 * cm))

    # スコアサマリー表
    story.append(Paragraph('■ 評価スコアサマリー', s_h2))
    summary_data = [
        ['評価項目', 'スコア', '満点', '評価'],
        ['地形・造成難易度', str(order.get('score_terrain', 0)), '20',
         _score_label(order.get('score_terrain', 0), 20)],
        ['地盤リスク',       str(order.get('score_soil', 0)), '20',
         _score_label(order.get('score_soil', 0), 20)],
        ['災害リスク',       str(order.get('score_disaster', 0)), '20',
         _score_label(order.get('score_disaster', 0), 20)],
        ['法規制',           str(order.get('score_regulation', 0)), '20',
         _score_label(order.get('score_regulation', 0), 20)],
        ['コスト',           str(order.get('score_cost', 0)), '20',
         _score_label(order.get('score_cost', 0), 20)],
        ['合　計',           str(score), '100', rank],
    ]
    tbl = Table(summary_data, colWidths=[6 * cm, 3 * cm, 2.5 * cm, 5.5 * cm])
    tbl.setStyle(score_table_style())
    tbl.setStyle(TableStyle([
        ('FONTNAME',    (0, 0), (-1, -1), F),
        ('FONTSIZE',    (0, 0), (-1, -1), 10),
        ('FONTSIZE',    (-1, -1), (-1, -1), 14),
        ('BACKGROUND',  (0,  0), (-1,  0), colors.HexColor('#283593')),
        ('TEXTCOLOR',   (0,  0), (-1,  0), colors.white),
        ('BACKGROUND',  (0, -1), (-1, -1), rank_bg),
        ('TEXTCOLOR',   (0, -1), (-1, -1), rank_color),
        ('GRID',        (0,  0), (-1, -1), 0.5, colors.HexColor('#9E9E9E')),
        ('ALIGN',       (1,  0), (2, -1),  'CENTER'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2),
         [colors.white, colors.HexColor('#F5F5F5')]),
        ('TOPPADDING',    (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING',   (0, 0), (-1, -1), 8),
    ]))
    story.append(tbl)
    story.append(PageBreak())

    # ============================================================
    # PAGE 2: 地形・地盤概況
    # ============================================================
    section_header('第1章　地形・地盤概況')

    story.append(Paragraph('■ 位置情報', s_h2))
    story.append(kv_table([
        ['緯度',       f"{order.get('latitude', 'N/A')}°"],
        ['経度',       f"{order.get('longitude', 'N/A')}°"],
        ['標高',       f"{order.get('elevation', 'N/A')} m"],
        ['周辺標高差', f"{ed:.1f} m（半径約500m圏内）"],
    ]))
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph('■ 地形評価', s_h2))
    terrain_comment = (
        '標高差が小さく、造成工事は比較的容易です。'     if ed <= 1 else
        '中程度の標高差があり、切盛土工事が必要です。'   if ed <= 3 else
        '標高差が大きく、大規模な造成工事が必要です。擁壁・盛土の設計に特別な注意が必要です。'
    )
    story.append(Paragraph(
        f'対象地の周辺標高差は約 <b>{ed:.1f} m</b> と評価されました。'
        f'{terrain_comment}'
        f'　地形スコアは <b>{order.get("score_terrain", 0)}/20点</b> です。',
        s_body))
    story.append(Spacer(1, 0.3 * cm))

    story.append(Paragraph('■ 地盤評価', s_h2))
    soil_imp_cost = order.get('soil_improvement_cost_per_sqm', 0) or 0
    soil_comment = (
        '地盤は良好であり、標準的な基礎工事で対応可能です。'
        if amp <= 1.2 else
        '地盤はやや軟弱であり、地盤調査の実施を推奨します。'
        if amp <= 1.5 else
        f'軟弱地盤の可能性が高く、地盤改良工事（推定 {soil_imp_cost:,}円/㎡）が必要と判断されます。'
        if amp <= 2.0 else
        f'軟弱地盤と判定されます。本格的な地盤改良工事（推定 {soil_imp_cost:,}円/㎡）が必要です。専門家による詳細調査を強く推奨します。'
    )
    story.append(Paragraph(
        f'J-SHIS（地震ハザードステーション）データによると、'
        f'地盤増幅率（AVS30）は <b>{amp:.2f}</b> と推定されます。'
        f'{soil_comment}'
        f'　地盤スコアは <b>{order.get("score_soil", 0)}/20点</b> です。',
        s_body))
    story.append(PageBreak())

    # ============================================================
    # PAGE 3: 災害リスク
    # ============================================================
    section_header('第2章　災害リスク評価')

    story.append(Paragraph('■ 洪水浸水リスク', s_h2))
    flood_comment = (
        '浸水リスクは低く、通常の設計で問題ありません。'   if fd == 0 else
        '軽微な浸水リスクがあります。排水計画の検討を推奨します。' if fd <= 0.5 else
        '中程度の浸水リスクがあります。防水対策・排水計画が必要です。' if fd <= 2 else
        '高い浸水リスクがあります。嵩上げ工事または十分な防水対策が必須です。'
    )
    story.append(Paragraph(
        f'想定最大浸水深は <b>{fd:.1f} m</b> と評価されます。{flood_comment}', s_body))
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph('■ 土砂災害リスク', s_h2))
    ls_comment = (
        '土砂災害警戒区域に該当する可能性があります。都道府県の指定区域を必ず確認し、開発許可の取得が必要な場合があります。'
        if ls > 0 else
        '現状では土砂災害の直接リスクは低いと判定されます。ただし、造成後の切土・盛土面の安定性については別途確認が必要です。'
    )
    story.append(Paragraph(
        f'土砂災害リスク判定：<b>{"リスクあり（警戒区域の可能性）" if ls > 0 else "リスクは低い"}</b>。{ls_comment}',
        s_body))
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph('■ 総合災害リスクスコア', s_h2))
    tbl = Table([
        ['リスク種別', '評価', '備考'],
        ['洪水浸水', f'{fd:.1f}m', '想定最大浸水深'],
        ['土砂災害', '有' if ls > 0 else '無', '警戒区域該当可能性'],
        ['スコア', f'{order.get("score_disaster", 0)}/20点',
         _score_label(order.get("score_disaster", 0), 20)],
    ], colWidths=[5 * cm, 4 * cm, 8 * cm])
    tbl.setStyle(score_table_style())
    story.append(tbl)
    story.append(PageBreak())

    # ============================================================
    # PAGE 4: 法規制
    # ============================================================
    section_header('第3章　法規制・開発許可')

    story.append(Paragraph('■ 関連法規の概要', s_h2))
    story.append(Paragraph(
        f'利用用途「{use}」における主な関連法規制を以下に示します。', s_body))
    story.append(Spacer(1, 0.2 * cm))

    reg_rows = [
        ['法令・規制', '主な内容', '確認事項'],
        ['都市計画法', '開発許可・用途地域', '1,000㎡以上は開発許可が必要'],
        ['建築基準法', '建築物の基準', '用途地域・建ぺい率・容積率'],
        ['農地法', '農地転用', '農地の場合は転用許可'],
        ['宅地造成等規制法', '造成工事の規制', '宅地造成等規制区域の確認'],
        ['砂防法・急傾斜地法', '土砂災害防止', '指定区域の確認'],
        ['森林法', '林地開発許可', '森林の場合は許可が必要'],
    ]
    if use == '太陽光':
        reg_rows += [
            ['農林水産省指針', '農地等での太陽光', '農業委員会への届出・許可'],
            ['FIT法', '固定価格買取制度', '認定申請・設備規模の確認'],
        ]
    tbl = Table(reg_rows, colWidths=[4.5 * cm, 5.5 * cm, 7 * cm])
    tbl.setStyle(score_table_style())
    story.append(tbl)
    story.append(Spacer(1, 0.4 * cm))

    story.append(Paragraph('■ 法規制スコア', s_h2))
    story.append(Paragraph(
        f'利用用途「{use}」における法規制複雑度スコア：'
        f'<b>{order.get("score_regulation", 0)}/20点</b>　'
        f'（{_score_comment("regulation", order.get("score_regulation", 0))}）',
        s_body))
    story.append(PageBreak())

    # ============================================================
    # PAGE 5: 造成費概算
    # ============================================================
    section_header('第4章　造成費概算')

    grading   = order.get('grading_cost_per_sqm', 0) or 0
    soil_imp  = order.get('soil_improvement_cost_per_sqm', 0) or 0
    total_c   = order.get('total_cost_per_sqm', 0) or 0

    story.append(Paragraph('■ 費用概算（単価）', s_h2))
    tbl = Table([
        ['費用項目', '概算単価（円/㎡）', '適用条件'],
        ['造成工事費（切盛土）', f'{grading:,}', f'標高差 {ed:.1f}m'],
        ['地盤改良費', f'{soil_imp:,}' if soil_imp > 0 else '対象外',
         '増幅率 > 1.5 の場合'],
        ['合計概算（基本）', f'{total_c:,}', '税抜・設計費別'],
    ], colWidths=[6 * cm, 5 * cm, 6 * cm])
    tbl.setStyle(score_table_style())
    story.append(tbl)
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph('■ 面積別費用試算', s_h2))
    trial = [['面積（㎡）', '造成費（万円）', '地盤改良費（万円）', '合計（万円）']]
    for area in [100, 200, 300, 500, 1000]:
        trial.append([
            f'{area:,}',
            f'{grading * area // 10000:,}',
            f'{soil_imp * area // 10000:,}' if soil_imp > 0 else '－',
            f'{total_c * area // 10000:,}',
        ])
    tbl2 = Table(trial, colWidths=[4 * cm, 4.5 * cm, 5 * cm, 4.5 * cm])
    tbl2.setStyle(score_table_style())
    story.append(tbl2)
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(
        '※ 上記は概算値です。実際の費用は詳細設計・地盤調査結果により大きく変動します。', s_small))
    story.append(PageBreak())

    # ============================================================
    # PAGE 6: レーダーチャート
    # ============================================================
    section_header('第5章　総合評価レーダーチャート')
    story.append(Spacer(1, 0.3 * cm))

    try:
        radar_bytes = generate_radar_chart(order)
        chart_img = Image(io.BytesIO(radar_bytes), width=11 * cm, height=11 * cm)
        chart_img.hAlign = 'CENTER'
        story.append(chart_img)
    except Exception as e:
        story.append(Paragraph(f'レーダーチャート生成エラー: {e}', s_body))

    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph('■ 各項目スコア詳細', s_h2))
    detail_data = [
        ['評価項目', 'スコア', '満点', '割合', 'コメント'],
        ['地形・造成難易度',
         str(order.get('score_terrain', 0)), '20',
         f'{order.get("score_terrain", 0) / 20 * 100:.0f}%',
         _score_comment('terrain', order.get('score_terrain', 0))],
        ['地盤リスク',
         str(order.get('score_soil', 0)), '20',
         f'{order.get("score_soil", 0) / 20 * 100:.0f}%',
         _score_comment('soil', order.get('score_soil', 0))],
        ['災害リスク',
         str(order.get('score_disaster', 0)), '20',
         f'{order.get("score_disaster", 0) / 20 * 100:.0f}%',
         _score_comment('disaster', order.get('score_disaster', 0))],
        ['法規制',
         str(order.get('score_regulation', 0)), '20',
         f'{order.get("score_regulation", 0) / 20 * 100:.0f}%',
         _score_comment('regulation', order.get('score_regulation', 0))],
        ['コスト',
         str(order.get('score_cost', 0)), '20',
         f'{order.get("score_cost", 0) / 20 * 100:.0f}%',
         _score_comment('cost', order.get('score_cost', 0))],
    ]
    tbl = Table(detail_data, colWidths=[4 * cm, 2.5 * cm, 2 * cm, 2.5 * cm, 6 * cm])
    tbl.setStyle(score_table_style())
    story.append(tbl)
    story.append(PageBreak())

    # ============================================================
    # PAGE 7: 推奨アクション
    # ============================================================
    section_header('第6章　推奨アクション')

    story.append(Paragraph('■ 優先度別アクションプラン', s_h2))
    for priority, title, content in _get_recommended_actions(order):
        bg = {'高': '#FFEBEE', '中': '#FFF3E0', '低': '#E8F5E9'}[priority]
        tc = {'高': '#B71C1C', '中': '#E65100', '低': '#1B5E20'}[priority]
        action_tbl = Table(
            [[f'【優先度：{priority}】 {title}'],
             [content]],
            colWidths=[17 * cm])
        action_tbl.setStyle(TableStyle([
            ('FONTNAME',      (0, 0), (-1, -1), F),
            ('FONTSIZE',      (0, 0), (0, 0),  10),
            ('FONTSIZE',      (0, 1), (-1, -1), 9),
            ('BACKGROUND',    (0, 0), (-1,  0), colors.HexColor(bg)),
            ('TEXTCOLOR',     (0, 0), (-1,  0), colors.HexColor(tc)),
            ('BOX',           (0, 0), (-1, -1), 1, colors.HexColor(tc)),
            ('TOPPADDING',    (0, 0), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
            ('LEFTPADDING',   (0, 0), (-1, -1), 10),
        ]))
        story.append(action_tbl)
        story.append(Spacer(1, 0.25 * cm))

    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph('■ 推奨ステップ', s_h2))
    for step in [
        '① 本レポートを持参し、地域の建築士・土地家屋調査士に相談する',
        '② 自治体の都市計画課にて用途地域・開発許可要件を確認する',
        '③ 地盤調査（ボーリング調査・スウェーデン式サウンディング）を実施する',
        '④ 詳細設計・見積を複数の施工業者から取得し比較検討する',
        '⑤ 資金計画・融資条件を金融機関と事前協議する',
    ]:
        story.append(Paragraph(step, s_body))
    story.append(PageBreak())

    # ============================================================
    # PAGE 8: 免責事項
    # ============================================================
    section_header('免責事項・データソース')

    story.append(Paragraph(
        '本レポートは、公開されている地理空間情報・ハザード情報をもとに作成された参考資料です。'
        '以下の点をご了承の上ご活用ください。', s_body))
    story.append(Spacer(1, 0.3 * cm))

    for item in [
        '【精度について】本レポートは現地調査を行っておらず、デジタルデータに基づく机上評価です。'
        '実際の地形・地質・法規制状況とは異なる場合があります。',
        '【専門家調査の必要性】土地の取得・開発・建築に際しては、必ず専門家（建築士・地盤調査会社・'
        '測量士・弁護士等）による詳細調査・確認を行ってください。',
        '【法的効力】本レポートは建築確認申請・開発許可申請等の公的手続きに使用できる法的効力を有しません。',
        '【賠償責任の制限】本レポートの情報に基づく判断・行動による損害について、当社は一切の責任を負いません。',
        '【データの時点】使用するAPIデータは定期的に更新されますが、最新状況を反映していない場合があります。'
        '特に法改正・ハザード区域の見直し等については自治体窓口にて最新情報をご確認ください。',
        '【著作権】本レポートの著作権は発行者に帰属します。無断転載・複製を禁じます。',
    ]:
        story.append(Paragraph(item, s_small))
        story.append(Spacer(1, 0.15 * cm))

    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph('■ 使用データソース', s_h2))
    tbl = Table([
        ['データソース', 'URL', '利用内容'],
        ['国土地理院', 'https://maps.gsi.go.jp/', '住所検索・標高データ'],
        ['J-SHIS（防災科研）', 'https://www.j-shis.bosai.go.jp/', '地盤増幅率・地震ハザード'],
        ['ハザードマップポータル（国交省）', 'https://disaportal.gsi.go.jp/', '洪水・土砂・津波リスク'],
        ['国土数値情報（国交省）', 'https://nlftp.mlit.go.jp/', '用途地域・法規制情報'],
    ], colWidths=[5 * cm, 7 * cm, 5 * cm])
    tbl.setStyle(score_table_style())
    story.append(tbl)

    story.append(Spacer(1, 1 * cm))
    footer = Table(
        [[f'発行日：{datetime.now().strftime("%Y年%m月%d日")}　　土地造成リスク診断サービス']],
        colWidths=[17 * cm])
    footer.setStyle(TableStyle([
        ('FONTNAME',    (0, 0), (-1, -1), F),
        ('FONTSIZE',    (0, 0), (-1, -1), 9),
        ('TEXTCOLOR',   (0, 0), (-1, -1), colors.grey),
        ('ALIGN',       (0, 0), (-1, -1), 'RIGHT'),
        ('TOPPADDING',  (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(footer)

    # ============================================================
    # PAGE 9: 無料相談CTA
    # ============================================================
    story.append(PageBreak())
    story.append(Spacer(1, 1.5 * cm))

    cta_header = Table([['■ より詳細な検討をご希望の方へ']], colWidths=[17 * cm])
    cta_header.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('FONTSIZE',      (0, 0), (-1, -1), 16),
        ('BACKGROUND',    (0, 0), (-1, -1), colors.HexColor('#1A237E')),
        ('TEXTCOLOR',     (0, 0), (-1, -1), colors.white),
        ('ALIGN',         (0, 0), (-1, -1), 'LEFT'),
        ('TOPPADDING',    (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('LEFTPADDING',   (0, 0), (-1, -1), 14),
    ]))
    story.append(cta_header)
    story.append(Spacer(1, 0.5 * cm))

    s_cta_body = S('CTABody', fontSize=11, leading=20, spaceAfter=6)
    s_cta_contact = S('CTAContact', fontSize=12, leading=22, spaceAfter=4,
                      textColor=colors.HexColor('#1A237E'))

    story.append(Paragraph(
        '本レポートの診断結果をもとに、現地調査・詳細設計・造成工事のご相談を承っております。',
        s_cta_body))
    story.append(Spacer(1, 0.6 * cm))

    contact_tbl = Table([
        ['▶ 無料相談はこちら'],
        ['　メール：deltatech666@gmail.com'],
        ['　担当：関根 寛人（土木造成設計 専門家）'],
    ], colWidths=[17 * cm])
    contact_tbl.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('FONTSIZE',      (0, 0), (0, 0),   13),
        ('FONTSIZE',      (0, 1), (-1, -1), 12),
        ('BACKGROUND',    (0, 0), (-1, -1), colors.HexColor('#E3F2FD')),
        ('TEXTCOLOR',     (0, 0), (-1, -1), colors.HexColor('#1A237E')),
        ('BOX',           (0, 0), (-1, -1), 2, colors.HexColor('#1565C0')),
        ('TOPPADDING',    (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING',   (0, 0), (-1, -1), 14),
    ]))
    story.append(contact_tbl)
    story.append(Spacer(1, 0.6 * cm))

    story.append(Paragraph('対応サービス内容：', s_h2))
    for item in [
        '・造成設計図の作成',
        '・工事費用の詳細見積もり',
        '・施工会社のご紹介',
        '・許認可申請サポート',
    ]:
        story.append(Paragraph(item, s_cta_body))

    story.append(Spacer(1, 0.8 * cm))
    closing_tbl = Table([['お気軽にお問い合わせください。']], colWidths=[17 * cm])
    closing_tbl.setStyle(TableStyle([
        ('FONTNAME',      (0, 0), (-1, -1), F),
        ('FONTSIZE',      (0, 0), (-1, -1), 13),
        ('BACKGROUND',    (0, 0), (-1, -1), colors.HexColor('#1A237E')),
        ('TEXTCOLOR',     (0, 0), (-1, -1), colors.white),
        ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING',    (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    story.append(closing_tbl)

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ============================================================
# 9. Email
# ============================================================
def send_report_email(to_email: str, to_name: str,
                      pdf_bytes: bytes, order_id: int) -> bool:
    try:
        sg = SendGridAPIClient(api_key=SENDGRID_API_KEY)
        message = Mail(
            from_email=FROM_EMAIL,
            to_emails=to_email,
            subject=f'【土地造成リスク診断】レポートをお届けします（受付番号: {order_id:04d}）',
            html_content=f'''
<div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">
  <div style="background:#1A237E; color:white; padding:20px; text-align:center;">
    <h1 style="margin:0; font-size:20px;">土地造成リスク診断レポート</h1>
  </div>
  <div style="padding:30px;">
    <p>{to_name} 様</p>
    <p>この度は土地造成リスク診断サービスをご利用いただき、ありがとうございます。</p>
    <p>診断レポートを添付PDFにてお届けします。</p>
    <div style="background:#f5f5f5; border-left:4px solid #1A237E; padding:15px; margin:20px 0;">
      <b>受付番号：</b>{order_id:04d}
    </div>
    <p style="color:#666; font-size:12px;">
      ※ 本レポートは参考資料です。実際の土地取得・開発にあたっては専門家にご相談ください。
    </p>
  </div>
  <div style="background:#f5f5f5; padding:15px; text-align:center; font-size:11px; color:#666;">
    土地造成リスク診断サービス
  </div>
</div>
'''
        )
        attachment = Attachment(
            FileContent(base64.b64encode(pdf_bytes).decode()),
            FileName(f'land_risk_report_{order_id:04d}.pdf'),
            FileType('application/pdf'),
            Disposition('attachment')
        )
        message.attachment = attachment
        resp = sg.send(message)
        print(f'メール送信 status={resp.status_code} to={to_email}')
        return resp.status_code in (200, 201, 202)
    except Exception as e:
        print(f'メール送信エラー: {e}')
        # SendGrid の詳細エラーを表示（認証エラー等の原因確認に使用）
        if hasattr(e, 'body'):
            print(f'SendGrid エラー詳細: {e.body}')
        if hasattr(e, 'status_code'):
            print(f'SendGrid HTTP ステータス: {e.status_code}')
        return False


# ============================================================
# 10. Main Routes
# ============================================================
@app.route('/')
def index():
    return render_template('index.html',
                           stripe_publishable_key=STRIPE_PUBLISHABLE_KEY)


@app.route('/tokutei')
def tokutei():
    return render_template('tokutei.html')


@app.route('/api/submit', methods=['POST'])
def submit_form():
    """フォーム送信 → Stripe Checkoutセッション作成"""
    data = request.get_json(force=True)
    for field, label in [('name', '依頼者名'), ('email', 'メールアドレス'),
                         ('address', '土地の住所'), ('land_use', '利用用途')]:
        if not data.get(field, '').strip():
            return jsonify({'error': f'{label} は必須です'}), 400

    try:
        checkout = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price_data': {
                    'currency': 'jpy',
                    'product_data': {
                        'name': '土地造成リスク診断レポート',
                        'description': f'対象地: {data["address"]}',
                    },
                    'unit_amount': PRICE_JPY,
                },
                'quantity': 1,
            }],
            mode='payment',
            success_url=request.host_url
                        + 'payment/success?session_id={CHECKOUT_SESSION_ID}',
            cancel_url=request.host_url + 'payment/cancel',
            customer_email=data['email'].strip(),
            metadata={
                'requester_name': data['name'].strip(),
                'email':          data['email'].strip(),
                'address':        data['address'].strip(),
                'land_use':       data['land_use'].strip(),
            }
        )
        return jsonify({'checkout_url': checkout.url})
    except stripe.error.StripeError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'サーバーエラー: {str(e)}'}), 500


@app.route('/payment/success')
def payment_success():
    sid = request.args.get('session_id', '')
    if not sid:
        return redirect('/')

    try:
        cs = stripe.checkout.Session.retrieve(sid)
        if cs.payment_status != 'paid':
            return render_template('index.html',
                                   stripe_publishable_key=STRIPE_PUBLISHABLE_KEY,
                                   error='決済が完了していません。')

        meta = cs.metadata

        # 既存チェック（リロード対策）
        with get_db() as db:
            existing = db.execute(
                'SELECT id FROM orders WHERE stripe_session_id = ?', (sid,)
            ).fetchone()
            if existing:
                return render_template('success.html',
                                       order_id=existing['id'],
                                       name=meta.get('requester_name', ''))

        # 外部APIでデータ取得
        address   = meta.get('address', '')
        land_use  = meta.get('land_use', '住宅')

        geo           = geocode_address(address)
        lat, lon      = geo.get('lat'), geo.get('lon')
        elevation     = get_elevation(lon, lat)        if lat and lon else None
        elevation_diff= get_elevation_diff(lon, lat) if lat and lon else 1.0
        jshis         = get_jshis_data(lon, lat)     if lat and lon else {}
        soil_amp      = jshis.get('amplification', 1.5)
        hazard        = get_hazard_data(lon, lat)    if lat and lon else {}
        flood_depth   = hazard.get('flood_depth', 0)
        landslide     = hazard.get('landslide', 0)

        assessment = assess_risk(elevation_diff, soil_amp, flood_depth,
                                 landslide, land_use)

        api_data = {
            'geocode': geo,
            'jshis':   {'amplification': soil_amp},
            'hazard':  {'flood_depth': flood_depth, 'landslide': landslide},
        }

        with get_db() as db:
            cur = db.execute('''
                INSERT INTO orders (
                    requester_name, email, address, land_use,
                    payment_status, stripe_session_id,
                    latitude, longitude, elevation, elevation_diff,
                    soil_amplification, flood_depth, landslide_risk,
                    overall_rank, total_score,
                    score_terrain, score_soil, score_disaster,
                    score_regulation, score_cost,
                    grading_cost_per_sqm, soil_improvement_cost_per_sqm,
                    total_cost_per_sqm, api_data
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', (
                meta.get('requester_name'), meta.get('email'),
                address, land_use, 'paid', sid,
                lat, lon, elevation, elevation_diff,
                soil_amp, flood_depth, landslide,
                assessment['overall_rank'], assessment['total_score'],
                assessment['score_terrain'], assessment['score_soil'],
                assessment['score_disaster'], assessment['score_regulation'],
                assessment['score_cost'],
                assessment['grading_cost_per_sqm'],
                assessment['soil_improvement_cost_per_sqm'],
                assessment['total_cost_per_sqm'],
                json.dumps(api_data, ensure_ascii=False)
            ))
            db.commit()
            order_id = cur.lastrowid

        return render_template('success.html',
                               order_id=order_id,
                               name=meta.get('requester_name', ''))

    except Exception as e:
        traceback.print_exc()
        return render_template('index.html',
                               stripe_publishable_key=STRIPE_PUBLISHABLE_KEY,
                               error=f'処理中にエラーが発生しました: {e}')


@app.route('/payment/cancel')
def payment_cancel():
    return render_template('index.html',
                           stripe_publishable_key=STRIPE_PUBLISHABLE_KEY,
                           message='決済がキャンセルされました。再度お試しください。')


@app.route('/webhook/stripe', methods=['POST'])
def stripe_webhook():
    payload    = request.get_data()
    sig_header = request.headers.get('Stripe-Signature', '')
    try:
        event = stripe.Webhook.construct_event(
            payload, sig_header, STRIPE_WEBHOOK_SECRET)
    except (ValueError, stripe.error.SignatureVerificationError):
        abort(400)

    if event['type'] == 'checkout.session.completed':
        obj = event['data']['object']
        with get_db() as db:
            db.execute(
                'UPDATE orders SET payment_status = "paid" '
                'WHERE stripe_session_id = ?', (obj['id'],))
            db.commit()
    return jsonify({'status': 'ok'})


# ============================================================
# 11. Admin Routes
# ============================================================
def is_admin():
    return session.get('admin_authenticated', False)


@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        if request.form.get('password', '') == ADMIN_PASSWORD:
            session['admin_authenticated'] = True
            return redirect('/admin')
        return render_template('admin_login.html', error='パスワードが違います')
    return render_template('admin_login.html')


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_authenticated', None)
    return redirect('/admin/login')


@app.route('/admin')
def admin_panel():
    if not is_admin():
        return redirect('/admin/login')
    with get_db() as db:
        orders = db.execute(
            'SELECT * FROM orders ORDER BY created_at DESC').fetchall()
    return render_template('admin.html', orders=[dict(o) for o in orders])


@app.route('/admin/approve/<int:order_id>', methods=['POST'])
def admin_approve(order_id):
    if not is_admin():
        return jsonify({'error': '認証が必要です'}), 401

    with get_db() as db:
        row = db.execute(
            'SELECT * FROM orders WHERE id = ?', (order_id,)).fetchone()
    if not row:
        return jsonify({'error': '案件が見つかりません'}), 404

    order = dict(row)
    try:
        pdf_bytes    = build_pdf(order)
        pdf_filename = (f'report_{order_id:04d}_'
                        f'{datetime.now().strftime("%Y%m%d%H%M%S")}.pdf')
        pdf_path = os.path.join(REPORTS_DIR, pdf_filename)
        with open(pdf_path, 'wb') as f:
            f.write(pdf_bytes)

        sent = send_report_email(
            order['email'], order['requester_name'], pdf_bytes, order_id)

        with get_db() as db:
            db.execute('''
                UPDATE orders
                SET report_status = ?, sent_at = ?, pdf_path = ?
                WHERE id = ?
            ''', (
                'sent' if sent else 'generated',
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                pdf_path, order_id
            ))
            db.commit()

        return jsonify({
            'success':    True,
            'email_sent': sent,
            'message': (f'PDFを生成し、メールを送信しました。'
                        if sent else
                        f'PDFを生成しましたが、メール送信に失敗しました。'),
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/admin/download/<int:order_id>')
def admin_download(order_id):
    if not is_admin():
        return redirect('/admin/login')

    with get_db() as db:
        row = db.execute(
            'SELECT * FROM orders WHERE id = ?', (order_id,)).fetchone()
    if not row:
        abort(404)

    order = dict(row)
    if order.get('pdf_path') and os.path.exists(order['pdf_path']):
        return send_file(order['pdf_path'], as_attachment=True,
                         download_name=f'report_{order_id:04d}.pdf',
                         mimetype='application/pdf')

    pdf_bytes = build_pdf(order)
    return send_file(
        io.BytesIO(pdf_bytes), as_attachment=True,
        download_name=f'report_{order_id:04d}.pdf',
        mimetype='application/pdf')


# ============================================================
# 12. Entry Point
# ============================================================
if __name__ == '__main__':
    init_db()
    register_japanese_font()
    print('=' * 50)
    print('土地造成リスク診断アプリ 起動中...')
    print(f'  DB      : {DB_PATH}')
    print(f'  Reports : {REPORTS_DIR}')
    print(f'  Font    : {FONT_NAME}')
    print('  URL     : http://localhost:5000')
    print('  Admin   : http://localhost:5000/admin')
    print('=' * 50)
    app.run(debug=True, host='0.0.0.0', port=5000)
