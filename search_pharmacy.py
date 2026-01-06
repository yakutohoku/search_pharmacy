"""
薬局マップ（Streamlit + Folium）

【主な機能】
1) Excel先頭シートの P列(緯度)・Q列(経度)を使って薬局を地図表示（青ピン）
   - 緯度/経度が空の行は表示しない
   - ピンをクリックすると
     薬局名(C列)、法人名（E列全文ラベル変更）、住所(I列) を表示
     Googleマップリンク・求人リンク（管理薬剤師K列, 常勤L列, パートM列, 派遣N列, 契約O列）

2) 検索
   - 住所・駅名で検索（半径指定 / ネットワークが必要）
   - 地図をクリックして指定（半径指定 / 確実）
   - 法人で検索（E列から法人名部分だけ抽出 + スペース無視で完全一致）
   - 社長名で検索（E列から「氏名」部分だけ抽出 + スペース無視で完全一致）

【反映済みの要望】
① 検索のデフォルトを「住所・駅名で検索」に
② 表記を「住所・駅名で検索」に
③ 半径のデフォルトを 2km に
④ ピンpopup：法人名の重複を解消
   - 「法人名:（抽出法人名）」を削除
   - 「開設者氏名:」ラベルを「法人名:」に変更（表示内容はE列全文）
⑤ タイトル「薬局検索マップ」が見切れないようレイアウト修正（st.title + CSS）
⑥ 一覧のチェックが1回目で消える問題を修正（data_editor入力DFを固定し、editedで上書きしない）
⑦ 駅名検索を安定させる改修
   - 住所/駅名の候補が複数取れた場合、「薬局データの中心に最も近い候補」を採用

"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import Optional, Tuple, List, Dict, Any

import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from branca.element import MacroElement, Template

# 住所 -> 緯度経度（外部通信が必要）
try:
    from geopy.geocoders import Nominatim, ArcGIS
    from geopy.extra.rate_limiter import RateLimiter
except Exception:
    Nominatim = None
    ArcGIS = None
    RateLimiter = None


# =============================================================================
# 設定
# =============================================================================
DEFAULT_LOCAL_FILE = r"\\file-tky\Section\薬剤師共有\★【支店・課】_東北\東北\東北 薬局リスト.xlsm"

JOB_BASE_URL = "https://yaku-job.com/preview/"
GOOGLE_MAPS_QUERY_URL = "https://www.google.com/maps/search/?api=1&query={lat},{lon}"

# Excel列（0-based index）
# C=2, E=4, I=8, K=10, L=11, M=12, N=13, O=14, P=15, Q=16
COL = {
    "pharmacy_name": 2,    # C: 薬局名
    "opener_name": 4,      # E: 厚生局開設者氏名（法人名+代表取締役+氏名）
    "address": 8,          # I: 住所
    "id_manager": 10,      # K: 管理薬剤師ID
    "id_fulltime": 11,     # L: 常勤ID
    "id_part": 12,         # M: パートID
    "id_temp": 13,         # N: 派遣ID
    "id_contract": 14,     # O: 契約社員ID
    "lat": 15,             # P: 緯度
    "lon": 16,             # Q: 経度
}


# =============================================================================
# データ構造
# =============================================================================
@dataclass
class SearchPoint:
    lat: float
    lon: float


# =============================================================================
# 基本ユーティリティ
# =============================================================================
def _safe_float(x) -> Optional[float]:
    """Excelセル値を float へ安全に変換。変換不能/空/NaN/inf は None。"""
    if x is None:
        return None
    if isinstance(x, (int, float)):
        if isinstance(x, float) and (math.isnan(x) or math.isinf(x)):
            return None
        return float(x)

    s = str(x).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _normalize_id(x) -> Optional[str]:
    """求人IDをURL用に正規化（Excelの 123.0 対策など）。"""
    if x is None:
        return None
    s = str(x).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    # 123.0 -> 123
    if s.endswith(".0"):
        s2 = s[:-2]
        if s2.isdigit():
            return s2
    return s


def _escape_html(text: str) -> str:
    """ポップアップHTMLの簡易エスケープ。"""
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """緯度経度2点間の距離(km)（ハーサイン）。"""
    r = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = (math.sin(dphi / 2) ** 2) + (math.cos(phi1) * math.cos(phi2) * (math.sin(dlambda / 2) ** 2))
    return 2 * r * math.asin(math.sqrt(a))


# =============================================================================
# 法人/社長検索用ユーティリティ
# =============================================================================
def normalize_space_ignored(text: str) -> str:
    """
    スペース（半角/全角）を無視して比較するための正規化
    - NFKC正規化
    - 全角スペースを半角へ
    - 空白（スペース/タブ等）を全削除
    """
    t = unicodedata.normalize("NFKC", text or "")
    t = t.replace("　", " ")
    t = re.sub(r"\s+", "", t)
    return t


def extract_corporation_part_from_opener(opener_raw: str) -> str:
    """
    E列（厚生局開設者氏名）から「法人名部分だけ」を切り出す。
    - 「代表取締役」等の手前までを法人名として採用
    """
    t = unicodedata.normalize("NFKC", opener_raw or "").replace("　", " ").strip()
    m = re.search(r"(代表取締役社長|代表取締役|代表社員|理事長|院長|代表者)", t)
    if m:
        t = t[: m.start()].strip()
    return t


def corporation_search_key_from_opener(opener_raw: str) -> str:
    """法人検索用キー（スペース無視で完全一致）。"""
    corp_part = extract_corporation_part_from_opener(opener_raw)
    return normalize_space_ignored(corp_part)


def extract_ceo_name_from_opener(opener_raw: str) -> str:
    """
    E列（厚生局開設者氏名）から「氏名部分」を切り出す。
    例: "株式会社 アルタックス 代表取締役 神林 明仁" -> "神林 明仁"
    """
    t = unicodedata.normalize("NFKC", opener_raw or "").replace("　", " ").strip()
    m = re.search(r"(代表取締役社長|代表取締役|代表社員|理事長|院長|代表者)\s*(.*)$", t)
    if not m:
        return ""
    name_part = (m.group(2) or "").strip()

    # 保険：先頭の肩書が混ざるケースを軽く除去
    for prefix in ["社長", "会長", "CEO", "ＣＥＯ"]:
        if name_part.startswith(prefix):
            name_part = name_part[len(prefix):].strip()
    return name_part


def ceo_search_key_from_opener(opener_raw: str) -> str:
    """社長名検索キー（苗字・名前の間の空白も無視、完全一致）。"""
    name_part = extract_ceo_name_from_opener(opener_raw)
    return normalize_space_ignored(name_part)


# =============================================================================
# UI（日本語化＆地図を大きく見せるCSS）
# =============================================================================
def _apply_ui_css() -> None:
    st.markdown(
        """
        <style>
        /* 余白を詰めて地図を大きく取る */
        .block-container { padding-top: 1.2rem; padding-bottom: 1.0rem; }
        header[data-testid="stHeader"] { height: 0.2rem; }

        /* タイトル（見切れ防止 + 2pt相当小さく） */
        div[data-testid="stAppViewContainer"] h1 {
            font-size: 30px !important;
            font-weight: 800 !important;
            margin: 0.2rem 0 0.2rem 0 !important;
            line-height: 1.2 !important;
            overflow: visible !important;
            display: inline-block !important;
            max-width: 100% !important;
            white-space: normal !important;
        }

        /* file_uploader の案内文（英語）を隠して、日本語を重ねる */
        div[data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderDropzoneInstructions"] { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] > div:nth-child(1) { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] p { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] small { display:none !important; }

        div[data-testid="stFileUploaderDropzone"]::before{
            content:"ここにExcelファイルをドラッグ＆ドロップ、または「ファイルを選択」を押してください";
            display:block;
            padding:0.25rem 0 0.6rem 0;
            font-size:0.95rem;
            color:rgba(49,51,63,0.9);
        }

        /* Browse files を「ファイルを選択」に（ボタン文字を置き換え） */
        div[data-testid="stFileUploaderDropzone"] button{
            position:relative;
        }
        div[data-testid="stFileUploaderDropzone"] button *{
            font-size:0px !important;
        }
        div[data-testid="stFileUploaderDropzone"] button::after{
            content:"ファイルを選択";
            font-size:14px;
            position:absolute;
            inset:0;
            display:flex;
            align-items:center;
            justify-content:center;
            pointer-events:none;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Excel読み込み（先頭シート固定）
# =============================================================================
@st.cache_data(show_spinner=False)
def load_pharmacy_data(file_bytes: bytes) -> pd.DataFrame:
    """
    Excelを読み込み、先頭（最左）シートからデータを取り出す。
    緯度(P列)・経度(Q列)が空の行は地図表示できないため除外する。
    """
    import openpyxl

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    if not wb.sheetnames:
        raise ValueError("Excelにシートが存在しません。")

    sheet = wb[wb.sheetnames[0]]  # 先頭シート固定

    records: List[Dict[str, Any]] = []

    # enumerate(..., start=2) で Excelの行番号（ヘッダ1行目の次）を保持
    for excel_row_no, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        lat = _safe_float(row[COL["lat"]] if len(row) > COL["lat"] else None)
        lon = _safe_float(row[COL["lon"]] if len(row) > COL["lon"] else None)

        # 緯度・経度が空の薬局は表示しない
        if lat is None or lon is None:
            continue

        pharmacy_name = row[COL["pharmacy_name"]] if len(row) > COL["pharmacy_name"] else None
        opener_name = row[COL["opener_name"]] if len(row) > COL["opener_name"] else None
        address = row[COL["address"]] if len(row) > COL["address"] else None

        opener_str = str(opener_name).strip() if opener_name is not None else ""

        corp_name = extract_corporation_part_from_opener(opener_str)
        ceo_name = extract_ceo_name_from_opener(opener_str)

        records.append(
            {
                # 一意ID（Excel行番号をそのまま使う）
                "UID": str(excel_row_no),
                "Excel行番号": int(excel_row_no),

                "薬局名": str(pharmacy_name).strip() if pharmacy_name is not None else "",
                "開設者氏名": opener_str,        # E列全文（表示用）
                "法人名": corp_name,             # 法人名（抽出）
                "社長名": ceo_name,              # 氏名（抽出）

                "法人検索キー": corporation_search_key_from_opener(opener_str),
                "社長検索キー": ceo_search_key_from_opener(opener_str),

                "住所": str(address).strip() if address is not None else "",
                "緯度": float(lat),
                "経度": float(lon),

                "管理薬剤師ID": _normalize_id(row[COL["id_manager"]] if len(row) > COL["id_manager"] else None),
                "常勤ID": _normalize_id(row[COL["id_fulltime"]] if len(row) > COL["id_fulltime"] else None),
                "パートID": _normalize_id(row[COL["id_part"]] if len(row) > COL["id_part"] else None),
                "派遣ID": _normalize_id(row[COL["id_temp"]] if len(row) > COL["id_temp"] else None),
                "契約社員ID": _normalize_id(row[COL["id_contract"]] if len(row) > COL["id_contract"] else None),
            }
        )

    df = pd.DataFrame.from_records(records)
    if df.empty:
        raise ValueError("緯度(P)・経度(Q)が入っている行が見つかりませんでした。")
    return df


# =============================================================================
# 地図：読み込みオーバーレイ（ブラウザ側）
# =============================================================================
def add_map_loading_overlay(m: folium.Map, message: str = "検索中・・・地図を読み込み中です") -> None:
    """Leaflet(ブラウザ側)でタイル読み込み中に、地図の上へメッセージを表示する。"""
    msg = _escape_html(message)
    map_name = m.get_name()

    tpl = Template(
        f"""
        {{% macro html(this, kwargs) %}}
        <style>
          .map-loading-overlay {{
            position: absolute;
            inset: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(255,255,255,0.75);
            z-index: 9999;
            font-size: 16px;
            font-weight: 600;
            color: #333;
            pointer-events: none;
          }}
        </style>
        {{% endmacro %}}

        {{% macro script(this, kwargs) %}}
        (function() {{
          var map = {map_name};
          if (!map) return;

          var container = map.getContainer();
          if (!container) return;

          container.style.position = "relative";
          var overlay = document.createElement("div");
          overlay.className = "map-loading-overlay";
          overlay.innerText = "{msg}";
          container.appendChild(overlay);

          function show() {{ overlay.style.display = "flex"; }}
          function hide() {{ overlay.style.display = "none"; }}

          show();

          var hooked = false;
          map.eachLayer(function(layer) {{
            if (layer && layer instanceof L.TileLayer) {{
              hooked = true;
              layer.on("loading", show);
              layer.on("load", hide);
            }}
          }});

          if (!hooked) {{
            window.setTimeout(hide, 1500);
          }}
        }})();
        {{% endmacro %}}
        """
    )

    macro = MacroElement()
    macro._template = tpl
    m.get_root().add_child(macro)


# =============================================================================
# 地図：ピンのポップアップ
# =============================================================================
def _make_popup_html(r: pd.Series) -> str:
    """
    ピンのクリック時に表示する内容をHTMLで作成。

    - 「法人名:（抽出法人名）」行は削除
    - 「開設者氏名:」ラベルを「法人名:」に変更（表示内容はE列全文のまま）
    """
    name = _escape_html(str(r.get("薬局名", "")))
    opener_full = _escape_html(str(r.get("開設者氏名", "")))  # E列全文
    addr = _escape_html(str(r.get("住所", "")))

    lat = float(r["緯度"])
    lon = float(r["経度"])
    gmap = GOOGLE_MAPS_QUERY_URL.format(lat=lat, lon=lon)

    def job_li(label: str, job_id: Optional[str]) -> str:
        if job_id:
            return f'<li><a href="{JOB_BASE_URL + job_id}" target="_blank" rel="noopener noreferrer">{label}</a></li>'
        return f"<li>{label}: IDなし</li>"

    jobs = (
        "<div style='margin-top:6px;'><div><b>求人リンク</b></div>"
        "<ul style='margin:4px 0 0 18px;padding:0;'>"
        f"{job_li('管理薬剤師求人', r.get('管理薬剤師ID'))}"
        f"{job_li('常勤求人', r.get('常勤ID'))}"
        f"{job_li('パート求人', r.get('パートID'))}"
        f"{job_li('派遣求人', r.get('派遣ID'))}"
        f"{job_li('契約社員', r.get('契約社員ID'))}"
        "</ul></div>"
    )

    return f"""
    <div style="font-size:13px;line-height:1.4;max-width:420px;">
      <div style="font-size:14px;"><b>{name}</b></div>
      <div>法人名: {opener_full}</div>
      <div>住所: {addr}</div>
      <div style="margin-top:6px;">
        <a href="{gmap}" target="_blank" rel="noopener noreferrer">Googleマップを開く</a>
      </div>
      {jobs}
    </div>
    """


@st.cache_data(show_spinner=False)
def prepare_points_for_map(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """地図表示に必要な最小情報を作る（キャッシュ対象）。"""
    pts: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        pts.append(
            {
                "lat": float(r["緯度"]),
                "lon": float(r["経度"]),
                "tooltip": str(r.get("薬局名", "")),
                "popup_html": _make_popup_html(r),
            }
        )
    return pts


def build_map(
    center: Tuple[float, float],
    zoom: int,
    points: List[Dict[str, Any]],
    search_point: Optional[SearchPoint],
    radius_km: Optional[float],
    pending_point: Optional[SearchPoint],
) -> folium.Map:
    """Folium地図を構築（薬局=青ピン、検索地点=赤ピン、未確定=橙ピン）。"""
    m = folium.Map(location=center, zoom_start=zoom, control_scale=True)

    # 未確定クリック地点（橙）
    if pending_point is not None:
        folium.Marker(
            location=(pending_point.lat, pending_point.lon),
            tooltip="選択中（未確定）",
            icon=folium.Icon(color="orange", icon="map-marker"),
        ).add_to(m)

    # 確定検索地点（赤）+ 円（半径検索のとき）
    if search_point is not None:
        sp_html = f"""
        <div style="font-size:13px;line-height:1.4;">
          <div style="font-size:14px;"><b>検索地点</b></div>
          <div>緯度: {search_point.lat:.6f}</div>
          <div>経度: {search_point.lon:.6f}</div>
          <div style="margin-top:6px;">
            <a href="{GOOGLE_MAPS_QUERY_URL.format(lat=search_point.lat, lon=search_point.lon)}"
               target="_blank" rel="noopener noreferrer">Googleマップを開く</a>
          </div>
        </div>
        """
        folium.Marker(
            location=(search_point.lat, search_point.lon),
            tooltip="検索地点（赤ピン）",
            popup=folium.Popup(sp_html, max_width=360),
            icon=folium.Icon(color="red", icon="map-marker"),
        ).add_to(m)

        if radius_km and radius_km > 0:
            folium.Circle(
                location=(search_point.lat, search_point.lon),
                radius=radius_km * 1000,
                fill=False,
                weight=2,
            ).add_to(m)

    # 薬局（青）をクラスタ表示
    cluster = MarkerCluster(name="薬局").add_to(m)
    for p in points:
        folium.Marker(
            location=(p["lat"], p["lon"]),
            tooltip=p["tooltip"],
            popup=folium.Popup(p["popup_html"], max_width=480),
            icon=folium.Icon(color="blue", icon="info-sign"),
        ).add_to(cluster)

    folium.LayerControl().add_to(m)
    return m


# =============================================================================
# 住所 -> 緯度経度（外部通信が必要）
# =============================================================================
def _address_candidates(address: str) -> List[str]:
    a = unicodedata.normalize("NFKC", address).replace("　", " ").strip()
    a = re.sub(r"^\s*〒?\s*\d{3}\s*-\s*\d{4}\s*", "", a)
    a = re.sub(r"^\s*〒?\s*\d{7}\s*", "", a).strip()

    b = re.sub(r"(\d+)\s*丁目", r"\1-", a)
    b = re.sub(r"(\d+)\s*番", r"\1-", b)
    b = re.sub(r"(\d+)\s*号", r"\1", b)
    b = re.sub(r"-{2,}", "-", b).strip("- ").strip()

    out: List[str] = []
    for q in [a, b, b + " 日本"]:
        q = q.strip()
        if q and q not in out:
            out.append(q)
    return out


def geocode_address(address: str, bias_df: Optional[pd.DataFrame] = None) -> Optional[SearchPoint]:
    """
    住所・駅名から緯度経度を取得する。
    ★駅名などの曖昧入力対策：候補が複数取れた場合、bias_df（薬局データ）の中心に最も近い候補を採用する。
    """
    if not address.strip():
        return None
    if Nominatim is None or ArcGIS is None or RateLimiter is None:
        return None

    nom = Nominatim(user_agent="tohoku-pharmacy-map-app/1.0", timeout=10)
    nom_geocode = RateLimiter(nom.geocode, min_delay_seconds=1, swallow_exceptions=True)

    arc = ArcGIS(timeout=10)
    arc_geocode = RateLimiter(arc.geocode, min_delay_seconds=1, swallow_exceptions=True)

    candidates: List[SearchPoint] = []

    for q in _address_candidates(address):
        loc = nom_geocode(q, country_codes="jp")
        if loc is not None:
            candidates.append(SearchPoint(lat=float(loc.latitude), lon=float(loc.longitude)))

        loc2 = arc_geocode(q)
        if loc2 is not None:
            candidates.append(SearchPoint(lat=float(loc2.latitude), lon=float(loc2.longitude)))

    if not candidates:
        return None

    # bias無しなら先頭
    if bias_df is None or bias_df.empty:
        return candidates[0]

    # 薬局データ中心に一番近い候補を採用（駅名の誤ヒット対策）
    c_lat = float(bias_df["緯度"].mean())
    c_lon = float(bias_df["経度"].mean())
    best = min(candidates, key=lambda sp: _haversine_km(sp.lat, sp.lon, c_lat, c_lon))
    return best


# =============================================================================
# フィルタ＆一覧リンク
# =============================================================================
def filter_within_radius(df: pd.DataFrame, center: SearchPoint, radius_km: float) -> pd.DataFrame:
    """検索地点から半径内の薬局だけに絞り、距離順で返す。"""
    dists: List[float] = []
    for _, r in df.iterrows():
        dists.append(_haversine_km(center.lat, center.lon, float(r["緯度"]), float(r["経度"])))

    out = df.copy()
    out["距離(km)"] = dists
    out = out[out["距離(km)"] <= radius_km].sort_values("距離(km)")
    return out


def filter_by_corporation(df: pd.DataFrame, corp_query_raw: str) -> pd.DataFrame:
    """法人名検索（完全一致 / スペース無視）。"""
    q = normalize_space_ignored(corp_query_raw)
    if not q:
        return df.iloc[0:0].copy()
    return df[df["法人検索キー"] == q].copy()


def filter_by_ceo_name(df: pd.DataFrame, ceo_query_raw: str) -> pd.DataFrame:
    """社長名検索（完全一致 / 苗字・名前の間のスペース無視）。"""
    q = normalize_space_ignored(ceo_query_raw)
    if not q:
        return df.iloc[0:0].copy()
    return df[df["社長検索キー"] == q].copy()


def add_link_columns(df: pd.DataFrame) -> pd.DataFrame:
    """一覧表示用リンク列を作る（管理薬剤師求人URLを追加）。"""
    out = df.copy()

    out["Googleマップ"] = out.apply(
        lambda r: GOOGLE_MAPS_QUERY_URL.format(lat=r["緯度"], lon=r["経度"]), axis=1
    )

    out["管理薬剤師求人"] = out["管理薬剤師ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["常勤求人"] = out["常勤ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["パート求人"] = out["パートID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["派遣求人"] = out["派遣ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["契約社員"] = out["契約社員ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")

    return out


# =============================================================================
# 選択データのExcel出力
# =============================================================================
def build_selected_excel_bytes(selected_df: pd.DataFrame) -> bytes:
    """選択された薬局（薬局名・法人名・住所）を .xlsx として出力する。"""
    out_cols = ["薬局名", "法人名", "住所"]
    export_df = selected_df[out_cols].copy()

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="選択薬局")
    return bio.getvalue()


# =============================================================================
# 上部ステータス帯
# =============================================================================
def _status_bar_html(
    total: int,
    shown: int,
    mode_label: str,
    radius_km: Optional[float],
    sp: Optional[SearchPoint],
    corp_query: Optional[str],
    ceo_query: Optional[str],
) -> str:
    if mode_label == "法人で検索":
        q = _escape_html(corp_query or "")
        detail = f"<b>法人検索</b>：{q}"
    elif mode_label == "社長名で検索":
        q = _escape_html(ceo_query or "")
        detail = f"<b>社長名検索</b>：{q}（スペース無視で完全一致）"
    else:
        sp_text = "（未指定）" if sp is None else f"緯度 {sp.lat:.5f} / 経度 {sp.lon:.5f}"
        r = "" if radius_km is None else f"<b>半径 {radius_km:.1f} km 以内</b>"
        detail = f"{r}&nbsp;&nbsp;&nbsp;<b>検索点</b>：{sp_text}"

    return f"""
    <div style="
        background:#E8F5E9;
        border:1px solid rgba(0,0,0,0.06);
        padding:10px 12px;
        border-radius:10px;
        font-size:14px;
        display:flex;
        gap:18px;
        flex-wrap:wrap;
        align-items:center;
    ">
      <div><b>読み込み完了</b>：{total:,}件（緯度・経度あり）</div>
      <div><b>検索モード</b>：{mode_label}</div>
      <div>{detail}</div>
      <div><b>表示件数</b>：{shown:,}件</div>
    </div>
    """


# =============================================================================
# メイン
# =============================================================================
def main() -> None:
    st.set_page_config(page_title="薬局マップ", layout="wide")
    _apply_ui_css()

    st.title("薬局検索マップ")
    st.caption("手順：1) Excelを読み込む → 2) 検索方法を選ぶ → 3) 地図と一覧で確認 → 4) ☑で選択しExcel出力")

    # -----------------------------
    # セッション状態（永続化したい値）
    # -----------------------------
    if "clicked_point" not in st.session_state:
        st.session_state.clicked_point = None
    if "pending_point" not in st.session_state:
        st.session_state.pending_point = None
    if "corp_query_raw" not in st.session_state:
        st.session_state.corp_query_raw = ""
    if "ceo_query_raw" not in st.session_state:
        st.session_state.ceo_query_raw = ""

    if "selected_uids" not in st.session_state:
        st.session_state.selected_uids = set()

    # ★ data_editor の状態保持（一覧用）
    if "list_editor_df" not in st.session_state:
        st.session_state.list_editor_df = None
    if "list_editor_uids" not in st.session_state:
        st.session_state.list_editor_uids = []

    # -------------------------------------------------------------------------
    # 1. データ読み込み（サイドバー）
    # -------------------------------------------------------------------------
    st.sidebar.header("1. データを読み込む")

    load_mode = st.sidebar.radio(
        "読み込み方法",
        ["Excelをアップロード", "PC上のファイルパスを指定"],
        index=0,
    )

    file_bytes: Optional[bytes] = None

    if load_mode == "Excelをアップロード":
        st.sidebar.write("Excelファイルを選択してください。")
        uploaded = st.sidebar.file_uploader(
            "Excelファイル",
            type=["xlsm", "xlsx"],
            label_visibility="collapsed",
        )
        if uploaded is not None:
            file_bytes = uploaded.getvalue()
    else:
        import os
        path = st.sidebar.text_input("ファイルパス", value=DEFAULT_LOCAL_FILE)
        if st.sidebar.button("このファイルを読み込む"):
            if not os.path.exists(path):
                st.sidebar.error("指定したパスにファイルが見つかりません。")
            else:
                with open(path, "rb") as f:
                    file_bytes = f.read()

    if file_bytes is None:
        st.info("左の「1. データを読み込む」からExcelを読み込んでください。")
        st.stop()

    try:
        df = load_pharmacy_data(file_bytes)
    except Exception as e:
        st.error(f"読み込みに失敗しました: {e}")
        st.stop()

    # -------------------------------------------------------------------------
    # 2. 検索（サイドバー）
    # -------------------------------------------------------------------------
    st.sidebar.header("2. 検索")

    search_mode = st.sidebar.radio(
        "検索方法",
        [
            "住所・駅名で検索（半径指定）",
            "地図をクリックして指定（半径指定）",
            "法人で検索",
            "社長名で検索",
        ],
        index=0,
    )

    radius_km = st.sidebar.number_input(
        "半径（km）",
        min_value=0.1,
        max_value=200.0,
        value=2.0,
        step=0.5,
        disabled=(search_mode in {"法人で検索", "社長名で検索"}),
    )

    if search_mode == "住所・駅名で検索（半径指定）":
        st.sidebar.caption("例：宮城県仙台市青葉区中央1-1-1 / 仙台駅")
        address = st.sidebar.text_input("住所・駅名", value="")
        if st.sidebar.button("この住所で検索"):
            with st.spinner("住所・駅名から緯度・経度を取得しています..."):
                # ★改修：薬局データ中心に近い候補を採用するため df を渡す
                sp = geocode_address(address, bias_df=df)

            if sp is None:
                st.warning("位置情報を取得できませんでした（社内ネットワーク等でブロックの可能性）。")
            else:
                st.session_state.clicked_point = sp
                st.session_state.pending_point = None
                st.session_state.corp_query_raw = ""
                st.session_state.ceo_query_raw = ""

    elif search_mode == "地図をクリックして指定（半径指定）":
        if st.sidebar.button("検索地点の確定を解除（全件表示）"):
            st.session_state.clicked_point = None
            st.session_state.pending_point = None
            st.session_state.corp_query_raw = ""
            st.session_state.ceo_query_raw = ""

    elif search_mode == "法人で検索":
        st.sidebar.caption("E列から「代表取締役」の前までを法人名として検索します（スペース無視・完全一致）")
        corp_raw = st.sidebar.text_input("法人名", value=st.session_state.corp_query_raw)
        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("この法人で検索"):
                st.session_state.corp_query_raw = corp_raw
                st.session_state.clicked_point = None
                st.session_state.pending_point = None
                st.session_state.ceo_query_raw = ""
        with c2:
            if st.button("解除"):
                st.session_state.corp_query_raw = ""

    else:
        st.sidebar.caption("E列の「代表取締役」の後ろを氏名として検索します（スペース無視・完全一致）")
        st.sidebar.caption("例：神林明仁（推奨） / 神林 明仁（空白は自動で無視）")
        ceo_raw = st.sidebar.text_input("社長名（苗字+名前）", value=st.session_state.ceo_query_raw)
        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("この社長名で検索"):
                st.session_state.ceo_query_raw = ceo_raw
                st.session_state.clicked_point = None
                st.session_state.pending_point = None
                st.session_state.corp_query_raw = ""
        with c2:
            if st.button("解除"):
                st.session_state.ceo_query_raw = ""

    # -------------------------------------------------------------------------
    # 絞り込み
    # -------------------------------------------------------------------------
    corp_query_raw = st.session_state.corp_query_raw.strip()
    ceo_query_raw = st.session_state.ceo_query_raw.strip()

    if search_mode == "法人で検索":
        show_df = filter_by_corporation(df, corp_query_raw) if corp_query_raw else df.iloc[0:0].copy()
        search_point = None
        mode_label = "法人で検索"

    elif search_mode == "社長名で検索":
        show_df = filter_by_ceo_name(df, ceo_query_raw) if ceo_query_raw else df.iloc[0:0].copy()
        search_point = None
        mode_label = "社長名で検索"

    else:
        search_point = st.session_state.clicked_point
        mode_label = "半径検索"
        show_df = df
        if search_point is not None:
            show_df = filter_within_radius(df, search_point, float(radius_km))

    # -------------------------------------------------------------------------
    # ステータス帯
    # -------------------------------------------------------------------------
    st.markdown(
        _status_bar_html(
            total=len(df),
            shown=len(show_df),
            mode_label=mode_label,
            radius_km=(None if mode_label in {"法人で検索", "社長名で検索"} else float(radius_km)),
            sp=search_point,
            corp_query=(corp_query_raw if mode_label == "法人で検索" else None),
            ceo_query=(ceo_query_raw if mode_label == "社長名で検索" else None),
        ),
        unsafe_allow_html=True,
    )

    # -------------------------------------------------------------------------
    # レイアウト：地図を大きく（左=地図、右=一覧）
    # -------------------------------------------------------------------------
    col_map, col_list = st.columns([5, 2], gap="large")

    # 地図中心
    if not show_df.empty:
        center = (float(show_df["緯度"].mean()), float(show_df["経度"].mean()))
    else:
        center = (float(df["緯度"].mean()), float(df["経度"].mean()))
    zoom = 11 if (mode_label in {"法人で検索", "社長名で検索"} or search_point is not None) else 8

    # -------------------------------------------------------------------------
    # 地図
    # -------------------------------------------------------------------------
    with col_map:
        st.subheader("地図")

        pending_point: Optional[SearchPoint] = st.session_state.pending_point

        with st.spinner("検索中・・・地図を表示しています"):
            points = prepare_points_for_map(show_df)

            fmap = build_map(
                center=center,
                zoom=zoom,
                points=points,
                search_point=search_point,
                radius_km=(None if mode_label in {"法人で検索", "社長名で検索"} else float(radius_km)),
                pending_point=pending_point,
            )

            add_map_loading_overlay(fmap, "検索中・・・地図を読み込み中です")

            map_data = st_folium(
                fmap,
                width=None,
                height=820,
                returned_objects=["last_clicked"],
                key="map",
            )

        if search_mode == "地図をクリックして指定（半径指定）":
            clicked = (map_data or {}).get("last_clicked")
            if clicked and "lat" in clicked and "lng" in clicked:
                st.session_state.pending_point = SearchPoint(
                    lat=float(clicked["lat"]),
                    lon=float(clicked["lng"]),
                )

            pending = st.session_state.pending_point
            if pending is not None:
                st.info(f"選択中（未確定）：緯度 {pending.lat:.6f} / 経度 {pending.lon:.6f}")
                if st.button("この地点を検索地点として確定する"):
                    st.session_state.clicked_point = pending
                    st.session_state.pending_point = None
                    st.session_state.corp_query_raw = ""
                    st.session_state.ceo_query_raw = ""

        if mode_label == "法人で検索" and corp_query_raw and show_df.empty:
            st.warning(
                "一致する法人名が見つかりませんでした。\n"
                f"入力（スペース無視）: {normalize_space_ignored(corp_query_raw)}"
            )

        if mode_label == "社長名で検索" and ceo_query_raw and show_df.empty:
            st.warning(
                "一致する社長名が見つかりませんでした。\n"
                f"入力（スペース無視）: {normalize_space_ignored(ceo_query_raw)}\n"
                "※E列に「代表取締役」が無い行は社長名検索の対象外になります。"
            )

    # -------------------------------------------------------------------------
    # 一覧（☑チェック + Excel出力）
    # -------------------------------------------------------------------------
    with col_list:
        st.subheader("一覧")

        view = add_link_columns(show_df).reset_index(drop=True)

        cols = [
            "☑",
            "薬局名",
            "法人名",
            "住所",
            "Googleマップ",
            "管理薬剤師求人",
            "常勤求人",
            "パート求人",
            "派遣求人",
            "契約社員",
        ]

        if view.empty:
            st.info("表示対象がありません。")
        else:
            # 表示中のUID（行対応のキー）
            display_uids = view["UID"].astype(str).tolist()

            # editor用のベース表（毎回組み立てる）
            base = view.drop(columns=["☑"], errors="ignore").copy()
            base.insert(0, "☑", [uid in st.session_state.selected_uids for uid in display_uids])
            base = base[cols].copy()

            # ★表示対象UIDが変わったら editor をリセットしてから差し替え
            if (st.session_state.list_editor_df is None) or (st.session_state.list_editor_uids != display_uids):
                st.session_state.list_editor_df = base
                st.session_state.list_editor_uids = display_uids
                st.session_state.pop("list_editor", None)  # widget内部状態の破棄

            # ★重要：data_editor に渡す DataFrame は固定（list_editor_df）
            edited = st.data_editor(
                st.session_state.list_editor_df,
                use_container_width=True,
                height=720,
                hide_index=True,
                column_config={
                    "☑": st.column_config.CheckboxColumn("☑"),
                    "Googleマップ": st.column_config.LinkColumn("Googleマップ", display_text="開く"),
                    "管理薬剤師求人": st.column_config.LinkColumn("管理", display_text="開く"),
                    "常勤求人": st.column_config.LinkColumn("常勤", display_text="開く"),
                    "パート求人": st.column_config.LinkColumn("パート", display_text="開く"),
                    "派遣求人": st.column_config.LinkColumn("派遣", display_text="開く"),
                    "契約社員": st.column_config.LinkColumn("契約", display_text="開く"),
                },
                key="list_editor",
            )

            # ★重要：ここで list_editor_df を edited で上書きしない（1回目消える原因になりがち）
            # st.session_state.list_editor_df = edited  # ←しない

            # ★チェック結果を selected_uids に反映（indexズレに依存しない）
            uids_for_zip = st.session_state.list_editor_uids
            checked_uids = {uid for uid, checked in zip(uids_for_zip, edited["☑"].tolist()) if checked}

            selected = set(st.session_state.selected_uids)
            selected -= set(uids_for_zip)   # 表示中ぶんを一旦除外
            selected |= checked_uids        # 表示中のチェック分を反映
            st.session_state.selected_uids = selected

        # -----------------------------
        # 選択結果のExcel出力
        # -----------------------------
        st.markdown("---")
        c1, c2 = st.columns([1, 1])
        with c1:
            st.write(f"☑ 選択中：{len(st.session_state.selected_uids):,}件")
        with c2:
            if st.button("☑選択をすべて解除"):
                st.session_state.selected_uids = set()
                st.session_state.list_editor_df = None
                st.session_state.list_editor_uids = []
                st.session_state.pop("list_editor", None)

        sel_df = df[df["UID"].astype(str).isin(st.session_state.selected_uids)].copy()
        if not sel_df.empty:
            xlsx_bytes = build_selected_excel_bytes(sel_df)
            st.download_button(
                label="選択した薬局をExcelでダウンロード",
                data=xlsx_bytes,
                file_name="選択薬局リスト.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("☑を付けると、ここからExcelをダウンロードできます。")


if __name__ == "__main__":
    main()
