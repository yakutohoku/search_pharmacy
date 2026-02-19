"""
薬局検索マップ（Streamlit + Folium）

ご要望対応（2026-02-19版）
① 一覧クリックの「1つ前が強調」問題を解消：
   - 一覧の選択状態を、地図描画“前”に session_state から読み取り反映
② 薬局を選択したら「2km圏内」へズーム：
   - 選択薬局を中心にズーム（既定ズーム=14）＋2km円を表示
③ Excel出力は複数選択（☑）を維持：
   - Excel出力タブの data_editor ☑ は複数選択可（従来通り）
④ ☑操作のたびに地図が広角へ戻る問題を解消：
   - st_folium から center/zoom を受け取り、セッションに保存して次回描画に反映（ズーム維持）

注意
- 住所/駅名検索（ジオコーディング）は外部通信が必要です（ネットワーク制限がある場合は失敗します）
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
JOB_BASE_URL = "https://yaku-job.com/preview/"
GOOGLE_MAPS_QUERY_URL = "https://www.google.com/maps/search/?api=1&query={lat},{lon}"

REFERENCE_FOLDER = r"\\file-tky\Section\薬剤師共有\★【支店・課】_東北\東北\薬局検索"
REFERENCE_FILE_LIST = [
    "東北 薬局リスト.xlsm",
    "北海道 薬局リスト.xlsm",
]

PREF_LIST = ["", "北海道", "青森", "秋田", "岩手", "宮城", "山形", "福島"]

# Excel列（0-based index）
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

# 薬局を選択した時のズーム（2km圏内目安）
FOCUS_ZOOM = 14
FOCUS_RADIUS_KM = 2.0


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
    if x is None:
        return None
    s = str(x).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    if s.endswith(".0"):
        s2 = s[:-2]
        if s2.isdigit():
            return s2
    return s


def _escape_html(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
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
    t = unicodedata.normalize("NFKC", text or "")
    t = t.replace("　", " ")
    t = re.sub(r"\s+", "", t)
    return t


def extract_corporation_part_from_opener(opener_raw: str) -> str:
    t = unicodedata.normalize("NFKC", opener_raw or "").replace("　", " ").strip()
    m = re.search(r"(代表取締役社長|代表取締役|代表社員|理事長|院長|代表者)", t)
    if m:
        t = t[: m.start()].strip()
    return t


def corporation_search_key_from_opener(opener_raw: str) -> str:
    return normalize_space_ignored(extract_corporation_part_from_opener(opener_raw))


def extract_ceo_name_from_opener(opener_raw: str) -> str:
    t = unicodedata.normalize("NFKC", opener_raw or "").replace("　", " ").strip()
    m = re.search(r"(代表取締役社長|代表取締役|代表社員|理事長|院長|代表者)\s*(.*)$", t)
    if not m:
        return ""
    name_part = (m.group(2) or "").strip()
    for prefix in ["社長", "会長", "CEO", "ＣＥＯ"]:
        if name_part.startswith(prefix):
            name_part = name_part[len(prefix):].strip()
    return name_part


def ceo_search_key_from_opener(opener_raw: str) -> str:
    return normalize_space_ignored(extract_ceo_name_from_opener(opener_raw))


# =============================================================================
# UI（日本語化CSS）
# =============================================================================
def _apply_ui_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.2rem; padding-bottom: 1.0rem; }
        header[data-testid="stHeader"] { height: 0.2rem; }

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

        /* file_uploader の英語を隠して日本語を重ねる */
        div[data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderDropzoneInstructions"] { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] > div:nth-child(1) { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] p { display:none !important; }
        div[data-testid="stFileUploaderDropzone"] small { display:none !important; }

        div[data-testid="stFileUploaderDropzone"]::before{
            content:"ここにExcelファイルをドラッグ＆ドロップ、または「ファイルを選択」を押してください";
            display:block;
            padding:0.25rem 0 0.6rem 0;
            font-size:0.95rem;
            font-weight:600;
            color:rgba(49,51,63,0.9);
        }

        div[data-testid="stFileUploaderDropzone"] button{ position:relative; }
        div[data-testid="stFileUploaderDropzone"] button *{ font-size:0px !important; }
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
    import openpyxl

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    if not wb.sheetnames:
        raise ValueError("Excelにシートが存在しません。")

    sheet = wb[wb.sheetnames[0]]

    records: List[Dict[str, Any]] = []
    for excel_row_no, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        lat = _safe_float(row[COL["lat"]] if len(row) > COL["lat"] else None)
        lon = _safe_float(row[COL["lon"]] if len(row) > COL["lon"] else None)
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
                "UID": str(excel_row_no),
                "Excel行番号": int(excel_row_no),
                "薬局名": str(pharmacy_name).strip() if pharmacy_name is not None else "",
                "開設者氏名": opener_str,
                "法人名": corp_name,
                "社長名": ceo_name,
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
# 地図：読み込みオーバーレイ
# =============================================================================
def add_map_loading_overlay(m: folium.Map, message: str = "検索中・・・地図を読み込み中です") -> None:
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
# ピンpopup
# =============================================================================
def _make_popup_html(r: pd.Series) -> str:
    name = _escape_html(str(r.get("薬局名", "")))
    opener_full = _escape_html(str(r.get("開設者氏名", "")))
    corp = _escape_html(str(r.get("法人名", "")))
    ceo = _escape_html(str(r.get("社長名", "")))
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
      <div>法人名: {corp}</div>
      <div>社長名: {ceo}</div>
      <div>開設者氏名: {opener_full}</div>
      <div>住所: {addr}</div>
      <div style="margin-top:6px;">
        <a href="{gmap}" target="_blank" rel="noopener noreferrer">Googleマップを開く</a>
      </div>
      {jobs}
    </div>
    """


@st.cache_data(show_spinner=False)
def prepare_points_for_map(df: pd.DataFrame) -> List[Dict[str, Any]]:
    pts: List[Dict[str, Any]] = []
    for _, r in df.iterrows():
        pts.append(
            {
                "uid": str(r.get("UID", "")),
                "lat": float(r["緯度"]),
                "lon": float(r["経度"]),
                "tooltip": str(r.get("薬局名", "")),
                "popup_html": _make_popup_html(r),
            }
        )
    return pts


# =============================================================================
# 住所→緯度経度
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

    if bias_df is None or bias_df.empty:
        return candidates[0]

    c_lat = float(bias_df["緯度"].mean())
    c_lon = float(bias_df["経度"].mean())
    best = min(candidates, key=lambda sp: _haversine_km(sp.lat, sp.lon, c_lat, c_lon))
    return best


# =============================================================================
# フィルタ
# =============================================================================
def filter_within_radius(df: pd.DataFrame, center: SearchPoint, radius_km: float) -> pd.DataFrame:
    dists: List[float] = []
    for _, r in df.iterrows():
        dists.append(_haversine_km(center.lat, center.lon, float(r["緯度"]), float(r["経度"])))

    out = df.copy()
    out["距離(km)"] = dists
    out = out[out["距離(km)"] <= radius_km].sort_values("距離(km)")
    return out


def filter_by_corporation(df: pd.DataFrame, corp_query_raw: str) -> pd.DataFrame:
    q = normalize_space_ignored(corp_query_raw)
    if not q:
        return df.iloc[0:0].copy()
    return df[df["法人検索キー"] == q].copy()


def filter_by_ceo_name(df: pd.DataFrame, ceo_query_raw: str) -> pd.DataFrame:
    q = normalize_space_ignored(ceo_query_raw)
    if not q:
        return df.iloc[0:0].copy()
    return df[df["社長検索キー"] == q].copy()


def filter_by_pharmacy(df: pd.DataFrame, name_query: str, pref: str) -> pd.DataFrame:
    out = df.copy()
    q = (name_query or "").strip()
    if q:
        out = out[out["薬局名"].str.contains(q, na=False)]
    if pref:
        out = out[out["住所"].str.contains(pref, na=False)]
    return out.copy()


def add_link_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Googleマップ"] = out.apply(lambda r: GOOGLE_MAPS_QUERY_URL.format(lat=r["緯度"], lon=r["経度"]), axis=1)
    out["管理薬剤師求人"] = out["管理薬剤師ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["常勤求人"] = out["常勤ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["パート求人"] = out["パートID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["派遣求人"] = out["派遣ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    out["契約社員"] = out["契約社員ID"].apply(lambda x: (JOB_BASE_URL + x) if x else "")
    return out


# =============================================================================
# Excel出力
# =============================================================================
def build_selected_excel_bytes(selected_df: pd.DataFrame) -> bytes:
    out_cols = ["薬局名", "法人名", "住所"]
    export_df = selected_df[out_cols].copy()

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="選択薬局")
    return bio.getvalue()


# =============================================================================
# ステータス帯
# =============================================================================
def _status_bar_html(
    total: int,
    shown: int,
    mode_label: str,
    radius_km: Optional[float],
    sp: Optional[SearchPoint],
    corp_query: Optional[str],
    ceo_query: Optional[str],
    pharmacy_query: Optional[str],
    pref_query: Optional[str],
) -> str:
    if mode_label == "法人で検索":
        detail = f"<b>法人検索</b>：{_escape_html(corp_query or '')}"
    elif mode_label == "社長名で検索":
        detail = f"<b>社長名検索</b>：{_escape_html(ceo_query or '')}（スペース無視で完全一致）"
    elif mode_label == "薬局で検索":
        q = _escape_html(pharmacy_query or "")
        p = _escape_html(pref_query or "")
        detail = f"<b>薬局名</b>：{q}&nbsp;&nbsp;&nbsp;<b>都道府県</b>：{p if p else '（指定なし）'}"
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
# 地図構築（強調UID対応）
# =============================================================================
def build_map(
    center: Tuple[float, float],
    zoom: int,
    points: List[Dict[str, Any]],
    search_point: Optional[SearchPoint],
    radius_km: Optional[float],
    pending_point: Optional[SearchPoint],
    highlight_uid: Optional[str] = None,
    focus_circle_center: Optional[SearchPoint] = None,
    focus_circle_radius_km: Optional[float] = None,
) -> folium.Map:
    m = folium.Map(location=center, zoom_start=zoom, control_scale=True)

    if pending_point is not None:
        folium.Marker(
            location=(pending_point.lat, pending_point.lon),
            tooltip="選択中（未確定）",
            icon=folium.Icon(color="orange", icon="map-marker"),
        ).add_to(m)

    if search_point is not None:
        folium.Marker(
            location=(search_point.lat, search_point.lon),
            tooltip="検索地点（赤ピン）",
            icon=folium.Icon(color="red", icon="map-marker"),
        ).add_to(m)

        if radius_km and radius_km > 0:
            folium.Circle(
                location=(search_point.lat, search_point.lon),
                radius=radius_km * 1000,
                fill=False,
                weight=2,
            ).add_to(m)

    # 選択薬局の2km円（ご要望②）
    if focus_circle_center is not None and focus_circle_radius_km and focus_circle_radius_km > 0:
        folium.Circle(
            location=(focus_circle_center.lat, focus_circle_center.lon),
            radius=focus_circle_radius_km * 1000,
            fill=False,
            weight=2,
        ).add_to(m)

    cluster = MarkerCluster(name="薬局").add_to(m)

    for p in points:
        uid = str(p.get("uid", ""))
        is_highlight = (highlight_uid is not None and uid == str(highlight_uid))
        icon_color = "red" if is_highlight else "blue"
        icon_name = "star" if is_highlight else "info-sign"

        folium.Marker(
            location=(p["lat"], p["lon"]),
            tooltip=p["tooltip"],
            popup=folium.Popup(p["popup_html"], max_width=480),
            icon=folium.Icon(color=icon_color, icon=icon_name),
        ).add_to(cluster)

    folium.LayerControl().add_to(m)
    return m


# =============================================================================
# 一覧クリックの選択状態を、地図描画の前に同期（ご要望①）
# =============================================================================
def sync_selected_pin_from_list_table(show_df: pd.DataFrame) -> None:
    state = st.session_state.get("list_click_df")
    if not state:
        return

    selection = None
    if isinstance(state, dict):
        selection = state.get("selection")
    else:
        selection = getattr(state, "selection", None)

    if not selection:
        return

    rows = selection.get("rows") if isinstance(selection, dict) else None
    if not rows:
        return

    idx = int(rows[0])
    if idx < 0:
        return

    try:
        uid = str(show_df.reset_index(drop=True).iloc[idx]["UID"])
        st.session_state.selected_pin_uid = uid
    except Exception:
        return


# =============================================================================
# main
# =============================================================================
def main() -> None:
    st.set_page_config(page_title="薬局マップ", layout="wide")
    _apply_ui_css()

    st.title("薬局検索マップ")
    st.caption("手順：1) Excelを読み込む → 2) 検索方法を選ぶ → 3) 地図と一覧で確認 → 4) ☑で選択しExcel出力")

    # セッション
    st.session_state.setdefault("clicked_point", None)
    st.session_state.setdefault("pending_point", None)
    st.session_state.setdefault("corp_query_raw", "")
    st.session_state.setdefault("ceo_query_raw", "")
    st.session_state.setdefault("pharmacy_query_raw", "")
    st.session_state.setdefault("pref_query", "")
    st.session_state.setdefault("selected_pin_uid", None)

    st.session_state.setdefault("selected_uids", set())
    st.session_state.setdefault("list_editor_df", None)
    st.session_state.setdefault("list_editor_uids", [])

    # 地図の表示状態を保持（ご要望④）
    st.session_state.setdefault("map_view_center", None)  # (lat, lon)
    st.session_state.setdefault("map_view_zoom", None)    # int

    # ---- 読み込み
    st.sidebar.header("1. データを読み込む")
    uploaded = st.sidebar.file_uploader("Excelファイル", type=["xlsm", "xlsx"], label_visibility="collapsed")
    file_bytes = uploaded.getvalue() if uploaded is not None else None

    # 未読込時のみ参照フォルダ表示
    if file_bytes is None:
        st.sidebar.markdown("---")
        st.sidebar.markdown("★参照フォルダ（表示のみ）")
        st.sidebar.code(REFERENCE_FOLDER, language="text")

        escaped_path = REFERENCE_FOLDER.replace("\\", "\\\\")
        st.sidebar.markdown(
            f"""
            <button onclick="navigator.clipboard.writeText('{escaped_path}')"
            style="
                background-color:#f0f2f6;
                border:1px solid #ccc;
                padding:6px 10px;
                border-radius:6px;
                cursor:pointer;
                font-size:13px;">
                フォルダのパスをコピー
            </button>
            """,
            unsafe_allow_html=True,
        )

        st.sidebar.markdown("★リスト一覧（Cloudでは手動定義）")
        for f in REFERENCE_FILE_LIST:
            st.sidebar.write(f"・{f}")

        st.info("左の「1. データを読み込む」からExcelを読み込んでください。")
        st.stop()

    df = load_pharmacy_data(file_bytes)

    # ---- 検索
    st.sidebar.header("2. 検索")
    search_mode = st.sidebar.radio(
        "検索方法",
        [
            "住所・駅名で検索（半径指定）",
            "地図をクリックして指定（半径指定）",
            "法人で検索",
            "社長名で検索",
            "薬局で検索",
        ],
        index=0,
    )

    radius_km = st.sidebar.number_input(
        "半径（km）",
        min_value=0.1,
        max_value=200.0,
        value=2.0,
        step=0.5,
        disabled=(search_mode in {"法人で検索", "社長名で検索", "薬局で検索"}),
    )

    if search_mode == "住所・駅名で検索（半径指定）":
        address = st.sidebar.text_input("住所・駅名", value="")
        if st.sidebar.button("この住所で検索"):
            with st.spinner("住所・駅名から緯度・経度を取得しています..."):
                sp = geocode_address(address, bias_df=df)
            if sp is None:
                st.warning("位置情報を取得できませんでした（ネットワーク制限の可能性）。")
            else:
                st.session_state.clicked_point = sp
                st.session_state.pending_point = None
                st.session_state.corp_query_raw = ""
                st.session_state.ceo_query_raw = ""
                st.session_state.pharmacy_query_raw = ""
                st.session_state.pref_query = ""
                st.session_state.selected_pin_uid = None

    elif search_mode == "地図をクリックして指定（半径指定）":
        if st.sidebar.button("検索地点の確定を解除（全件表示）"):
            st.session_state.clicked_point = None
            st.session_state.pending_point = None
            st.session_state.selected_pin_uid = None

    elif search_mode == "法人で検索":
        corp_raw = st.sidebar.text_input("法人名", value=st.session_state.corp_query_raw)
        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("この法人で検索"):
                st.session_state.corp_query_raw = corp_raw
                st.session_state.clicked_point = None
                st.session_state.pending_point = None
                st.session_state.selected_pin_uid = None
        with c2:
            if st.button("解除"):
                st.session_state.corp_query_raw = ""

    elif search_mode == "社長名で検索":
        ceo_raw = st.sidebar.text_input("社長名（苗字+名前）", value=st.session_state.ceo_query_raw)
        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("この社長名で検索"):
                st.session_state.ceo_query_raw = ceo_raw
                st.session_state.clicked_point = None
                st.session_state.pending_point = None
                st.session_state.selected_pin_uid = None
        with c2:
            if st.button("解除"):
                st.session_state.ceo_query_raw = ""

    else:
        name_q = st.sidebar.text_input("薬局名（部分一致）", value=st.session_state.pharmacy_query_raw)

        default_index = 0
        if st.session_state.pref_query in PREF_LIST:
            default_index = PREF_LIST.index(st.session_state.pref_query)

        pref_q = st.sidebar.selectbox("都道府県", PREF_LIST, index=default_index)

        c1, c2 = st.sidebar.columns(2)
        with c1:
            if st.button("この条件で検索"):
                st.session_state.pharmacy_query_raw = name_q
                st.session_state.pref_query = pref_q
                st.session_state.selected_pin_uid = None
        with c2:
            if st.button("解除"):
                st.session_state.pharmacy_query_raw = ""
                st.session_state.pref_query = ""

    # ---- 絞り込み
    corp_query_raw = st.session_state.corp_query_raw.strip()
    ceo_query_raw = st.session_state.ceo_query_raw.strip()
    pharmacy_query_raw = st.session_state.pharmacy_query_raw.strip()
    pref_query = (st.session_state.pref_query or "").strip()

    if search_mode == "法人で検索":
        show_df = filter_by_corporation(df, corp_query_raw) if corp_query_raw else df.iloc[0:0].copy()
        search_point = None
        mode_label = "法人で検索"
    elif search_mode == "社長名で検索":
        show_df = filter_by_ceo_name(df, ceo_query_raw) if ceo_query_raw else df.iloc[0:0].copy()
        search_point = None
        mode_label = "社長名で検索"
    elif search_mode == "薬局で検索":
        show_df = filter_by_pharmacy(df, pharmacy_query_raw, pref_query) if (pharmacy_query_raw or pref_query) else df.iloc[0:0].copy()
        search_point = None
        mode_label = "薬局で検索"
    else:
        search_point = st.session_state.clicked_point
        mode_label = "半径検索"
        show_df = df
        if search_point is not None:
            show_df = filter_within_radius(df, search_point, float(radius_km))

    # ①対策：一覧のクリック選択状態を、地図描画前に反映
    sync_selected_pin_from_list_table(show_df)

    # ---- ステータス
    st.markdown(
        _status_bar_html(
            total=len(df),
            shown=len(show_df),
            mode_label=mode_label,
            radius_km=None if mode_label in {"法人で検索", "社長名で検索", "薬局で検索"} else float(radius_km),
            sp=search_point,
            corp_query=corp_query_raw if mode_label == "法人で検索" else None,
            ceo_query=ceo_query_raw if mode_label == "社長名で検索" else None,
            pharmacy_query=pharmacy_query_raw if mode_label == "薬局で検索" else None,
            pref_query=pref_query if mode_label == "薬局で検索" else None,
        ),
        unsafe_allow_html=True,
    )

    # ---- レイアウト
    col_map, col_list = st.columns([5, 2], gap="large")

    highlight_uid = st.session_state.selected_pin_uid

    # ②&④：中心/ズームの決定ロジック
    focus_center: Optional[SearchPoint] = None
    focus_radius_km: Optional[float] = None

    if highlight_uid:
        row = show_df[show_df["UID"].astype(str) == str(highlight_uid)]
        if not row.empty:
            focus_center = SearchPoint(lat=float(row.iloc[0]["緯度"]), lon=float(row.iloc[0]["経度"]))
            focus_radius_km = FOCUS_RADIUS_KM
            center = (focus_center.lat, focus_center.lon)
            zoom = FOCUS_ZOOM
            st.session_state.map_view_center = center
            st.session_state.map_view_zoom = zoom
        else:
            center = st.session_state.map_view_center or (float(df["緯度"].mean()), float(df["経度"].mean()))
            zoom = st.session_state.map_view_zoom or (11 if (mode_label in {"法人で検索", "社長名で検索", "薬局で検索"} or search_point is not None) else 8)
    else:
        default_center = (float(show_df["緯度"].mean()), float(show_df["経度"].mean())) if not show_df.empty else (float(df["緯度"].mean()), float(df["経度"].mean()))
        default_zoom = 11 if (mode_label in {"法人で検索", "社長名で検索", "薬局で検索"} or search_point is not None) else 8
        center = st.session_state.map_view_center or default_center
        zoom = st.session_state.map_view_zoom or default_zoom

    # ---- 地図
    with col_map:
        st.subheader("地図")

        pending_point: Optional[SearchPoint] = st.session_state.pending_point
        with st.spinner("検索中・・・地図を表示しています"):
            points = prepare_points_for_map(show_df)
            fmap = build_map(
                center=center,
                zoom=int(zoom),
                points=points,
                search_point=search_point,
                radius_km=None if mode_label in {"法人で検索", "社長名で検索", "薬局で検索"} else float(radius_km),
                pending_point=pending_point,
                highlight_uid=highlight_uid,
                focus_circle_center=focus_center,
                focus_circle_radius_km=focus_radius_km,
            )
            add_map_loading_overlay(fmap, "検索中・・・地図を読み込み中です")

            map_data = st_folium(
                fmap,
                width=None,
                height=820,
                returned_objects=["last_clicked", "center", "zoom"],
                key="map",
            )

        if map_data:
            c = map_data.get("center")
            z = map_data.get("zoom")
            if c and "lat" in c and "lng" in c:
                st.session_state.map_view_center = (float(c["lat"]), float(c["lng"]))
            if z is not None:
                try:
                    st.session_state.map_view_zoom = int(z)
                except Exception:
                    pass

        if search_mode == "地図をクリックして指定（半径指定）":
            clicked = (map_data or {}).get("last_clicked")
            if clicked and "lat" in clicked and "lng" in clicked:
                st.session_state.pending_point = SearchPoint(lat=float(clicked["lat"]), lon=float(clicked["lng"]))

            pending = st.session_state.pending_point
            if pending is not None:
                st.info(f"選択中（未確定）：緯度 {pending.lat:.6f} / 経度 {pending.lon:.6f}")
                if st.button("この地点を検索地点として確定する"):
                    st.session_state.clicked_point = pending
                    st.session_state.pending_point = None
                    st.session_state.selected_pin_uid = None

    # ---- 一覧
    with col_list:
        st.subheader("一覧")
        if show_df.empty:
            st.info("表示対象がありません。")
            st.stop()

        tab_list, tab_export = st.tabs(["一覧（クリックで地図で強調）", "Excel出力（☑で選択）"])

        with tab_list:
            st.caption("行をクリックすると、地図上の該当ピンを赤で強調し、2km圏内へズームします。")

            list_view = show_df[["UID", "薬局名", "法人名", "住所"]].copy().reset_index(drop=True)

            try:
                st.dataframe(
                    list_view.drop(columns=["UID"]),
                    use_container_width=True,
                    height=520,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key="list_click_df",
                )
            except TypeError:
                st.info("この環境のStreamlitでは『行クリック選択』が使えません。必要ならStreamlitの更新をご検討ください。")

            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button("地図の強調を解除"):
                    st.session_state.selected_pin_uid = None
            with c2:
                if highlight_uid:
                    st.write("強調中：1件")

        with tab_export:
            st.caption("☑で複数選択できます。選択した薬局は下からExcelダウンロードできます。")
            view = add_link_columns(show_df).reset_index(drop=True)

            cols = ["☑", "薬局名", "法人名", "住所", "Googleマップ", "管理薬剤師求人", "常勤求人", "パート求人", "派遣求人", "契約社員"]
            display_uids = view["UID"].astype(str).tolist()

            base = view.drop(columns=["☑"], errors="ignore").copy()
            base.insert(0, "☑", [uid in st.session_state.selected_uids for uid in display_uids])
            base = base[cols].copy()

            if (st.session_state.list_editor_df is None) or (st.session_state.list_editor_uids != display_uids):
                st.session_state.list_editor_df = base
                st.session_state.list_editor_uids = display_uids
                st.session_state.pop("list_editor", None)

            edited = st.data_editor(
                st.session_state.list_editor_df,
                use_container_width=True,
                height=520,
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

            uids_for_zip = st.session_state.list_editor_uids
            checked_uids = {uid for uid, checked in zip(uids_for_zip, edited["☑"].tolist()) if checked}

            selected = set(st.session_state.selected_uids)
            selected -= set(uids_for_zip)
            selected |= checked_uids
            st.session_state.selected_uids = selected

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
