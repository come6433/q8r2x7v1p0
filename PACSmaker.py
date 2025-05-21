import openpyxl
import pandas as pd
import folium
import base64
import os
from folium.plugins import Draw, LocateControl
import datetime
import requests
import sys
import re

CURRENT_VERSION = "1.4"
UPDATE_DATE = "2025-05-21"
GITHUB_RAW_URL = "https://raw.githubusercontent.com/come6433/q8r2x7v1p0/main/PACSmaker.py"

def get_version_from_text(text):
    m = re.search(r'CURRENT_VERSION\s*=\s*["\']([\d\.]+)["\']', text)
    return m.group(1) if m else None

def normalize_version(v):
    parts = v.split(".")
    if len(parts) == 2 and parts[1].endswith("0"):
        parts[1] = parts[1].rstrip("0")
        if parts[1] == "":
            parts[1] = "0"
    return tuple(map(int, parts))

def version_compare(a, b):
    na = normalize_version(a)
    nb = normalize_version(b)
    return (na > nb) - (na < nb)  # 1: a>b, -1: a<b, 0: a==b

def check_and_update():
    try:
        print("업데이트 확인 중 ...")
        r = requests.get(GITHUB_RAW_URL, timeout=5)
        if r.status_code == 200:
            remote_text = r.text
            remote_version = get_version_from_text(remote_text)
            if remote_version and version_compare(remote_version, CURRENT_VERSION) > 0:
                print(f"\n새 버전({remote_version})이 있습니다. 자동 업데이트를 진행합니다.")
                try:
                    os.rename(__file__, __file__ + ".bak")
                except Exception:
                    pass
                with open(__file__, "w", encoding="utf-8") as f:
                    f.write(remote_text)
                print("업데이트 완료! 프로그램을 다시 실행해 주세요.")
                sys.exit(0)
            else:
                print("최신 버전입니다.")
        else:
            print("업데이트 서버 연결 실패:", r.status_code)
    except Exception as e:
        print("업데이트 확인 중 오류:", e)

check_and_update()

print("=" * 40)
print("      PACS 저상게시대 지도 생성기")
print("=" * 40)
print("버전:        ", CURRENT_VERSION)
print("업데이트:    ", UPDATE_DATE)
print("- 관리번호 세분화")
print("- 검색 기능 보완")
print("- 특이사항 항목 추가")
print("=" * 40)
print("관리목록.xlsx 파일을 읽는 중...\n")

# --- 병합 셀 포함된 Excel 파일 읽기 ---
wb = openpyxl.load_workbook('관리목록.xlsx', data_only=True)
ws = wb.active
data = []
for row in ws.iter_rows(min_row=3, values_only=True):
    data.append(row)
df = pd.DataFrame(data)
col_names = []
for cell in ws[2]:
    col_names.append(cell.value)
df.columns = col_names

base_cols = ['설치장소', '단수', '관리번호', '위도', '경도']
extra_cols = [c for c in df.columns if c not in ['설치장소', '단수', '위도', '경도', '순번'] and c is not None]
df = df.dropna(subset=['설치장소', '관리번호', '위도', '경도'])
print("지도 작성 중 ...")

center_lat = df.iloc[0]['위도']
center_lon = df.iloc[0]['경도']

def image_to_base64(path):
    with open(path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

m = folium.Map(location=[center_lat, center_lon], zoom_start=13, tiles=None)
LocateControl(auto_start=False, flyTo=True, keepCurrentZoomLevel=True).add_to(m)

vworld_base = "https://xdworld.vworld.kr/2d/Base/service/{z}/{x}/{y}.png"
folium.TileLayer(
    tiles=vworld_base,
    attr="VWorld Base",
    name="VWorld 일반지도",
    overlay=False,
    control=True
).add_to(m)

vworld_sat = "https://xdworld.vworld.kr/2d/Satellite/service/{z}/{x}/{y}.jpeg"
folium.TileLayer(
    tiles=vworld_sat,
    attr="VWorld Satellite",
    name="VWorld 위성지도",
    overlay=False,
    control=True
).add_to(m)

naver_tile = "https://map.pstatic.net/nrs/api/v1/raster/satellite/{z}/{x}/{y}.jpg?version=6.03"
naver_layer = folium.TileLayer(
    tiles=naver_tile,
    attr="Naver Satellite",
    name="네이버 위성지도",
    overlay=False,
    control=True
)
naver_layer.add_to(m)

def get_color(단수, marker_no):
    if str(marker_no).startswith('특'):
        return 'pink'
    try:
        if int(단수) == 1:
            return 'blue'
        else:
            return 'red'
    except Exception:
        return 'blue'

fg1 = folium.FeatureGroup(name='1단 (파랑)').add_to(m)
fg2 = folium.FeatureGroup(name='2단 (빨강)').add_to(m)
fg_special = folium.FeatureGroup(name='특이(핑크)').add_to(m)

grouped = df.groupby('마커번호')
for marker_no, group in grouped:
    first = group.iloc[0]
    lat, lon = first['위도'], first['경도']
    단수 = first['단수'] if (pd.notnull(first['단수']) and not str(marker_no).startswith('특')) else 1
    설치장소 = first['설치장소'] if '설치장소' in group.columns else ""
    popup_html = f"<div style='text-align:center;'><b class='popup-title'>{설치장소}</b><br>"
    for _, row in group.iterrows():
        관리번호 = str(row['관리번호'])
        image_path = f"images/{관리번호}.jpg"
        if os.path.exists(image_path):
            img_base64 = image_to_base64(image_path)
            popup_html += (
                f"<img src='data:image/jpeg;base64,{img_base64}' width='200' class='popup-img' "
                "style='cursor:zoom-in;display:block;margin:0 auto;'><br>"
            )
            table_width = "200px"
        else:
            table_width = "180px"
        popup_html += f"<table style='border-collapse:collapse; width:{table_width}; margin:8px auto 0 auto;'>"
        popup_html += (
            "<tr>"
            "<td style='border:1px solid #000; padding:4px 8px; background:#f0f0f0; font-weight:bold; width:70px;'>관리번호</td>"
            f"<td style='border:1px solid #000; padding:4px 8px;'>{관리번호}</td>"
            "</tr>"
        )
        for col in df.columns:
            if col in ['마커번호', '관리번호', '위도', '경도', '설치장소', '단수', '순번']:
                continue
            val = row[col] if pd.notnull(row[col]) else ""
            popup_html += (
                "<tr>"
                f"<td style='border:1px solid #000; padding:4px 8px; background:#f0f0f0; font-weight:bold; width:70px;'>{col}</td>"
                f"<td style='border:1px solid #000; padding:4px 8px;'>{val}</td>"
                "</tr>"
            )
        popup_html += "</table><br>"
    popup_html += "</div>"
    icon_html = f"""
        <div style="
            background-color:{get_color(단수, marker_no)};
            color:white;
            border-radius:50%;
            text-align:center;
            width:24px;
            height:24px;
            line-height:24px;
            font-size:12px;">
            {marker_no}
        </div>
    """
    marker = folium.Marker(
        location=[lat, lon],
        icon=folium.DivIcon(html=icon_html),
        popup=folium.Popup(popup_html, max_width=250)
    )
    if str(marker_no).startswith('특'):
        fg_special.add_child(marker)
    elif 단수 == 1:
        fg1.add_child(marker)
    else:
        fg2.add_child(marker)

count_1 = (df['단수'] == 1).sum()
count_2 = (df['단수'] == 2).sum()
count_special = df['마커번호'].apply(lambda x: str(x).startswith('특')).sum()
legend_html = f"""
<div id="legend" style="
    position: fixed; 
    bottom: 50px; left: 50px; width: 200px; height: 100px; 
    background-color: white; 
    border:2px solid grey; 
    z-index:9999; 
    font-size:14px;
    padding: 10px;
    ">
    <b>범례</b><br>
    <i style="background:blue; width:15px; height:15px; display:inline-block; border-radius:50%;"></i> 1단 - {count_1}개<br>
    <i style="background:red; width:15px; height:15px; display:inline-block; border-radius:50%;"></i> 2단 - {count_2}개<br>
    <i style="background:pink; width:15px; height:15px; display:inline-block; border-radius:50%;"></i> 특이(핑크) - {count_special}개<br>
</div>
"""
m.get_root().html.add_child(folium.Element(legend_html))
folium.LayerControl(collapsed=False).add_to(m)
filename = "PACS.html"

custom_js_css = r"""
<style>
#showLatLngBtn {
    position: fixed;
    top: 20px;
    left: 50px;
    z-index: 9999;
    background: #1976d2;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 8px 16px;
    font-size: 14px;
    cursor: pointer;
    box-shadow: 1px 2px 8px #888;
}
#toggleBtns {
    position: fixed;
    top: 20px;
    left: 210px;
    z-index: 9999;
    display: flex;
    gap: 8px;
}
#hideAllBtn, #showAllBtn {
    background: #e53935;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 8px 16px;
    font-size: 14px;
    cursor: pointer;
    box-shadow: 1px 2px 8px #888;
}
#showAllBtn {
    background: #1976d2;
    margin-left: 0;
}
#searchBox {
    position: fixed;
    top: 70px;
    left: 50px;
    z-index: 9999;
    background: white;
    border: 1px solid #aaa;
    border-radius: 5px;
    padding: 8px 12px;
    width: 300px;
    box-shadow: 1px 2px 8px #888;
}
#imgOverlay {
    display: none;
    position: fixed;
    z-index: 10000;
    left: 0; top: 0; width: 100vw; height: 100vh;
    background: rgba(0,0,0,0.7);
    justify-content: center; align-items: center;
}
#imgOverlay img {
    max-width: 90vw; max-height: 80vh; border: 5px solid #fff; border-radius: 8px;
}
#imgOverlayClose {
    position: absolute; top: 30px; right: 40px; color: #fff; font-size: 2em; cursor: pointer;
}
</style>
<button id="showLatLngBtn" class="search-btn">위/경도 표시</button>
<span id="toggleBtns">
    <button id="hideAllBtn">모두 감추기</button>
    <button id="showAllBtn" style="display:none;">모두 보이기</button>
</span>
<div id="searchBox">
    <input id="searchInput" type="text" placeholder="설치장소, 관리부서, 관리번호 검색">
    <button id="searchBtn" class="search-btn">검색</button>
    <button id="resetBtn" class="search-btn" style="margin-left:6px;">필터초기화</button>
</div>
<div id="imgOverlay" onclick="this.style.display='none'">
    <span id="imgOverlayClose" onclick="document.getElementById('imgOverlay').style.display='none';event.stopPropagation();">&times;</span>
    <img id="imgOverlayImg" src="">
</div>
<script>
var latlngPopupActive = false;
var latlngPopup;
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(function() {
        for (var key in window) {
            if (key.startsWith("map_") && window[key] instanceof L.Map) {
                window.map = window[key];
            }
        }
        if (window.map) {
            window.map.on('click', function(e) {
                if (!latlngPopupActive) return;
                if (latlngPopup) window.map.closePopup(latlngPopup);
                latlngPopup = L.popup()
                    .setLatLng(e.latlng)
                    .setContent("위도: " + e.latlng.lat.toFixed(6) + "<br>경도: " + e.latlng.lng.toFixed(6))
                    .openOn(window.map);
            });
        }
        var showLatLngBtn = document.getElementById('showLatLngBtn');
        if (showLatLngBtn) {
            showLatLngBtn.onclick = function() {
                latlngPopupActive = !latlngPopupActive;
                this.innerText = latlngPopupActive ? "위/경도 끄기" : "위/경도 표시";
                if (!latlngPopupActive && latlngPopup && window.map) {
                    window.map.closePopup(latlngPopup);
                }
            };
        }
        var allMarkers = [];
        if (window.map) {
            window.map.eachLayer(function(layer) {
                if (layer instanceof L.Marker && layer._popup) {
                    allMarkers.push(layer);
                }
            });
        }
        var searchBtn = document.getElementById('searchBtn');
        var resetBtn = document.getElementById('resetBtn');
        var searchInput = document.getElementById('searchInput');
        function filterMarkers() {
            var q = searchInput.value.trim();
            var isNumber = /^\d+$/.test(q);
            allMarkers.forEach(function(marker) {
                var html = marker._popup.getContent();
                if (typeof html !== "string" && html && html.innerHTML) {
                    html = html.innerHTML;
                }
                html = String(html).toLowerCase();
                var show = false;
                if (!q) {
                    show = true;
                } else if (isNumber) {
                    var matches = html.match(/<td[^>]*>관리번호<\/td>\s*<td[^>]*>([^<]+)<\/td>/g);
                    if (matches) {
                        for (var i = 0; i < matches.length; i++) {
                            var val = matches[i].replace(/.*<td[^>]*>관리번호<\/td>\s*<td[^>]*>([^<]+)<\/td>.*/, "$1").trim();
                            if (val === q || val.startsWith(q + "-")) {
                                show = true;
                                break;
                            }
                        }
                    }
                } else {
                    show = html.indexOf(q.toLowerCase()) !== -1;
                }
                if (show) {
                    if (!window.map.hasLayer(marker)) marker.addTo(window.map);
                } else {
                    if (window.map.hasLayer(marker)) window.map.removeLayer(marker);
                }
            });
        }
        if (searchBtn) searchBtn.onclick = filterMarkers;
        if (searchInput) searchInput.onkeydown = function(e) {
            if (e.key === "Enter") filterMarkers();
        };
        if (resetBtn) resetBtn.onclick = function() {
            searchInput.value = "";
            allMarkers.forEach(function(marker) {
                if (!window.map.hasLayer(marker)) marker.addTo(window.map);
            });
        };
    }, 300);
    var hideBtn = document.getElementById('hideAllBtn');
    var showBtn = document.getElementById('showAllBtn');
    hideBtn.onclick = function() {
        document.getElementById('showLatLngBtn').style.display = 'none';
        document.getElementById('searchBox').style.display = 'none';
        var legend = document.getElementById('legend');
        if (legend) legend.style.display = 'none';
        var layerControls = document.getElementsByClassName('leaflet-control-layers');
        for (var i = 0; i < layerControls.length; i++) {
            layerControls[i].style.display = 'none';
        }
        hideBtn.style.display = 'none';
        showBtn.style.display = '';
    };
    showBtn.onclick = function() {
        document.getElementById('showLatLngBtn').style.display = '';
        document.getElementById('searchBox').style.display = '';
        var legend = document.getElementById('legend');
        if (legend) legend.style.display = '';
        var layerControls = document.getElementsByClassName('leaflet-control-layers');
        for (var i = 0; i < layerControls.length; i++) {
            layerControls[i].style.display = '';
        }
        hideBtn.style.display = '';
        showBtn.style.display = 'none';
    };
});
document.addEventListener('click', function(e) {
    if (e.target.tagName === 'IMG' && e.target.classList.contains('popup-img')) {
        var overlay = document.getElementById('imgOverlay');
        var overlayImg = document.getElementById('imgOverlayImg');
        overlayImg.src = e.target.src;
        overlay.style.display = 'flex';
        e.stopPropagation();
    }
});
</script>
"""

m.get_root().html.add_child(folium.Element(custom_js_css))
m.save(filename)

print("\nHTML 파일 저장 완료:", filename)

# --- GitHub 업로드 자동화 ---
from github import Github
import os

GITHUB_TOKEN = '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'
REPO_NAME = 'come6433/q8r2x7v1p0'
SHARE_URL = f'come6433.github.io/q8r2x7v1p0/{filename}'
LOCAL_HTML = filename
REMOTE_HTML = filename
IMAGES_DIR = 'images'

g = Github(GITHUB_TOKEN)
repo = g.get_repo(REPO_NAME)

def upload_or_update(path_local, path_remote):
    print(f"업로드 시도: {path_local} -> {path_remote}")
    with open(path_local, "rb") as f:
        content = f.read()
    try:
        contents = repo.get_contents(path_remote)
        repo.update_file(path_remote, "자동 업로드", content, contents.sha)
        print(f"업데이트: {path_remote}")
    except Exception as e:
        print(f"신규 생성 시도: {path_remote} (사유: {e})")
        repo.create_file(path_remote, "자동 업로드", content)
        print(f"생성: {path_remote}")

print("\nHTML 파일 업로드 시작")
upload_or_update(LOCAL_HTML, REMOTE_HTML)
print("\n서버 업로드 완료!")
print(f"공유주소: {SHARE_URL}")
print("=" * 40)
input("아무 키나 누르면 종료합니다.")