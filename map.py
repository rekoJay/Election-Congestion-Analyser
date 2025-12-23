import pandas as pd
import folium
from folium.features import DivIcon
import requests
import json
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# [필수] 카카오 REST API 키 입력
KAKAO_API_KEY = "55204575d02bdd0ba57fcb1db49b5cb2"
# ==========================================

# --- [기존 함수들: 좌표변환, 경계선 그리기] ---
def kakao_geocode(address):
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        response = requests.get(url, headers=headers, params={"query": address.split('(')[0].strip()})
        if response.status_code == 200 and response.json()['documents']:
            return float(response.json()['documents'][0]['y']), float(response.json()['documents'][0]['x'])
    except:
        pass
    return None, None

def add_yuseong_boundary(m):
    url = 'https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_municipalities_geo_simple.json'
    try:
        resp = requests.get(url, verify=False)
        data = json.loads(resp.text)
        yuseong = next((f for f in data['features'] if f['properties']['name'] == '유성구'), None)
        if yuseong:
            folium.GeoJson(yuseong, name='유성구', style_function=lambda x: {'fillColor':'#ffff00', 'color':'#ff0000', 'weight':4, 'fillOpacity':0.2}).add_to(m)
    except:
        pass

# --- [새로 추가된 함수: HTML을 이미지로 변환] ---
def save_map_as_image(html_file, image_file):
    print(f"이미지 변환 중... ({html_file} -> {image_file})")
    
    # 1. 크롬 브라우저 설정 (화면 없이 실행, 고해상도 설정)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # 창을 띄우지 않음
    options.add_argument('--window-size=1920,1080') # 해상도 설정 (필요하면 더 키우세요)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # 2. HTML 파일 열기 (절대 경로 필요)
    abs_path = os.path.abspath(html_file)
    driver.get(f"file:///{abs_path}")
    
    # 3. 지도가 다 로딩될 때까지 잠시 대기 (중요!)
    time.sleep(3) 
    
    # 4. 스크린샷 저장
    driver.save_screenshot(image_file)
    driver.quit()
    print(f"✅ 이미지 저장 완료! '{image_file}' 파일을 내부망으로 옮기세요.")

def create_complete_map(file_path):
    # ... (파일 읽기 및 데이터 처리 로직은 동일) ...
    # 편의를 위해 파일 읽기 부분 요약
    try: df = pd.read_excel(file_path, header=4)
    except: 
        try: df = pd.read_csv(file_path, header=4, encoding='cp949')
        except: return

    df = df.dropna(subset=['소재지'])
    coords_list = [kakao_geocode(addr) for addr in df['소재지']]
    df['coords'] = coords_list
    df_clean = df[df['coords'].apply(lambda x: x[0] is not None)]

    if df_clean.empty: return

    # [수정] 모든 투표소 좌표의 평균(무게중심)을 구해서 지도의 중심으로 설정
    mean_lat = df_clean['coords'].apply(lambda x: x[0]).mean()
    mean_lon = df_clean['coords'].apply(lambda x: x[1]).mean()

    # 지도 생성 (평균 좌표를 중심으로 설정)
    m = folium.Map(location=[mean_lat, mean_lon], zoom_start=12)
    add_yuseong_boundary(m)

    for _, row in df_clean.iterrows():
        lat, lon = row['coords']
        # 별표 마커
        folium.Marker([lat, lon], icon=folium.Icon(color='red', icon='star')).add_to(m)
        # 텍스트 라벨
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(
                icon_size=(150, 36),
                icon_anchor=(75, -5),
                html=f'<div style="font-size: 10pt; font-weight: bold; text-align: center; text-shadow: -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff, 1px 1px 0 #fff;">{row["사전투표소명"]}</div>'
            )
        ).add_to(m)

    # HTML 저장
    html_file = "kakao_voting_map_final.html"
    m.save(html_file)
    print("HTML 생성 완료.")
    
    # ★★★ 여기서 이미지로 변환 시작! ★★★
    image_file = "voting_map_final.png"
    save_map_as_image(html_file, image_file)

# --- 실행 ---
input_file = "data.csv"
if os.path.exists(input_file):
    create_complete_map(input_file)