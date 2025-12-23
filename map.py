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
import glob  # <--- [추가] 파일 목록을 조회하는 라이브러리

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

def add_dynamic_boundaries(m, df):
    # 1. 엑셀 파일의 주소에서 '시/군/구' 이름 추출 (예: '대전광역시 유성구 xxx' -> '유성구')
    # 주소의 두 번째 단어를 행정구역 명으로 가정합니다.
    target_districts = set()
    for addr in df['소재지']:
        parts = addr.split()
        if len(parts) >= 2:
            target_districts.add(parts[1]) # 두 번째 단어 저장 (중복 제거됨)
    
    print(f"감지된 행정구역: {target_districts}")

    # 2. 대한민국 전체 행정구역 데이터 가져오기
    url = 'https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_municipalities_geo_simple.json'
    try:
        resp = requests.get(url, verify=False)
        geo_data = json.loads(resp.text)
        
        # 3. 감지된 지역과 이름이 일치하는 경계선만 필터링
        filtered_features = []
        for feature in geo_data['features']:
            # GeoJSON의 지역명과 엑셀에서 뽑은 지역명이 같은지 확인
            if feature['properties']['name'] in target_districts:
                filtered_features.append(feature)
        
        # 4. 지도에 표시 (지역이 있으면 그림)
        if filtered_features:
            folium.GeoJson(
                {"type": "FeatureCollection", "features": filtered_features},
                name='행정구역',
                style_function=lambda x: {
                    'fillColor': '#ffff00', # 노란색 채우기
                    'color': '#ff0000',     # 빨간색 테두리
                    'weight': 3,
                    'fillOpacity': 0.1
                }
            ).add_to(m)
            print("행정구역 경계선 그리기 완료.")
        else:
            print("일치하는 행정구역 경계선을 찾지 못했습니다. (주소 형식 확인 필요)")
            
    except Exception as e:
        print(f"행정구역 데이터 로딩 중 오류 발생: {e}")

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

def create_map_from_combined_data():
    # 1. 현재 폴더의 모든 csv, xlsx 파일 찾기
    all_files = glob.glob("*.csv") + glob.glob("*.xlsx")
    print(f"발견된 파일 목록: {all_files}")

    if not all_files:
        print("오류: 처리할 데이터 파일(.csv, .xlsx)이 없습니다.")
        return

    # 2. 모든 파일을 읽어서 하나로 합치기
    df_list = []
    for filename in all_files:
        try:
            # 엑셀인 경우
            if filename.endswith('.xlsx'):
                temp_df = pd.read_excel(filename, header=4) # header 위치 주의
            # CSV인 경우
            else:
                try: temp_df = pd.read_csv(filename, header=4, encoding='cp949')
                except: temp_df = pd.read_csv(filename, header=4, encoding='utf-8')
            
            # '소재지' 컬럼이 있는지 확인 (데이터 유효성 검사)
            if '소재지' in temp_df.columns:
                df_list.append(temp_df)
                print(f"  - {filename}: 데이터 로드 성공 ({len(temp_df)}건)")
            else:
                print(f"  - {filename}: '소재지' 컬럼이 없어 건너뜁니다.")
        except Exception as e:
            print(f"  - {filename} 읽기 실패: {e}")

    if not df_list:
        print("유효한 데이터가 하나도 없습니다.")
        return

    # 데이터 합치기 (Concat)
    df = pd.concat(df_list, ignore_index=True)
    print(f"총 {len(df)}건의 데이터를 통합했습니다.")

    # --- 여기서부터는 기존 로직과 동일 ---
    df = df.dropna(subset=['소재지'])
    
    print("주소 변환(Geocoding) 시작...")
    coords_list = []
    for addr in df['소재지']:
        coords_list.append(kakao_geocode(addr))
        
    df['coords'] = coords_list
    df_clean = df[df['coords'].apply(lambda x: x[0] is not None)]

    if df_clean.empty:
        print("좌표 변환된 데이터가 없습니다.")
        return

    # 중심점 계산 및 지도 생성
    mean_lat = df_clean['coords'].apply(lambda x: x[0]).mean()
    mean_lon = df_clean['coords'].apply(lambda x: x[1]).mean()

    m = folium.Map(location=[mean_lat, mean_lon], zoom_start=11) # 여러 지역일 수 있으니 줌을 약간 뺌
    
    # 동적 경계선 그리기
    add_dynamic_boundaries(m, df_clean)
    
    # 마커 찍기
    for _, row in df_clean.iterrows():
        lat, lon = row['coords']
        folium.Marker([lat, lon], icon=folium.Icon(color='red', icon='star')).add_to(m)
        
        # 텍스트 라벨 (사전투표소명 컬럼이 있다고 가정)
        name = row.get('사전투표소명', '투표소') # 컬럼 없으면 '투표소'로 표시
        folium.map.Marker(
            [lat, lon],
            icon=DivIcon(
                icon_size=(150, 36),
                icon_anchor=(75, -5),
                html=f'<div style="font-size: 10pt; font-weight: bold; text-align: center; text-shadow: -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff, 1px 1px 0 #fff;">{name}</div>'
            )
        ).add_to(m)

    # 저장
    html_file = "voting_map_combined.html"
    m.save(html_file)
    print(f"HTML 지도 생성 완료: {html_file}")
    
    image_file = "voting_map_combined.png"
    save_map_as_image(html_file, image_file)

# --- 실행 ---
create_map_from_combined_data()
