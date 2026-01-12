import sys
import pandas as pd
import re
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
import platform
import numpy as np 
from mpl_toolkits.axes_grid1 import make_axes_locatable
from matplotlib.colors import ListedColormap 

class ElectionAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사전투표장비 배분 최적화 시스템")
        self.root.geometry("680x920") 
        self.root.resizable(False, True) 
        self.vote_files = []
        self.cached_data = {} 
        self.equipment_file = None
        self.file_past_elect = None   
        self.file_recent_elect = None 
        
        self.region_name = "" # [추가] 지역명 저장 변수 (예: 서울 성동구)

        self.last_reserve_count = 5
        self.station_data = {} 
        
        self.create_widgets()

    # [수정] 통합 조정률 텍스트 생성 헬퍼
    def _get_merged_rate_text(self, r_intra, r_extra):
        def _fmt(val):
            # [변경] 화살표 제거하고 +, - 기호로 통일 (가독성 향상)
            if val > 0: return f"+ {val}%"        
            elif val < 0: return f"- {abs(val)}%" 
            else: return "-"
            
        if r_intra == r_extra:
            return _fmt(r_intra)
        else:
            return f"관내:{_fmt(r_intra)} / 관외:{_fmt(r_extra)}"
            
    def create_widgets(self):
        # 메인 스크롤 프레임 설정
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        # [수정] 캔버스에 프레임을 그릴 때 ID를 변수(frame_id)에 저장합니다.
        frame_id = main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # [핵심 수정] 창 크기가 변할 때(Configure), 내용물(frame_id)의 너비를 창 너비(e.width)에 강제로 맞춥니다.
        def _on_canvas_configure(e):
            main_canvas.itemconfig(frame_id, width=e.width)
        
        # [기존 기능 유지] 내용물이 변할 때 스크롤 범위 갱신
        def _on_frame_configure(e):
            main_canvas.configure(scrollregion=main_canvas.bbox("all"))

        main_canvas.bind("<Configure>", _on_canvas_configure)
        scrollable_frame.bind("<Configure>", _on_frame_configure)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        content_frame = ttk.Frame(scrollable_frame, padding="20")
        content_frame.pack(fill="both", expand=True)

        # 1. 기초 데이터 로드
        frame_data = ttk.LabelFrame(content_frame, text=" 1. 기초 데이터 로드 ", padding="10")
        frame_data.pack(fill="x", pady=(0, 10))
        
        btn_files = ttk.Button(frame_data, text="📂 투표 데이터 파일 업로드", command=self.select_vote_files)
        btn_files.pack(fill="x", ipady=3)
        self.lbl_file_count = ttk.Label(frame_data, text="파일 없음", foreground="gray", font=("맑은 고딕", 8))
        self.lbl_file_count.pack(pady=(2, 5))

        btn_equip = ttk.Button(frame_data, text="📂 장비 현황 파일 업로드", command=self.select_equip_file)
        btn_equip.pack(fill="x", ipady=3)
        self.lbl_equip_status = ttk.Label(frame_data, text="파일 미선택 (기본값: 1대 적용)", foreground="gray", font=("맑은 고딕", 8))
        self.lbl_equip_status.pack(pady=(2, 5))

        # [수정] 선거인수 파일 버튼 분리 (과거/최근)
        frame_elect = ttk.Frame(frame_data)
        frame_elect.pack(fill="x", pady=(5, 0))
        
        # 과거 파일
        btn_past = ttk.Button(frame_elect, text="📂 ① 과거 선거인수", command=self.select_past_file)
        btn_past.pack(side="left", fill="x", expand=True, padx=(0, 2))
        
        # 최근 파일
        btn_recent = ttk.Button(frame_elect, text="📂 ② 최근 선거인수", command=self.select_recent_file)
        btn_recent.pack(side="right", fill="x", expand=True, padx=(2, 0))
        
        self.lbl_elect_status = ttk.Label(frame_data, text="파일 미선택 (변동률 미적용)", foreground="gray", font=("맑은 고딕", 8))
        self.lbl_elect_status.pack(pady=(2, 0))
        
        # 2. 시뮬레이션 설정
        frame_sim = ttk.LabelFrame(content_frame, text=" 2. 시뮬레이션 설정 (데이터 튜닝) ", padding="10")
        frame_sim.pack(fill="x", pady=(0, 10))
        
        # 2-1. 투표율 조정 슬라이더
        frame_rate = ttk.Frame(frame_sim)
        frame_rate.pack(fill="x", pady=(0, 10))
        ttk.Label(frame_rate, text="📉 전체 투표자 증가율: ").pack(side="left")
        
        self.var_rate = tk.DoubleVar(value=0.0)
        self.lbl_rate = ttk.Label(frame_rate, text="0% (변동 없음)", foreground="blue", font=("맑은 고딕", 9, "bold"))
        
        scale = ttk.Scale(frame_sim, from_=-30, to=30, variable=self.var_rate, command=self.on_slider_change)
        scale.pack(fill="x", padx=10, pady=(0,10))
        self.lbl_rate.pack(side="right")

        # 2-2. 장비 및 개별 증가율 리스트
        # [추가] 오토 밸런싱 버튼 영역
        frame_balance = ttk.Frame(frame_sim)
        frame_balance.pack(fill="x", pady=(0, 5))

        ttk.Label(frame_balance, text="📋 투표소별 설정 (수정: 더블클릭)", font=("맑은 고딕", 9, "bold")).pack(side="left")
        btn_balance = ttk.Button(frame_balance, text="⚖️ 장비 자동 배분", command=self.open_balance_popup)
        btn_balance.pack(side="right")

        tree_frame = ttk.Frame(frame_sim)
        tree_frame.pack(fill="both", expand=True, pady=5)
        
        # [수정] 조정률 컬럼 통합 (rate_merged)
        columns = ("station", "elect_diff", "intra", "extra", "rate_merged")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)
        
        self.tree.heading("station", text="투표소명")
        self.tree.heading("elect_diff", text="선거인수")
        self.tree.heading("intra", text="관내장비")
        self.tree.heading("extra", text="관외장비")
        # [변경] 헤더 텍스트 수정
        self.tree.heading("rate_merged", text="조정률(관내/외)") 
        
        self.tree.column("station", width=150)
        self.tree.column("elect_diff", width=90, anchor="center")
        self.tree.column("intra", width=60, anchor="center")
        self.tree.column("extra", width=60, anchor="center")
        # [변경] 글자가 길어지므로 너비를 120 -> 150으로 확대
        self.tree.column("rate_merged", width=150, anchor="center")
        
        scrollbar_tree = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_tree.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")
        
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # 3. 보기 옵션
        frame_option = ttk.LabelFrame(content_frame, text=" 3. 보기 옵션 ", padding="10")
        frame_option.pack(fill="x", pady=(0, 10))
        
        self.var_day1 = tk.BooleanVar(value=True)
        self.var_day2 = tk.BooleanVar(value=True)
        self.var_intra = tk.BooleanVar(value=True)
        self.var_extra = tk.BooleanVar(value=True)
        
        chk_frame = ttk.Frame(frame_option)
        chk_frame.pack(fill="x")
        
        # [수정] '전체' 옵션 변수 추가
        self.var_day_all = tk.BooleanVar(value=True) 

        ttk.Label(chk_frame, text="기간: ").pack(side="left")
        ttk.Checkbutton(chk_frame, text="1일차", variable=self.var_day1).pack(side="left", padx=5)
        ttk.Checkbutton(chk_frame, text="2일차", variable=self.var_day2).pack(side="left", padx=5)
        # [추가] 전체(평균) 체크박스 생성
        ttk.Checkbutton(chk_frame, text="전체(평균)", variable=self.var_day_all).pack(side="left", padx=5)
        
        ttk.Separator(chk_frame, orient="vertical").pack(side="left", fill="y", padx=15)
        ttk.Label(chk_frame, text="구분: ").pack(side="left")
        ttk.Checkbutton(chk_frame, text="관내", variable=self.var_intra).pack(side="left", padx=5)
        ttk.Checkbutton(chk_frame, text="관외", variable=self.var_extra).pack(side="left", padx=5)
        
        # 4. 실행 버튼
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=10)
        btn_run = ttk.Button(content_frame, text="🚀 시뮬레이션 / 분석 실행", command=self.run_simulation)
        btn_run.pack(fill="x", ipady=12)
        
        # 5. 로그창
        log_frame = ttk.LabelFrame(content_frame, text=" 시스템 로그 ", padding="10")
        log_frame.pack(fill="x", pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=6, state='disabled', bg="#F0F0F0", font=("맑은 고딕", 9))
        self.log_text.pack(fill="both", expand=True)

    def select_past_file(self):
        file = filedialog.askopenfilename(title="과거 선거인수 파일 (A열:동명, B열:인수)", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_past_elect = file
            self._update_elect_status()

    def select_recent_file(self):
        file = filedialog.askopenfilename(title="최근 선거인수 파일 (A열:동명, B열:인수)", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.file_recent_elect = file
            self._update_elect_status()

    def _update_elect_status(self):
        # 두 파일 상태 확인 및 스캔 트리거
        p = "✅" if self.file_past_elect else "❌"
        r = "✅" if self.file_recent_elect else "❌"
        
        self.lbl_elect_status.config(text=f"과거: {p} / 최근: {r}", foreground="blue" if (p=="✅" and r=="✅") else "red")
        
        if self.file_past_elect and self.file_recent_elect:
            self.log("두 선거인수 파일 준비됨. 변동률 계산 시작...")
            self.scan_stations()

    def log(self, msg):
        def _update():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.root.after(0, _update)

    def on_slider_change(self, val):
        rate = int(float(val))
        text = f"{rate}% "
        if rate > 0:
            text += "(증가)"
            color = "red"
        elif rate < 0:
            text += "(감소)"
            color = "blue"
        else:
            text += "(동일)"
            color = "black"
        self.lbl_rate.config(text=text, foreground=color)
        
        for item_id in self.tree.get_children():
            st_name = self.tree.item(item_id)['values'][0]
            if st_name in self.station_data:
                # [수정] 데이터 업데이트 (관내/관외 각각 저장)
                self.station_data[st_name]['rate_intra'] = rate
                self.station_data[st_name]['rate_extra'] = rate
                
                # 화면 갱신용 데이터 준비
                elect_disp = self.tree.item(item_id)['values'][1] # 선거인수 컬럼 유지
                curr_intra = self.station_data[st_name]['intra']
                curr_extra = self.station_data[st_name]['extra']
                org_intra = self.station_data[st_name]['org_intra']
                org_extra = self.station_data[st_name]['org_extra']
                
                disp_intra = f"{org_intra} → {curr_intra}" if curr_intra != org_intra else str(curr_intra)
                disp_extra = f"{org_extra} → {curr_extra}" if curr_extra != org_extra else str(curr_extra)

                # [수정] 통합 텍스트 적용
                rate_txt = self._get_merged_rate_text(rate, rate)

                # 컬럼 5개 반영
                self.tree.item(item_id, values=(st_name, elect_disp, disp_intra, disp_extra, rate_txt))

    def select_vote_files(self):
        files = filedialog.askopenfilenames(title="투표 데이터 선택", filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if files:
            self.vote_files = files
            self.cached_data = {} # [최적화] 새 파일 선택 시 캐시 초기화
            self.lbl_file_count.config(text=f"✅ {len(files)}개 파일 로드됨", foreground="blue")
            self.log(f"{len(files)}개 파일 선택됨. 데이터 로드 및 스캔 시작...")
            self.scan_stations()

    def select_equip_file(self):
        file = filedialog.askopenfilename(title="장비현황 파일 선택", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.equipment_file = file
            self.lbl_equip_status.config(text=f"✅ {os.path.basename(file)}", foreground="blue")
            self.log(f"장비 파일 로드됨. 목록 업데이트 중...")
            self.scan_stations() 

    def _ensure_data_loaded(self):
        # [최적화] 파일이 캐시에 없으면 읽어서 저장
        for file in self.vote_files:
            if file in self.cached_data:
                continue
                
            try:
                day, time, header_row = self.get_file_info_header(file)
                if day is None: continue

                if file.endswith('.csv'):
                    try: df = pd.read_csv(file, header=header_row, encoding='cp949')
                    except: df = pd.read_csv(file, header=header_row, encoding='utf-8')
                else:
                    df = pd.read_excel(file, header=header_row)

                if '사전투표소명' in df.columns:
                    df = df.dropna(subset=['사전투표소명'])
                    # 공통 전처리: 합계/소계 제거
                    if '읍면동명' in df.columns:
                        temp_col = df['읍면동명'].astype(str).str.replace(' ', '')
                        mask = temp_col.str.contains('합계|소계|총계|누계', na=False)
                        df = df[~mask].copy()
                    
                    # 숫자 변환 미리 수행
                    for col in ['관내사전투표자수', '관외사전투표자수']:
                        if col in df.columns and df[col].dtype == 'object':
                            df[col] = df[col].astype(str).str.replace(',', '').str.strip()
                            df[col] = pd.to_numeric(df[col], errors='coerce')

                    self.cached_data[file] = (df, day, time)
            except Exception as e:
                self.log(f"파일 로드 실패({os.path.basename(file)}): {e}")

    def scan_stations(self):
        if not self.vote_files:
            return

        self._ensure_data_loaded() # [최적화] 데이터 로드 보장

        station_list = []  
        seen = set()       
        
        # 캐시된 데이터에서 투표소 추출
        for file in self.vote_files:
            if file not in self.cached_data: continue
            
            df, _, _ = self.cached_data[file]
            stations = df['사전투표소명'].unique()
            
            for s in stations:
                s_str = str(s).strip()
                if s_str and s_str != 'nan':
                    if s_str not in seen:
                        seen.add(s_str)
                        station_list.append(s_str)

        # [수정] 사용자 지정 서식(C열:이름, F열:관내, G열:관외) 맞춤 로직
        equip_map = {}
        if self.equipment_file:
            try:
                # 1. 파일 읽기 (헤더 없이 읽음)
                if self.equipment_file.endswith('.csv'):
                    try: df_raw = pd.read_csv(self.equipment_file, header=None, encoding='cp949')
                    except: df_raw = pd.read_csv(self.equipment_file, header=None, encoding='utf-8')
                else:
                    df_raw = pd.read_excel(self.equipment_file, header=None)

                # [추가] 3행(Index 2)에서 지역 이름 추출 로직
                try:
                    # 3행의 데이터 중 '값'이 있는 첫 번째 칸을 지역 이름으로 간주
                    row_3_vals = df_raw.iloc[2].astype(str).values
                    for v in row_3_vals:
                        v_clean = v.strip().replace('nan', '')
                        if v_clean:
                            self.region_name = v_clean
                            break
                except:
                    self.region_name = ""

                # 2. 데이터 시작 행 찾기 (C열에 '읍면동'이나 '투표소'가 나오는 줄)
                start_row_idx = 0
                for idx, row in df_raw.head(15).iterrows():
                    # 엑셀 C열은 인덱스 2
                    c_col_val = str(row[2]).replace(" ", "")
                    if "읍면동" in c_col_val or "투표소" in c_col_val:
                        start_row_idx = idx + 1 # 헤더 다음 줄부터 데이터
                        break
                
                # 3. 데이터 추출 (C열=2, F열=5, G열=6)
                for idx in range(start_row_idx, len(df_raw)):
                    row = df_raw.iloc[idx]
                    
                    # C열: 투표소명
                    st_name = str(row[2]).strip()
                    if st_name == 'nan' or not st_name: continue
                    if '합계' in st_name or '소계' in st_name: continue

                    # 숫자 정제 함수
                    def parse_count(val):
                        try:
                            txt = str(val).split('(')[0].replace(',', '').replace('대', '').strip()
                            return int(float(txt))
                        except:
                            return 1 

                    # F열(5): 관내 장비수, G열(6): 관외 장비수
                    intra_count = parse_count(row[5])
                    extra_count = parse_count(row[6])

                    equip_map[st_name] = {'intra': intra_count, 'extra': extra_count}

                self.log(f"장비 파일 로드 완료: {len(equip_map)}개소 (C,F,G열 기준)")

            except Exception as e:
                self.log(f"장비 파일 읽기 오류: {e}")
        
        # [수정] 선거인수 변동률 계산 로직 (두 파일 병합)
        electorate_rates = {}
        electorate_diffs = {}
        
        if self.file_past_elect and self.file_recent_elect:
            try:
                # 데이터 추출 내부 함수 (A열: 동명, B열: 숫자 라고 가정)
                # [수정] 데이터 추출 내부 함수 (A열: 읍면동명, D열: 선거인수)
                def load_elect_data(path):
                    data_map = {}
                    try:
                        # 1. 헤더 없이 읽어서 데이터 위치 찾기
                        if path.endswith('.csv'):
                            try: df = pd.read_csv(path, header=None, encoding='cp949')
                            except: df = pd.read_csv(path, header=None, encoding='utf-8')
                        else:
                            df = pd.read_excel(path, header=None)
                        
                        start_row = 0
                        # 2. '읍면동명'이 있는 행 찾기 (헤더 위치 검색)
                        for idx, row in df.head(15).iterrows():
                            # A열(0번 인덱스) 확인
                            if "읍면동명" in str(row[0]):
                                start_row = idx + 1
                                break
                        
                        # 3. 데이터 추출
                        for idx in range(start_row, len(df)):
                            row = df.iloc[idx]
                            
                            # A열(0): 동 이름
                            k_raw = str(row[0])
                            if pd.isna(row[0]) or k_raw.strip() == '' or k_raw == 'nan': continue
                            
                            k = k_raw.strip().replace(" ", "")
                            if '합계' in k or '소계' in k: continue # 합계 행 제외
                            
                            # D열(3): 선거인수 (예: "21,412\n(25, 12)")
                            v_raw = str(row[3])
                            
                            try:
                                # 줄바꿈(\n)이나 괄호(() 앞부분의 숫자만 가져오기
                                v_str = v_raw.split('\n')[0].split('(')[0]
                                v_str = v_str.replace(',', '').strip()
                                v = float(v_str)
                                
                                if v > 0:
                                    data_map[k] = v
                            except:
                                continue
                                
                    except Exception as e:
                        print(f"선거인수 파일 읽기 실패({path}): {e}")
                        
                    return data_map

                past_map = load_elect_data(self.file_past_elect)
                recent_map = load_elect_data(self.file_recent_elect)
                
                # 두 맵을 비교하여 증감률 및 차이 계산
                count_matched = 0
                for dong_name, recent_val in recent_map.items():
                    if dong_name in past_map:
                        past_val = past_map[dong_name]
                        if past_val > 0:
                            # 1. 시뮬레이션용 비율 계산 (유지)
                            rate_val = ((recent_val - past_val) / past_val) * 100
                            electorate_rates[dong_name] = int(round(rate_val))
                            
                            # 2. [추가] 화면 표시용 인원 차이 계산
                            diff_val = int(recent_val - past_val)
                            electorate_diffs[dong_name] = diff_val
                            
                            count_matched += 1
                
                self.log(f"변동률 계산 완료: {count_matched}개 동 매칭됨")
                
            except Exception as e:
                self.log(f"선거인수 파일 처리 오류: {e}")

        for item in self.tree.get_children():
            self.tree.delete(item)
        
        sorted_stations = station_list
        self.station_data = {} 
        current_global_rate = int(self.var_rate.get())

        for st in sorted_stations:
            # 1. 장비 매칭
            matched_data = None
            if st in equip_map:
                matched_data = equip_map[st]
            else:
                for k, v in equip_map.items():
                    if str(k) in st or st in str(k): 
                        matched_data = v
                        break
            
            if matched_data:
                intra = matched_data['intra']
                extra = matched_data['extra']
            else:
                intra = 1
                extra = 1
            
            # 2. 선거인수 증감률 매칭 (별도 변수 elect_rate에 저장)
            elect_display = "-"
            elect_rate = 0 # 선거인수 변동률 기본값 0
            
            if electorate_rates:
                st_clean = st.replace(" ", "")
                for dong_name, e_rate in electorate_rates.items():
                    if dong_name in st_clean:
                        elect_rate = e_rate
                        
                        diff = electorate_diffs.get(dong_name, 0)
                        
                        # [변경] 화살표 대신 직관적인 +, - 기호 사용
                        if diff > 0:
                            elect_display = f"+ {diff:,}명" 
                        elif diff < 0:
                            elect_display = f"- {abs(diff):,}명"
                        else:
                            elect_display = "-" # 변동 없음
                        break
            
            # [수정] 데이터 저장: rate를 intra/extra로 분리
            self.station_data[st] = {
                'intra': intra, 'extra': extra, 
                'rate_intra': current_global_rate,  # 관내 사용자 설정
                'rate_extra': current_global_rate,  # 관외 사용자 설정
                'elect_rate': elect_rate,           # 선거인수 변동 (고정)
                'org_intra': intra, 'org_extra': extra
            }
            
            # [수정] tags 옵션 삭제 (원래대로 복구)
            rate_txt = self._get_merged_rate_text(current_global_rate, current_global_rate)

            # 맨 뒤에 있었던 tags=(row_tag,) 부분을 지우세요.
            self.tree.insert("", "end", iid=st, values=(st, elect_display, intra, extra, rate_txt))
            
        self.log(f"목록 갱신 완료: 총 {len(sorted_stations)}개 투표소")

    def on_tree_double_click(self, event):
        try:
            # 1. 클릭한 위치(행/열) 파악
            region = self.tree.identify("region", event.x, event.y)
            if region != "cell": return 
            
            item_id = self.tree.identify_row(event.y)
            column = self.tree.identify_column(event.x)
            
            if not item_id: return
            
            # 2. 안전하게 투표소명 가져오기
            # (IID가 아니라 실제 표에 적힌 첫 번째 값을 기준으로 함)
            item_values = self.tree.item(item_id)['values']
            if not item_values: return
            st_name = str(item_values[0])
            
            if st_name not in self.station_data:
                # 혹시나 해서 IID로 한 번 더 시도
                if item_id in self.station_data: st_name = item_id
                else: return

            # 3. 데이터 가져오기 (여기서 'rate'를 찾던 코드를 삭제하고 분리된 변수를 가져옵니다)
            data = self.station_data[st_name]
            curr_intra = data['intra']
            curr_extra = data['extra']
            org_intra = data['org_intra']
            org_extra = data['org_extra']
            
            # [수정] 'rate' 키는 이제 없으므로 rate_intra, rate_extra를 가져옴
            val_rate_intra = data['rate_intra']
            val_rate_extra = data['rate_extra']
            
            elect_disp = item_values[1] # 선거인수 표기는 그대로 유지

            # 화면 표시용 텍스트 생성 내부함수
            def get_display_text(val, org_val):
                return f"{org_val} → {val}" if val != org_val else str(val)

            # 4. 컬럼별 수정 로직
            if column == '#3': # 관내 장비
                new_intra = simpledialog.askinteger("관내 장비 수정", f"[{st_name}]\n관내 장비 수:", 
                                                  initialvalue=curr_intra, minvalue=1, maxvalue=50, parent=self.root)
                if new_intra is not None:
                    self.station_data[st_name]['intra'] = new_intra
                    disp_intra = get_display_text(new_intra, org_intra)
                    disp_extra = get_display_text(curr_extra, org_extra)
                    self.tree.item(item_id, values=(st_name, elect_disp, disp_intra, disp_extra, val_rate_intra, val_rate_extra))
                    self.log(f"{st_name} 관내 장비 변경: {new_intra}대")
                    
            elif column == '#4': # 관외 장비
                new_extra = simpledialog.askinteger("관외 장비 수정", f"[{st_name}]\n관외 장비 수:", 
                                                  initialvalue=curr_extra, minvalue=1, maxvalue=50, parent=self.root)
                if new_extra is not None:
                    self.station_data[st_name]['extra'] = new_extra
                    disp_intra = get_display_text(curr_intra, org_intra)
                    disp_extra = get_display_text(new_extra, org_extra)
                    self.tree.item(item_id, values=(st_name, elect_disp, disp_intra, disp_extra, val_rate_intra, val_rate_extra))
                    self.log(f"{st_name} 관외 장비 변경: {new_extra}대")
                    
            elif column == '#5': # 조정률(통합) 수정 -> 팝업 호출
                self._open_rate_input_dialog(st_name, item_id, elect_disp, curr_intra, curr_extra, org_intra, org_extra)
            
            else:
                messagebox.showinfo("알림", "수정 가능한 항목(장비 수, 조정률)을 더블 클릭해주세요.", parent=self.root)

        except Exception as e:
            # 에러 발생 시 로그에 남김
            print(f"더블 클릭 오류: {e}")
            import traceback
            traceback.print_exc()

    def get_file_info_header(self, file_path):
        try:
            if file_path.endswith('.csv'):
                try: df_meta = pd.read_csv(file_path, header=None, nrows=10, encoding='cp949')
                except: df_meta = pd.read_csv(file_path, header=None, nrows=10, encoding='utf-8')
            else:
                df_meta = pd.read_excel(file_path, header=None, nrows=10)
            
            day, time = None, None
            header_idx = 3

            for idx, row in df_meta.iterrows():
                row_str = " ".join(row.astype(str).values)
                if day is None:
                    match_day = re.search(r'\[(\d+)일차\]', row_str)
                    match_time = re.search(r'\[(\d{1,2}):(\d{2})\]', row_str)
                    if match_day: day = int(match_day.group(1))
                    if match_time: time = int(match_time.group(1))
                if "사전투표소명" in row_str or "읍면동명" in row_str:
                    header_idx = idx
            return day, time, header_idx
        except:
            return None, None, 3

    def run_simulation(self):
        if not self.vote_files:
            messagebox.showwarning("주의", "투표 데이터 파일이 없습니다.")
            return

        # 1. 로딩 팝업창 생성
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("처리 중")
        self.loading_win.geometry("300x100")
        self.loading_win.resizable(False, False)
        # 팝업이 떠 있는 동안 메인 창 조작 금지 (모달)
        self.loading_win.grab_set() 
        
        # 화면 중앙 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 50
        self.loading_win.geometry(f"+{x}+{y}")

        lbl = ttk.Label(self.loading_win, text="데이터 분석 및 시각화 중입니다...\n잠시만 기다려 주세요.", justify="center")
        lbl.pack(pady=20)
        
        # 프로그레스바 (왔다갔다 하는 모드)
        pb = ttk.Progressbar(self.loading_win, mode='indeterminate')
        pb.pack(fill="x", padx=20, pady=(0, 20))
        pb.start(10)

        # 2. 별도 스레드에서 무거운 작업 실행
        # daemon=True로 설정하여 메인 프로그램 종료 시 같이 종료되게 함
        t = threading.Thread(target=self._execute_simulation, daemon=True)
        t.start()

    def _execute_simulation(self):
        try:
            # [수정] 백엔드 설정 및 경고 무시
            import matplotlib
            matplotlib.use('Agg')
            import warnings
            warnings.simplefilter(action='ignore', category=FutureWarning) # 경고 메시지 숨기기

            label = "통합 분석"
            self.log(f"시뮬레이션 시작: {label}")
            
            self._ensure_data_loaded() 
            
            all_data = []
            
            for file in self.vote_files:
                if file not in self.cached_data: continue
                
                try:
                    org_df, day, time = self.cached_data[file]
                    df = org_df.copy() 
                    
                    df['사전투표소명'] = df['사전투표소명'].astype(str).str.strip()
                    
                    user_rate_intra_map = {name: data.get('rate_intra', 0) for name, data in self.station_data.items()}
                    user_rate_extra_map = {name: data.get('rate_extra', 0) for name, data in self.station_data.items()}
                    
                    user_rates_intra = df['사전투표소명'].map(user_rate_intra_map).fillna(0)
                    user_rates_extra = df['사전투표소명'].map(user_rate_extra_map).fillna(0)
                    
                    elect_rate_map = {name: data.get('elect_rate', 0) for name, data in self.station_data.items()}
                    elect_rates = df['사전투표소명'].map(elect_rate_map).fillna(0)
                    
                    factor_intra = (1 + (elect_rates / 100.0)) * (1 + (user_rates_intra / 100.0))
                    factor_extra = (1 + (user_rates_extra / 100.0))
                    
                    df['관내사전투표자수'] = df['관내사전투표자수'] * factor_intra
                    df['관외사전투표자수'] = df['관외사전투표자수'] * factor_extra
                            
                    df['일차'] = day
                    df['시간대'] = time
                    all_data.append(df)
                except Exception as e:
                    self.log(f"데이터 처리 오류({os.path.basename(file)}): {e}")

            if not all_data:
                self.root.after(0, lambda: messagebox.showerror("오류", "유효한 데이터가 없습니다."))
                self.root.after(0, self.loading_win.destroy)
                return

            final_df = pd.concat(all_data, ignore_index=True)

            original_order = []
            seen = set()
            for temp_df in all_data:
                stats = temp_df['사전투표소명'].unique()
                for s in stats:
                    if s not in seen:
                        seen.add(s)
                        original_order.append(s)
            
            final_df['사전투표소명'] = pd.Categorical(
                final_df['사전투표소명'], categories=original_order, ordered=True
            )
            
            duplicates = final_df[final_df.duplicated(subset=['사전투표소명', '일차', '시간대'], keep=False)]
            if not duplicates.empty:
                final_df = final_df.drop_duplicates(subset=['사전투표소명', '일차', '시간대'])

            final_df = final_df.sort_values(by=['사전투표소명', '일차', '시간대'])
            
            # [수정] 경고 해결을 위해 observed=True 추가 (없어도 동작은 함)
            final_df['시간대별_관내투표자수'] = final_df.groupby(['사전투표소명', '일차'], observed=True)['관내사전투표자수'].diff()
            final_df['시간대별_관외투표자수'] = final_df.groupby(['사전투표소명', '일차'], observed=True)['관외사전투표자수'].diff()
            
            for (st, day), group in final_df.groupby(['사전투표소명', '일차'], observed=True):
                first_idx = group.index[0]
                final_df.loc[first_idx, '시간대별_관내투표자수'] = final_df.loc[first_idx, '관내사전투표자수']
                final_df.loc[first_idx, '시간대별_관외투표자수'] = final_df.loc[first_idx, '관외사전투표자수']

            def get_equip_info(row, type_):
                st = row['사전투표소명']
                if st in self.station_data:
                    return self.station_data[st][type_], self.station_data[st][f'org_{type_}']
                return 1, 1

            final_df[['관내장비수', '원본_관내장비수']] = final_df.apply(lambda x: pd.Series(get_equip_info(x, 'intra')), axis=1)
            final_df[['관외장비수', '원본_관외장비수']] = final_df.apply(lambda x: pd.Series(get_equip_info(x, 'extra')), axis=1)

            final_df['관내_혼잡도'] = final_df['시간대별_관내투표자수'] / final_df['관내장비수']
            final_df['관외_혼잡도'] = final_df['시간대별_관외투표자수'] / final_df['관외장비수']

            final_df = final_df.loc[:, ~final_df.columns.str.contains('^Unnamed')]
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            if getattr(sys, 'frozen', False):
                script_dir = os.path.dirname(os.path.abspath(sys.executable))
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))

            excel_name = f"시뮬레이션_결과_{timestamp}.xlsx"
            full_excel_path = os.path.join(script_dir, excel_name)

            final_df.to_excel(full_excel_path, index=False)
            self.log(f"엑셀 저장 완료: {full_excel_path}")
            
            self.log("그래프 생성 중...")
            
            png_name = f"시뮬레이션_{timestamp}.png"
            full_png_path = os.path.join(script_dir, png_name)

            self.visualize_results(final_df, timestamp, full_png_path, mode='screen')
            
            def _finish():
                if hasattr(self, 'loading_win'): self.loading_win.destroy()
                messagebox.showinfo("완료", f"분석 완료!\n\n파일이 저장되었습니다:\n{full_png_path}")
                if platform.system() == 'Windows':
                    try: os.startfile(full_png_path)
                    except: pass
            
            self.root.after(0, _finish)

        except Exception as e:
            # [핵심 수정] 에러 메시지를 문자열(err_msg)로 미리 저장해둡니다.
            # 이렇게 해야 나중에 _error 함수가 실행될 때 e가 사라져도 내용을 알 수 있습니다.
            err_msg = str(e)
            import traceback
            traceback.print_exc() # 콘솔에 에러 상세 내용 출력

            def _error():
                if hasattr(self, 'loading_win'): self.loading_win.destroy()
                self.log(f"치명적 오류: {err_msg}")
                messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n{err_msg}")
                
            self.root.after(0, _error)

    def visualize_results(self, df, timestamp, save_name, mode='screen'):
        # 1. 폰트 설정
        system_name = platform.system()
        font_family = 'Malgun Gothic' if system_name == 'Windows' else 'AppleGothic'
        plt.rc('font', family=font_family)
        plt.rc('axes', unicode_minus=False)

        # 2. 이름 정리
        df['short_name'] = df['사전투표소명'].astype(str).str.replace('사전투표소', '').str.strip()
        df['label_clean'] = df['short_name'] 
        
        # 3. 전체(평균) 데이터 생성
        numeric_cols = ['관내_혼잡도', '관외_혼잡도', '관내장비수', '원본_관내장비수', '관외장비수', '원본_관외장비수']
        try:
            df_mean = df.groupby(['사전투표소명', '시간대', 'short_name', 'label_clean'], observed=True)[numeric_cols].mean().reset_index()
            df_mean['일차'] = '전체'
            df = pd.concat([df, df_mean], ignore_index=True)
        except Exception as e:
            print(f"평균 데이터 생성 중 오류(무시됨): {e}")

        # 4. 시나리오 설정 (체크박스 값 반영)
        all_scenarios = [
            (1, '관내', 'label_clean', '관내_혼잡도', '관내장비수', '원본_관내장비수', self.var_day1.get() and self.var_intra.get()),
            (1, '관외', 'label_clean', '관외_혼잡도', '관외장비수', '원본_관외장비수', self.var_day1.get() and self.var_extra.get()),
            (2, '관내', 'label_clean', '관내_혼잡도', '관내장비수', '원본_관내장비수', self.var_day2.get() and self.var_intra.get()),
            (2, '관외', 'label_clean', '관외_혼잡도', '관외장비수', '원본_관외장비수', self.var_day2.get() and self.var_extra.get()),
            ('전체', '관내', 'label_clean', '관내_혼잡도', '관내장비수', '원본_관내장비수', self.var_day_all.get() and self.var_intra.get()),
            ('전체', '관외', 'label_clean', '관외_혼잡도', '관외장비수', '원본_관외장비수', self.var_day_all.get() and self.var_extra.get())
        ]
        
        # 활성화된 시나리오만 필터링
        active_scenarios = [s for s in all_scenarios if s[6]]
        if not active_scenarios: return

        unique_stations = df['사전투표소명'].unique()
        
        # [핵심] 여기서 새로 만드신 _plot_page 함수를 호출합니다!
        # save_name을 filename이라는 이름으로 넘겨줍니다.
        return self._plot_page(df, active_scenarios, unique_stations, filename=save_name, is_pdf=False)

    def _read_equip_summary(self):
        """
        장비현황 파일의 D7(총 장비수), H7(예비수) 셀을 읽어옵니다.
        엑셀/CSV 모두 헤더 없이 읽어서 좌표로 접근합니다.
        D7 -> Row 6, Col 3
        H7 -> Row 6, Col 7
        """
        if not self.equipment_file:
            return None, None
            
        try:
            # 헤더 없이 읽어서 절대 좌표(행/열)로 접근
            if self.equipment_file.endswith('.csv'):
                try: df = pd.read_csv(self.equipment_file, header=None, encoding='cp949')
                except: df = pd.read_csv(self.equipment_file, header=None, encoding='utf-8')
            else:
                df = pd.read_excel(self.equipment_file, header=None)
            
            # 파일 크기가 D7, H7을 읽을 수 있는지 확인 (행 7개 이상, 열 8개 이상)
            if df.shape[0] < 7 or df.shape[1] < 8:
                return None, None
                
            # D7 (Index: [6, 3]) -> 총 장비수
            raw_total = str(df.iloc[6, 3])
            # H7 (Index: [6, 7]) -> 예비수
            raw_reserve = str(df.iloc[6, 7])
            
            def _clean_num(val):
                # 숫자 외 문자 제거 (콤마, '대' 등)
                import re
                txt = re.sub(r'[^0-9]', '', val)
                return int(txt) if txt else 0
                
            total = _clean_num(raw_total)
            reserve = _clean_num(raw_reserve)
            
            return total, reserve
            
        except Exception as e:
            print(f"장비 요약 정보 로드 실패: {e}")
            return None, None
        
    def open_balance_popup(self):
        if not self.vote_files:
            messagebox.showwarning("주의", "먼저 투표 데이터 파일을 로드해주세요.")
            return
            
        # [수정] 기본값 설정 로직 변경
        # 1순위: 장비 파일의 D7(총보유), H7(예비) 값 사용
        # 2순위: 파일 없으면 기존 로직(화면 합계 + 5) 사용
        
        file_total, file_reserve = self._read_equip_summary()
        
        if file_total is not None and file_total > 0:
            default_total_assets = file_total
            default_reserve = file_reserve if file_reserve is not None else 5
            self.last_reserve_count = default_reserve # 읽어온 예비값 기억
            
            source_msg = "※ 장비 파일(D7, H7)에서 정보를 불러왔습니다."
        else:
            # [기존 로직 유지] 파일이 없거나 읽기 실패 시
            curr_allocated = sum([item['intra'] + item['extra'] for item in self.station_data.values()])
            default_total_assets = curr_allocated + self.last_reserve_count
            default_reserve = self.last_reserve_count
            
            source_msg = "※ 장비 파일 정보 없음 (기존 설정값 사용)"
        
        # 팝업창 생성
        pop = tk.Toplevel(self.root)
        pop.title("장비 자동 배분 (통합 모드)")
        pop.geometry("350x300") # 높이 약간 증가
        pop.resizable(False, False)
        
        # 화면 중앙 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 150
        pop.geometry(f"+{x}+{y}")
        
        ttk.Label(pop, text="보유한 [전체 장비 수]를 입력하세요.\n관내/관외 구분 없이 혼잡도에 따라 통합 배분합니다.", 
                  justify="center", foreground="gray").pack(pady=(15, 5))
        
        # 출처 표시 (파일에서 왔는지 확인용)
        ttk.Label(pop, text=source_msg, justify="center", foreground="blue", font=("맑은 고딕", 8)).pack(pady=(0, 10))
        
        frame_input = ttk.Frame(pop, padding="20")
        frame_input.pack(fill="both", expand=True)
        
        # 입력 필드 생성 함수
        def create_entry(parent, label, default_val):
            frame = ttk.Frame(parent)
            frame.pack(fill="x", pady=8)
            ttk.Label(frame, text=label, width=15, font=("맑은 고딕", 10, "bold")).pack(side="left")
            entry = ttk.Entry(frame, justify="right", font=("맑은 고딕", 10))
            entry.insert(0, str(default_val))
            entry.pack(side="right", expand=True, fill="x")
            return entry
            
        entry_total = create_entry(frame_input, "총 보유 장비:", default_total_assets)
        entry_reserve = create_entry(frame_input, "예비 장비:", default_reserve)
        
        def _run():
            try:
                total_assets = int(entry_total.get())
                total_reserve = int(entry_reserve.get())
                
                self.last_reserve_count = total_reserve
                
                available = total_assets - total_reserve
                min_req = len(self.station_data) * 2
                
                if available < min_req:
                    msg = f"장비가 부족합니다!\n\n투표소 수: {len(self.station_data)}개\n최소 필요 장비: {min_req}대 (관내1+관외1)\n현재 가용 장비: {available}대"
                    messagebox.showerror("배분 불가", msg)
                    return
                    
                self.run_auto_balance(total_assets, total_reserve)
                pop.destroy()
                
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자를 입력해주세요.")

        ttk.Button(pop, text="최적 배분 실행", command=_run).pack(fill="x", padx=20, pady=20)

    def run_auto_balance(self, total_assets, total_reserve):
        self._ensure_data_loaded()
        
        # 1. 사용할 수 있는 실제 장비 수
        target_count = total_assets - total_reserve
        num_stations = len(self.station_data)
        
        # 2. 기초 데이터 집계 (투표소별 총 투표자 수 예측)
        station_stats = {}
        
        for file in self.vote_files:
            if file not in self.cached_data: continue
            df, _, _ = self.cached_data[file]
            
            for idx, row in df.iterrows():
                st_name = str(row['사전투표소명']).strip()
                if st_name not in self.station_data: continue 
                
                if st_name not in station_stats:
                    station_stats[st_name] = {'intra_voters': 0, 'extra_voters': 0}
                
                # 분리된 설정값 가져오기
                user_rate_intra = self.station_data[st_name]['rate_intra']
                user_rate_extra = self.station_data[st_name]['rate_extra']
                elect_rate = self.station_data[st_name].get('elect_rate', 0)
                
                # 시뮬레이션과 동일한 공식 적용
                factor_intra = (1 + (elect_rate / 100.0)) * (1 + (user_rate_intra / 100.0))
                factor_extra = (1 + (user_rate_extra / 100.0))
                
                try:
                    station_stats[st_name]['intra_voters'] += float(row['관내사전투표자수']) * factor_intra
                    station_stats[st_name]['extra_voters'] += float(row['관외사전투표자수']) * factor_extra
                except: pass
        
        # 3. 배분 알고리즘 시작
        # (1) 기본 할당: 모든 투표소의 관내/관외에 1대씩 강제 할당
        current_alloc = {} # 장비 배분 현황을 저장할 빈 딕셔너리를 생성합니다.
        for st in self.station_data:
            # 모든 투표소(st)를 순회하며 초기값을 관내 1대, 관외 1대로 설정합니다.
            current_alloc[st] = {'intra': 1, 'extra': 1}

        # 남은 장비(remaining) 계산:
        # 전체 가용 장비(target_count)에서 방금 나눠준 기본 장비(투표소 수 * 2대)를 뺍니다.
        # 이제 이 'remaining' 개수만큼 추가 배분을 진행할 수 있습니다.    
        remaining = target_count - (num_stations * 2)
        
        # (2) Greedy Algorithm
        while remaining > 0: # 남은 장비가 0이 될 때까지 이 반복문을 계속 돌립니다.
            # 이번 턴(장비 1대)을 받을 '가장 바쁜 곳'을 찾기 위한 비교 변수 초기화
            max_load = -1 # 현재까지 발견된 '최대 부하(혼잡도)' 값
            target_info = None # 장비를 받을 대상 정보 (투표소명, 장비타입)
            
            # 모든 투표소를 하나씩 확인하며 어디가 제일 급한지 비교합니다.
            for st in current_alloc:
                # 관외 업무 가중치 (1.156)
                weight_extra = 1.156

                # 현재 이 투표소에 배정된 관내 장비 수 확인
                curr_intra = current_alloc[st]['intra']

                if curr_intra > 0:
                    # 관내 부하 = (예측된 관내 투표자 수) / (현재 관내 장비 수)
                    # 예: 투표자 1000명 / 장비 2대 = 부하 500
                    load_intra = station_stats[st]['intra_voters'] / curr_intra
                else:
                    load_intra = float('inf') # 0대면 무조건 최우선 배정

                # 만약 이 투표소의 관내 부하가 지금까지 찾은 최대 부하보다 크다면?
                if load_intra > max_load:
                    max_load = load_intra # 최대 부하 값을 갱신하고
                    target_info = (st, 'intra') # "현재 1등은 이 투표소의 '관내' 쪽입니다"라고 기록
                
                # 현재 이 투표소에 배정된 관외 장비 수 확인
                curr_extra = current_alloc[st]['extra']
                if curr_extra > 0:
                    # 관외 부하 = (예측된 관외 투표자 수 * 1.156) / (현재 관외 장비 수)
                    # *중요: 관외 투표자 수에 가중치를 곱해 부하를 더 높게(더 힘들게) 평가함
                    load_extra = (station_stats[st]['extra_voters'] * weight_extra) / curr_extra
                else:
                    load_extra = float('inf') # 0대면 무조건 최우선 배정

                # 만약 이 투표소의 관외 부하가 지금까지 찾은 최대 부하(관내 포함)보다 더 크다면?
                if load_extra > max_load:
                    max_load = load_extra # 최대 부하 값을 갱신하고
                    target_info = (st, 'extra') # "현재 1등은 이 투표소의 '관외' 쪽입니다"라고 기록 (덮어쓰기)
            
            # for문이 끝나고 최종적으로 선정된 곳(target_info)이 있다면
            if target_info:
                st_name, r_type = target_info # 이름과 타입(intra/extra)을 꺼냅니다.
                
                # 해당 투표소의 해당 타입 장비를 1대 늘려줍니다. (즉시 반영)
                # *중요: 여기서 +1 된 값은 다음 while 루프가 돌 때 'curr_intra/extra'가 되어
                # 분모를 키우므로 부하를 낮추는 역할을 합니다.
                current_alloc[st_name][r_type] += 1
                # 사용할 수 있는 남은 장비 수를 1개 줄입니다.
                remaining -= 1
            else:
                # 만약 더 이상 줄 곳이 없거나 오류가 있다면 반복을 종료합니다.
                break

        # 4. 결과 집계 및 UI 반영 (여기가 에러 났던 부분)
        total_intra_used = 0
        total_extra_used = 0
        
        for item_id in self.tree.get_children():
            # 안전하게 투표소명 가져오기
            item_values = self.tree.item(item_id)['values']
            if not item_values: continue
            st_name = str(item_values[0])
            
            if st_name in self.station_data:
                new_intra = current_alloc[st_name]['intra']
                new_extra = current_alloc[st_name]['extra']
                
                # 데이터 저장
                self.station_data[st_name]['intra'] = new_intra
                self.station_data[st_name]['extra'] = new_extra
                
                total_intra_used += new_intra
                total_extra_used += new_extra
                
                # [수정] UI 업데이트 시 필요한 변수들 가져오기
                org_intra = self.station_data[st_name]['org_intra']
                org_extra = self.station_data[st_name]['org_extra']
                
                # KeyError: 'rate' 해결 -> rate_intra, rate_extra 가져오기
                val_rate_intra = self.station_data[st_name]['rate_intra']
                val_rate_extra = self.station_data[st_name]['rate_extra']
                
                # 기존 선거인수 표시값 유지 (Treeview에서 가져옴)
                elect_disp = item_values[1]

                # 표시 텍스트 결정
                disp_intra = f"{org_intra} → {new_intra}" if new_intra != org_intra else str(new_intra)
                disp_extra = f"{org_extra} → {new_extra}" if new_extra != org_extra else str(new_extra)
                
                # [수정] 통합 텍스트 적용
                rate_txt = self._get_merged_rate_text(val_rate_intra, val_rate_extra)
                
                # 5개 컬럼 구조에 맞춰 업데이트
                self.tree.item(item_id, values=(st_name, elect_disp, disp_intra, disp_extra, rate_txt))
        
        # 5. 결과 메시지
        final_used = total_intra_used + total_extra_used
        msg = (f"배분 완료!\n\n"
               f"■ 총 보유 장비: {total_assets}대\n"
               f"■ 실제 배치: {final_used}대 (관내 {total_intra_used} / 관외 {total_extra_used})\n"
               f"■ 예비 장비: {total_reserve}대")
               
        self.log(f"[자동 배분] 총 {total_assets}대 중 {final_used}대 배치 완료. (예비 {total_reserve})")
        messagebox.showinfo("배분 완료", msg)
    def _open_rate_input_dialog(self, st_name, item_id, elect_disp, curr_intra, curr_extra, org_intra, org_extra):
        # 현재 값 가져오기
        cur_r_intra = self.station_data[st_name]['rate_intra']
        cur_r_extra = self.station_data[st_name]['rate_extra']

        # 팝업창 생성
        pop = tk.Toplevel(self.root)
        pop.title("조정률 개별 설정")
        pop.geometry("260x180")
        pop.resizable(False, False)
        
        # 중앙 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 130
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 90
        pop.geometry(f"+{x}+{y}")

        ttk.Label(pop, text=f"[{st_name}]", font=("맑은 고딕", 10, "bold")).pack(pady=(15, 10))

        frame_in = ttk.Frame(pop)
        frame_in.pack(fill="x", padx=30, pady=5)
        ttk.Label(frame_in, text="관내 조정(%):").pack(side="left")
        entry_intra = ttk.Entry(frame_in, width=10, justify="right")
        entry_intra.insert(0, str(cur_r_intra))
        entry_intra.pack(side="right")

        frame_out = ttk.Frame(pop)
        frame_out.pack(fill="x", padx=30, pady=5)
        ttk.Label(frame_out, text="관외 조정(%):").pack(side="left")
        entry_extra = ttk.Entry(frame_out, width=10, justify="right")
        entry_extra.insert(0, str(cur_r_extra))
        entry_extra.pack(side="right")

        def _apply():
            try:
                new_r_intra = int(entry_intra.get())
                new_r_extra = int(entry_extra.get())
                
                # 데이터 업데이트
                self.station_data[st_name]['rate_intra'] = new_r_intra
                self.station_data[st_name]['rate_extra'] = new_r_extra
                
                # 화면 갱신
                rate_txt = self._get_merged_rate_text(new_r_intra, new_r_extra)
                
                # 장비대수 표시 텍스트 생성 (변경여부 확인)
                disp_intra = f"{org_intra} → {curr_intra}" if curr_intra != org_intra else str(curr_intra)
                disp_extra = f"{org_extra} → {curr_extra}" if curr_extra != org_extra else str(curr_extra)

                self.tree.item(item_id, values=(st_name, elect_disp, disp_intra, disp_extra, rate_txt))
                self.log(f"{st_name} 조정률 변경: 내 {new_r_intra}% / 외 {new_r_extra}%")
                pop.destroy()
            except ValueError:
                messagebox.showerror("오류", "숫자만 입력해주세요.", parent=pop)

        ttk.Button(pop, text="적용", command=_apply).pack(pady=15, fill='x', padx=30)

    def _plot_page(self, df, scenarios, stations_list, filename=None, is_pdf=False):
        count = len(scenarios)
        
        # 1. 기본 단위 높이 계산
        if is_pdf:
            unit_h = 13 
        else:
            # 투표소 개수에 따라 유동적으로 높이 조절
            unit_h = max(7, 4 + (len(stations_list) * 0.6))

        # 2. [수정] 그래프 개수에 따른 행/열 및 전체 크기 자동 계산 (최대 6개 대응)
        if count == 1: 
            nrows, ncols = 1, 1
            figsize = (12, unit_h)
        elif count == 2: 
            nrows, ncols = 1, 2
            figsize = (20, unit_h)
        elif count <= 4: 
            nrows, ncols = 2, 2
            figsize = (20, unit_h * 2) # 2줄 높이
        else: 
            # 5개~6개인 경우 (3행 2열) -> 1,2일차+전체 선택 시 여기 해당
            nrows, ncols = 3, 2
            figsize = (20, unit_h * 3) # 3줄 높이

        # 3. 서브플롯 생성
        fig, axes = plt.subplots(nrows, ncols, figsize=figsize)
        
        # axes 배열을 1차원 리스트로 펴서 인덱싱하기 쉽게 변환
        if count == 1: axes_flat = [axes]
        else: axes_flat = axes.flatten()

        max_val = max(df['관내_혼잡도'].max(), df['관외_혼잡도'].max()) if not df.empty else 1
        
        for idx, (day, type_name, label_col, value_col, eq_col, org_eq_col, _) in enumerate(scenarios):
            ax = axes_flat[idx]
            
            # [전체]와 [일반 일차] 구분하여 데이터 필터링
            if str(day) == '전체':
                df_day = df[df['일차'] == '전체']
            else:
                df_day = df[df['일차'] == day]
            
            if df_day.empty:
                ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center')
                continue
            
            pivot = df_day.pivot_table(index=label_col, columns='시간대', values=value_col)
            
            # 평균 행/열 생성
            avg_label = '' 
            pivot[avg_label] = pivot.mean(axis=1) 
            avg_row = pivot.mean(axis=0)
            pivot.loc[avg_label] = avg_row
            
            # 정렬
            time_cols = sorted([c for c in pivot.columns if c != avg_label])
            new_cols = [avg_label] + time_cols
            pivot = pivot[new_cols]
            
            target_labels = [s.replace('사전투표소','') for s in stations_list]
            valid_labels = [l for l in target_labels if l in pivot.index]
            new_rows = [avg_label] + valid_labels
            pivot = pivot.reindex(new_rows)

            # 장비 데이터 매칭
            # '전체'일 경우 장비 수는 평균이 아니라 그냥 해당 투표소의 설정값을 따라가야 함 (중복 제거)
            if str(day) == '전체':
                # 전체 평균 데이터에는 장비수 컬럼이 평균내져 있을 수 있으므로, 원본 매핑을 다시 참조하거나
                # 이미 df_mean 생성 시 장비수도 평균냈으므로(같은 값이면 평균도 같음) 그대로 사용
                equip_data = df_day.drop_duplicates(subset=[label_col]).set_index(label_col)[[eq_col, org_eq_col]]
            else:
                equip_data = df_day.drop_duplicates(subset=[label_col]).set_index(label_col)[[eq_col, org_eq_col]]

            annot_labels = []
            for row_label in new_rows:
                if row_label == avg_label:
                    annot_labels.append("") 
                else:
                    try:
                        # 장비대수는 소수점이 나올 수 없으므로 int 처리 (전체 평균인 경우에도 장비수는 동일)
                        curr = int(equip_data.loc[row_label, eq_col])
                        org = int(equip_data.loc[row_label, org_eq_col])
                        if curr != org: txt = f"{org} → {curr}"
                        else: txt = f"{curr}"
                        annot_labels.append(txt)
                    except: annot_labels.append("?")

            equip_df = pd.DataFrame(1, index=new_rows, columns=['장비']) 
            equip_df.iloc[0] = 0 

            annot_matrix = pd.DataFrame(annot_labels, index=new_rows, columns=['장비'])

            divider = make_axes_locatable(ax)
            ax_equip = divider.append_axes("left", size="7%", pad=0.08) 
            
            custom_cmap = ListedColormap(['white', '#F0F4F8'])

            sns.heatmap(equip_df, annot=annot_matrix, fmt='', 
                        cmap=custom_cmap, vmin=0, vmax=1,
                        cbar=False, xticklabels=False,
                        linewidths=0.5, linecolor='white', ax=ax_equip)
            
            ax_equip.set_xlabel("")
            ax_equip.set_ylabel("사전투표소", fontsize=11, fontweight='bold')
            ax_equip.tick_params(axis='y', rotation=0, length=0)

            ax_equip.text(0.5, 0.95, "장비수", ha='center', va='bottom', fontsize=10, fontweight='bold', color='black')
            ax_equip.text(0.95, 0.5, "시간대별 평균 →", ha='right', va='center', fontsize=9, fontweight='bold', color='#3B5BDB')

            # [추가] 1. 주석(Annotation)용 데이터프레임 생성 (문자열 포맷)
            annot_df = pivot.applymap(lambda x: f"{x:.1f}")

            # [추가] 2. 11시~18시 컬럼 필터링 및 평균 계산
            # pivot의 컬럼 중 정수형이면서 11 이상 18 이하인 것만 추출
            target_hours = [c for c in pivot.columns if isinstance(c, (int, float)) and 11 <= c <= 18]
            
            if target_hours:
                # avg_label('') 행은 '시간대별 전체 평균'을 담고 있음. 여기서 11~18시 데이터만 뽑아서 다시 평균 계산
                mean_11_18 = pivot.loc[avg_label, target_hours].mean()
                
                # [추가] 3. 좌측 상단(전체 평균) 셀 텍스트 수정
                # 기존 값(전체 평균) 아래에 괄호로 11~18시 평균 추가
                original_text = annot_df.iloc[0, 0]
                annot_df.iloc[0, 0] = f"{original_text}\n({mean_11_18:.1f})"

            # [수정] annot에 True 대신 직접 만든 문자열 DF(annot_df) 전달, fmt는 비움('')
            sns.heatmap(pivot, annot=annot_df, fmt='', cmap='Greens', cbar=False, 
                        linewidths=0.5, linecolor='white', vmin=0, vmax=max_val, ax=ax)
            
            ax.text(0.5, -0.2, "↓ 투표소별\n평균", ha='center', va='bottom', fontsize=10, fontweight='bold', color='#3B5BDB', clip_on=False)
            
            rect_row = patches.Rectangle((0, 0), len(pivot.columns), 1, linewidth=3, edgecolor='#3B5BDB', facecolor='none', clip_on=False)
            ax.add_patch(rect_row)
            rect_col = patches.Rectangle((0, 0), 1, len(pivot), linewidth=3, edgecolor='#3B5BDB', facecolor='none', clip_on=False)
            ax.add_patch(rect_col)

            ax.set_ylabel("") 
            ax.set_yticks([]) 
            ax.xaxis.tick_top()
            ax.xaxis.set_label_position('top')
            ax.tick_params(axis='x', length=0)

            # [엄격 모드] 데이터 무결성 검사 (순서대로 붙여넣으세요)
            if time_cols:
                try:
                    # 1. 시간대 숫자 변환 및 범위 계산
                    times = [int(c) for c in time_cols]
                    start_t, end_t = min(times), max(times)
                    expected_count = end_t - start_t + 1 # 예: 6시~9시면 4개여야 함
                    
                    # 2. [검증] 실제 데이터 칸 수 vs 계산된 칸 수 비교
                    if len(times) != expected_count:
                        # 여기서 에러를 발생시켜 프로그램이 경고창을 띄우게 함
                        raise ValueError(
                            f"데이터 오류 발견! [{day}일차]\n"
                            f"시간대가 연속되지 않거나 중복 파일이 있습니다.\n"
                            f"- 범위: {start_t}시 ~ {end_t}시 (필요: {expected_count}칸)\n"
                            f"- 실제: {len(times)}칸 (중복/누락 확인 필요)"
                        )
                        
                    # 3. 검증 통과 시, 엄격한 기준으로 라벨 생성
                    labels = [''] + list(range(start_t, end_t + 1))
                    
                except ValueError as ve:
                    raise ve # 위에서 만든 에러 메시지를 그대로 상위로 전달
                except Exception:
                    # 숫자가 아닌 컬럼이 섞여있을 경우 (예외 처리)
                    labels = [''] + time_cols
            else:
                labels = ['']

            # 4. 틱(눈금) 위치 설정 (데이터 개수에 정확히 맞춤)
            # 0.5, 1.5, 2.5... 위치에 라벨을 찍어 정확도 향상
            ticks = np.arange(len(pivot.columns)) + 0.5
            ax.set_xticks(ticks)
            ax.set_xticklabels(labels, rotation=0)

            # [재수정] 제목 포맷 변경: {관내/관외} 사전투표 ({기간})
            # 예: 관내 사전투표 (전체(평균)) 또는 관외 사전투표 (1일차)
            day_str = "전체(평균)" if str(day) == '전체' else f"{day}일차"
            title_txt = f"{type_name} 사전투표 ({day_str})"
            
            ax.set_title(title_txt, fontsize=14, fontweight='bold', pad=20)
            ax.set_xlabel('시간대', fontsize=11, fontweight='bold')

        # [추가] 만들어진 칸보다 그래프가 적을 때 빈 칸 숨기기 (예: 6칸 만들었는데 5개만 그릴 때)
        for i in range(count, len(axes_flat)):
            axes_flat[i].axis('off')

        # [재수정] 메인 타이틀 포맷 변경: {지역명} 사전투표소 (예상) 혼잡도
        # self.region_name에 값이 있으면 넣고, 없으면 기본 텍스트 출력
        if self.region_name:
            main_title = f"{self.region_name} 사전투표소 (예상) 혼잡도"
        else:
            main_title = "사전투표소 (예상) 혼잡도"

        fig.suptitle(main_title, fontsize=20, fontweight='bold')
        # [수정] 하단 설명 문구 개선 (가독성 높임)
        fig.text(0.5, 0.02, 
                    "※ 각 셀의 수치: 장비 1대당 1시간 동안의 투표자 수 (혼잡도)\n"
                    "파란색 테두리: 전체 시간 평균  |  ( 괄호 안 숫자 ): 11~18시 집중평균  |  장비: [기존] → [변경]", 
                    ha='center', fontsize=11, fontweight='bold', color='#333333')
        
        plt.tight_layout(rect=[0, 0.05, 1, 0.95]) 
        
        # _plot_page 함수가 받은 인자인 filename을 사용해야 합니다.
        if filename and not is_pdf:
            plt.savefig(filename)
            plt.close(fig)
            
        return fig

if __name__ == "__main__":
    root = tk.Tk()
    app = ElectionAnalyzerApp(root)
    root.mainloop()

