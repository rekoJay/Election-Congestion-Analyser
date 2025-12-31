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
        self.cached_data = {} # [최적화] 읽어들인 데이터를 메모리에 저장 (경로: (df, day, time))
        self.equipment_file = None

        self.last_reserve_count = 5
        
        # 데이터 구조: { '투표소명': {'intra': 1, 'extra': 1, 'rate': 0.0, 'org_intra': 1, 'org_extra': 1} }
        self.station_data = {} 
        
        self.create_widgets()
        
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
        self.lbl_equip_status.pack(pady=(2, 0))
        
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
        
        columns = ("station", "intra", "extra", "rate")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=12)
        self.tree.heading("station", text="투표소명")
        self.tree.heading("intra", text="관내 장비")
        self.tree.heading("extra", text="관외 장비")
        self.tree.heading("rate", text="증가율(%)")
        
        self.tree.column("station", width=180)
        self.tree.column("intra", width=70, anchor="center")
        self.tree.column("extra", width=70, anchor="center")
        self.tree.column("rate", width=80, anchor="center")
        
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
        
        ttk.Label(chk_frame, text="기간: ").pack(side="left")
        ttk.Checkbutton(chk_frame, text="1일차", variable=self.var_day1).pack(side="left", padx=5)
        ttk.Checkbutton(chk_frame, text="2일차", variable=self.var_day2).pack(side="left", padx=5)
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
                # 데이터 업데이트
                self.station_data[st_name]['rate'] = rate
                
                # [수정] 화면 갱신 시 기존 화살표 상태 유지 로직
                curr_intra = self.station_data[st_name]['intra']
                curr_extra = self.station_data[st_name]['extra']
                org_intra = self.station_data[st_name]['org_intra']
                org_extra = self.station_data[st_name]['org_extra']
                
                disp_intra = f"{org_intra} → {curr_intra}" if curr_intra != org_intra else str(curr_intra)
                disp_extra = f"{org_extra} → {curr_extra}" if curr_extra != org_extra else str(curr_extra)

                self.tree.item(item_id, values=(st_name, disp_intra, disp_extra, rate))

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
                import traceback
                traceback.print_exc()

        for item in self.tree.get_children():
            self.tree.delete(item)
        
        sorted_stations = station_list
        self.station_data = {} 
        current_global_rate = int(self.var_rate.get())

        for st in sorted_stations:
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
            
            self.station_data[st] = {
                'intra': intra, 'extra': extra, 'rate': current_global_rate,
                'org_intra': intra, 'org_extra': extra
            }
            self.tree.insert("", "end", iid=st, values=(st, intra, extra, current_global_rate))
            
        self.log(f"목록 갱신 완료: 총 {len(sorted_stations)}개 투표소")

    # [복구된 함수] 더블 클릭 이벤트 핸들러
    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x) 
        
        if not item_id: return
        
        st_name = item_id
        # [수정] 표에 적힌 글자(vals) 대신 실제 데이터(self.station_data)에서 값을 가져옵니다.
        # 이렇게 해야 "3 → 5" 같은 문자가 있어도 숫자 5를 정확히 가져옵니다.
        if st_name not in self.station_data: return
        
        data = self.station_data[st_name]
        curr_intra = data['intra']
        curr_extra = data['extra']
        curr_rate = data['rate']
        org_intra = data['org_intra']
        org_extra = data['org_extra']
        
        # 화면 표시용 텍스트 생성 함수 (내부 함수)
        def get_display_text(val, org_val):
            return f"{org_val} → {val}" if val != org_val else str(val)
        
        if column == '#2': # 관내
            new_intra = simpledialog.askinteger("관내 장비 수정", f"[{st_name}]\n관내 장비 수:", 
                                              initialvalue=curr_intra, minvalue=1, maxvalue=50)
            if new_intra is not None:
                self.station_data[st_name]['intra'] = new_intra
                # UI 업데이트 (화살표 반영)
                disp_intra = get_display_text(new_intra, org_intra)
                disp_extra = get_display_text(curr_extra, org_extra)
                self.tree.item(item_id, values=(st_name, disp_intra, disp_extra, curr_rate))
                self.log(f"{st_name} 관내 장비 변경: {new_intra}대")
                
        elif column == '#3': # 관외
            new_extra = simpledialog.askinteger("관외 장비 수정", f"[{st_name}]\n관외 장비 수:", 
                                              initialvalue=curr_extra, minvalue=1, maxvalue=50)
            if new_extra is not None:
                self.station_data[st_name]['extra'] = new_extra
                # UI 업데이트 (화살표 반영)
                disp_intra = get_display_text(curr_intra, org_intra)
                disp_extra = get_display_text(new_extra, org_extra)
                self.tree.item(item_id, values=(st_name, disp_intra, disp_extra, curr_rate))
                self.log(f"{st_name} 관외 장비 변경: {new_extra}대")
                
        elif column == '#4': # 증가율
            new_rate = simpledialog.askinteger("증가율 수정", f"[{st_name}]\n투표자 증가율(%):", 
                                             initialvalue=curr_rate, minvalue=-100, maxvalue=200)
            if new_rate is not None:
                # UI 업데이트 (기존 화살표 유지)
                disp_intra = get_display_text(curr_intra, org_intra)
                disp_extra = get_display_text(curr_extra, org_extra)
                self.tree.item(item_id, values=(st_name, disp_intra, disp_extra, new_rate))
                self.station_data[st_name]['rate'] = new_rate
                self.log(f"{st_name} 증가율 변경: {new_rate}%")

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
        # [실제 작업 로직] 기존 run_simulation 내용을 이곳으로 이동
        try:
            label = "통합 분석"
            self.log(f"시뮬레이션 시작: {label}")
            
            # 메인 스레드가 아닌 곳에서 GUI를 업데이트 하려면 invoke나 after를 써야 하지만,
            # 데이터 로드와 계산은 백그라운드에서 안전하게 수행됨.
            self._ensure_data_loaded() 
            
            all_data = []
            
            for file in self.vote_files:
                if file not in self.cached_data: continue
                
                try:
                    org_df, day, time = self.cached_data[file]
                    df = org_df.copy() 
                    
                    df['사전투표소명'] = df['사전투표소명'].astype(str).str.strip()
                    
                    rate_map = {name: data.get('rate', 0) for name, data in self.station_data.items()}
                    rates = df['사전투표소명'].map(rate_map).fillna(0)
                    factor = 1 + (rates / 100.0)
                    
                    df['관내사전투표자수'] = df['관내사전투표자수'] * factor
                    df['관외사전투표자수'] = df['관외사전투표자수'] * factor
                            
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
            
            final_df['시간대별_관내투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관내사전투표자수'].diff()
            final_df['시간대별_관외투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관외사전투표자수'].diff()
            
            mask_start = final_df['시간대'] == 7
            final_df.loc[mask_start, '시간대별_관내투표자수'] = final_df.loc[mask_start, '관내사전투표자수']
            final_df.loc[mask_start, '시간대별_관외투표자수'] = final_df.loc[mask_start, '관외사전투표자수']

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
            # [수정] exe로 실행될 때와 파이썬 스크립트로 실행될 때의 경로 차이 해결
            if getattr(sys, 'frozen', False):
                # exe 실행 시: 실행 파일이 있는 위치를 저장 경로로 설정
                script_dir = os.path.dirname(os.path.abspath(sys.executable))
            else:
                # 코드 실행 시: 파일이 있는 위치 설정
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
                self.loading_win.destroy() # 로딩창 닫기
                messagebox.showinfo("완료", f"분석 완료!\n\n파일이 저장되었습니다:\n{full_png_path}")
                if platform.system() == 'Windows':
                    try: os.startfile(full_png_path)
                    except: pass
            
            self.root.after(0, _finish)

        except Exception as e:
            # 에러 발생 시 처리
            def _error():
                if hasattr(self, 'loading_win'): self.loading_win.destroy()
                self.log(f"치명적 오류: {e}")
                import traceback
                traceback.print_exc()
                messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n{e}")
            self.root.after(0, _error)

    def visualize_results(self, df, timestamp, save_name, mode='screen'):
        system_name = platform.system()
        font_family = 'Malgun Gothic' if system_name == 'Windows' else 'AppleGothic'
        plt.rc('font', family=font_family)
        plt.rc('axes', unicode_minus=False)

        df['short_name'] = df['사전투표소명'].str.replace('사전투표소', '')
        df['label_clean'] = df['short_name'] 

        all_scenarios = [
            (1, '관내', 'label_clean', '관내_혼잡도', '관내장비수', '원본_관내장비수', self.var_day1.get() and self.var_intra.get()),
            (1, '관외', 'label_clean', '관외_혼잡도', '관외장비수', '원본_관외장비수', self.var_day1.get() and self.var_extra.get()),
            (2, '관내', 'label_clean', '관내_혼잡도', '관내장비수', '원본_관내장비수', self.var_day2.get() and self.var_intra.get()),
            (2, '관외', 'label_clean', '관외_혼잡도', '관외장비수', '원본_관외장비수', self.var_day2.get() and self.var_extra.get())
        ]
        
        active_scenarios = [s for s in all_scenarios if s[6]]
        if not active_scenarios: return

        unique_stations = df['사전투표소명'].unique()
        total_stations = len(unique_stations)
        
        if mode == 'screen':
            # === 화면용: 길게 한 장으로 ===
            self._plot_page(df, active_scenarios, unique_stations, save_name, is_pdf=False)
    
    def open_balance_popup(self):
        if not self.vote_files:
            messagebox.showwarning("주의", "먼저 투표 데이터 파일을 로드해주세요.")
            return
            
        # [수정] 현재 화면에 배치된 실제 장비 수 합산
        curr_allocated = sum([item['intra'] + item['extra'] for item in self.station_data.values()])
        
        # [수정] 팝업에 띄울 '총 보유 장비' 초기값 = (현재 배치된 장비) + (창고에 있는 예비 장비)
        # 이렇게 해야 99(배치) + 1(예비) = 100(총보유)으로 올바르게 표시됩니다.
        default_total_assets = curr_allocated + self.last_reserve_count
        
        # 팝업창 생성
        pop = tk.Toplevel(self.root)
        pop.title("장비 자동 배분 (통합 모드)")
        pop.geometry("350x280") 
        pop.resizable(False, False)
        
        # 화면 중앙 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 140
        pop.geometry(f"+{x}+{y}")
        
        ttk.Label(pop, text="보유한 [전체 장비 수]를 입력하세요.\n관내/관외 구분 없이 혼잡도에 따라 통합 배분합니다.", 
                  justify="center", foreground="gray").pack(pady=15)
        
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
            
        # [수정] 계산된 값(default_total_assets)을 입력창 초기값으로 설정
        entry_total = create_entry(frame_input, "총 보유 장비:", default_total_assets)
        
        # 기억해둔 예비 장비 값 사용
        entry_reserve = create_entry(frame_input, "예비 장비:", self.last_reserve_count) 
        
        def _run():
            try:
                total_assets = int(entry_total.get())
                total_reserve = int(entry_reserve.get())
                
                # 입력한 예비 장비 수를 변수에 저장 (다음 번을 위해 기억)
                self.last_reserve_count = total_reserve
                
                # 가용 장비 = 총 보유 - 예비
                available = total_assets - total_reserve
                
                # 최소 요구량: 투표소 수 * 2 (관내1 + 관외1)
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
        
        # 2. 기초 데이터 집계 (투표소별 총 투표자 수)
        # 구조: {'투표소명': {'intra_voters': 1000, 'extra_voters': 200, ...}}
        station_stats = {}
        
        for file in self.vote_files:
            if file not in self.cached_data: continue
            df, _, _ = self.cached_data[file]
            
            for idx, row in df.iterrows():
                st_name = str(row['사전투표소명']).strip()
                if st_name not in self.station_data: continue # 리스트에 없는 투표소 건너뜀
                
                if st_name not in station_stats:
                    station_stats[st_name] = {'intra_voters': 0, 'extra_voters': 0}
                
                rate = self.station_data[st_name]['rate']
                factor = 1 + (rate / 100.0)
                
                try:
                    station_stats[st_name]['intra_voters'] += float(row['관내사전투표자수']) * factor
                    station_stats[st_name]['extra_voters'] += float(row['관외사전투표자수']) * factor
                except: pass
        
        # 3. 배분 알고리즘 시작
        # (1) 모든 투표소의 관내/관외에 1대씩 강제 할당
        current_alloc = {}
        for st in self.station_data:
            current_alloc[st] = {'intra': 1, 'extra': 1}
            
        # 남은 장비 수 계산 (총 가용 - (투표소수 * 2))
        remaining = target_count - (num_stations * 2)
        
        # (2) Greedy Algorithm: 남은 장비를 하나씩 '가장 혼잡한 곳(관내/관외 불문)'에 투입
        while remaining > 0:
            max_load = -1
            target_info = None # (st_name, 'intra' or 'extra')
            
            for st in current_alloc:
                # [수정] 관외 업무 가중치 (1.156 = 관외가 관내보다 처리 시간이 1.156배 걸린다고 가정)
                # 이 값을 높일수록 관외에 장비가 더 많이 배정됩니다.
                weight_extra = 1.156

                # 관내 혼잡도 계산
                load_intra = station_stats[st]['intra_voters'] / current_alloc[st]['intra']
                if load_intra > max_load:
                    max_load = load_intra
                    target_info = (st, 'intra')
                    
                # 관외 혼잡도 계산 (가중치 적용)
                # 관외 투표자 수에 가중치를 곱해 부하를 높게 산출 -> 장비 우선 할당 유도
                load_extra = (station_stats[st]['extra_voters'] * weight_extra) / current_alloc[st]['extra']
                if load_extra > max_load:
                    max_load = load_extra
                    target_info = (st, 'extra')
            
            if target_info:
                st_name, r_type = target_info
                current_alloc[st_name][r_type] += 1
                remaining -= 1
            else:
                break

        # 4. 결과 집계 및 UI 반영
        total_intra_used = 0
        total_extra_used = 0
        
        for item_id in self.tree.get_children():
            st_name = self.tree.item(item_id)['values'][0]
            if st_name in self.station_data:
                new_intra = current_alloc[st_name]['intra']
                new_extra = current_alloc[st_name]['extra']
                
                # 데이터 저장
                self.station_data[st_name]['intra'] = new_intra
                self.station_data[st_name]['extra'] = new_extra
                
                total_intra_used += new_intra
                total_extra_used += new_extra
                
                # [수정] UI 업데이트 시 변경된 값은 화살표로 표시
                org_intra = self.station_data[st_name]['org_intra']
                org_extra = self.station_data[st_name]['org_extra']
                rate = self.station_data[st_name]['rate']
                
                # 표시 텍스트 결정 (다르면 "원래값 → 새값", 같으면 "값")
                disp_intra = f"{org_intra} → {new_intra}" if new_intra != org_intra else str(new_intra)
                disp_extra = f"{org_extra} → {new_extra}" if new_extra != org_extra else str(new_extra)
                
                self.tree.item(item_id, values=(st_name, disp_intra, disp_extra, rate))
        
        # 5. 결과 메시지
        final_used = total_intra_used + total_extra_used
        msg = (f"배분 완료!\n\n"
               f"■ 총 보유 장비: {total_assets}대\n"
               f"■ 실제 배치: {final_used}대 (관내 {total_intra_used} / 관외 {total_extra_used})\n"
               f"■ 예비 장비: {total_reserve}대")
               
        self.log(f"[자동 배분] 총 {total_assets}대 중 {final_used}대 배치 완료. (예비 {total_reserve})")
        messagebox.showinfo("배분 완료", msg)

    def _plot_page(self, df, scenarios, stations_list, filename=None, is_pdf=False):
        # 내부적으로 사용하는 그리기 함수
        count = len(scenarios)
        
        # 높이 계산 (PDF는 고정 A4 비율 권장, 화면용은 동적)
        if is_pdf:
            # A4 Landscape 느낌의 비율 (가로 20, 세로 12 고정)
            figsize_h = 13 
        else:
            # 화면용은 길게 (여백 + 투표소당 높이)
            figsize_h = max(7, 4 + (len(stations_list) * 0.6))

        if count == 1: nrows, ncols, figsize = 1, 1, (12, figsize_h)
        elif count == 2: nrows, ncols, figsize = 1, 2, (20, figsize_h)
        elif count == 3: nrows, ncols, figsize = 1, 3, (22, figsize_h)
        else: nrows, ncols, figsize = 2, 2, (20, figsize_h * 2) # 2줄이면 높이 2배

        fig, axes = plt.subplots(nrows, ncols, figsize=figsize)
        if count == 1: axes_flat = [axes]
        else: axes_flat = axes.flatten()

        max_val = max(df['관내_혼잡도'].max(), df['관외_혼잡도'].max()) if not df.empty else 1
        
        for idx, (day, type_name, label_col, value_col, eq_col, org_eq_col, _) in enumerate(scenarios):
            ax = axes_flat[idx]
            df_day = df[df['일차'] == day]
            
            if df_day.empty:
                ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center')
                continue
            
            pivot = df_day.pivot_table(index=label_col, columns='시간대', values=value_col)
            

            # === [수정됨] 평균 행/열 생성 및 텍스트 수동 배치 코드 시작 ===
            
            # 1. 평균 계산 (라벨을 빈 문자열 ''로 설정하여 Y축 이름이 안 겹치게 함)
            avg_label = '' 
            pivot[avg_label] = pivot.mean(axis=1) 
            avg_row = pivot.mean(axis=0)
            pivot.loc[avg_label] = avg_row
            
            # 2. 정렬 (평균을 맨 앞으로)
            time_cols = sorted([c for c in pivot.columns if c != avg_label])
            new_cols = [avg_label] + time_cols
            pivot = pivot[new_cols]
            
            target_labels = [s.replace('사전투표소','') for s in stations_list]
            valid_labels = [l for l in target_labels if l in pivot.index]
            new_rows = [avg_label] + valid_labels
            pivot = pivot.reindex(new_rows)

            # 3. 장비 데이터 준비
            equip_data = df_day.drop_duplicates(subset=[label_col]).set_index(label_col)[[eq_col, org_eq_col]]
            annot_labels = []
            
            for row_label in new_rows:
                if row_label == avg_label:
                    annot_labels.append("") # 텍스트를 수동으로 넣기 위해 빈칸으로 둠
                else:
                    try:
                        curr = equip_data.loc[row_label, eq_col]
                        org = equip_data.loc[row_label, org_eq_col]
                        if curr != org: txt = f"{int(org)} → {int(curr)}"
                        else: txt = f"{int(curr)}"
                        annot_labels.append(txt)
                    except: annot_labels.append("?")

            # [수정] 데이터프레임 생성 시 값 구분 (1: 데이터 행, 0: 헤더 행)
            equip_df = pd.DataFrame(1, index=new_rows, columns=['장비']) 
            equip_df.iloc[0] = 0 # 첫 번째 행(헤더)은 0으로 설정

            annot_matrix = pd.DataFrame(annot_labels, index=new_rows, columns=['장비'])

            divider = make_axes_locatable(ax)
            ax_equip = divider.append_axes("left", size="7%", pad=0.08) 
            
            # [수정] 컬러맵 정의: 0 -> 흰색(헤더), 1 -> 연회색(데이터)
            custom_cmap = ListedColormap(['white', '#F0F4F8'])

            # 4. 왼쪽 장비수 히트맵 (vmin=0, vmax=1로 색상 고정)
            sns.heatmap(equip_df, annot=annot_matrix, fmt='', 
                        cmap=custom_cmap, vmin=0, vmax=1,
                        cbar=False, xticklabels=False,
                        linewidths=0.5, linecolor='white', ax=ax_equip)
            
            ax_equip.set_xlabel("")
            ax_equip.set_ylabel("사전투표소", fontsize=11, fontweight='bold')
            # [수정] length=0 을 추가하여 이름 옆의 눈금(-) 표시 제거
            ax_equip.tick_params(axis='y', rotation=0, length=0)

            # [텍스트 추가 1] 왼쪽 바닥 중앙 "장비수" (x=0.5, y=0.95)
            ax_equip.text(0.5, 0.95, "장비수", 
                         ha='center', va='bottom', 
                         fontsize=10, fontweight='bold', color='black')

            # 5. 오른쪽 메인 히트맵
            sns.heatmap(pivot, annot=True, fmt='.1f', cmap='Greens', cbar=False, 
                        linewidths=0.5, linecolor='white', vmin=0, vmax=max_val, ax=ax)
            
            # [수정 1] 왼쪽 바닥 중앙 "장비수"
            ax_equip.text(0.5, 0.95, "장비수", 
                         ha='center', va='bottom', 
                         fontsize=10, fontweight='bold', color='black')

            # [수정 2] "시간대별 평균 →" 라벨을 장비 그래프(ax_equip) 영역 안으로 이동
            # 이렇게 하면 왼쪽으로 잘리지 않고, 장비수 칸 안에서 오른쪽을 가리키게 됩니다.
            # x=0.95 (장비칸의 오른쪽 끝), y=0.5 (첫 번째 행의 중앙)
            ax_equip.text(0.95, 0.5, "시간대별 평균 →", 
                         ha='right', va='center', 
                         fontsize=9, fontweight='bold', color='#3B5BDB')

            # 5. 오른쪽 메인 히트맵
            sns.heatmap(pivot, annot=True, fmt='.1f', cmap='Greens', cbar=False, 
                        linewidths=0.5, linecolor='white', vmin=0, vmax=max_val, ax=ax)
            
            # [수정 3] 글자가 길어 숫자 '6'을 가리는 문제 해결 -> 두 줄로 분리
            # "↓ 투표소별" (윗줄) / "평균" (아랫줄)
            ax.text(0.5, -0.2, "↓ 투표소별\n평균", 
                    ha='center', va='bottom', 
                    fontsize=10, fontweight='bold', color='#3B5BDB',
                    clip_on=False)
            
            rect_row = patches.Rectangle((0, 0), len(pivot.columns), 1, linewidth=3, edgecolor='#3B5BDB', facecolor='none', clip_on=False)
            ax.add_patch(rect_row)
            rect_col = patches.Rectangle((0, 0), 1, len(pivot), linewidth=3, edgecolor='#3B5BDB', facecolor='none', clip_on=False)
            ax.add_patch(rect_col)

            ax.set_ylabel("") 
            ax.set_yticks([]) 
            ax.xaxis.tick_top()
            ax.xaxis.set_label_position('top')
            
            # [수정] X축(시간대) 눈금(-) 길이 0으로 설정하여 제거
            ax.tick_params(axis='x', length=0)

            ticks = [0.5] + list(range(1, len(pivot.columns) + 1))
            ax.set_xticks(ticks)
            
            if time_cols:
                start_time = int(time_cols[0]) - 1
                end_time = int(time_cols[-1])
                labels = [''] + list(range(start_time, end_time + 1))
                ax.set_xticklabels(labels, rotation=0)

            ax.set_title(f'{type_name} 사전투표 {day}일차 (예상) 혼잡도', fontsize=14, fontweight='bold', pad=20)
            ax.set_xlabel('시간대', fontsize=11, fontweight='bold')

        if count == 3 and nrows * ncols > 3: axes_flat[3].axis('off')

        fig.suptitle(f"사전투표 운용장비 산출 시뮬레이션 결과", fontsize=20, fontweight='bold')
        fig.text(0.5, 0.02, 
                    "각 셀의 수치는 1시간 동안 사전투표 장비 1대당 투표용지 발급자 수를 나타냄.\n"
                    "장비 열 표기: [기존] → [변경] / 파란색 테두리: 평균값", 
                    ha='center', fontsize=11, color='gray')
        
        plt.tight_layout(rect=[0, 0.05, 1, 0.95]) 
        
        if filename and not is_pdf:
            plt.savefig(filename)
            plt.close(fig) # PNG 저장 후 닫기
            
        return fig # PDF 저장을 위해 figure 객체 반환

if __name__ == "__main__":
    root = tk.Tk()
    app = ElectionAnalyzerApp(root)
    root.mainloop()

