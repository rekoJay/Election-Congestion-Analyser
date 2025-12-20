import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from datetime import datetime

# 시각화 및 시스템 관련
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
import platform
import numpy as np 

class ElectionAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사전투표 운용장비 산출 프로그램")
        self.root.geometry("680x920") 
        self.root.resizable(False, True) 
        
        self.vote_files = []
        self.equipment_file = None
        # 데이터 구조: { '투표소명': {'intra': 1, 'extra': 1, 'rate': 0.0} }
        self.station_data = {} 
        
        self.create_widgets()
        
    def create_widgets(self):
        # 메인 스크롤 프레임 설정
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

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
        ttk.Label(frame_sim, text="📋 투표소별 설정 (수정할 항목을 더블클릭하세요)", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        
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
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

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
            current_vals = self.tree.item(item_id)['values']
            st_name = current_vals[0]
            self.tree.item(item_id, values=(st_name, current_vals[1], current_vals[2], rate))
            
            if st_name in self.station_data:
                self.station_data[st_name]['rate'] = rate

    def select_vote_files(self):
        files = filedialog.askopenfilenames(title="투표 데이터 선택", filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv")])
        if files:
            self.vote_files = files
            self.lbl_file_count.config(text=f"✅ {len(files)}개 파일 로드됨", foreground="blue")
            self.log(f"{len(files)}개 파일 선택됨. 투표소 목록 스캔 시작...")
            self.scan_stations() 

    def select_equip_file(self):
        file = filedialog.askopenfilename(title="장비현황 파일 선택", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.equipment_file = file
            self.lbl_equip_status.config(text=f"✅ {os.path.basename(file)}", foreground="blue")
            self.log(f"장비 파일 로드됨. 목록 업데이트 중...")
            self.scan_stations() 

    def get_column_config(self):
        return { "equip_cols_idx": [0, 7, 8] }

    def scan_stations(self):
        if not self.vote_files:
            return

        station_set = set()
        
        for file in self.vote_files:
            try:
                _, _, header_row = self.get_file_info_header(file)
                
                if file.endswith('.csv'):
                    try: df = pd.read_csv(file, header=header_row, encoding='cp949')
                    except: df = pd.read_csv(file, header=header_row, encoding='utf-8')
                else:
                    df = pd.read_excel(file, header=header_row)
                
                if '사전투표소명' in df.columns:
                    df = df.dropna(subset=['사전투표소명'])
                    
                    if '읍면동명' in df.columns:
                        temp_col = df['읍면동명'].astype(str).str.replace(' ', '')
                        mask = temp_col.str.contains('합계|소계|총계|누계', na=False)
                        df = df[~mask]

                    stations = df['사전투표소명'].unique()
                    for s in stations:
                        s_str = str(s).strip()
                        if s_str and s_str != 'nan':
                            station_set.add(s_str)
                            
            except Exception as e:
                self.log(f"스캔 경고({os.path.basename(file)}): {e}")

        # 장비 파일 읽기
        equip_map = {}
        if self.equipment_file:
            try:
                df_eq = pd.read_excel(self.equipment_file)
                df_eq.columns = [str(c).replace(" ", "").strip() for c in df_eq.columns]
                
                name_col, intra_col, extra_col = None, None, None
                for col in df_eq.columns:
                    if '투표소' in col or '읍면동' in col: name_col = col
                    if '관내' in col and '수' in col: intra_col = col
                    if '관외' in col and '수' in col: extra_col = col
                
                if not (name_col and intra_col and extra_col):
                    config = self.get_column_config()
                    cols_idx = config['equip_cols_idx']
                    raw = pd.read_excel(self.equipment_file, header=None)
                    df_eq = raw.iloc[2:, cols_idx].copy()
                    df_eq.columns = ['name', 'intra', 'extra']
                    name_col, intra_col, extra_col = 'name', 'intra', 'extra'

                for _, row in df_eq.iterrows():
                    name = str(row[name_col]).strip()
                    try: intra = int(row[intra_col])
                    except: intra = 1
                    try: extra = int(row[extra_col])
                    except: extra = 1
                    equip_map[name] = {'intra': intra, 'extra': extra}
                
                self.log(f"장비 파일 인식 성공: {len(equip_map)}개")
            except Exception as e:
                self.log(f"장비 파일 읽기 오류: {e}")

        for item in self.tree.get_children():
            self.tree.delete(item)
        
        sorted_stations = sorted(list(station_set))
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
            
            self.station_data[st] = {'intra': intra, 'extra': extra, 'rate': current_global_rate}
            self.tree.insert("", "end", iid=st, values=(st, intra, extra, current_global_rate))
            
        self.log(f"목록 갱신 완료: 총 {len(sorted_stations)}개 투표소")

    def on_tree_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x) 
        
        if not item_id: return
        
        st_name = item_id
        vals = self.tree.item(item_id)['values']
        
        curr_intra = vals[1]
        curr_extra = vals[2]
        curr_rate = vals[3]
        
        if column == '#2': 
            new_intra = simpledialog.askinteger("관내 장비 수정", f"[{st_name}]\n관내 장비 수:", 
                                              initialvalue=curr_intra, minvalue=1, maxvalue=50)
            if new_intra is not None:
                self.tree.item(item_id, values=(st_name, new_intra, curr_extra, curr_rate))
                self.station_data[st_name]['intra'] = new_intra
                self.log(f"{st_name} 관내 장비 변경: {new_intra}대")
                
        elif column == '#3': 
            new_extra = simpledialog.askinteger("관외 장비 수정", f"[{st_name}]\n관외 장비 수:", 
                                              initialvalue=curr_extra, minvalue=1, maxvalue=50)
            if new_extra is not None:
                self.tree.item(item_id, values=(st_name, curr_intra, new_extra, curr_rate))
                self.station_data[st_name]['extra'] = new_extra
                self.log(f"{st_name} 관외 장비 변경: {new_extra}대")
                
        elif column == '#4': 
            new_rate = simpledialog.askinteger("증가율 수정", f"[{st_name}]\n투표자 증가율(%):", 
                                             initialvalue=curr_rate, minvalue=-100, maxvalue=200)
            if new_rate is not None:
                self.tree.item(item_id, values=(st_name, curr_intra, curr_extra, new_rate))
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

        label = "통합 분석"

        self.log(f"시뮬레이션 시작: {label}")
        all_data = []
        
        for file in self.vote_files:
            try:
                day, time, header_row = self.get_file_info_header(file)
                if day is None: continue
                
                if file.endswith('.csv'):
                    try: df = pd.read_csv(file, header=header_row, encoding='cp949')
                    except: df = pd.read_csv(file, header=header_row, encoding='utf-8')
                else:
                    df = pd.read_excel(file, header=header_row)

                if '사전투표소명' not in df.columns: continue
                
                df = df.dropna(subset=['사전투표소명'])
                
                if '읍면동명' in df.columns:
                    temp_col = df['읍면동명'].astype(str).str.replace(' ', '')
                    mask = temp_col.str.contains('합계|소계|총계|누계', na=False)
                    df = df[~mask].copy()
                
                df['사전투표소명'] = df['사전투표소명'].astype(str).str.strip()

                for col in ['관내사전투표자수', '관외사전투표자수']:
                    if col in df.columns:
                        if df[col].dtype == 'object':
                            df[col] = df[col].astype(str).str.replace(',', '').str.strip()
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                
                def apply_rate(row, col_name):
                    st = row['사전투표소명']
                    original_val = row[col_name]
                    rate = 0
                    if st in self.station_data:
                        rate = self.station_data[st]['rate']
                    factor = 1 + (rate / 100.0)
                    return original_val * factor

                df['관내사전투표자수'] = df.apply(lambda x: apply_rate(x, '관내사전투표자수'), axis=1)
                df['관외사전투표자수'] = df.apply(lambda x: apply_rate(x, '관외사전투표자수'), axis=1)
                        
                df['일차'] = day
                df['시간대'] = time
                all_data.append(df)
            except Exception as e:
                self.log(f"데이터 처리 오류({os.path.basename(file)}): {e}")

        if not all_data:
            messagebox.showerror("오류", "유효한 데이터가 없습니다. 로그를 확인해주세요.")
            return

        final_df = pd.concat(all_data, ignore_index=True)
        
        duplicates = final_df[final_df.duplicated(subset=['사전투표소명', '일차', '시간대'], keep=False)]
        if not duplicates.empty:
            problem_stations = duplicates['사전투표소명'].unique()
            messagebox.showwarning("중복 데이터 경고", f"중복 데이터가 있습니다 (같은 시간/투표소).\n파일을 중복 선택했는지 확인하세요.\n{problem_stations[:3]}...")
            final_df = final_df.drop_duplicates(subset=['사전투표소명', '일차', '시간대'])

        final_df = final_df.sort_values(by=['사전투표소명', '일차', '시간대'])
        
        final_df['시간대별_관내투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관내사전투표자수'].diff()
        final_df['시간대별_관외투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관외사전투표자수'].diff()
        
        mask_start = final_df['시간대'] == 7
        final_df.loc[mask_start, '시간대별_관내투표자수'] = final_df.loc[mask_start, '관내사전투표자수']
        final_df.loc[mask_start, '시간대별_관외투표자수'] = final_df.loc[mask_start, '관외사전투표자수']

        def get_equip_cnt(row, type_):
            st = row['사전투표소명']
            if st in self.station_data:
                return self.station_data[st][type_]
            return 1

        final_df['관내장비수'] = final_df.apply(lambda x: get_equip_cnt(x, 'intra'), axis=1)
        final_df['관외장비수'] = final_df.apply(lambda x: get_equip_cnt(x, 'extra'), axis=1)

        final_df['관내_혼잡도'] = final_df['시간대별_관내투표자수'] / final_df['관내장비수']
        final_df['관외_혼잡도'] = final_df['시간대별_관외투표자수'] / final_df['관외장비수']
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        save_name = f"시뮬레이션_결과_{timestamp}.xlsx"
        final_df.to_excel(save_name, index=False)
        self.log(f"엑셀 저장 완료: {save_name}")
        
        self.log("그래프 생성 중...")
        try:
            self.visualize_results(final_df, timestamp, label, save_name)
        except Exception as e:
            self.log(f"시각화 실패: {e}")
            messagebox.showerror("오류", str(e))

    def visualize_results(self, df, timestamp, label_text, save_name):
        system_name = platform.system()
        font_family = 'Malgun Gothic' if system_name == 'Windows' else 'AppleGothic'
        plt.rc('font', family=font_family)
        plt.rc('axes', unicode_minus=False)

        df['short_name'] = df['사전투표소명'].str.replace('사전투표소', '')
        df['label_intra'] = df['short_name'] + "(" + df['관내장비수'].astype(int).astype(str) + ")"
        df['label_extra'] = df['short_name'] + "(" + df['관외장비수'].astype(int).astype(str) + ")"

        all_scenarios = [
            (1, '관내', 'label_intra', '관내_혼잡도', self.var_day1.get() and self.var_intra.get()),
            (1, '관외', 'label_extra', '관외_혼잡도', self.var_day1.get() and self.var_extra.get()),
            (2, '관내', 'label_intra', '관내_혼잡도', self.var_day2.get() and self.var_intra.get()),
            (2, '관외', 'label_extra', '관외_혼잡도', self.var_day2.get() and self.var_extra.get())
        ]
        
        active_scenarios = [s for s in all_scenarios if s[4]]
        
        count = len(active_scenarios)
        if count == 0:
            messagebox.showwarning("알림", "옵션을 선택해주세요.")
            return

        if count == 1: nrows, ncols, figsize = 1, 1, (10, 7)
        elif count == 2: nrows, ncols, figsize = 1, 2, (18, 7)
        elif count == 3: nrows, ncols, figsize = 1, 3, (20, 7)
        else: nrows, ncols, figsize = 2, 2, (18, 14)

        fig, axes = plt.subplots(nrows, ncols, figsize=figsize)
        if count == 1: axes_flat = [axes]
        else: axes_flat = axes.flatten()
        
        # 전체 데이터 기준 최대값 계산 (스케일 통일)
        max_val = max(df['관내_혼잡도'].max(), df['관외_혼잡도'].max()) if not df.empty else 1
        
        for idx, (day, type_name, label_col, value_col, _) in enumerate(active_scenarios):
            ax = axes_flat[idx]
            df_day = df[df['일차'] == day]
            
            if df_day.empty:
                ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center')
                continue
            
            # 피벗 테이블 생성
            pivot = df_day.pivot_table(index=label_col, columns='시간대', values=value_col)
            
            # [수정] 1. 평균 열 (투표소별 평균) 추가
            pivot['평균'] = pivot.mean(axis=1)
            
            # [수정] 2. 평균 행 (시간대별 평균) 추가 (방금 추가한 평균 열도 포함하여 계산)
            avg_row = pivot.mean(axis=0)
            pivot.loc['평균'] = avg_row
            
            # [수정] 3. 컬럼 순서 재배치 ('평균'이 제일 앞으로)
            time_cols = sorted([c for c in pivot.columns if c != '평균'])
            new_cols = ['평균'] + time_cols
            pivot = pivot[new_cols]
            
            # [수정] 4. 행 순서 재배치 ('평균'이 제일 위로)
            row_labels = sorted([r for r in pivot.index if r != '평균'])
            new_rows = ['평균'] + row_labels
            pivot = pivot.reindex(new_rows)
            
            # 히트맵 그리기
            sns.heatmap(pivot, annot=True, fmt='.1f', cmap='Greens', cbar=False, linewidths=.5, vmin=0, vmax=max_val, ax=ax)
            
            # 축 설정
            ax.xaxis.tick_top()
            ax.xaxis.set_label_position('top')
            
            # [수정] 5. X축 눈금 및 라벨 설정
            # 첫 번째 셀('평균')은 가운데 정렬, 나머지(시간대)는 경계선(Line)에 맞춤
            
            # 눈금 위치: 0.5 ('평균' 중앙), 그리고 1부터 끝까지 정수 (시간대 경계선)
            ticks = [0.5] + list(range(1, len(pivot.columns) + 1))
            ax.set_xticks(ticks)
            
            # 라벨 텍스트: '평균' + 시작시간-1 부터 끝시간까지
            if time_cols:
                start_time = int(time_cols[0]) - 1 # 7시 데이터면 6시부터 시작
                end_time = int(time_cols[-1])
                labels = ['평균'] + list(range(start_time, end_time + 1))
                ax.set_xticklabels(labels, rotation=0)

            ax.set_title(f'{day}일차 {type_name} 혼잡도 (시간당 처리인원)', fontsize=14, fontweight='bold', pad=20)
            ax.set_ylabel('사전투표소(장비수)', fontsize=11, fontweight='bold')
            ax.set_xlabel('시간대', fontsize=11, fontweight='bold')

        if count == 3 and nrows * ncols > 3: axes_flat[3].axis('off')

        plt.suptitle(f"시뮬레이션 결과 - {label_text}", fontsize=20, fontweight='bold')
        plt.figtext(0.5, 0.02, "각 셀의 수치는 1시간 동안 사전투표 장비 1대당 투표용지 발급자 수를 나타냄.", ha='center', fontsize=12)
        plt.tight_layout(rect=[0, 0.05, 1, 0.95]) 
        
        img_name = f"시뮬레이션_{timestamp}.png"
        plt.savefig(img_name)
        
        messagebox.showinfo("완료", f"시뮬레이션 완료!\n\n📊 {img_name}")
        if system_name == 'Windows':
            try: os.startfile(img_name)
            except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ElectionAnalyzerApp(root)
    root.mainloop()
