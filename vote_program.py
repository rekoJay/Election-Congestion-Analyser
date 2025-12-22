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
from mpl_toolkits.axes_grid1 import make_axes_locatable
from matplotlib.colors import ListedColormap 

class ElectionAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사전투표 운용장비 산출 프로그램")
        self.root.geometry("680x920") 
        self.root.resizable(False, True) 
        
        self.vote_files = []
        self.equipment_file = None
        
        # 데이터 구조: { '투표소명': {'intra': 1, 'extra': 1, 'rate': 0.0, 'org_intra': 1, 'org_extra': 1} }
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

    def scan_stations(self):
        if not self.vote_files:
            return

        # (새로 넣을 코드)
        station_list = []  # 순서 유지를 위한 리스트
        seen = set()       # 중복 체크를 위한 집합
        
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
                            # [수정] 순서를 유지하면서 중복만 제거
                            if s_str not in seen:
                                seen.add(s_str)
                                station_list.append(s_str)
                            
            except Exception as e:
                self.log(f"스캔 경고({os.path.basename(file)}): {e}")

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

        # [추가] 1. 원본 엑셀에 등장한 투표소 순서 추출 (중복 제거하되 순서 유지)
        original_order = []
        seen = set()
        
        # 읽어들인 데이터프레임들을 순회하며 투표소 등장 순서 수집
        for temp_df in all_data:
            # 해당 파일에 있는 투표소명들 (순서 유지됨)
            stats = temp_df['사전투표소명'].unique()
            for s in stats:
                if s not in seen:
                    seen.add(s)
                    original_order.append(s)
        
        # [추가] 2. '사전투표소명' 컬럼을 단순 글자가 아니라 '순서가 있는 카테고리'로 변환
        # 이렇게 하면 나중에 sort_values를 해도 가나다순이 아니라 위에서 만든 순서대로 정렬됨
        final_df['사전투표소명'] = pd.Categorical(
            final_df['사전투표소명'], 
            categories=original_order, 
            ordered=True
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

        final_df[['관내장비수', '원본_관내장비수']] = final_df.apply(
            lambda x: pd.Series(get_equip_info(x, 'intra')), axis=1
        )
        final_df[['관외장비수', '원본_관외장비수']] = final_df.apply(
            lambda x: pd.Series(get_equip_info(x, 'extra')), axis=1
        )

        final_df['관내_혼잡도'] = final_df['시간대별_관내투표자수'] / final_df['관내장비수']
        final_df['관외_혼잡도'] = final_df['시간대별_관외투표자수'] / final_df['관외장비수']
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')

        # [수정] 현재 py 파일이 있는 '진짜' 폴더 경로 찾기
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # 1. 엑셀 저장 (경로 결합)
        excel_name = f"시뮬레이션_결과_{timestamp}.xlsx"
        full_excel_path = os.path.join(script_dir, excel_name)

        final_df.to_excel(full_excel_path, index=False)
        self.log(f"엑셀 저장 완료: {full_excel_path}")
        
        self.log("그래프 생성 중...")
        try:
            # 2. 이미지 저장 (경로 결합)
            png_name = f"시뮬레이션_{timestamp}.png"
            full_png_path = os.path.join(script_dir, png_name)

            # [핵심] visualize_results에 우리가 만든 '전체 경로'를 넘겨줌
            self.visualize_results(final_df, timestamp, full_png_path, mode='screen')
            
            messagebox.showinfo("완료", f"분석 완료!\n\n파일이 저장되었습니다:\n{full_png_path}")
            if platform.system() == 'Windows':
                try: os.startfile(full_png_path)
                except: pass
        except Exception as e:
            self.log(f"시각화 실패: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", str(e))

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
        
        # [모드 분기] 화면용(PNG) vs 인쇄용(PDF)
        if mode == 'screen':
            # === 화면용: 길게 한 장으로 ===
            self._plot_page(df, active_scenarios, unique_stations, save_name, is_pdf=False)

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
            
            # 평균행 (PDF 페이지별 평균이 아니라, 전체 평균을 보여주고 싶다면 
            # 외부에서 계산해서 넘겨야 하지만, 여기서는 "해당 페이지 내 평균"이 표기됨)
            # -> 통일성을 위해 빈 문자열로 평균행 처리
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
