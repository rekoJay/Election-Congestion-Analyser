import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

# 시각화 및 시스템 관련
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
import platform

class ElectionAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("선거 사전투표 혼잡도 분석기 (Smart Ver.)")
        self.root.geometry("620x750")
        self.root.resizable(False, False) 
        
        self.vote_files = []
        self.equipment_file = None
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True)

        # 1. 선거 유형
        frame_type = ttk.LabelFrame(main_frame, text=" 1. 선거 유형 선택 ", padding="15")
        frame_type.pack(fill="x", pady=(0, 15))
        
        self.election_type = tk.StringVar(value="president")
        
        radio_frame = ttk.Frame(frame_type)
        radio_frame.pack(fill="x")
        
        ttk.Radiobutton(radio_frame, text="대통령선거", 
                        variable=self.election_type, value="president").pack(anchor="w", pady=2)
        ttk.Radiobutton(radio_frame, text="국회의원선거", 
                        variable=self.election_type, value="general").pack(anchor="w", pady=2)
        ttk.Radiobutton(radio_frame, text="지방선거", 
                        variable=self.election_type, value="local").pack(anchor="w", pady=2)
        
        # 2. 투표 데이터 업로드
        frame_data = ttk.LabelFrame(main_frame, text=" 2. 투표 데이터 업로드 ", padding="15")
        frame_data.pack(fill="x", pady=(0, 15))
        
        btn_files = ttk.Button(frame_data, text="📂 투표 데이터 파일 선택", command=self.select_vote_files)
        btn_files.pack(fill="x", ipady=5)
        
        self.lbl_file_count = ttk.Label(frame_data, text="선택된 파일 없음", foreground="gray")
        self.lbl_file_count.pack(pady=(5, 0))
        
        # 3. 장비 현황
        frame_equip = ttk.LabelFrame(main_frame, text=" 3. 장비 현황 (선택사항) ", padding="15")
        frame_equip.pack(fill="x", pady=(0, 15))
        
        btn_frame = ttk.Frame(frame_equip)
        btn_frame.pack(fill="x")

        btn_template = ttk.Button(btn_frame, text="💾 양식 다운로드", command=self.create_template)
        btn_template.pack(side="left", fill="x", expand=True, padx=(0, 5), ipady=3)
        
        btn_equip_file = ttk.Button(btn_frame, text="📂 작성 파일 업로드", command=self.select_equip_file)
        btn_equip_file.pack(side="right", fill="x", expand=True, padx=(5, 0), ipady=3)
        
        self.lbl_equip_status = ttk.Label(frame_equip, text="파일 미선택 (기본값: 1대 적용)", foreground="gray")
        self.lbl_equip_status.pack(pady=(5, 0))
        
        # 4. 실행 버튼
        ttk.Separator(main_frame, orient="horizontal").pack(fill="x", pady=10)
        
        btn_run = ttk.Button(main_frame, text="🚀 분석 및 시각화 실행", command=self.run_analysis)
        btn_run.pack(fill="x", ipady=10, pady=5)
        
        # 5. 로그창
        log_frame = ttk.LabelFrame(main_frame, text=" 진행 상황 ", padding="10")
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=8, state='disabled', bg="#F0F0F0", relief="flat", font=("맑은 고딕", 9))
        self.log_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

    def select_vote_files(self):
        files = filedialog.askopenfilenames(title="투표 데이터 선택", filetypes=[("Excel Files", "*.xlsx *.xls *.csv")])
        if files:
            self.vote_files = files
            self.lbl_file_count.config(text=f"✅ {len(files)}개 파일 준비됨", foreground="blue")
            self.log(f"파일 {len(files)}개 선택됨.")

    def create_template(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="장비현황_양식.xlsx")
        if save_path:
            df = pd.DataFrame(columns=["사전투표소명", "관내장비수", "관외장비수"])
            df.loc[0] = ["예시: 서울종로구사전투표소", 3, 2]
            df.to_excel(save_path, index=False)
            messagebox.showinfo("완료", "양식이 저장되었습니다.")

    def select_equip_file(self):
        file = filedialog.askopenfilename(title="장비현황 파일 선택", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.equipment_file = file
            self.lbl_equip_status.config(text=f"✅ 선택됨: {os.path.basename(file)}", foreground="blue")
            self.log(f"장비 파일 로드됨: {file}")

    def get_column_config(self):
        e_type = self.election_type.get()
        if e_type == "president":
            return { "equip_cols_idx": [0, 7, 8] }
        elif e_type == "general":
            return { "equip_cols_idx": [0, 7, 8] } 
        else:
            return { "equip_cols_idx": [0, 4, 5] }

    def get_file_info(self, file_path):
        """
        파일의 상단(5줄)을 읽어서 [X일차]와 [HH:MM] 정보를 추출하고,
        데이터가 시작되는 헤더 인덱스도 함께 찾습니다.
        """
        try:
            # 파일 포맷에 따라 상단 읽기
            if file_path.endswith('.csv'):
                try:
                    df_meta = pd.read_csv(file_path, header=None, nrows=10, encoding='cp949')
                except:
                    df_meta = pd.read_csv(file_path, header=None, nrows=10, encoding='utf-8')
            else:
                df_meta = pd.read_excel(file_path, header=None, nrows=10)
            
            day = None
            time = None
            header_idx = 3 # 기본값

            # 메타데이터 스캔
            for idx, row in df_meta.iterrows():
                row_str = " ".join(row.astype(str).values)
                
                # 1. 일차/시간 찾기 (예: [1일차], [07:00])
                if day is None:
                    match_day = re.search(r'\[(\d+)일차\]', row_str)
                    match_time = re.search(r'\[(\d{1,2}):(\d{2})\]', row_str)
                    
                    if match_day:
                        day = int(match_day.group(1))
                    if match_time:
                        time = int(match_time.group(1)) # 07:00 -> 7

                # 2. 헤더 위치 찾기 ('읍면동명'이 있는 줄)
                if "읍면동명" in row_str:
                    header_idx = idx

            return day, time, header_idx

        except Exception as e:
            print(f"File Read Error: {e}")
            return None, None, 3

    def run_analysis(self):
        if not self.vote_files:
            messagebox.showwarning("주의", "투표 데이터 파일을 먼저 선택해주세요.")
            return

        e_type = self.election_type.get()
        if e_type == 'president':
            congestion_threshold = 120
            label_text = "대통령선거"
        elif e_type == 'general':
            congestion_threshold = 100
            label_text = "국회의원선거"
        else:
            congestion_threshold = 60
            label_text = "지방선거"
            
        self.log(f"분석 시작: {label_text} (기준: {congestion_threshold}명)")
        
        all_data = []
        config = self.get_column_config()
        
        success_count = 0

        for file in self.vote_files:
            try:
                # 파일 내부에서 정보 추출
                day, time, header_row = self.get_file_info(file)
                
                if day is None or time is None:
                    self.log(f"⚠️ 정보 인식 실패 (건너뜀): {os.path.basename(file)}")
                    continue
                
                # 데이터 로드
                if file.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=header_row, encoding='cp949')
                    except:
                        df = pd.read_csv(file, header=header_row, encoding='utf-8')
                else:
                    df = pd.read_excel(file, header=header_row)

                if '읍면동명' not in df.columns:
                    continue
                    
                df = df.dropna(subset=['읍면동명'])
                df = df[df['읍면동명'].astype(str).str.strip() != '합계'].copy()
                
                cols_to_fix = ['사전투표자수', '관내사전투표자수', '관외사전투표자수']
                for col in cols_to_fix:
                    if col in df.columns and df[col].dtype == 'object':
                        df[col] = df[col].astype(str).str.replace(',', '').str.strip()
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                        
                df['일차'] = day
                df['시간대'] = time
                all_data.append(df)
                success_count += 1
                
            except Exception as e:
                self.log(f"에러 발생 ({os.path.basename(file)}): {e}")
                pass

        if not all_data:
            self.log("❌ 유효한 데이터가 없습니다. 파일을 확인해주세요.")
            messagebox.showerror("오류", "데이터를 읽을 수 없습니다.\n파일 내부에 [1일차][07:00] 형식의 정보가 있는지 확인해주세요.")
            return

        self.log(f"총 {success_count}개 파일 처리 완료. 데이터 병합 중...")

        final_df = pd.concat(all_data, ignore_index=True)
        final_df = final_df.sort_values(by=['사전투표소명', '일차', '시간대'])
        
        final_df['시간대별_관내투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관내사전투표자수'].diff()
        final_df['시간대별_관외투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관외사전투표자수'].diff()
        
        mask_start = final_df['시간대'] == 7
        final_df.loc[mask_start, '시간대별_관내투표자수'] = final_df.loc[mask_start, '관내사전투표자수']
        final_df.loc[mask_start, '시간대별_관외투표자수'] = final_df.loc[mask_start, '관외사전투표자수']

        # 장비 데이터 병합
        if self.equipment_file:
            try:
                equip_df = pd.read_excel(self.equipment_file)
                if "관내장비수" in equip_df.columns:
                    equip_df = equip_df[['사전투표소명', '관내장비수', '관외장비수']]
                else:
                    cols_idx = config['equip_cols_idx']
                    equip_raw = pd.read_excel(self.equipment_file, header=None)
                    equip_df = equip_raw.iloc[2:, cols_idx].copy()
                    equip_df.columns = ['사전투표소명', '관내장비수', '관외장비수']

                equip_df['사전투표소명'] = equip_df['사전투표소명'].astype(str).str.strip()
                final_df['사전투표소명'] = final_df['사전투표소명'].astype(str).str.strip()
                final_df = pd.merge(final_df, equip_df, on='사전투표소명', how='left')
                final_df['관내장비수'] = pd.to_numeric(final_df['관내장비수'], errors='coerce').fillna(1)
                final_df['관외장비수'] = pd.to_numeric(final_df['관외장비수'], errors='coerce').fillna(1)
            except:
                final_df['관내장비수'] = 1
                final_df['관외장비수'] = 1
        else:
            final_df['관내장비수'] = 1
            final_df['관외장비수'] = 1

        final_df['관내_혼잡도'] = final_df['시간대별_관내투표자수'] / final_df['관내장비수']
        final_df['관외_혼잡도'] = final_df['시간대별_관외투표자수'] / final_df['관외장비수']
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        save_name = f"결과_{e_type}_{timestamp}.xlsx"
        final_df.to_excel(save_name, index=False)
        self.log(f"엑셀 저장 완료: {save_name}")
        
        self.log("그래프 생성 중...")
        try:
            self.visualize_results(final_df, timestamp, congestion_threshold, label_text, save_name)
        except Exception as e:
            self.log(f"시각화 실패: {e}")
            messagebox.showerror("오류", str(e))

    def visualize_results(self, df, timestamp, threshold, label_text, save_name):
        system_name = platform.system()
        font_family = 'Malgun Gothic' if system_name == 'Windows' else 'AppleGothic'
        plt.rc('font', family=font_family)
        plt.rc('axes', unicode_minus=False)

        df['short_name'] = df['사전투표소명'].str.replace('사전투표소', '')
        df['label_intra'] = df['short_name'] + "(" + df['관내장비수'].astype(int).astype(str) + ")"
        df['label_extra'] = df['short_name'] + "(" + df['관외장비수'].astype(int).astype(str) + ")"

        fig, axes = plt.subplots(2, 2, figsize=(18, 14))
        scenarios = [
            (1, '관내', 'label_intra', '관내_혼잡도', axes[0,0]),
            (1, '관외', 'label_extra', '관외_혼잡도', axes[0,1]),
            (2, '관내', 'label_intra', '관내_혼잡도', axes[1,0]),
            (2, '관외', 'label_extra', '관외_혼잡도', axes[1,1])
        ]
        
        max_val = max(df['관내_혼잡도'].max(), df['관외_혼잡도'].max()) if not df.empty else 1
        
        for day, type_name, label_col, value_col, ax in scenarios:
            df_day = df[df['일차'] == day]
            if df_day.empty:
                ax.text(0.5, 0.5, '데이터 없음', ha='center', va='center')
                continue
                
            pivot = df_day.pivot_table(index=label_col, columns='시간대', values=value_col)
            sns.heatmap(pivot, annot=True, fmt='.1f', cmap='Reds', linewidths=.5, vmin=0, vmax=max_val, ax=ax)
            ax.set_title(f'{day}일차 {type_name} 혼잡도', fontsize=14, fontweight='bold')
            
            # [수정됨] 축 제목 설정 (글자 크기 살짝 키움)
            ax.set_ylabel('사전투표소(장비수)', fontsize=11, fontweight='bold')
            
            rows, cols = pivot.shape
            for y in range(rows):
                for x in range(cols):
                    val = pivot.iloc[y, x]
                    if pd.notna(val) and val >= threshold:
                        rect = patches.Rectangle((x, y), 1, 1, linewidth=3, edgecolor='#00FF00', facecolor='none')
                        ax.add_patch(rect)

        # [수정됨] 범례 텍스트 수정
        plt.suptitle(f"사전투표 혼잡도 분석 - {label_text}\n(녹색 테두리: 혼잡도 {threshold} 이상)", fontsize=20, fontweight='bold')
        plt.tight_layout()
        
        img_name = f"시각화_{self.election_type.get()}_{timestamp}.png"
        plt.savefig(img_name)
        self.log(f"시각화 완료: {img_name}")
        
        messagebox.showinfo("완료", f"분석 끝!\n\n📄 {save_name}\n📊 {img_name}")
        
        if system_name == 'Windows':
            try:
                os.startfile(img_name)
            except:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ElectionAnalyzerApp(root)
    root.mainloop()