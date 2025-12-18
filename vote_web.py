import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import seaborn as sns
import platform
import io

# 한글 폰트 설정
system_name = platform.system()
font_family = 'Malgun Gothic' if system_name == 'Windows' else 'AppleGothic'
plt.rc('font', family=font_family)
plt.rc('axes', unicode_minus=False)

# 페이지 기본 설정
st.set_page_config(page_title="사전투표 혼잡도 분석기", layout="wide")

st.title("🗳️ 선거 사전투표 혼잡도 분석기 (Web Ver.)")
st.markdown("---")

# 함수 정의: 엑셀 양식 생성
def get_template_byte():
    df_temp = pd.DataFrame({
        "사전투표소명": ["예시: 서울종로구사전투표소", "예시: 00동사전투표소"],
        "관내장비수": [3, 5],
        "관외장비수": [2, 4]
    })
    buffer = io.BytesIO()
    df_temp.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer

# 1. 사이드바: 설정 및 업로드
with st.sidebar:
    st.header("1. 설정 및 업로드")
    
    # 선거 유형 선택
    e_type_label = st.radio(
        "선거 유형 선택",
        ('대통령선거', '국회의원선거', '지방선거')
    )
    
    if e_type_label == '대통령선거':
        e_type = 'president'
        threshold = 120
    elif e_type_label == '국회의원선거':
        e_type = 'general'
        threshold = 100
    else:
        e_type = 'local'
        threshold = 60
        
    st.info(f"혼잡도 기준: {threshold}명 이상 (녹색 테두리)")
    
    # 파일 업로드
    st.subheader("투표 데이터 파일")
    uploaded_files = st.file_uploader("엑셀/CSV 파일을 드래그하세요", accept_multiple_files=True, type=['xlsx', 'xls', 'csv'])
    
    st.markdown("---")
    
    # 장비 현황 (양식 다운로드 추가됨)
    st.subheader("장비 현황 파일 (선택)")
    
    # [추가된 기능] 양식 다운로드 버튼
    st.download_button(
        label="💾 장비현황 양식 다운로드 (.xlsx)",
        data=get_template_byte(),
        file_name="장비현황_양식.xlsx",
        mime="application/vnd.ms-excel",
        help="클릭하면 장비 입력을 위한 엑셀 양식을 다운로드합니다."
    )
    
    equip_file = st.file_uploader("작성한 장비 파일 업로드", type=['xlsx', 'xls'])

# 함수 정의 (파일 정보 읽기)
def get_file_info(file_obj):
    try:
        # Streamlit의 파일 객체는 바로 read 가능
        if file_obj.name.endswith('.csv'):
            try:
                df_meta = pd.read_csv(file_obj, header=None, nrows=10, encoding='cp949')
            except:
                file_obj.seek(0)
                df_meta = pd.read_csv(file_obj, header=None, nrows=10, encoding='utf-8')
        else:
            df_meta = pd.read_excel(file_obj, header=None, nrows=10)
        
        day = None
        time = None
        header_idx = 3

        for idx, row in df_meta.iterrows():
            row_str = " ".join(row.astype(str).values)
            if day is None:
                match_day = re.search(r'\[(\d+)일차\]', row_str)
                match_time = re.search(r'\[(\d{1,2}):(\d{2})\]', row_str)
                if match_day: day = int(match_day.group(1))
                if match_time: time = int(match_time.group(1))
            if "읍면동명" in row_str:
                header_idx = idx
        
        # 파일 포인터 초기화
        file_obj.seek(0)
        return day, time, header_idx
    except Exception as e:
        return None, None, 3

# 2. 메인 분석 로직
if st.button("🚀 분석 시작하기", type="primary"):
    if not uploaded_files:
        st.error("투표 데이터 파일을 먼저 업로드해주세요.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_data = []
        
        for i, file in enumerate(uploaded_files):
            status_text.text(f"분석 중... {file.name}")
            progress_bar.progress((i + 1) / len(uploaded_files))
            
            try:
                day, time, header_row = get_file_info(file)
                
                if day is None or time is None:
                    continue
                
                if file.name.endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=header_row, encoding='cp949')
                    except:
                        file.seek(0)
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
            except:
                pass

        if not all_data:
            st.error("유효한 데이터를 찾지 못했습니다.")
        else:
            # 데이터 병합 및 계산
            final_df = pd.concat(all_data, ignore_index=True)
            final_df = final_df.sort_values(by=['사전투표소명', '일차', '시간대'])
            
            final_df['시간대별_관내투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관내사전투표자수'].diff()
            final_df['시간대별_관외투표자수'] = final_df.groupby(['사전투표소명', '일차'])['관외사전투표자수'].diff()
            
            mask_start = final_df['시간대'] == 7
            final_df.loc[mask_start, '시간대별_관내투표자수'] = final_df.loc[mask_start, '관내사전투표자수']
            final_df.loc[mask_start, '시간대별_관외투표자수'] = final_df.loc[mask_start, '관외사전투표자수']

            # 장비 데이터 처리
            if equip_file:
                try:
                    equip_df = pd.read_excel(equip_file)
                    # 파일 컬럼명 유연성 확보 (사용자가 양식을 안쓰고 대충 만들었을 경우 대비)
                    if "관내장비수" in equip_df.columns:
                        equip_df = equip_df[['사전투표소명', '관내장비수', '관외장비수']]
                    else:
                        # 컬럼 이름이 다르면 첫번째 시트의 0, 1, 2번째 컬럼을 가져옴
                        equip_df = equip_df.iloc[:, [0, 1, 2]]
                        equip_df.columns = ['사전투표소명', '관내장비수', '관외장비수']
                    
                    equip_df['사전투표소명'] = equip_df['사전투표소명'].astype(str).str.strip()
                    final_df['사전투표소명'] = final_df['사전투표소명'].astype(str).str.strip()
                    final_df = pd.merge(final_df, equip_df, on='사전투표소명', how='left')
                    final_df['관내장비수'] = pd.to_numeric(final_df['관내장비수'], errors='coerce').fillna(1)
                    final_df['관외장비수'] = pd.to_numeric(final_df['관외장비수'], errors='coerce').fillna(1)
                except:
                    st.warning("장비 파일 형식이 맞지 않아 기본값(1대)으로 처리했습니다.")
                    final_df['관내장비수'] = 1
                    final_df['관외장비수'] = 1
            else:
                final_df['관내장비수'] = 1
                final_df['관외장비수'] = 1

            final_df['관내_혼잡도'] = final_df['시간대별_관내투표자수'] / final_df['관내장비수']
            final_df['관외_혼잡도'] = final_df['시간대별_관외투표자수'] / final_df['관외장비수']

            st.success("분석 완료!")
            
            # 3. 결과 탭 구성
            tab1, tab2 = st.tabs(["📊 시각화 결과", "💾 데이터 다운로드"])
            
            with tab1:
                # 시각화 로직
                final_df['short_name'] = final_df['사전투표소명'].str.replace('사전투표소', '')
                final_df['label_intra'] = final_df['short_name'] + "(" + final_df['관내장비수'].astype(int).astype(str) + ")"
                final_df['label_extra'] = final_df['short_name'] + "(" + final_df['관외장비수'].astype(int).astype(str) + ")"

                fig, axes = plt.subplots(2, 2, figsize=(18, 14))
                scenarios = [
                    (1, '관내', 'label_intra', '관내_혼잡도', axes[0,0]),
                    (1, '관외', 'label_extra', '관외_혼잡도', axes[0,1]),
                    (2, '관내', 'label_intra', '관내_혼잡도', axes[1,0]),
                    (2, '관외', 'label_extra', '관외_혼잡도', axes[1,1])
                ]
                
                max_val = max(final_df['관내_혼잡도'].max(), final_df['관외_혼잡도'].max()) if not final_df.empty else 1
                
                for day, type_name, label_col, value_col, ax in scenarios:
                    df_day = final_df[final_df['일차'] == day]
                    if df_day.empty: continue
                        
                    pivot = df_day.pivot_table(index=label_col, columns='시간대', values=value_col)
                    sns.heatmap(pivot, annot=True, fmt='.1f', cmap='Reds', linewidths=.5, vmin=0, vmax=max_val, ax=ax)
                    ax.set_title(f'{day}일차 {type_name} 혼잡도', fontsize=14, fontweight='bold')
                    ax.set_ylabel('사전투표소(장비수)', fontsize=11, fontweight='bold')
                    
                    rows, cols = pivot.shape
                    for y in range(rows):
                        for x in range(cols):
                            val = pivot.iloc[y, x]
                            if pd.notna(val) and val >= threshold:
                                rect = patches.Rectangle((x, y), 1, 1, linewidth=3, edgecolor='#00FF00', facecolor='none')
                                ax.add_patch(rect)

                plt.suptitle(f"사전투표 혼잡도 분석 - {e_type_label}\n(녹색 테두리: 혼잡도 {threshold} 이상)", fontsize=20, fontweight='bold')
                plt.tight_layout()
                st.pyplot(fig)

            with tab2:
                st.dataframe(final_df) 
                
                # 엑셀 다운로드 버튼
                buffer = io.BytesIO()
                final_df.to_excel(buffer, index=False)
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=buffer.getvalue(),
                    file_name=f"분석결과_{e_type}.xlsx",
                    mime="application/vnd.ms-excel"
                )
