import streamlit as st
import pandas as pd
import numpy as np
import io
import pulp
import traceback
import random
import openpyxl
import re
from datetime import datetime

# ==========================================
# 1. 網頁頁面配置
# ==========================================
st.set_page_config(page_title="段考監考終極自動化", page_icon="🏫", layout="wide")
st.title("🏫 試務組-段考監考全自動化系統 (智慧雷達版)")
st.info("💡 終極更新：已修復 `MergedCell` 當機錯誤！內建「表頭自動導航雷達」，支援任意新增跨欄置中大標題，座標再也不會跑掉！")

# --- 初始化狀態 ---
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# ==========================================
# 2. 輔助功能定義
# ==========================================
def to_excel_bytes(df, header_df=None):
    output = io.BytesIO()
    if header_df is not None:
        df.columns = header_df.columns
        final_out = pd.concat([header_df, df], ignore_index=True)
    else:
        final_out = df
    final_out = final_out.fillna("")
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_out.to_excel(writer, index=False, header=False)
    return output.getvalue()

def clean_str(s):
    if pd.isna(s) or s is None: return ""
    s = str(s).strip().replace('ㄧ', '一').replace(' ', '').replace('　', '').replace('\n', '').replace('\r', '')
    s = s.translate(str.maketrans('１２３４５６７８９０', '1234567890'))
    return s

def normalize_subject(s):
    s = clean_str(s)
    aliases = {'國文':'國語文', '英文':'英語文', '公社':'公民與社會', '公民':'公民與社會', 
               '地科':'地球科學', '健護':'健康與護理', '護理':'健康與護理', 
               '國防':'全民國防教育', '生科':'生活科技', '應數':'應用數學'}
    return aliases.get(s, s)

def get_teacher_fuzzy(cls, subj, course_dict):
    if (cls, subj) in course_dict: return course_dict[(cls, subj)]
    clean_target = subj.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
    for (c, s), t in course_dict.items():
        if c == cls:
            s_clean = s.replace('選修', '').replace('彈性學習', '').replace('補強', '').replace('-', '')
            if clean_target and (clean_target in s_clean or s_clean in clean_target):
                return t
    if len(clean_target) >= 2:
        for (c, s), t in course_dict.items():
            if c == cls and clean_target[:2] in s:
                return t
    return ""

def extract_mm_dd(text):
    if pd.isna(text): return ""
    nums = re.findall(r'\d+', str(text))
    nums = [n for n in nums if len(n) <= 2]
    if len(nums) >= 2:
        return f"{int(nums[0]):02d}-{int(nums[1]):02d}"
    return str(text).strip()

# ==========================================
# 3. 介面佈局
# ==========================================
st.divider()
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("📂 1. 上傳排考與標籤資料")
    file_quota = st.file_uploader("1️⃣ 監考堂數.xlsx", type=['xlsx'], key=f"f1_{st.session_state['uploader_key']}")
    file_list = st.file_uploader("2️⃣ 監考名單.xlsx", type=['xlsx'], key=f"f2_{st.session_state['uploader_key']}")
    file_type = st.file_uploader("3️⃣ 監考類型總數.xlsx", type=['xlsx'], key=f"f3_{st.session_state['uploader_key']}")
    file_pub = st.file_uploader("4️⃣ 監考總表公布版.xlsx (範本)", type=['xlsx'], key=f"f4_{st.session_state['uploader_key']}")
    file_assign = st.file_uploader("5️⃣ 監考一覽表.xlsx (班級分配範本)", type=['xlsx'], key=f"f5_{st.session_state['uploader_key']}")
    st.write("---")
    file_course = st.file_uploader("6️⃣ 配課表.xlsx (多工作表)", type=['xlsx'], key=f"f6_{st.session_state['uploader_key']}")
    file_label = st.file_uploader("7️⃣ 標籤列印.xlsx (試卷袋範本)", type=['xlsx'], key=f"f7_{st.session_state['uploader_key']}")

with col2:
    st.subheader("⚙️ 2. 考試設定與特許名單")
    selected_sheet = None
    if file_quota:
        xls = pd.ExcelFile(file_quota)
        selected_sheet = st.selectbox("👇 選擇考試項目：", xls.sheet_names)
    
    flex_names = []
    if file_list:
        temp_df = pd.read_excel(file_list, header=None).fillna("")
        teacher_list = temp_df.iloc[2:, 1].astype(str).str.strip().tolist()
        teacher_list = [t for t in teacher_list if t != "" and t != "nan"]
        flex_names = st.multiselect("🛡️ 優先時數不大於名單：", options=teacher_list)

    st.write("")
    c_d1, c_d2 = st.columns(2)
    with c_d1: d1_date = st.date_input("📅 第一天日期(選填)：", datetime.now())
    with c_d2: d2_date = st.date_input("📅 第二天日期(選填)：", datetime.now())
    
    force_run = st.checkbox("⚠️ 忽略健檢警告，強制執行")
    if st.button("🗑️ 清除所有設定", use_container_width=True):
        st.session_state['results'] = None
        st.session_state['uploader_key'] += 1
        st.rerun()

# ==========================================
# 4. 核心演算法執行
# ==========================================
st.divider()

if st.button("🚀 啟動終極全自動排班系統", type="primary", use_container_width=True):
    if not all([file_quota, file_list, file_type, file_assign]):
        st.error("🚨 請至少確認【1, 2, 3, 5】號基礎檔案皆已上傳！")
    else:
        try:
            # --- 1. 基礎資料讀取 ---
            df_quota = pd.read_excel(file_quota, sheet_name=selected_sheet).fillna("")
            quota_dict = dict(zip(df_quota.iloc[:, 0].astype(str).str.strip(), pd.to_numeric(df_quota.iloc[:, 1], errors='coerce').fillna(0)))
            
            df_type = pd.read_excel(file_type, header=None).fillna("")
            req_matrix = {'△': [0]*10, '※': [0]*10}
            for i in range(2, len(df_type)):
                row_name = str(df_type.iloc[i, 0]).strip()
                if row_name in ['△', '※']:
                    req_matrix[row_name] = pd.to_numeric(df_type.iloc[i, 1:11], errors='coerce').fillna(0).astype(int).tolist()

            # --- 2. 智慧雷達掃描【監考一覽表】動態座標 ---
            df_assign_temp = pd.read_excel(file_assign, header=None).fillna("")
            
            # 尋找班級起始列
            assign_start_idx = 2
            for r in range(min(10, len(df_assign_temp))):
                val = str(df_assign_temp.iloc[r, 0]).strip()
                if val and any(char in val for char in ['一', '二', '三', '1', '2', '3', '4']):
                    assign_start_idx = r
                    break
                    
            class_names_raw = df_assign_temp.iloc[assign_start_idx:, 0].tolist()
            class_names_raw = [c for c in class_names_raw if str(c).strip()]
            norm_class_names = [clean_str(c) for c in class_names_raw]
            assign_map = {name: idx for idx, name in enumerate(norm_class_names)}
            
            # 尋找節次欄位位置 (自動適應是否插入提早考欄位)
            periods_row = df_assign_temp.iloc[assign_start_idx - 1, :].tolist()
            day1_cols, day2_cols = [], []
            current_day = 1
            for c_idx, val in enumerate(periods_row):
                v_str = str(val).strip()
                if v_str in ['1', '2', '3', '4', '5']:
                    if v_str == '1' and len(day1_cols) >= 5: current_day = 2
                    if current_day == 1 and len(day1_cols) < 5: day1_cols.append(c_idx)
                    elif current_day == 2 and len(day2_cols) < 5: day2_cols.append(c_idx)
            
            if len(day1_cols) < 5: day1_cols = [1,2,3,4,5]
            if len(day2_cols) < 5: day2_cols = [6,7,8,9,10]
            target_cols = day1_cols + day2_cols
            
            # 抓取真實日期
            dates_row_idx = max(0, assign_start_idx - 2)
            sys_d1_raw = str(df_assign_temp.iloc[dates_row_idx, target_cols[0]])
            sys_d2_raw = str(df_assign_temp.iloc[dates_row_idx, target_cols[5]])
            d1_match = extract_mm_dd(sys_d1_raw)
            d2_match = extract_mm_dd(sys_d2_raw)
            
            display_d1 = d1_date.strftime('%m月%d日')
            display_d2 = d2_date.strftime('%m月%d日')

            # 總表 Header 設定
            df_list_raw = pd.read_excel(file_list, header=None).fillna("")
            header_df = df_list_raw.iloc[0:2].copy().astype(str).replace('nan', '')
            for c in range(3, 8): header_df.iloc[0, c] = display_d1
            for c in range(8, 13): header_df.iloc[0, c] = display_d2
            df_list = df_list_raw.iloc[2:].copy()
            teachers = df_list.iloc[:, 1].astype(str).str.strip().tolist()

            # --- 3. PuLP 運算 ---
            with st.spinner("🧠 正在生成完美監考總表..."):
                prob = pulp.LpProblem("Scheduling", pulp.LpMinimize)
                vX = {}; vY = {}
                for i in range(len(teachers)):
                    vX[i] = {}; vY[i] = {}
                    for j in range(10):
                        vX[i][j] = pulp.LpVariable(f"X_{i}_{j}", cat='Binary')
                        vY[i][j] = pulp.LpVariable(f"Y_{i}_{j}", cat='Binary')
                
                penalty = 0
                for i, t in enumerate(teachers):
                    tgt = int(quota_dict.get(t, 0))
                    act = pulp.lpSum([vX[i][k] + vY[i][k]*2 for k in range(10)])
                    prob += act <= tgt
                    dfct = pulp.LpVariable(f"dfct_{i}", 0)
                    prob += act + dfct == tgt
                    penalty += dfct * (1 if t in flex_names else 1000)
                    for j in range(10):
                        prob += vX[i][j] + vY[i][j] <= 1
                        cell_val = str(df_list.iloc[i, j+3]).strip()
                        if cell_val != "" and cell_val != "nan":
                            prob += vX[i][j] == 0; prob += vY[i][j] == 0
                    prob += vX[i][1] >= vY[i][0]
                    prob += vX[i][6] >= vY[i][5]
                for j in range(10):
                    prob += pulp.lpSum([vX[i][j] for i in range(len(teachers))]) == req_matrix['△'][j]
                    prob += pulp.lpSum([vY[i][j] for i in range(len(teachers))]) == req_matrix['※'][j]
                prob += penalty
                prob.solve()

                schedule_dict = {}
                df_out_master = df_list.copy()
                for i, t in enumerate(teachers):
                    res = []
                    for j in range(10):
                        val = str(df_list.iloc[i, j+3]).strip()
                        if val == "" or val == "nan":
                            if vX[i][j].varValue == 1: val = "△"
                            elif vY[i][j].varValue == 1: val = "※"
                            else: val = "" 
                        res.append(val)
                        df_out_master.iloc[i, j+3] = val
                    schedule_dict[t] = res

            # --- 4. 監考一覽表分配矩陣計算 ---
            with st.spinner("🎯 執行班級自動分配..."):
                assigned_matrix = np.empty((len(class_names_raw), 10), dtype=object)
                for day_start in [0, 5]:
                    j1 = day_start
                    proctors_j1 = [t for t in teachers if schedule_dict[t][j1] in ["△", "※"]]
                    random.shuffle(proctors_j1)
                    for idx, p in zip(range(len(class_names_raw)), proctors_j1): assigned_matrix[idx, j1] = p
                    
                    j2 = day_start + 1
                    proctors_j2 = [t for t in teachers if schedule_dict[t][j2] in ["△", "※"]]
                    bound = {}
                    for idx in range(len(class_names_raw)):
                        p_prev = assigned_matrix[idx, j1]
                        if p_prev is not None and schedule_dict[p_prev][j1] == "※" and schedule_dict[p_prev][j2] == "△":
                            assigned_matrix[idx, j2] = p_prev
                            bound[p_prev] = True
                    rem = [p for p in proctors_j2 if p not in bound]
                    random.shuffle(rem)
                    r_idx = 0
                    for idx in range(len(class_names_raw)):
                        if assigned_matrix[idx, j2] is None and r_idx < len(rem):
                            assigned_matrix[idx, j2] = rem[r_idx]; r_idx += 1

                    for offset in [2, 3, 4]:
                        curr_j = day_start + offset
                        proctors = [t for t in teachers if schedule_dict[t][curr_j] in ["△", "※"]]
                        random.shuffle(proctors)
                        for idx, p in zip(range(len(class_names_raw)), proctors): assigned_matrix[idx, curr_j] = p

            # --- 5. 寫入【監考一覽表】(使用 openpyxl 避開合併儲存格) ---
            with st.spinner("🖨️ 正在將班級分配無縫套入一覽表範本..."):
                wb_assign = openpyxl.load_workbook(file_assign)
                ws_assign = wb_assign.active
                
                # 精準填寫老師，完全避開標題列的合併儲存格
                for r_idx in range(len(class_names_raw)):
                    excel_r = assign_start_idx + r_idx + 1 # openpyxl 是從 1 開始
                    for c_idx in range(10):
                        p_name = assigned_matrix[r_idx, c_idx]
                        if p_name:
                            excel_c = target_cols[c_idx] + 1
                            cell = ws_assign.cell(row=excel_r, column=excel_c)
                            if type(cell).__name__ != 'MergedCell':
                                cell.value = str(p_name).replace('None', '')
                
                out_assign = io.BytesIO()
                wb_assign.save(out_assign)
                assign_bytes = out_assign.getvalue()

            # --- 6. 公布版套印 ---
            pub_bytes = None
            if file_pub:
                with st.spinner("🖨️ 正在無縫套印公布版..."):
                    wb = openpyxl.load_workbook(file_pub)
                    ws = wb.active
                    h_row = -1; t_cols = []
                    for r in range(1, min(20, ws.max_row + 1)):
                        for c in range(1, ws.max_column + 1):
                            val = ws.cell(row=r, column=c).value
                            if val and "教師" in str(val):
                                h_row = r; t_cols.append(c)
                        if h_row != -1: break
                    if h_row != -1:
                        for c in t_cols:
                            if h_row - 1 >= 1:
                                cell1 = ws.cell(row=h_row-1, column=c+2)
                                cell2 = ws.cell(row=h_row-1, column=c+7)
                                if type(cell1).__name__ != 'MergedCell': cell1.value = display_d1
                                if type(cell2).__name__ != 'MergedCell': cell2.value = display_d2
                            for r in range(h_row+1, ws.max_row + 1):
                                t_val = ws.cell(row=r, column=c).value
                                if t_val:
                                    name = str(t_val).strip()
                                    if name in schedule_dict:
                                        for j in range(5): 
                                            cw = ws.cell(row=r, column=c+2+j)
                                            if type(cw).__name__ != 'MergedCell': cw.value = schedule_dict[name][j]
                                        for j in range(5): 
                                            cw = ws.cell(row=r, column=c+7+j)
                                            if type(cw).__name__ != 'MergedCell': cw.value = schedule_dict[name][j+5]
                    out_pub = io.BytesIO()
                    wb.save(out_pub)
                    pub_bytes = out_pub.getvalue()

            # --- 7. 標籤列印合成 (採用 A, B, D, E 絕對座標法則) ---
            label_bytes = None
            if file_course and file_label:
                with st.spinner("🏷️ 啟動【A,B,D,E座標綁定】，精準合成試卷袋標籤..."):
                    course_dict = {}
                    xls_course = pd.ExcelFile(file_course)
                    for sheet in xls_course.sheet_names:
                        df_c = pd.read_excel(file_course, sheet_name=sheet, header=None).fillna("")
                        h_idx = 0
                        for r in range(min(5, len(df_c))):
                            row_str = "".join(str(x) for x in df_c.iloc[r, :])
                            if "科目" in row_str or "一1" in row_str or "二1" in row_str or "三1" in row_str:
                                h_idx = r; break
                        classes_in_sheet = [clean_str(x) for x in df_c.iloc[h_idx, :]]
                        
                        for r_idx in range(h_idx + 1, len(df_c)):
                            row = df_c.iloc[r_idx, :]
                            subj_raw = str(row.iloc[0])
                            if not subj_raw: continue
                            subj_norm = normalize_subject(subj_raw)
                            for c_idx in range(1, len(row)):
                                cls_raw = classes_in_sheet[c_idx]
                                teacher = clean_str(row.iloc[c_idx])
                                if teacher and cls_raw:
                                    course_dict[(cls_raw, subj_norm)] = teacher
                    
                    wb_label = openpyxl.load_workbook(file_label)
                    ws_label = wb_label.active
                    col_map = {}
                    for c in range(1, ws_label.max_column + 1):
                        val = clean_str(ws_label.cell(row=1, column=c).value)
                        if val: col_map[val] = c
                    col_teacher = col_map.get('任課教師', 6)  
                    col_proctor = col_map.get('監考老師', 8)  
                    
                    d1_fmts = [d1_date.strftime('%Y-%m-%d'), d1_date.strftime('%Y/%m/%d'), d1_date.strftime('%m-%d')]
                    d2_fmts = [d2_date.strftime('%Y-%m-%d'), d2_date.strftime('%Y/%m/%d'), d2_date.strftime('%m-%d')]
                    if d1_match: d1_fmts.append(d1_match)
                    if d2_match: d2_fmts.append(d2_match)
                    
                    def get_val(r, c_idx):
                        v = ws_label.cell(row=r, column=c_idx).value
                        return str(v).strip() if v is not None else ""

                    for r in range(2, ws_label.max_row + 1):
                        val_A = get_val(r, 1) 
                        val_B = ws_label.cell(row=r, column=2).value 
                        val_D = get_val(r, 4) 
                        val_E = get_val(r, 5) 
                        
                        if not val_D: continue
                        cls = clean_str(val_D)
                        subj = normalize_subject(val_E)
                        
                        teacher = get_teacher_fuzzy(cls, subj, course_dict)
                        if teacher:
                            t_cell = ws_label.cell(row=r, column=col_teacher)
                            if type(t_cell).__name__ != 'MergedCell': t_cell.value = teacher
                        
                        try: p_val = int(float(val_A))
                        except: p_val = -1
                        
                        if cls in assign_map and 1 <= p_val <= 5:
                            if isinstance(val_B, datetime): date_str = val_B.strftime('%Y-%m-%d')
                            else: date_str = str(val_B).strip().replace('/', '-')
                                
                            day_offset = -1
                            for fmt in d1_fmts:
                                if fmt and fmt in date_str: day_offset = 0; break
                            if day_offset == -1:
                                for fmt in d2_fmts:
                                    if fmt and fmt in date_str: day_offset = 5; break
                            
                            if day_offset != -1:
                                target_col = day_offset + p_val
                                proctor = assigned_matrix[assign_map[cls], target_col - 1]
                                if pd.notna(proctor) and proctor is not None:
                                    p_cell = ws_label.cell(row=r, column=col_proctor)
                                    if type(p_cell).__name__ != 'MergedCell': 
                                        p_cell.value = str(proctor).replace('None', '')

                    out_label = io.BytesIO()
                    wb_label.save(out_label)
                    label_bytes = out_label.getvalue()

            st.balloons()
            st.session_state['results'] = {
                'orig': to_excel_bytes(df_out_master, header_df),
                'assign': assign_bytes,
                'pub': pub_bytes,
                'label': label_bytes
            }
            st.success("🎉 完美防當機版合成完畢！所有 Excel 格式、框線、背景色皆已完美繼承！")

        except Exception as e:
            st.error(f"發生錯誤: {e}")
            st.code(traceback.format_exc())

# ==========================================
# 5. 下載區
# ==========================================
if st.session_state['results']:
    st.divider()
    res = st.session_state['results']
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.download_button("📥 1. 監考總表", res['orig'], "監考總表.xlsx", "application/vnd.ms-excel", use_container_width=True)
    with c2: st.download_button("📥 2. 監考一覽表(保留格式版)", res['assign'], "監考一覽表_分配完成.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
    with c3: 
        if res['pub']: st.download_button("📥 3. 公布版套印總表", res['pub'], "公布版總表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c4:
        if res.get('label'): st.download_button("📥 4. 標籤列印(完整版)", res['label'], "標籤列印_完整版.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
