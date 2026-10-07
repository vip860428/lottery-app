import streamlit as st
import pandas as pd
import re
import io

# 頁面基本設定
st.set_page_config(
    page_title="抽獎資料彙整系統",
    page_icon="⭐",
    layout="wide"
)

# 自訂 CSS 樣式：美化按鈕顏色與文字大小
st.markdown("""
<style>
    /* 上傳按鈕樣式修改為淺綠色 */
    [data-testid="stFileUploader"] button {
        background-color: #e8f5e9 !important;
        color: #2e7d32 !important;
        border: 1px solid #a5d6a7 !important;
        font-weight: bold !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #c8e6c9 !important;
        border-color: #81c784 !important;
    }
    
    /* 上傳標題文字大小設定為 16px */
    .upload-label {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #333333;
        margin-bottom: 8px;
    }
    
    /* 增加段落間距 */
    .spacer {
        height: 24px;
    }
</style>
""", unsafe_allow_html=True)

# 1. 大標題與簡介
st.title("⭐ 抽獎資料彙整系統")
st.write("請上傳會員管理與會員銷售業績檔案，系統將自動比對並統計符合標準格式的抽獎資料。")

# 間隔約 2 行寬
st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)

# 2. 上傳區域 (兩欄並排)
col1, col2 = st.columns(2)

with col1:
    st.markdown('<p class="upload-label">1. 上傳【會員管理】文件 (.xls/.xlsx)</p>', unsafe_allow_html=True)
    file_m = st.file_uploader("請選擇會員管理檔案", type=['xls', 'xlsx'], label_visibility="collapsed")

with col2:
    st.markdown('<p class="upload-label">2. 上傳【會員銷售業績】文件 (.xls/.xlsx)</p>', unsafe_allow_html=True)
    file_s = st.file_uploader("請選擇會員銷售業績檔案", type=['xls', 'xlsx'], label_visibility="collapsed")

# 間隔約 2 行寬
st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)

# 清理 Excel 匯出外包裝公式語法（如 ="TWB0000269"）
def clean_val(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    m = re.match(r'^="(.*)"$', val_str)
    if m:
        return m.group(1)
    return val_str

# 數值轉整數格式函數 (去除小數點)
def to_int_str(val):
    if pd.isna(val) or val == "":
        return ""
    try:
        # 先轉浮點數再四捨五入轉整數
        num = float(val)
        return int(round(num))
    except (ValueError, TypeError):
        return str(val)

# 讀取 Excel/HTML 工具函數
def load_data(uploaded_file):
    if uploaded_file is None:
        return None
    content = uploaded_file.read()
    try:
        tables = pd.read_html(io.BytesIO(content))
        if tables:
            return tables[0]
    except Exception:
        pass
    try:
        return pd.read_excel(io.BytesIO(content))
    except Exception as e:
        st.error(f"檔案解析失敗：{e}")
        return None

# 3. 處理與比對邏輯
if file_m and file_s:
    with st.spinner("正在自動比對並彙整資料中，請稍候..."):
        df_m = load_data(file_m)
        df_s = load_data(file_s)
        
        if df_m is not None and df_s is not None:
            # 資料清理
            if hasattr(df_m, 'map'):
                df_m = df_m.map(clean_val)
                df_s = df_s.map(clean_val)
            else:
                df_m = df_m.applymap(clean_val)
                df_s = df_s.applymap(clean_val)
            
            # 準備關聯資料表
            m_cols = ['會員編號', '身分證字號', '行動電話']
            available_m_cols = [c for c in m_cols if c in df_m.columns]
            df_m_sub = df_m[available_m_cols].drop_duplicates(subset=['會員編號'])
            
            # 以銷售業績為主表做左邊界合併
            df_merged = pd.merge(df_s, df_m_sub, on='會員編號', how='left')
            
            # 重命名欄位
            rename_dict = {
                '身分證字號': '身份證字號',
                '行動電話': '手機號碼'
            }
            df_merged = df_merged.rename(columns=rename_dict)
            
            # 目標 14 個標準欄位
            target_columns = [
                '會員編號', '會員名稱', '訂購單號', '訂單日期', '收件中心',
                '產品編號', '產品名稱', '數量', '單價', 'DP',
                '金額小計', 'DP值小計', '身份證字號', '手機號碼'
            ]
            
            for col in target_columns:
                if col not in df_merged.columns:
                    df_merged[col] = ""
            
            df_final = df_merged[target_columns].copy()
            
            # 數值欄位去除小數點處理 (包含 L 欄：DP值小計)
            int_columns = ['數量', '單價', 'DP', '金額小計', 'DP值小計']
            for col in int_columns:
                if col in df_final.columns:
                    df_final[col] = df_final[col].apply(to_int_str)
            
            st.success(f"🎉 資料比對成功！共處理 {len(df_final)} 筆記錄。")
            
            # 間隔約 2 行寬
            st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)
            
            # 提供 Excel 下載按鈕
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Sheet1")
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載比對完成 Excel 檔",
                data=excel_data,
                file_name="抽獎資料彙整結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            # 預覽前 100 筆資料
            st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)
            st.subheader("📊 資料預覽（前 100 筆）")
            st.dataframe(df_final.head(100), use_container_width=True)
else:
    st.info("請上傳上述兩份文件以進行自動比對與彙整。")
