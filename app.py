import streamlit as st
import pandas as pd
import re
import io

# 頁面標題與風格設定
st.set_page_config(page_title="會員抽獎資料自動比對彙整系統", layout="wide")

st.title("📊 會員銷售與抽獎資料比對匯出系統")
st.markdown("請上傳 **會員管理** 與 **會員銷售業績** 檔案，系統將自動比對並產出符合標準格式的抽獎資料。")

# 欄位清理函數：去除 Excel HTML 格式外包裝 ="..."
def clean_val(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    m = re.match(r'^="(.*)"$', val_str)
    if m:
        return m.group(1)
    return val_str

# 讀取 Excel / HTML 表格
def load_data(uploaded_file):
    try:
        # 優先嘗試以 HTML 方式讀取 (.xls 匯出檔常見格式)
        dfs = pd.read_html(uploaded_file)
        df = dfs[0]
    except Exception:
        # 若失敗則以標準 Excel 讀取
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file)
    
    # 清理所有文字欄位外加符號 (新版 pandas 使用 map，舊版使用 applymap 做相容處理)
    if hasattr(df, 'map'):
        return df.map(clean_val)
    else:
        return df.applymap(clean_val)

# 檔案上傳區塊
col1, col2 = st.columns(2)

with col1:
    file_member = st.file_uploader("1. 上傳【會員管理】檔案 (.xls/.xlsx)", type=["xls", "xlsx"])

with col2:
    file_sales = st.file_uploader("2. 上傳【會員銷售業績】檔案 (.xls/.xlsx)", type=["xls", "xlsx"])

if file_member and file_sales:
    st.divider()
    with st.spinner("資料比對與彙整中..."):
        try:
            df_m = load_data(file_member)
            df_s = load_data(file_sales)

            # 檢查關鍵欄位是否存在
            if '會員編號' not in df_m.columns or '會員編號' not in df_s.columns:
                st.error("兩份檔案均必須包含【會員編號】欄位！")
            else:
                # 取出會員管理所需欄位
                m_cols = ['會員編號', '身分證字號', '行動電話']
                available_m_cols = [c for c in m_cols if c in df_m.columns]
                
                df_m_sub = df_m[available_m_cols].drop_duplicates(subset=['會員編號'])

                # 合併主表 (銷售業績) 與 副表 (會員資訊)
                df_merged = pd.merge(df_s, df_m_sub, on='會員編號', how='left')

                # 欄位重新命名對照
                rename_dict = {
                    '身分證字號': '身份證字號',
                    '行動電話': '手機號碼'
                }
                df_merged = df_merged.rename(columns=rename_dict)

                # 抽獎資料標準欄位順序 (完全比照【抽獎資料.xlsx】)
                target_columns = [
                    '會員編號', '會員名稱', '訂購單號', '訂單日期', '收件中心',
                    '產品編號', '產品名稱', '數量', '單價', 'DP',
                    '金額小計', 'DP值小計', '身份證字號', '手機號碼'
                ]

                # 補齊可能缺失的欄位並重排順序
                for col in target_columns:
                    if col not in df_merged.columns:
                        df_merged[col] = ""

                df_final = df_merged[target_columns]

                # 顯示統計與預覽
                st.success(f"✅ 比對成功！共處理 {len(df_final)} 筆銷售明細。")
                
                st.subheader("📋 彙整結果預覽")
                # 已修正參數名稱為 use_container_width
                st.dataframe(df_final.head(100), use_container_width=True)

                # 產生 Excel 下載
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='抽獎資料')
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 下載【抽獎資料】彙整檔案 (Excel)",
                    data=excel_data,
                    file_name="抽獎資料_彙整檔.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"處理過程發生錯誤：{str(e)}")
else:
    st.info("請上傳上述兩份檔案以進行自動比對與彙整。")