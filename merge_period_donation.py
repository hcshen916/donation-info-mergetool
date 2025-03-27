import streamlit as st
import pandas as pd
import re
import io
import openpyxl

st.title('定期定額捐款資料合併工具')

# 上傳檔案
official_file = st.file_uploader("請上傳官網下載-定期定額捐款紀錄", type=['xlsx', 'csv'])
newebpay_file = st.file_uploader("請上傳藍新下載-銷售紀錄查詢", type=['xlsx', 'csv'])

if official_file and newebpay_file:
    # 讀取檔案
    try:
        if official_file.name.endswith('.csv'):
            df_official = pd.read_csv(official_file)
        else:
            df_official = pd.read_excel(official_file)
            
        if newebpay_file.name.endswith('.csv'):
            df_newebpay = pd.read_csv(newebpay_file)
        else:
            df_newebpay = pd.read_excel(newebpay_file)
        
        # 檢查必要的欄位是否存在
        required_official_cols = ['委託單號', '付款時間', '委託金額', '身份', '姓名', '收據抬頭', 
                                 '收據統編或身分證號', '電話', '收據寄送地址', 'Email', 
                                 '收據選項', '指定地方黨部', '指定用途']
        
        required_newebpay_cols = ['商店訂單編號', '預計撥款日']
        
        missing_official_cols = [col for col in required_official_cols if col not in df_official.columns]
        missing_newebpay_cols = [col for col in required_newebpay_cols if col not in df_newebpay.columns]
        
        if missing_official_cols:
            st.error(f"官網檔案缺少必要欄位: {', '.join(missing_official_cols)}")
            st.write("如欄位名稱不同，請確認檔案格式或聯絡管理員。")
            st.stop()
            
        if missing_newebpay_cols:
            st.error(f"藍新檔案缺少必要欄位: {', '.join(missing_newebpay_cols)}")
            st.write("如欄位名稱不同，請確認檔案格式或聯絡管理員。")
            st.stop()
            
        # 處理藍新資料中的="xxx"格式
        for column in df_newebpay.columns:
            df_newebpay[column] = df_newebpay[column].astype(str).apply(lambda x: x.strip('="') if x.startswith('="') else x)
        
        # 建立新的DataFrame
        result_df = pd.DataFrame(columns=range(21))
        
        # 處理藍新資料的訂單編號
        df_newebpay['編號'] = df_newebpay['商店訂單編號'].apply(lambda x: x.split('_')[0])
        
        # 移除重複的委託單號，只保留第一筆
        df_official = df_official.drop_duplicates(subset=['委託單號'], keep='first')
        
        # 逐筆處理藍新資料
        for idx, newebpay_row in df_newebpay.iterrows():
            order_id = newebpay_row['編號']
            
            # 在官網資料中尋找對應的委託單號
            official_match = df_official[df_official['委託單號'] == order_id]
            
            if not official_match.empty:
                official_row = official_match.iloc[0]
                
                new_row = pd.Series(index=range(21))  # 修改為22個元素
                
                # 填入資料
                new_row[0] = official_row['付款時間']
                new_row[1] = str(order_id)  # 確保為文字格式
                new_row[2] = '金錢'
                new_row[3] = official_row['委託金額']
                new_row[4] = official_row['身份']
                
                # 處理收據抬頭
                new_row[5] = official_row['姓名'] if pd.isna(official_row['收據抬頭']) else official_row['收據抬頭']
                
                new_row[6] = official_row['收據統編或身分證號']
                
                # 處理電話格式
                phone = str(official_row['電話']).strip()
                if phone.startswith('0'):
                    new_row[7] = f"{phone[:4]}-{phone[4:]}"
                else:
                    phone = '0' + phone
                    new_row[7] = f"{phone[:4]}-{phone[4:]}"
                    
                new_row[8] = '藍新'
                new_row[9] = official_row['收據寄送地址']
                new_row[10] = official_row['收據寄送地址']
                new_row[11] = official_row['姓名']
                new_row[12] = official_row['Email']
                new_row[13] = official_row['收據選項']
                new_row[14] = '定期定額'  # 固定填入「定期定額」
                new_row[15] = ''  # 黨部留空
                new_row[16] = official_row['委託金額']
                new_row[17] = official_row['指定地方黨部']
                new_row[18] = ''  # 指定捐款留空
                new_row[19] = official_row['指定用途']
                new_row[20] = newebpay_row['預計撥款日']
                
                result_df = pd.concat([result_df, new_row.to_frame().T], ignore_index=True)
        
        # 設定欄位名稱
        column_names = ['付款時間', '編號', '捐款類別', '收入金額', '收支科目', 
                        '收據開立對象', '統一編號', '電話', '入帳方式', '收據地址',
                        '收據寄送地址', '付款人', '電子郵件', '收據與否', '捐款連結',
                        '黨部', '會計認列金額', '會計認列黨部', '指定捐款', '留言', '預計撥款日']
        result_df.columns = column_names
        
        # 下載按鈕
        if not result_df.empty:
            # 將DataFrame轉換為Excel檔案格式
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Sheet1')
            excel_data = output.getvalue()
            
            st.download_button(
                label="下載合併後的檔案",
                data=excel_data,
                file_name="定期定額-捐款資料.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning('沒有找到符合的資料')
    
    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {str(e)}")
        st.write("請確認上傳的檔案格式是否正確。如問題持續發生，請聯絡管理員。")
