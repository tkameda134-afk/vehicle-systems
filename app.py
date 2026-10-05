import streamlit as st
import pandas as pd
import os
from datetime import date
from dateutil.relativedelta import relativedelta

# 💡 画面全体のワイド化と余白調整
st.set_page_config(layout="wide")
st.markdown("""
    <style>
        .block-container { padding-top: 2rem !important; }
        [data-testid="stMetricValue"] { font-size: 24px !important; }
        [data-testid="stMetricLabel"] { font-size: 14px !important; margin-bottom: -5px !important; }
        .element-container { margin-bottom: 0.5rem !important; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "vehicle_data.xlsx"
st.title("🚗 中里運送車両管理システム（2025年度）")
ADMIN_PASSWORD = "1234"

def safe_to_int(val):
    try: return int(float(str(val).strip().replace(",", "")))
    except: return 0

# 🔴【エラー完全解消の核心】古い処理パーツでも確実に動く保存ロジック（openpyxlを強制指定）
def save_to_excel(df_v, df_h):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as w:
        df_v.to_excel(w, sheet_name="Vehicle_Master", index=False)
        df_h.to_excel(w, sheet_name="Maintenance_History", index=False)

if not os.path.exists(EXCEL_FILE):
    st.error("Excelファイルが見つかりません。")
else:
    try:
        df_v = pd.read_excel(EXCEL_FILE, sheet_name="Vehicle_Master", dtype=object, engine="openpyxl")
        df_h = pd.read_excel(EXCEL_FILE, sheet_name="Maintenance_History", dtype=object, engine="openpyxl")
        df_v["Mileage"] = df_v["Mileage"].apply(safe_to_int)
        df_v["Car_Number"] = df_v["Car_Number"].astype(str).str.strip()
        df_v["Vehicle_ID"] = df_v["Vehicle_ID"].astype(str).str.strip()
        df_h["Vehicle_ID"] = df_h["Vehicle_ID"].astype(str).str.strip()

        col_left, col_right = st.columns([1, 3]) # 👈 左1：右3の黄金比で右側をワイド化
        
        if "Office" not in df_v.columns: df_v["Office"] = "千葉営業所"
        if "Type" not in df_v.columns: df_v["Type"] = "トラック"
        if "Registration_Date" not in df_v.columns: df_v["Registration_Date"] = "2020-01-01"

        # 👈 【左側エリア】：車両選択ツリー
        with col_left:
            st.subheader("🌲 車両選択")
            selected_office = st.selectbox("🏢 営業所を選択", options=["千葉営業所", "埼玉営業所"])
            type_opts = ["トラック", "トレーラー"] if selected_office == "千葉営業所" else ["トラック"]
            sel_type = st.radio("区分", type_opts, horizontal=True)
            st.write("---")
            
            df_f = df_v[(df_v["Office"].astype(str).str.strip() == selected_office) & (df_v["Type"].astype(str).str.strip() == sel_type)]
            options_list = df_f["Car_Number"].dropna().tolist()
            if not options_list:
                st.warning("対象車両がありません。"); st.stop()
                
            sel_car = st.radio("🔢 車番を選択してください", options_list)
            v_rows = df_v[df_v["Car_Number"] == sel_car]
            if v_rows.empty: v_rows = df_v[df_v["Car_Number"] == options_list]
            
            row_num = int(v_rows.index[0]) # 🔴確実な整数型で一本釣り
            
            selected_vehicle_id = str(df_v.at[row_num, "Vehicle_ID"]).strip()
            car_name_val = str(df_v.at[row_num, "Car_Name"])
            next_insp_val = str(df_v.at[row_num, "Next_Inspection"])
            mileage_val = int(df_v.at[row_num, "Mileage"])
            reg_date_val = df_v.at[row_num, "Registration_Date"]

            try:
                rd = pd.to_datetime(reg_date_val)
                delta = relativedelta(date.today(), rd.date())
                elp = f"{delta.years}年 {delta.months}ヶ月"
                r_dp = rd.strftime('%Y-%m')
            except: elp = "-"; r_dp = str(reg_date_val)

            st.write("---")
            st.subheader("🔒 管理者メニュー")
            pwd_input = st.text_input("編集用パスワードを入力", type="password")
            is_admin = (pwd_input == ADMIN_PASSWORD)
            if is_admin: st.success("🔓 編集権限が有効です")
            elif pwd_input != "": st.error("パスワードが違います")

        # 👉 【右側エリア】：メイン表示
        with col_right:
            st.subheader("🔍 車両詳細情報")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("車名", car_name_val)
            c2.metric("次回車検", str(next_insp_val)[:10])
            if sel_type == "トラック":
                c3.metric("現在の走行距離", f"{mileage_val:,} km")
                c4.metric("初度登録 (経過)", f"{r_dp} ({elp})")
            else:
                c3.write("")
                c4.metric("初度登録 (経過)", f"{r_dp} ({elp})")
            
            st.write("---")
            df_hf = df_h[df_h["Vehicle_ID"] == selected_vehicle_id]
            
            col_seibi, col_shuri = st.columns(2)
            
            with col_seibi:
                st.subheader("🔧 整備記録")
                df_s = df_hf[df_hf["Category"].astype(str).str.strip() != "修理"]
                if not df_s.empty:
                    st.dataframe(df_s[["History_ID", "Date", "Category", "Details", "Cost"]].sort_values(by="Date", ascending=False).rename(
                        columns={"History_ID": "管理ID", "Date": "実施日", "Category": "区分", "Details": "作業内容", "Cost": "費用(円)"}
                    ), use_container_width=True, height=250, hide_index=True)
                    st.info(f"🔧 整備費用 累計： {df_s['Cost'].apply(safe_to_int).sum():,} 円")
                else: st.caption("整備記録はありません。")
            
            with col_shuri:
                st.subheader("🔨 修理記録")
                df_r = df_hf[df_hf["Category"].astype(str).str.strip() == "修理"]
                if not df_r.empty:
                    st.dataframe(df_r[["History_ID", "Date", "Category", "Details", "Cost"]].sort_values(by="Date", ascending=False).rename(
                        columns={"History_ID": "管理ID", "Date": "実施日", "Category": "区分", "Details": "作業内容", "Cost": "費用(円)"}
                    ), use_container_width=True, height=250, hide_index=True)
                    st.success(f"🔨 修理費用 累計： {df_r['Cost'].apply(safe_to_int).sum():,} 円")
                else: st.caption("修理記録はありません。")
            
            st.write("---")
            if not df_hf.empty:
                total_cost = df_hf['Cost'].apply(safe_to_int).sum()
                st.warning(f"💰 総メンテナンス費用 (整備 ＋ 修理)： **{total_cost:,} 円**")
            
            if is_admin:
                st.write("---")
                with st.expander("🔄 過去の記録を修正・削除する"):
                    if not df_hf.empty:
                        target_id = safe_to_int(st.selectbox("修正対象の管理ID：", options=df_hf["History_ID"].tolist()))
                        t_rows = df_h[df_h["History_ID"].apply(safe_to_int) == target_id]
                        if not t_rows.empty:
                            t_row = t_rows.iloc[0]
                            ec1, ec2 = st.columns(2)
                            try: edit_date = pd.to_datetime(t_row["Date"])
                            except: edit_date = date.today()
                            cats = ["オイル交換", "定期点検", "車検", "修理", "その他"]
                            edit_cat = ec1.selectbox("修正区分", cats, index=cats.index(str(t_row["Category"]).strip()) if str(t_row["Category"]).strip() in cats else 0)
                            edit_cost = ec2.text_input("修正費用", value=str(t_row["Cost"]))
                            edit_det = ec2.text_area("修正詳細・作業内容", value=str(t_row["Details"]))
                            
                            b1, b2 = st.columns(2)
                            if b1.button("💾 上書き保存"):
                                df_h.loc[df_h["History_ID"].apply(safe_to_int) == target_id, ["Date", "Category", "Details", "Cost"]] = [str(edit_date), edit_cat, edit_det, safe_to_int(edit_cost)]
                                save_to_excel(df_v, df_h); st.success("修正完了"); st.rerun()
                            if b2.button("🗑️ 削除する"):
                                df_h = df_h[df_h["History_ID"].apply(safe_to_int) != target_id]
                                save_to_excel(df_v, df_h); st.warning("削除完了"); st.rerun()

                with st.expander("➕ 新しい整備・修理記録を追加する"):
                    with st.form("add_form", clear_on_submit=True):
                        fc1, fc2 = st.columns(2)
                        i_date = fc1.date_input("日付", date.today())
                        i_cat = fc1.selectbox("区分", ["オイル交換", "定期点検", "車検", "修理", "その他"])
                        i_mile = fc2.text_input("最新走行距離 (km)", value=str(mileage_val)) if sel_type == "トラック" else str(mileage_val)
                        i_cost = fc2.text_input("費用 (円)", value="0")
                        i_det = st.text_area("作業内容・詳細")
                        
                        if st.form_submit_button("🚀 記録を保存する", use_container_width=True):
                            new_id = df_h["History_ID"].apply(safe_to_int).max() + 1 if not df_h.empty else 1001
                            new_row = {"History_ID": new_id, "Vehicle_ID": selected_vehicle_id, "Date": str(i_date), "Category": i_cat, "User": "", "Details": i_det, "Cost": safe_to_int(i_cost)}
                            df_h = pd.concat([df_h, pd.DataFrame([new_row])], ignore_index=True)
                            if sel_type == "トラック": df_v.at[row_num, "Mileage"] = safe_to_int(i_mile)
                            save_to_excel(df_v, df_h); st.success("保存完了"); st.rerun()
            else:
                st.caption("🔒 データの追加・修正・削除を行うには、左メニュー最下部でパスワードを入力してください（一般社員は閲覧のみ可能です）。")
                
    except Exception as e:
        st.error("エラーが発生しました。Excelの列名を確認してください。")
        st.exception(e)