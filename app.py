import streamlit as st
import pandas as pd
import os
from datetime import date
from dateutil.relativedelta import relativedelta

st.set_page_config(layout="wide")
st.markdown("<style>.block-container{padding-top:1.5rem!important;}[data-testid='stMetricValue']{font-size:22px!important;}[data-testid='stMetricLabel']{font-size:13px!important;margin-bottom:-5px!important;}.element-container{margin-bottom:0.4rem!important;}</style>", unsafe_allow_html=True)

EXCEL_FILE = "vehicle_data.xlsx"
st.title("🚗 中里運送車両管理システム")
ADMIN_PASSWORD = "1234"

def safe_to_int(val):
    try: return int(float(str(val).strip().replace(",", "")))
    except: return 0

def save_to_excel(df_v, df_h):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as w:
        df_v.to_excel(w, sheet_name="Vehicle_Master", index=False)
        df_h.to_excel(w, sheet_name="Maintenance_History", index=False)

def get_fiscal_year(d):
    try:
        dt = pd.to_datetime(d)
        if pd.isna(dt): return ""
        return f"{dt.year - 1 if dt.month < 4 else dt.year}年度"
    except: return ""

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

        col_left, col_right = st.columns([1, 4.5])
        if "Office" not in df_v.columns: df_v["Office"] = "千葉営業所"
        if "Type" not in df_v.columns: df_v["Type"] = "トラック"
        if "Registration_Date" not in df_v.columns: df_v["Registration_Date"] = "2020-01-01"

        with col_left:
            st.subheader("🌲 車両選択")
            selected_office = st.selectbox("🏢 営業所", ["千葉営業所", "埼玉営業所"])
            type_opts = ["トラック", "トレーラー"] if selected_office == "千葉営業所" else ["トラック"]
            sel_type = st.radio("区分", type_opts, horizontal=True)
            st.write("---")
            df_f = df_v[(df_v["Office"].astype(str).str.strip() == selected_office) & (df_v["Type"].astype(str).str.strip() == sel_type)]
            options_list = df_f["Car_Number"].dropna().tolist()
            if not options_list: st.warning("対象車両なし"); st.stop()
            sel_car = st.radio("🔢 車番一覧", options_list)
            v_rows = df_v[df_v["Car_Number"] == sel_car]
            if v_rows.empty: v_rows = df_v[df_v["Car_Number"] == options_list]
            row_num = int(v_rows.index[0])
            selected_vehicle_id = str(df_v.at[row_num, "Vehicle_ID"]).strip()
            car_name_val = str(df_v.at[row_num, "Car_Name"])
            next_insp_val = str(df_v.at[row_num, "Next_Inspection"])
            mileage_val = int(df_v.at[row_num, "Mileage"])
            reg_date_val = df_v.at[row_num, "Registration_Date"]
            try:
                rd = pd.to_datetime(reg_date_val)
                delta = relativedelta(date.today(), rd.date())
                elp = f"{delta.years}年{delta.months}ヶ月"
                r_dp = rd.strftime('%Y-%m')
            except: elp = "-"; r_dp = str(reg_date_val)
            st.write("---")
            pwd_input = st.text_input("🔒 編集パスワード", type="password")
            is_admin = (pwd_input == ADMIN_PASSWORD)
            if is_admin: st.success("🔓 編集可")
            elif pwd_input != "": st.error("パスワード違い")
        with col_right:
            st.subheader("🔍 車両詳細情報")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("車名", car_name_val)
            c2.metric("次回車検", str(next_insp_val)[:10])
            if sel_type == "トラック":
                c3.metric("走行距離", f"{mileage_val:,} km")
                c4.metric("初度登録 (経過)", f"{r_dp} ({elp})")
            else:
                c3.write("")
                c4.metric("初度登録 (経過)", f"{r_dp} ({elp})")
            st.write("---")
            sel_fy = st.radio("📅 表示年度", ["2026年度", "2025年度", "全期間"], horizontal=True)
            df_hf = df_h[df_h["Vehicle_ID"] == selected_vehicle_id].copy()
            if not df_hf.empty and sel_fy != "全期間":
                df_hf["FY"] = df_hf["Date"].apply(get_fiscal_year)
                df_hf = df_hf[df_hf["FY"] == sel_fy]
            cat_s = df_hf["Category"].astype(str).str.strip() if not df_hf.empty else pd.Series(dtype=str)
            col_s, col_r, col_t = st.columns(3)
            rename_dict = {"History_ID": "ID", "Date": "日付", "Category": "区分", "Details": "内容", "Cost": "費用"}
            with col_s:
                st.subheader("🔧 整備記録")
                df_s = df_hf[~cat_s.isin(["修理", "タイヤ購入"])] if not df_hf.empty else pd.DataFrame()
                if not df_s.empty:
                    st.dataframe(df_s[["History_ID", "Date", "Category", "Details", "Cost"]].sort_values("Date", ascending=False).rename(columns=rename_dict), use_container_width=True, height=230, hide_index=True)
                    st.info(f"🔧 整備累計 ({sel_fy})： **{df_s['Cost'].apply(safe_to_int).sum():,} 円**")
                else: st.caption("記録なし")
            with col_r:
                st.subheader("🔨 修理記録")
                df_r = df_hf[cat_s == "修理"] if not df_hf.empty else pd.DataFrame()
                if not df_r.empty:
                    st.dataframe(df_r[["History_ID", "Date", "Details", "Cost"]].sort_values("Date", ascending=False).rename(columns=rename_dict), use_container_width=True, height=230, hide_index=True)
                    st.success(f"🔨 修理累計 ({sel_fy})： **{df_r['Cost'].apply(safe_to_int).sum():,} 円**")
                else: st.caption("記録なし")
            with col_t:
                st.subheader("🛞 タイヤ購入記録")
                df_t = df_hf[cat_s == "タイヤ購入"] if not df_hf.empty else pd.DataFrame()
                if not df_t.empty:
                    st.dataframe(df_t[["History_ID", "Date", "Details", "Cost"]].sort_values("Date", ascending=False).rename(columns=rename_dict), use_container_width=True, height=230, hide_index=True)
                    st.info(f"🛞 タイヤ累計 ({sel_fy})： **{df_t['Cost'].apply(safe_to_int).sum():,} 円**")
                else: st.caption("記録なし")
            st.write("---")
            total_cost = df_hf['Cost'].apply(safe_to_int).sum() if not df_hf.empty else 0
            st.warning(f"💰 総メンテナンス費用【{sel_fy}】 (整備 ＋ 修理 ＋ タイヤ)： **{total_cost:,} 円**")
            if is_admin:
                st.write("---")
                cats = ["オイル交換", "定期点検", "車検", "修理", "タイヤ購入", "その他"]
                df_h_all = df_h[df_h["Vehicle_ID"] == selected_vehicle_id]
                with st.expander("🔄 過去の記録を修正・削除する"):
                    if not df_h_all.empty:
                        target_id = safe_to_int(st.selectbox("修正ID：", df_h_all["History_ID"].tolist()))
                        t_rows = df_h[df_h["History_ID"].apply(safe_to_int) == target_id]
                        if not t_rows.empty:
                            t_row = t_rows.iloc[0]
                            ec1, ec2 = st.columns(2)
                            try: edit_date = ec1.date_input("修正日", pd.to_datetime(t_row["Date"]))
                            except: edit_date = date.today()
                            c_cat = str(t_row["Category"]).strip()
                            edit_cat = ec1.selectbox("修正区分", cats, index=cats.index(c_cat) if c_cat in cats else 0)
                            edit_cost = ec2.text_input("修正費用", value=str(t_row["Cost"]))
                            edit_det = ec2.text_area("修正内容", value=str(t_row["Details"]))
                            b1, b2 = st.columns(2)
                            if b1.button("💾 上書き保存"):
                                df_h.loc[df_h["History_ID"].apply(safe_to_int) == target_id, ["Date", "Category", "Details", "Cost"]] = [str(edit_date), edit_cat, edit_det, safe_to_int(edit_cost)]
                                save_to_excel(df_v, df_h); st.success("修正完了"); st.rerun()
                            if b2.button("🗑️ 削除する"):
                                df_h = df_h[df_h["History_ID"].apply(safe_to_int) != target_id]
                                save_to_excel(df_v, df_h); st.warning("削除完了"); st.rerun()
                with st.expander("➕ 新しい記録を追加する"):
                    with st.form("add_form", clear_on_submit=True):
                        fc1, fc2 = st.columns(2)
                        i_date = fc1.date_input("日付", date.today())
                        i_cat = fc1.selectbox("区分", cats)
                        i_mile = fc2.text_input("走行距離 (km)", value=str(mileage_val)) if sel_type == "トラック" else str(mileage_val)
                        i_cost = fc2.text_input("費用 (円)", value="0")
                        i_det = st.text_area("作業内容・購入タイヤ詳細")
                        if st.form_submit_button("🚀 記録を保存する", use_container_width=True):
                            new_id = df_h["History_ID"].apply(safe_to_int).max() + 1 if not df_h.empty else 1001
                            new_row = {"History_ID": new_id, "Vehicle_ID": selected_vehicle_id, "Date": str(i_date), "Category": i_cat, "User": "", "Details": i_det, "Cost": safe_to_int(i_cost)}
                            df_h = pd.concat([df_h, pd.DataFrame([new_row])], ignore_index=True)
                            if sel_type == "トラック": df_v.at[row_num, "Mileage"] = safe_to_int(i_mile)
                            save_to_excel(df_v, df_h); st.success("保存完了"); st.rerun()
                with open(EXCEL_FILE, "rb") as f:
                    st.download_button("📥 最新Excelをダウンロード", f, file_name=f"vehicle_data_{date.today().strftime('%Y%m%d')}.xlsx", use_container_width=True)
    except Exception as e:
        st.error("エラーが発生しました。")
        st.exception(e)
