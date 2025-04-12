import streamlit as st
import pandas as pd
import re

st.title("DSC Hour Calculator")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if "Datum" not in df.columns:
        st.error("Error: The uploaded file must contain a 'Datum' column.")
    else:
        df["Datum"] = pd.to_datetime(df["Datum"], dayfirst=True, errors='coerce')
        df = df.dropna(subset=["Datum"])

        min_date = df["Datum"].min()
        max_date = df["Datum"].max()

        start_date = st.date_input("Start Date", min_date)
        end_date = st.date_input("End Date", max_date)

        st.subheader("Set Time Values (in minutes)")
        t1 = st.number_input("T1 (gestanzt)", min_value=1, value=5)
        t2 = st.number_input("T2 (sample cutter)", min_value=1, value=10)
        t3 = st.number_input("T3 (Für die Auswertung)", min_value=1, value=15)

        if start_date > end_date:
            st.error("Error: Start Date cannot be after End Date.")
        else:
            filtered_df = df[(df["Datum"] >= pd.to_datetime(start_date)) & 
                             (df["Datum"] <= pd.to_datetime(end_date))]

            selected_columns = ["Fortlaufende Nummer", "Projekt", "Datum", 
                                "Probenform", "Messung Durchgeführt", "Auswertung Durchgeführt"]
            existing_columns = [col for col in selected_columns if col in filtered_df.columns]
            filtered_df = filtered_df[existing_columns]

            st.subheader("Filtered Data")
            st.dataframe(filtered_df)

            gestanzt_pattern = r"gestanzt|gestanzz|gestanztz|gestantzt|gestanzr"
            gestanzt_count = filtered_df["Probenform"].astype(str).str.lower().str.contains(gestanzt_pattern, regex=True).sum()
            sample_cutter_count = filtered_df["Probenform"].astype(str).str.lower().str.contains(r"sample cutter", regex=True).sum()

            mh_time_messung = ak_time_messung = hd_time_messung = 0
            mh_total_time_auswertung = ak_total_time_auswertung = hd_total_time_auswertung = 0

            project_time_dict = {}
            original_names = {}

            for _, row in filtered_df.iterrows():
                projekt_original = str(row["Projekt"]).strip()
                projekt_key = re.sub(r'\\s+', ' ', projekt_original.upper().replace('\u00A0', ' ')).strip()

                if projekt_key not in original_names:
                    original_names[projekt_key] = projekt_original

                probenform = str(row["Probenform"]).lower()
                messung_person = str(row["Messung Durchgeführt"]).strip()
                auswertung_person = str(row["Auswertung Durchgeführt"]).strip()

                messung_time = 0
                if pd.notna(probenform):
                    if any(x in probenform for x in ["gestanzt", "gestanzz", "gestanztz", "gestantzt", "gestanzr"]):
                        messung_time = t1
                    elif "sample cutter" in probenform:
                        messung_time = t2

                auswertung_time = t3

                if messung_person == "MH":
                    mh_time_messung += messung_time
                elif messung_person == "AK":
                    ak_time_messung += messung_time
                elif messung_person == "HD":
                    hd_time_messung += messung_time

                if auswertung_person == "MH":
                    mh_total_time_auswertung += t3
                elif auswertung_person == "AK":
                    ak_total_time_auswertung += t3
                elif auswertung_person == "HD":
                    hd_total_time_auswertung += t3

                if pd.notna(projekt_key):
                    total_time = messung_time + auswertung_time
                    project_time_dict[projekt_key] = project_time_dict.get(projekt_key, 0) + total_time

            summary_data = {
                "Category": [
                    "Total 'gestanzt' samples",
                    "Total 'sample cutter' samples",
                    "MH (Messung Durchgeführt)", "AK (Messung Durchgeführt)", "HD (Messung Durchgeführt)",
                    "MH (Auswertung Durchgeführt)", "AK (Auswertung Durchgeführt)", "HD (Auswertung Durchgeführt)"
                ],
                "Count": [
                    gestanzt_count,
                    sample_cutter_count,
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("MH", 0),
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("AK", 0),
                    filtered_df["Messung Durchgeführt"].astype(str).value_counts().get("HD", 0),
                    filtered_df["Auswertung Durchgeführt"].astype(str).value_counts().get("MH", 0),
                    filtered_df["Auswertung Durchgeführt"].astype(str).value_counts().get("AK", 0),
                    filtered_df["Auswertung Durchgeführt"].astype(str).value_counts().get("HD", 0)
                ]
            }

            summary_df = pd.DataFrame(summary_data)
            st.subheader("Summary Table")
            st.dataframe(summary_df)

            def format_time(minutes):
                return f"{minutes // 60}:{minutes % 60:02d}"

            st.subheader("Total Time each person spent at the DSC machine (Hour:Min)")
            st.write(f"MH Total (Messung): {format_time(mh_time_messung)}")
            st.write(f"AK Total (Messung): {format_time(ak_time_messung)}")
            st.write(f"HD Total (Messung): {format_time(hd_time_messung)}")
            st.write(f"MH Total (Auswertung): {format_time(mh_total_time_auswertung)}")
            st.write(f"AK Total (Auswertung): {format_time(ak_total_time_auswertung)}")
            st.write(f"HD Total (Auswertung): {format_time(hd_total_time_auswertung)}")

            st.subheader("Total Time per Projekt (Hour:Min)")
            project_time_df = pd.DataFrame([
                {"Projekt": original_names[k], "Total Time": format_time(v)}
                for k, v in sorted(project_time_dict.items(), key=lambda x: -x[1])
            ])
            st.dataframe(project_time_df, use_container_width=True)