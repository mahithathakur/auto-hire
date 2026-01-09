import pandas as pd
import os
from pymongo import MongoClient
import subprocess
import platform

def open_file(filepath):
    abs_path = os.path.abspath(filepath)
    if os.path.exists(abs_path):
        if platform.system() == "Windows":
            os.startfile(abs_path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", abs_path])
        else:
            subprocess.run(["xdg-open", abs_path])
    else:
        print(f"❌ File not found: {abs_path}")

def export_to_excel(data=None, file_path="Candidate_Report.xlsx"):
    client = MongoClient("mongodb://localhost:27017/")
    db = client["resume_db"]

    try:
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        print(f"⚠️ Could not delete existing file: {e}")
        return

    if data is None:
        data = list(db["jd_evaluation"].find({}))

    if not data:
        print("⚠️ No data to export.")
        return

    df = pd.DataFrame(data)

    if "score" in df.columns:
        df["jd_match_score"] = df["score"] / 100.0

    df["jd_match_score"] = (df["jd_match_score"] * 100).round(2)
    df["is_duplicate"] = df.get("is_duplicate", False).apply(lambda x: "Yes" if x else "No")

    df.rename(columns={
        "name": "Name",
        "email": "Email",
        "jd_match_score": "JD Match (%)",
        "is_duplicate": "Duplicate",
        "status": "Status",
        "timestamp": "Processed At"
    }, inplace=True)

    if "Processed At" in df.columns:
        df["Processed At"] = pd.to_datetime(df["Processed At"]).dt.strftime('%Y-%m-%d %H:%M')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Candidates')
        workbook = writer.book
        worksheet = writer.sheets['Candidates']

        header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)

        status_format = {
            "Approved": "#C6EFCE",
            "Rejected": "#FFC7CE",
            "Pending": "#FFEB9C"
        }
        if "Status" in df.columns:
            for status, color in status_format.items():
                worksheet.conditional_format(1, df.columns.get_loc("Status"),
                    len(df), df.columns.get_loc("Status"),
                    {'type': 'text', 'criteria': 'containing', 'value': status,
                     'format': workbook.add_format({'bg_color': color})})

        if "Duplicate" in df.columns:
            dup_col = df.columns.get_loc("Duplicate")
            worksheet.conditional_format(1, dup_col, len(df), dup_col, {
                'type': 'text',
                'criteria': 'containing',
                'value': "Yes",
                'format': workbook.add_format({'bg_color': "#E19CFF"})
            })

        for status in ["Approved", "Rejected", "Pending"]:
            if "Status" in df.columns:
                status_df = df[df["Status"] == status]
                if not status_df.empty:
                    status_df.to_excel(writer, index=False, sheet_name=status)

        summary_data = {
            "Total Candidates": len(df),
            "Approved": (df["Status"] == "Approved").sum(),
            "Rejected": (df["Status"] == "Rejected").sum(),
            "Pending": (df["Status"] == "Pending").sum(),
            "Avg JD Match (%)": df["JD Match (%)"].mean().round(2)
        }
        summary_df = pd.DataFrame(list(summary_data.items()), columns=["Metric", "Value"])
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        summary_sheet = writer.sheets["Summary"]

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Candidate Status',
            'categories': ['Summary', 1, 0, 3, 0],
            'values':     ['Summary', 1, 1, 3, 1],
        })
        chart.set_title({'name': 'Candidate Status Distribution'})
        chart.set_x_axis({'name': 'Status'})
        chart.set_y_axis({'name': 'Count'})
        summary_sheet.insert_chart('D2', chart)

        # 📊 Unique vs Duplicate Resume Comparison
        try:
            dup_data = list(db["duplicate_candidates"].find({}))
            dup_count = len(dup_data)
            unique_count = len(df)

            graph_df = pd.DataFrame({
                "Type": ["Unique", "Duplicate"],
                "Count": [unique_count, dup_count]
            })

            graph_df.to_excel(writer, sheet_name="Duplicates", index=False)
            graph_sheet = writer.sheets["Duplicates"]

            pie_chart = workbook.add_chart({'type': 'pie'})
            pie_chart.add_series({
                'name': 'Resume Types',
                'categories': ['Duplicates', 1, 0, 2, 0],
                'values':     ['Duplicates', 1, 1, 2, 1],
            })
            pie_chart.set_title({'name': 'Unique vs Duplicate Resumes'})
            graph_sheet.insert_chart('E2', pie_chart)

            bar_chart = workbook.add_chart({'type': 'column'})
            bar_chart.add_series({
                'name': 'Resume Counts',
                'categories': ['Duplicates', 1, 0, 2, 0],
                'values':     ['Duplicates', 1, 1, 2, 1],
            })
            bar_chart.set_title({'name': 'Resume Distribution'})
            bar_chart.set_x_axis({'name': 'Type'})
            bar_chart.set_y_axis({'name': 'Count'})
            graph_sheet.insert_chart('E18', bar_chart)
        except Exception as e:
            print(f"⚠️ Error creating duplicates graph sheet: {e}")

        # 🟣 Detailed Duplicate Candidate Sheet
        try:
            if dup_data:
                dup_detail_df = pd.DataFrame(dup_data)
                dup_detail_df["is_duplicate"] = dup_detail_df.get("is_duplicate", True).apply(lambda x: "Yes" if x else "No")
                dup_detail_df.rename(columns={
                    "name": "Name",
                    "email": "Email",
                    "score": "JD Match (%)",
                    "status": "Status",
                    "timestamp": "Processed At",
                    "duplicate_reason": "Duplicate Reason"
                }, inplace=True)
                dup_detail_df["Processed At"] = pd.to_datetime(dup_detail_df["Processed At"]).dt.strftime('%Y-%m-%d %H:%M')
                dup_detail_df.to_excel(writer, index=False, sheet_name="Duplicate Candidates")
                dup_candidate_sheet = writer.sheets["Duplicate Candidates"]
                dup_candidate_sheet.autofilter(0, 0, len(dup_detail_df), len(dup_detail_df.columns) - 1)
                dup_candidate_sheet.freeze_panes(1, 0)

                header_format = workbook.add_format({'bold': True, 'bg_color': '#FFD6D6'})
                for col_num, value in enumerate(dup_detail_df.columns.values):
                    dup_candidate_sheet.write(0, col_num, value, header_format)
        except Exception as e:
            print(f"⚠️ Failed to load duplicate candidate details: {e}")

    print(f"✅ Excel report exported successfully → {file_path}")
    #open_file(file_path)

if __name__ == "__main__":
    export_to_excel()

