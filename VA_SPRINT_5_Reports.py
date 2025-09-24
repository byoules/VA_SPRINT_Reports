import pandas as pd
import plotly.express as px
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from collections import Counter
from wordcloud import WordCloud
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk


# === Progress Bar Window ===
class ProgressWindow:
    def __init__(self, title, max_value):
        self.root = tk.Toplevel()
        self.root.title(title)
        self.progress = ttk.Progressbar(self.root, length=400, mode='determinate', maximum=max_value)
        self.progress.pack(padx=20, pady=20)
        self.label = tk.Label(self.root, text="Starting...")
        self.label.pack(padx=20, pady=(0, 20))
        self.root.update()

    def update(self, value, text=""):
        self.progress['value'] = value
        if text:
            self.label.config(text=text)
        self.root.update()

    def close(self):
        self.root.destroy()


# === Shared Excel loader ===
def load_excel():
    root = tk.Tk()
    root.withdraw()

    # 🔹 Welcome popup
    messagebox.showinfo(
        "SPRINT API Dataset",
        "Please select the SPRINT API dataset Excel file."
    )

    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        messagebox.showerror("No File Selected", "Please select an Excel file.")
        return None, None
    try:
        df = pd.read_excel(file_path, dtype=str)

        # 🔹 Clean up field names
        df.columns = [c.strip() for c in df.columns]             # remove leading/trailing spaces
        df.columns = [" ".join(c.split()) for c in df.columns]   # collapse multiple spaces

        return file_path, df
    except Exception as e:
        messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
        return None, None


# === Column picker helper ===
def get_or_select_column(df, expected_name, title="Select Column"):
    if expected_name in df.columns:
        return expected_name

    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Column Not Found",
                        f"Could not find '{expected_name}'. Please choose manually.")

    selected = simpledialog.askstring(
        title,
        f"Available columns:\n\n" + "\n".join(df.columns) +
        "\n\nEnter the exact column name:"
    )
    if selected in df.columns:
        return selected
    else:
        messagebox.showerror("Invalid Selection", "Column not found. Skipping this analysis.")
        return None


# === Analysis 1: Funding Department ===
def analyze_funding_department(df, file_path):
    col_name = get_or_select_column(df, "Funding Department", "Select Funding Column")
    if not col_name:
        return None

    progress = ProgressWindow("Funding Department Report", 3)

    df[col_name] = df[col_name].str.strip().replace("other", "Other")
    progress.update(1, "Cleaning data...")

    missing_count = df[col_name].isna().sum() + (df[col_name] == "NR").sum()
    include_note = messagebox.askyesno(
        "Funding Department Report",
        "Include note about missing values in the Funding Department report?"
    )

    df_clean = df[~df[col_name].isin([None, "NR"]) & df[col_name].notna()]
    group_counts = df_clean[col_name].value_counts().reset_index()
    group_counts.columns = ['Funding Department', '# of Projects']

    progress.update(2, "Building chart...")

    total_n = df.shape[0]
    chart_title = f"<b>SPRINT API: Number of Projects by Funder (N = {total_n})</b>"

    fig = px.bar(group_counts, x='Funding Department', y='# of Projects', title=chart_title)
    if include_note:
        fig.add_annotation(
            text=f"<i>Note: {missing_count} missing values</i>",
            xref="paper", yref="paper", x=0.99, y=0.99,
            xanchor="right", yanchor="top", showarrow=False,
            font=dict(size=11, color="gray")
        )

    progress.update(3, "Saving outputs...")

    html_path = os.path.join(os.path.dirname(file_path), "1_SPRINT_API_by_Funder.html")
    png_path = os.path.join(os.path.dirname(file_path), "1_SPRINT_API_by_Funder.png")
    fig.write_html(html_path)
    fig.write_image(png_path, scale=2)

    progress.close()
    return "Funding Department"


# === Analysis 2: Study Type ===
def analyze_study_type(df, file_path):
    col_name = get_or_select_column(df, "Study Type", "Select Study Type Column")
    if not col_name:
        return None

    progress = ProgressWindow("Study Type Report", 3)

    df[col_name] = df[col_name].str.strip().replace("other", "Other")
    progress.update(1, "Cleaning data...")

    missing_count = df[col_name].isna().sum() + (df[col_name] == "NR").sum()
    include_note = messagebox.askyesno(
        "Study Type Report",
        "Include note about missing values in the Study Type report?"
    )

    df_clean = df[~df[col_name].isin([None, "NR"]) & df[col_name].notna()]
    group_counts = df_clean[col_name].value_counts().reset_index()
    group_counts.columns = ['Study Type', '# of Projects']

    progress.update(2, "Building chart...")

    total_n = df.shape[0]
    chart_title = f"<b>SPRINT API: Study Type Distribution (N = {total_n})</b>"

    fig = px.pie(group_counts, names='Study Type', values='# of Projects', title=chart_title)
    fig.update_traces(textinfo='percent+label', textfont_size=16)

    if include_note:
        fig.add_annotation(
            text=f"<i>Note: {missing_count} missing values</i>",
            xref="paper", yref="paper", x=0.35, y=-.05,
            xanchor="right", yanchor="top", showarrow=False,
            font=dict(size=11, color="gray")
        )

    progress.update(3, "Saving outputs...")

    html_path = os.path.join(os.path.dirname(file_path), "2_SPRINT_API_Study_Type_Pie_Chart.html")
    png_path = os.path.join(os.path.dirname(file_path), "2_SPRINT_API_Study_Type_Pie_Chart.png")
    fig.write_html(html_path)
    fig.write_image(png_path, scale=2)

    progress.close()
    return "Study Type"


# === Analysis 3: Public Health Approach ===
def analyze_public_health_approach(df, file_path):
    col_name = get_or_select_column(df, "Public Health Approach", "Select Public Health Column")
    if not col_name:
        return None

    progress = ProgressWindow("Public Health Report", 3)

    df[col_name] = df[col_name].str.strip().replace({"other": "Other", "selective": "Selective"})
    progress.update(1, "Cleaning data...")

    missing_count = df[col_name].isna().sum() + (df[col_name] == "NR").sum()
    include_note = messagebox.askyesno(
        "Public Health Report",
        "Include note about missing values in the Public Health report?"
    )

    df_clean = df[~df[col_name].isin([None, "NR"]) & df[col_name].notna()]
    group_counts = df_clean[col_name].value_counts().reset_index()
    group_counts.columns = ['Public Health Approach', '# of Projects']

    progress.update(2, "Building chart...")

    total_n = df.shape[0]
    chart_title = f"<b>SPRINT API: Public Health Approach Distribution (N = {total_n})</b>"

    fig = px.bar(group_counts, x='# of Projects', y='Public Health Approach',
                 orientation='h', title=chart_title)

    if include_note:
        fig.add_annotation(
            text=f"<i>Note: {missing_count} missing values</i>",
            xref="paper", yref="paper", x=0, y=-0.25,
            xanchor="left", yanchor="top", showarrow=False,
            font=dict(size=12, color="gray")
        )

    progress.update(3, "Saving outputs...")

    html_path = os.path.join(os.path.dirname(file_path), "3_SPRINT_API_public_health_approach_chart.html")
    png_path = os.path.join(os.path.dirname(file_path), "3_SPRINT_API_public_health_approach_chart.png")
    fig.write_html(html_path)
    fig.write_image(png_path, scale=2)

    progress.close()
    return "Public Health Approach"


# === Analysis 4: PI Facility Map ===
def analyze_pi_facility(df, file_path):
    col_name = get_or_select_column(df, "P.I. Facility and Location", "Select Facility Column")
    if not col_name:
        return None

    df[col_name] = df[col_name].str.strip().str.split(";").str[0].str.strip()
    df[col_name] = df[col_name].replace({"other": "Other", "Aurora, GA": "Aurora, CO"})

    group_counts = df[col_name].value_counts().reset_index()
    group_counts.columns = ['Location', '# of Projects']

    progress = ProgressWindow("PI Facility Map Report", len(group_counts))

    missing_count = df[col_name].isna().sum() + (df[col_name] == "NR").sum()
    include_note = messagebox.askyesno(
        "PI Facility Map Report",
        "Include note about missing values in the PI Facility Map report?"
    )

    geolocator = Nominatim(user_agent="project_mapper")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

    lats, longs = [], []
    for i, loc in enumerate(group_counts['Location'], 1):
        try:
            location = geocode(loc + ", USA")
            if location:
                lats.append(location.latitude)
                longs.append(location.longitude)
            else:
                lats.append(None)
                longs.append(None)
        except:
            lats.append(None)
            longs.append(None)

        progress.update(i, f"Geocoding {i}/{len(group_counts)}: {loc}")

    group_counts["Latitude"] = lats
    group_counts["Longitude"] = longs
    geo_df = group_counts.dropna(subset=["Latitude", "Longitude"])

    total_n = df.shape[0]
    total_in_visual = total_n - missing_count
    chart_title = f"<b>SPRINT API: Project PI Facilities and Locations (N = {total_in_visual})</b>"

    fig = px.scatter_geo(geo_df, lat="Latitude", lon="Longitude", size="# of Projects",
                         hover_name="Location", projection="albers usa", scope="usa", title=chart_title)

    if include_note:
        fig.add_annotation(
            text=f"<i>Note: {missing_count} missing values</i>",
            xref="paper", yref="paper", x=0, y=-0.25,
            xanchor="left", yanchor="top", showarrow=False,
            font=dict(size=13, color="gray")
        )

    html_path = os.path.join(os.path.dirname(file_path), "4_SPRINT_API_pi_facility_map.html")
    png_path = os.path.join(os.path.dirname(file_path), "4_SPRINT_API_pi_facility_map.png")
    fig.write_html(html_path)
    fig.write_image(png_path, scale=2)

    progress.close()
    return "PI Facility Map"


# === Analysis 5: Keyword Analysis ===
def analyze_keywords(df, file_path):
    keyword_cols = ["Key Word 1", "Key Word 2", "Key Word 3", "Key Word 4"]
    all_keywords = []

    progress = ProgressWindow("Keyword Analysis Report", len(keyword_cols))

    for i, expected in enumerate(keyword_cols, 1):
        col = get_or_select_column(df, expected, f"Select Column for {expected}")
        if not col:
            continue
        cleaned = df[col].dropna().str.strip()
        cleaned = cleaned[~cleaned.isin(["NR", ""])]
        all_keywords.extend(cleaned.tolist())
        progress.update(i, f"Processed {expected}")

    keyword_counts = Counter(all_keywords)
    top_20 = keyword_counts.most_common(20)

    wordcloud = WordCloud(width=800, height=400, background_color="white")
    wordcloud_img = wordcloud.generate_from_frequencies(keyword_counts)

    output_dir = os.path.dirname(file_path)
    wordcloud_img_path = os.path.join(output_dir, "wordcloud.png")
    wordcloud_img.to_file(wordcloud_img_path)

    df_top = pd.DataFrame(top_20, columns=["Keyword", "# of Projects"])

    html_path = os.path.join(output_dir, "5_SPRINT_API_Keyword_Analysis.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write("<html><head><title>SPRINT API: Keyword Analysis</title></head>")
        f.write("<body style='font-family:sans-serif;'>")
        f.write("<h1>SPRINT API: Keyword Analysis</h1>")
        f.write("<div style='display:flex; gap:30px;'>")
        f.write("<div style='flex:3;'><h2>Word Map</h2>")
        f.write("<img src='wordcloud.png' style='width:100%; max-width:800px;'></div>")
        f.write("<div style='flex:2;'><h2>Top 20 Keywords</h2>")
        f.write(df_top.to_html(index=False, border=0))
        f.write("</div></div></body></html>")

    progress.close()
    return "Keyword Analysis"


# === Main launcher ===
def main():
    file_path, df = load_excel()
    if df is None:
        return

    root = tk.Tk()
    root.withdraw()
    choice = simpledialog.askstring(
        "Select Report(s) to Run",
        "Please select which report(s) to generate:\n\n"
        "1 = Funding Department\n"
        "2 = Study Type\n"
        "3 = Public Health Approach\n"
        "4 = PI Facility Map\n"
        "5 = Keyword Analysis\n"
        "all = Run All Reports"
    )

    completed_reports = []

    if choice in ("1", "all"):
        result = analyze_funding_department(df, file_path)
        if result: completed_reports.append(result)
    if choice in ("2", "all"):
        result = analyze_study_type(df, file_path)
        if result: completed_reports.append(result)
    if choice in ("3", "all"):
        result = analyze_public_health_approach(df, file_path)
        if result: completed_reports.append(result)
    if choice in ("4", "all"):
        result = analyze_pi_facility(df, file_path)
        if result: completed_reports.append(result)
    if choice in ("5", "all"):
        result = analyze_keywords(df, file_path)
        if result: completed_reports.append(result)

    if completed_reports:
        messagebox.showinfo("Done", f"Generated reports:\n- " + "\n- ".join(completed_reports))
    else:
        messagebox.showwarning("No Reports", "No reports were generated.")


if __name__ == "__main__":
    main()
