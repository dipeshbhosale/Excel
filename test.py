import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import numpy as np

# === Placeholders for advanced modules ===
# (Implement each in its own function for modularity)

def data_upload_module():
    st.header("Step 1: Upload Data")
    uploaded_files = st.file_uploader(
        "Upload Excel (.xlsx) or CSV (.csv) files", type=["xlsx", "csv"], accept_multiple_files=True
    )
    return uploaded_files

def data_cleaning_module(df):
    st.header("Step 2: Data Cleaning Suggestions")
    return df

def data_analysis_module(df):
    st.header("Step 3: Data Analysis & Insights")
    summary = {}
    insights = []
    num_cols = df.select_dtypes(include='number').columns
    cat_cols = df.select_dtypes(include='object').columns

    # Numerical summary
    if len(num_cols) > 0:
        desc = df[num_cols].describe().T
        summary['numerical'] = desc
        for col in num_cols:
            col_data = df[col].dropna()
            if len(col_data) == 0:
                continue
            mean = col_data.mean()
            median = col_data.median()
            std = col_data.std()
            min_ = col_data.min()
            max_ = col_data.max()
            outliers = col_data[(col_data < mean - 3*std) | (col_data > mean + 3*std)]
            if len(outliers) > 0:
                insights.append(f"Column '{col}' has {len(outliers)} outlier(s).")
            insights.append(f"Column '{col}': mean={mean:.2f}, median={median:.2f}, min={min_}, max={max_}.")
    # Categorical summary
    if len(cat_cols) > 0:
        cat_summary = {}
        for col in cat_cols:
            vc = df[col].value_counts()
            cat_summary[col] = vc
            top = vc.index[0] if not vc.empty else None
            if top:
                insights.append(f"Most frequent value in '{col}' is '{top}' ({vc.iloc[0]} times).")
        summary['categorical'] = cat_summary
    return summary, insights

def plot_and_save(df, sheet_name):
    images = []
    captions = []
    num_cols = df.select_dtypes(include='number').columns
    cat_cols = df.select_dtypes(include='object').columns

    # Numerical columns: histograms and line plots
    for col in num_cols:
        fig, ax = plt.subplots()
        sns.histplot(df[col].dropna(), kde=True, ax=ax)
        ax.set_title(f"Distribution of {col}")
        buf = BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        images.append(buf.read())
        captions.append(f"Distribution of {col} in {sheet_name}.")
        plt.close(fig)

        # Line plot if time-like index
        if pd.api.types.is_datetime64_any_dtype(df.index):
            fig, ax = plt.subplots()
            df[col].plot(ax=ax)
            ax.set_title(f"{col} over time")
            buf = BytesIO()
            plt.savefig(buf, format='png')
            buf.seek(0)
            images.append(buf.read())
            captions.append(f"Trend of {col} over time in {sheet_name}.")
            plt.close(fig)

    # Categorical columns: bar plots and pie charts
    for col in cat_cols:
        vc = df[col].value_counts().head(10)
        if len(vc) > 1:
            fig, ax = plt.subplots()
            sns.barplot(x=vc.values, y=vc.index, ax=ax)
            ax.set_title(f"Top categories in {col}")
            buf = BytesIO()
            plt.savefig(buf, format='png')
            buf.seek(0)
            images.append(buf.read())
            captions.append(f"Top categories in {col} in {sheet_name}.")
            plt.close(fig)

            fig, ax = plt.subplots()
            vc.plot.pie(autopct='%1.1f%%', ax=ax)
            ax.set_ylabel('')
            ax.set_title(f"Category distribution in {col}")
            buf = BytesIO()
            plt.savefig(buf, format='png')
            buf.seek(0)
            images.append(buf.read())
            captions.append(f"Category distribution in {col} in {sheet_name}.")
            plt.close(fig)
    return images, captions

def add_title_slide(prs, file_name):
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Data Analysis Report"
    subtitle.text = f"File: {file_name}\nDate: {datetime.now().strftime('%Y-%m-%d')}"

def add_chart_slide(prs, image_bytes, caption):
    slide_layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = caption
    left = Inches(1)
    top = Inches(1.5)
    pic = slide.shapes.add_picture(BytesIO(image_bytes), left, top, width=Inches(6))
    # Optionally add caption as text box
    # ...

def add_summary_slide(prs, insights):
    slide_layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Key Findings"
    tf = slide.placeholders[1].text_frame
    for insight in insights:
        p = tf.add_paragraph()
        p.text = insight
        p.level = 0

def create_pptx(file_name, all_images, all_captions, insights):
    prs = Presentation()
    add_title_slide(prs, file_name)
    for img, cap in zip(all_images, all_captions):
        add_chart_slide(prs, img, cap)
    add_summary_slide(prs, insights)
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def reporting_module(selected_charts, insights):
    st.header("Step 5: Reporting & Export")
    st.info("Reporting module is under development.")
    # Placeholder for reporting functionality

def collaboration_module():
    st.header("Step 6: Collaboration & Sharing")
    st.info("Collaboration module is under development.")
    # Placeholder for collaboration functionality

def ai_ml_module(df):
    st.header("Step 7: AI/ML Features")
    st.info("AI/ML module is under development.")
    # Placeholder for AI/ML functionality

def analyze_and_clean(df, inplace=False, fill_missing='auto', drop_threshold=0.5, return_flagged=True):
    """
    Perform ground-level data analysis and cleanup on a pandas DataFrame.
    Args:
        df: pandas DataFrame
        inplace: if True, modify df in place; else work on a copy
        fill_missing: 'auto', 'mean', 'median', 'mode', 'drop', or None
        drop_threshold: if fraction of missing in a column > threshold, drop column
        return_flagged: if True, return (cleaned_df, flagged_df)
    Returns:
        cleaned_df, flagged_df (if return_flagged), else just cleaned_df
    """
    if not inplace:
        df = df.copy()
    print("=== Initial DataFrame shape:", df.shape)

    flagged_rows = pd.DataFrame(index=df.index)
    # 1. Check for missing values per column
    print("\n=== Missing Values Per Column ===")
    missing = df.isnull().sum()
    print(missing)
    # Suggest drop columns with too many missing
    cols_to_drop = missing[missing / len(df) > drop_threshold].index.tolist()
    if cols_to_drop:
        print(f"\nColumns with >{int(drop_threshold*100)}% missing values: {cols_to_drop} (will be dropped)")
        df.drop(columns=cols_to_drop, inplace=True)
    # Fill or drop missing values
    for col in df.columns:
        if df[col].isnull().any():
            if fill_missing == 'auto':
                if df[col].dtype.kind in 'biufc':  # numeric
                    fill_val = df[col].median()
                    print(f"Filling missing in '{col}' with median: {fill_val}")
                    df[col].fillna(fill_val, inplace=True)
                elif np.issubdtype(df[col].dtype, np.datetime64):
                    fill_val = df[col].mode().iloc[0] if not df[col].mode().empty else None
                    print(f"Filling missing in '{col}' with mode: {fill_val}")
                    df[col].fillna(fill_val, inplace=True)
                else:
                    fill_val = df[col].mode().iloc[0] if not df[col].mode().empty else ""
                    print(f"Filling missing in '{col}' with mode: {fill_val}")
                    df[col].fillna(fill_val, inplace=True)
            elif fill_missing == 'drop':
                print(f"Dropping rows with missing in '{col}'")
                df = df[df[col].notnull()]
            elif fill_missing in ['mean', 'median', 'mode']:
                if fill_missing == 'mean' and df[col].dtype.kind in 'biufc':
                    fill_val = df[col].mean()
                elif fill_missing == 'median' and df[col].dtype.kind in 'biufc':
                    fill_val = df[col].median()
                else:
                    fill_val = df[col].mode().iloc[0] if not df[col].mode().empty else ""
                print(f"Filling missing in '{col}' with {fill_missing}: {fill_val}")
                df[col].fillna(fill_val, inplace=True)
    # Flag rows with any missing values (after filling)
    flagged_rows['missing'] = df.isnull().any(axis=1)

    # 2. Remove duplicate rows
    before = len(df)
    df_nodup = df.drop_duplicates()
    after = len(df_nodup)
    print(f"\n=== Duplicate Removal ===\nRemoved {before - after} duplicate rows.")
    flagged_rows['duplicate'] = ~df.index.isin(df_nodup.index)
    df = df_nodup

    # 3. Convert column data types
    print("\n=== Data Type Conversion ===")
    for col in df.columns:
        orig_dtype = df[col].dtype
        # Try numeric
        if df[col].dtype == object:
            try:
                df[col] = pd.to_numeric(df[col])
                print(f"Converted '{col}' to numeric.")
            except Exception:
                # Try datetime
                try:
                    df[col] = pd.to_datetime(df[col])
                    print(f"Converted '{col}' to datetime.")
                except Exception:
                    pass
        # Try datetime for columns with 'date' in name
        if 'date' in col.lower() and not np.issubdtype(df[col].dtype, np.datetime64):
            try:
                df[col] = pd.to_datetime(df[col])
                print(f"Converted '{col}' to datetime (by name).")
            except Exception:
                pass
        if orig_dtype != df[col].dtype:
            print(f"Column '{col}': {orig_dtype} -> {df[col].dtype}")

    # 4. Standardize string columns
    print("\n=== Standardizing String Columns ===")
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip().str.lower()
        print(f"Standardized '{col}' (strip, lower).")

    # 5. Print summaries
    print("\n=== Column Data Types ===")
    print(df.dtypes)
    print("\n=== Basic Statistics ===")
    print(df.describe(include='all', datetime_is_numeric=True))
    print("\n=== Value Counts for Categorical Columns ===")
    for col in df.select_dtypes(include='object').columns:
        print(f"\nValue counts for '{col}':")
        print(df[col].value_counts())

    # 6. Detect and highlight outliers (IQR method)
    print("\n=== Outlier Detection (IQR method) ===")
    outlier_flags = pd.DataFrame(False, index=df.index, columns=df.select_dtypes(include=np.number).columns)
    for col in df.select_dtypes(include=np.number).columns:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1
        lower = Q1 - 1.5 * IQR
        upper = Q3 + 1.5 * IQR
        outliers = (df[col] < lower) | (df[col] > upper)
        n_out = outliers.sum()
        print(f"Column '{col}': {n_out} outlier(s) flagged.")
        outlier_flags[col] = outliers
    flagged_rows['outlier'] = outlier_flags.any(axis=1)

    # 7. Flag suspicious rows
    flagged_rows['suspicious'] = flagged_rows.any(axis=1)
    print(f"\n=== Suspicious Rows: {flagged_rows['suspicious'].sum()} flagged ===")

    # 8. Print before/after row counts
    print(f"\n=== Row Count: Before: {before}, After: {len(df)} ===")

    # 9. Return cleaned DataFrame and flagged DataFrame
    if return_flagged:
        flagged_df = df[flagged_rows['suspicious']]
        return df, flagged_df
    else:
        return df

def generate_clear_chart(
    df,
    chart_type="bar",
    x=None,
    y=None,
    title=None,
    save_path=None,
    return_fig=False,
    max_pie_slices=8
):
    """
    Generate a clear, simple, and sorted chart from a pandas DataFrame.
    Supports: bar, line, pie.
    """
    import matplotlib.pyplot as plt
    import seaborn as sns
    from matplotlib.ticker import FuncFormatter
    import matplotlib.dates as mdates

    # --- Style ---
    plt.rcParams.update({
        "axes.titlesize": 18,
        "axes.labelsize": 14,
        "xtick.labelsize": 13,
        "ytick.labelsize": 13,
        "axes.titlepad": 18,
        "axes.labelpad": 10,
        "axes.edgecolor": "#aaa",
        "axes.linewidth": 1.2,
        "figure.facecolor": "#fff",
        "axes.facecolor": "#fff",
        "legend.fontsize": 13,
        "font.family": "DejaVu Sans"
    })
    fig, ax = plt.subplots(figsize=(8, 5))

    def auto_title(chart_type, x, y):
        if chart_type == "bar":
            return f"{y or 'Value'} by {x or 'Category'}"
        elif chart_type == "line":
            return f"{y or 'Value'} over {x or 'Time'}"
        elif chart_type == "pie":
            return f"{x or 'Category'} Distribution"
        else:
            return "Chart"

    # --- Bar Chart ---
    if chart_type == "bar":
        if x is None or y is None:
            raise ValueError("x and y must be specified for bar chart")
        # Group and sort
        data = df.groupby(x, dropna=False)[y].sum().sort_values(ascending=False)
        data = data.head(15)  # Top 15 for clarity
        bars = ax.barh(
            data.index.astype(str),
            data.values,
            color=sns.color_palette("Blues_r", len(data)),
            edgecolor="#333",
            height=0.6
        )
        ax.set_xlabel(y)
        ax.set_ylabel(x)
        ax.set_title(title or auto_title("bar", x, y))
        ax.invert_yaxis()
        for bar in bars:
            width = bar.get_width()
            ax.text(width + max(data.values)*0.01, bar.get_y() + bar.get_height()/2,
                    f"{width:,.0f}", va="center", fontsize=13, color="#222")
        ax.spines[['top', 'right', 'left']].set_visible(False)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{int(x):,}"))
        ax.grid(axis='x', linestyle='--', alpha=0.3)
        plt.tight_layout()

    # --- Line Chart ---
    elif chart_type == "line":
        if x is None or y is None:
            raise ValueError("x and y must be specified for line chart")
        x_data = pd.to_datetime(df[x], errors="coerce")
        sorted_df = df.assign(_x=x_data).sort_values("_x")
        ax.plot(sorted_df["_x"], sorted_df[y], marker="o", color="#4F8DFD", linewidth=2)
        ax.set_xlabel(x)
        ax.set_ylabel(y)
        ax.set_title(title or auto_title("line", x, y))
        ax.spines[['top', 'right']].set_visible(False)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        # Date formatting
        if pd.api.types.is_datetime64_any_dtype(sorted_df["_x"]):
            locator = mdates.AutoDateLocator()
            formatter = mdates.ConciseDateFormatter(locator)
            ax.xaxis.set_major_locator(locator)
            ax.xaxis.set_major_formatter(formatter)
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right")
        plt.tight_layout()

    # --- Pie Chart ---
    elif chart_type == "pie":
        if x is None:
            raise ValueError("x must be specified for pie chart")
        counts = df[x].value_counts(dropna=False)
        if max_pie_slices and len(counts) > max_pie_slices:
            top = counts[:max_pie_slices-1]
            others = counts[max_pie_slices-1:].sum()
            counts = top.append(pd.Series({"Others": others}))
        colors = sns.color_palette("pastel", len(counts))
        wedges, texts, autotexts = ax.pie(
            counts.values,
            labels=[str(i) for i in counts.index],
            autopct=lambda pct: f"{pct:.1f}%",
            startangle=90,
            colors=colors,
            textprops={'fontsize': 13, 'color': "#222"}
        )
        ax.set_title(title or auto_title("pie", x, None))
        plt.tight_layout()
    else:
        raise ValueError("chart_type must be one of: 'bar', 'line', 'pie'")

    # --- Clean up ---
    ax.set_facecolor("#fff")
    fig.patch.set_facecolor("#fff")
    plt.subplots_adjust(left=0.18, right=0.98, top=0.88, bottom=0.18)

    # Save or return
    if save_path:
        plt.savefig(save_path, bbox_inches="tight", dpi=150)
    if return_fig:
        return fig, ax
    else:
        plt.show()
        plt.close(fig)

# Example usage:
# generate_clear_chart(df, chart_type="bar", x="Department", y="Expenses")
# generate_clear_chart(df, chart_type="line", x="Date", y="Sales")

def main():
    st.set_page_config(page_title="Data-Driven SaaS Assistant", layout="wide")
    st.title("Excel Data-Driven SaaS Assistant")
    uploaded_file = st.file_uploader("Upload an Excel (.xlsx) or CSV (.csv) file", type=["xlsx", "csv"])
    if uploaded_file is not None:
        try:
            file_name = uploaded_file.name
            file_bytes = uploaded_file.read()
            if not file_bytes:
                st.error("Uploaded file is empty.")
                return
            if file_name.lower().endswith(".xlsx"):
                excel_buffer = BytesIO(file_bytes)
                xls = pd.ExcelFile(excel_buffer, engine="openpyxl")
                st.write(f"Detected sheets: {xls.sheet_names}")
                all_images = []
                all_captions = []
                all_insights = []
                for sheet in xls.sheet_names:
                    st.subheader(f"Sheet: {sheet}")
                    df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
                    # --- Editable DataFrame ---
                    edited_df = st.data_editor(df, num_rows="dynamic", key=f"edit_{sheet}")
                    summary, insights = data_analysis_module(edited_df)
                    all_insights.extend(insights)
                    st.write("Summary statistics:")
                    if 'numerical' in summary:
                        st.write(summary['numerical'])
                    if 'categorical' in summary:
                        for col, vc in summary['categorical'].items():
                            st.write(f"Value counts for {col}:")
                            st.write(vc)
                    images, captions = plot_and_save(edited_df, sheet)
                    for img, cap in zip(images, captions):
                        st.image(img, caption=cap)
                    all_images.extend(images)
                    all_captions.extend(captions)
                    # --- Data Q&A ---
                    st.markdown("**Ask a question about this sheet:**")
                    user_query = st.text_input(f"Ask about '{sheet}'", key=f"query_{sheet}")
                    if user_query:
                        try:
                            # Simple pandas eval for demo (no OpenAI)
                            # e.g., "mean of column A", "max of sales"
                            # For safety, only allow simple queries
                            import re
                            colnames = list(edited_df.columns)
                            # Try to extract column and operation
                            found = False
                            for col in colnames:
                                if col.lower() in user_query.lower():
                                    if "mean" in user_query.lower():
                                        st.info(f"Mean of '{col}': {edited_df[col].mean()}")
                                        found = True
                                    elif "sum" in user_query.lower():
                                        st.info(f"Sum of '{col}': {edited_df[col].sum()}")
                                        found = True
                                    elif "max" in user_query.lower():
                                        st.info(f"Max of '{col}': {edited_df[col].max()}")
                                        found = True
                                    elif "min" in user_query.lower():
                                        st.info(f"Min of '{col}': {edited_df[col].min()}")
                                        found = True
                                    elif "count" in user_query.lower():
                                        st.info(f"Count of '{col}': {edited_df[col].count()}")
                                        found = True
                            if not found:
                                st.warning("Sorry, only simple queries like mean/sum/max/min/count of a column are supported in this demo.")
                        except Exception as e:
                            st.error(f"Error answering question: {e}")
                if st.button("Generate PowerPoint Report"):
                    pptx_io = create_pptx(file_name, all_images, all_captions, all_insights)
                    st.success("PowerPoint generated!")
                    st.download_button(
                        label="Download PowerPoint",
                        data=pptx_io,
                        file_name=f"{file_name}_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            elif file_name.lower().endswith(".csv"):
                csv_buffer = BytesIO(file_bytes)
                df = pd.read_csv(csv_buffer)
                st.subheader(f"CSV File: {file_name}")
                # --- Editable DataFrame ---
                edited_df = st.data_editor(df, num_rows="dynamic", key="edit_csv")
                summary, insights = data_analysis_module(edited_df)
                st.write("Summary statistics:")
                if 'numerical' in summary:
                    st.write(summary['numerical'])
                if 'categorical' in summary:
                    for col, vc in summary['categorical'].items():
                        st.write(f"Value counts for {col}:")
                        st.write(vc)
                images, captions = plot_and_save(edited_df, file_name)
                for img, cap in zip(images, captions):
                    st.image(img, caption=cap)
                # --- Data Q&A ---
                st.markdown("**Ask a question about this file:**")
                user_query = st.text_input("Ask about CSV", key="query_csv")
                if user_query:
                    try:
                        colnames = list(edited_df.columns)
                        found = False
                        for col in colnames:
                            if col.lower() in user_query.lower():
                                if "mean" in user_query.lower():
                                    st.info(f"Mean of '{col}': {edited_df[col].mean()}")
                                    found = True
                                elif "sum" in user_query.lower():
                                    st.info(f"Sum of '{col}': {edited_df[col].sum()}")
                                    found = True
                                elif "max" in user_query.lower():
                                    st.info(f"Max of '{col}': {edited_df[col].max()}")
                                    found = True
                                elif "min" in user_query.lower():
                                    st.info(f"Min of '{col}': {edited_df[col].min()}")
                                    found = True
                                elif "count" in user_query.lower():
                                    st.info(f"Count of '{col}': {edited_df[col].count()}")
                                    found = True
                        if not found:
                            st.warning("Sorry, only simple queries like mean/sum/max/min/count of a column are supported in this demo.")
                    except Exception as e:
                        st.error(f"Error answering question: {e}")
                if st.button("Generate PowerPoint Report"):
                    pptx_io = create_pptx(file_name, images, captions, insights)
                    st.success("PowerPoint generated!")
                    st.download_button(
                        label="Download PowerPoint",
                        data=pptx_io,
                        file_name=f"{file_name}_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            else:
                st.error("Please upload a valid .xlsx or .csv file.")
                return
        except Exception as e:
            st.error(f"Error reading file: {e}")

if __name__ == "__main__":
    main()