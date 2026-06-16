import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import StringIO
import numpy as np

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Phân Tích Bảng Cân Đối Kế Toán",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Phân Tích Bảng Cân Đối Kế Toán")
st.markdown("Nhập dữ liệu BCĐKT → Tự động tính chỉ số tài chính → Biểu đồ trực quan")

# ============================================================
# 1. DATA INPUT
# ============================================================
st.header("1️⃣ Nhập dữ liệu BCĐKT")

input_mode = st.radio("Chọn cách nhập dữ liệu:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True)

df_input = None

if input_mode == "📂 Upload file Excel":
    uploaded = st.file_uploader("Chọn file Excel (.xlsx / .xls)", type=["xlsx", "xls", "csv"])
    if uploaded:
        try:
            if uploaded.name.lower().endswith(".csv"):
                df_input = pd.read_csv(uploaded, header=None)
            else:
                df_input = pd.read_excel(uploaded, header=None, engine="openpyxl" if uploaded.name.lower().endswith(".xlsx") else None)
            st.success(f"Đã đọc file: {uploaded.name} ({df_input.shape[0]} dòng × {df_input.shape[1]} cột)")
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}\n\n💡 Thử lưu file dưới dạng .xlsx (Excel 2007+) rồi upload lại.")

elif input_mode == "✏️ Nhập tay":
    st.markdown("""
    Nhập các chỉ tiêu theo định dạng: `Mã, Tên chỉ tiêu, Số đầu kỳ, Số cuối kỳ`  
    Mỗi dòng một chỉ tiêu.
    """)
    default_data = """01,Tổng tài sản,100000,120000
02,Tài sản ngắn hạn,60000,75000
03,Tài sản dài hạn,40000,45000
10,Tổng nợ phải trả,60000,70000
11,Nợ ngắn hạn,40000,45000
12,Nợ dài hạn,20000,25000
30,Vốn chủ sở hữu,40000,50000
31,Vốn điều lệ,30000,30000"""
    text = st.text_area("Dữ liệu BCĐKT", value=default_data, height=220)
    try:
        df_input = pd.read_csv(StringIO(text), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
    except Exception:
        st.warning("Dữ liệu nhập tay chưa hợp lệ. Kiểm tra lại định dạng.")

# ============================================================
# 2. PARSE & DISPLAY RAW DATA
# ============================================================
if df_input is not None and not df_input.empty:
    # Normalize column names if from Excel
    df_input.columns = [str(c) for c in df_input.columns]

    st.header("2️⃣ Dữ liệu gốc")
    st.dataframe(df_input, use_container_width=True)

    # ============================================================
    # 3. EXTRACT KEY FIGURES
    # ============================================================
    st.header("3️⃣ Trích xuất chỉ tiêu chính")

    def find_value(df, keywords, col_name):
        """Search for a keyword in all string columns, return value from col_name."""
        for col in df.columns:
            try:
                mask = df[col].astype(str).str.lower().str.contains("|".join(keywords), na=False, regex=True)
                if mask.any():
                    idx = df[mask].index[0]
                    val = df.loc[idx, col_name]
                    try:
                        return float(val)
                    except (ValueError, TypeError):
                        return None
            except Exception:
                continue
        return None

    # Try to find columns that look like period values
    numeric_cols = [c for c in df_input.columns if df_input[c].dtype in [np.float64, np.int64, np.float32, np.int32] or pd.api.types.is_numeric_dtype(df_input[c])]
    text_cols = [c for c in df_input.columns if c not in numeric_cols]

    if len(numeric_cols) >= 2:
        col_start = numeric_cols[0]
        col_end = numeric_cols[1]
    elif len(numeric_cols) == 1:
        col_start = col_end = numeric_cols[0]
    else:
        col_start = df_input.columns[-2] if len(df_input.columns) >= 2 else df_input.columns[0]
        col_end = df_input.columns[-1]

    # Manual entry mode already has clean columns
    if input_mode == "✏️ Nhập tay":
        col_start = "Số đầu kỳ"
        col_end = "Số cuối kỳ"
        text_col = "Tên chỉ tiêu"
    else:
        text_col = text_cols[0] if text_cols else df_input.columns[0]

    # Define key items to extract
    key_items = {
        "Tổng tài sản": ["tổng tài sản", "tổng ts"],
        "Tài sản ngắn hạn": ["tài sản ngắn hạn", "ts ngắn hạn", "tài sản ngắn"],
        "Tài sản dài hạn": ["tài sản dài hạn", "ts dài hạn", "tài sản dài"],
        "Tiền và tương đương tiền": ["tiền", "tiền và tương đương"],
        "Phải thu ngắn hạn": ["phải thu ngắn hạn", "phải thu khách hàng"],
        "Hàng tồn kho": ["hàng tồn kho", "htk"],
        "Tổng nợ phải trả": ["tổng nợ phải trả", "tổng nợ", "nợ phải trả"],
        "Nợ ngắn hạn": ["nợ ngắn hạn", "nợ ngắn"],
        "Nợ dài hạn": ["nợ dài hạn", "nợ dài"],
        "Vốn chủ sở hữu": ["vốn chủ sở hữu", "vcsh", "vốn chủ"],
        "Lợi nhuận chưa phân phối": ["lợi nhuận chưa phân phối", "ln cpt", "lợi nhuận"],
    }

    extracted = {}
    for label, kws in key_items.items():
        val_start = find_value(df_input, kws, col_start)
        val_end = find_value(df_input, kws, col_end)
        extracted[label] = {"Đầu kỳ": val_start, "Cuối kỳ": val_end}

    df_extracted = pd.DataFrame(extracted).T
    df_extracted = df_extracted.reset_index().rename(columns={"index": "Chỉ tiêu"})
    df_extracted = df_extracted.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

    if df_extracted.empty:
        st.warning("⚠️ Không tự động trích được chỉ tiêu. Bạn có thể chỉnh tên chỉ tiêu trong file cho khớp từ khóa, hoặc dùng chế độ nhập tay.")
        st.stop()

    st.dataframe(df_extracted.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

    # ============================================================
    # 4. FINANCIAL RATIOS
    # ============================================================
    st.header("4️⃣ Chỉ số tài chính")

    def safe_div(a, b):
        if a is None or b is None or b == 0:
            return None
        return round(a / b, 4)

    ratios = {}

    # --- THANH KHOẢN ---
    ts_nh = extracted.get("Tài sản ngắn hạn", {}).get("Cuối kỳ")
    no_nh = extracted.get("Nợ ngắn hạn", {}).get("Cuối kỳ")
    tien = extracted.get("Tiền và tương đương tiền", {}).get("Cuối kỳ")
    htk = extracted.get("Hàng tồn kho", {}).get("Cuối kỳ")
    pt_nh = extracted.get("Phải thu ngắn hạn", {}).get("Cuối kỳ")

    ts_nh_dk = extracted.get("Tài sản ngắn hạn", {}).get("Đầu kỳ")
    no_nh_dk = extracted.get("Nợ ngắn hạn", {}).get("Đầu kỳ")
    tien_dk = extracted.get("Tiền và tương đương tiền", {}).get("Đầu kỳ")
    htk_dk = extracted.get("Hàng tồn kho", {}).get("Đầu kỳ")
    pt_nh_dk = extracted.get("Phải thu ngắn hạn", {}).get("Đầu kỳ")

    ratios["Hệ số thanh toán hiện hành"] = safe_div(ts_nh, no_nh)
    ratios["Hệ số thanh toán nhanh"] = safe_div((ts_nh or 0) - (htk or 0), no_nh)
    ratios["Hệ số thanh toán tức thời"] = safe_div(tien, no_nh)

    # --- ĐÒN BẨY ---
    tong_ts = extracted.get("Tổng tài sản", {}).get("Cuối kỳ")
    tong_no = extracted.get("Tổng nợ phải trả", {}).get("Cuối kỳ")
    vcsh = extracted.get("Vốn chủ sở hữu", {}).get("Cuối kỳ")
    no_dh = extracted.get("Nợ dài hạn", {}).get("Cuối kỳ")

    ratios["Tỷ lệ nợ/VCSH"] = safe_div(tong_no, vcsh)
    ratios["Tỷ lệ nợ/Tổng tài sản"] = safe_div(tong_no, tong_ts)
    ratios["Đòn bẩy tài chính"] = safe_div(tong_ts, vcsh)

    # --- HIỆU QUẢ ---
    ln_cpt = extracted.get("Lợi nhuận chưa phân phối", {}).get("Cuối kỳ")
    ln_cpt_dk = extracted.get("Lợi nhuận chưa phân phối", {}).get("Đầu kỳ")

    ratios["ROA (tạm tính)"] = safe_div(ln_cpt, tong_ts) if ln_cpt and tong_ts else None
    ratios["ROE (tạm tính)"] = safe_div(ln_cpt, vcsh) if ln_cpt and vcsh else None

    # --- CƠ CẤU ---
    ts_dh = extracted.get("Tài sản dài hạn", {}).get("Cuối kỳ")
    ratios["Tỷ lệ TS ngắn hạn/TS tổng"] = safe_div(ts_nh, tong_ts)
    ratios["Tỷ lệ TS dài hạn/TS tổng"] = safe_div(ts_dh, tong_ts)
    ratios["Tỷ lệ Nợ ngắn hạn/Tổng nợ"] = safe_div(no_nh, tong_no)
    ratios["Tỷ lệ Nợ dài hạn/Tổng nợ"] = safe_div(no_dh, tong_no)

    # Display
    ratio_labels_vi = {
        "Hệ số thanh toán hiện hành": "Thanh khoản hiện hành (TS ngắn hạn / Nợ ngắn hạn)",
        "Hệ số thanh toán nhanh": "Thanh khoản nhanh ((TS ngắn hạn - HTK) / Nợ ngắn hạn)",
        "Hệ số thanh toán tức thời": "Thanh khoản tức thời (Tiền / Nợ ngắn hạn)",
        "Tỷ lệ nợ/VCSH": "Nợ / Vốn chủ sở hữu",
        "Tỷ lệ nợ/Tổng tài sản": "Nợ / Tổng tài sản",
        "Đòn bẩy tài chính": "Tổng tài sản / Vốn chủ sở hữu",
        "ROA (tạm tính)": "Lợi nhuận / Tổng tài sản",
        "ROE (tạm tính)": "Lợi nhuận / Vốn chủ sở hữu",
        "Tỷ lệ TS ngắn hạn/TS tổng": "TS ngắn hạn / Tổng tài sản",
        "Tỷ lệ TS dài hạn/TS tổng": "TS dài hạn / Tổng tài sản",
        "Tỷ lệ Nợ ngắn hạn/Tổng nợ": "Nợ ngắn hạn / Tổng nợ",
        "Tỷ lệ Nợ dài hạn/Tổng nợ": "Nợ dài hạn / Tổng nợ",
    }

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("💧 Thanh khoản")
        for k in ["Hệ số thanh toán hiện hành", "Hệ số thanh toán nhanh", "Hệ số thanh toán tức thời"]:
            v = ratios.get(k)
            label = ratio_labels_vi.get(k, k)
            val_str = f"{v:.2f}" if v is not None else "N/A"
            st.metric(label=label, value=val_str)

    with col2:
        st.subheader("🏗️ Đòn bẩy & Hiệu quả")
        for k in ["Tỷ lệ nợ/VCSH", "Tỷ lệ nợ/Tổng tài sản", "Đòn bẩy tài chính", "ROA (tạm tính)", "ROE (tạm tính)"]:
            v = ratios.get(k)
            label = ratio_labels_vi.get(k, k)
            if v is not None and k not in ["Đòn bẩy tài chính"]:
                val_str = f"{v*100:.1f}%" if k in ["Tỷ lệ nợ/Tổng tài sản", "Tỷ lệ nợ/VCSH", "ROA (tạm tính)", "ROE (tạm tính)"] else f"{v:.2f}"
            elif v is not None:
                val_str = f"{v:.2f}x"
            else:
                val_str = "N/A"
            st.metric(label=label, value=val_str)

    # ============================================================
    # 5. CHARTS
    # ============================================================
    st.header("5️⃣ Biểu đồ")

    chart_cols = st.columns(2)

    # 5a. Balance comparison bar chart
    with chart_cols[0]:
        st.subheader("📊 So sánh Đầu kỳ vs Cuối kỳ")
        fig_bar = go.Figure()
        for _, row in df_extracted.iterrows():
            dk = row["Đầu kỳ"]
            ck = row["Cuối kỳ"]
            if dk is not None or ck is not None:
                fig_bar.add_trace(go.Bar(
                    name=row["Chỉ tiêu"],
                    x=["Đầu kỳ", "Cuối kỳ"],
                    y=[dk or 0, ck or 0],
                ))
        fig_bar.update_layout(barmode="group", height=400, xaxis_title="Kỳ", yaxis_title="Giá trị")
        st.plotly_chart(fig_bar, use_container_width=True)

    # 5b. Asset structure pie chart
    with chart_cols[1]:
        st.subheader("🥧 Cơ cấu tài sản (Cuối kỳ)")
        pie_vals = []
        pie_labels = []
        for label, key in [("TS ngắn hạn", "Tài sản ngắn hạn"), ("TS dài hạn", "Tài sản dài hạn")]:
            v = extracted.get(key, {}).get("Cuối kỳ")
            if v:
                pie_vals.append(v)
                pie_labels.append(label)
        # Add "Khác" if total assets is bigger
        total_pie = sum(pie_vals)
        if tong_ts and total_pie < tong_ts:
            pie_vals.append(tong_ts - total_pie)
            pie_labels.append("Khác")
        if pie_vals:
            fig_pie = px.pie(values=pie_vals, names=pie_labels, hole=0.4)
            fig_pie.update_layout(height=400)
            st.plotly_chart(fig_pie, use_container_width=True)

    # 5c. Ratio radar chart
    st.subheader("🕸️ Radar chỉ số tài chính")
    radar_labels = []
    radar_vals = []
    for k, label in ratio_labels_vi.items():
        v = ratios.get(k)
        if v is not None:
            # Normalize to 0-1 scale for radar
            if "%" in label or k in ["Tỷ lệ nợ/VCSH", "Tỷ lệ nợ/Tổng tài sản", "ROA (tạm tính)", "ROE (tạm tính)"]:
                radar_vals.append(min(v if v < 1 else v, 2.0))  # cap at 2 for display
            elif k == "Đòn bẩy tài chính":
                radar_vals.append(min(v / 5, 1.0))  # normalize leverage
            else:
                radar_vals.append(min(v / 3, 1.0))  # normalize ratios
            radar_labels.append(k)

    if radar_labels:
        fig_radar = go.Figure(data=go.Scatterpolar(
            r=radar_vals,
            theta=radar_labels,
            fill="toself",
            name="Cuối kỳ",
        ))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 1.5])), showlegend=False, height=500)
        st.plotly_chart(fig_radar, use_container_width=True)

    # 5d. Nợ vs VCSH stacked bar
    st.subheader("📈 Cơ cấu Nguồn vốn")
    fig_stack = go.Figure()
    if tong_no and vcsh:
        fig_stack.add_trace(go.Bar(x=["Cuối kỳ"], y=[tong_no], name="Nợ phải trả", marker_color="#ef5545"))
        fig_stack.add_trace(go.Bar(x=["Cuối kỳ"], y=[vcsh], name="Vốn chủ sở hữu", marker_color="#636efa"))
        # Also show đầu kỳ if available
        tong_no_dk = extracted.get("Tổng nợ phải trả", {}).get("Đầu kỳ")
        vcsh_dk = extracted.get("Vốn chủ sở hữu", {}).get("Đầu kỳ")
        if tong_no_dk and vcsh_dk:
            fig_stack.add_trace(go.Bar(x=["Đầu kỳ"], y=[tong_no_dk], name="Nợ phải trả", marker_color="#ef5545", showlegend=False))
            fig_stack.add_trace(go.Bar(x=["Đầu kỳ"], y=[vcsh_dk], name="Vốn chủ sở hữu", marker_color="#636efa", showlegend=False))
        fig_stack.update_layout(barmode="stack", height=400, yaxis_title="Giá trị")
        st.plotly_chart(fig_stack, use_container_width=True)

    # ============================================================
    # 6. EXPORT
    # ============================================================
    st.header("6️⃣ Xuất báo cáo")

    # Build export dataframe
    export_rows = []
    for label, kws in key_items.items():
        row = {"Chỉ tiêu": label}
        row["Đầu kỳ"] = extracted.get(label, {}).get("Đầu kỳ")
        row["Cuối kỳ"] = extracted.get(label, {}).get("Cuối kỳ")
        export_rows.append(row)
    for k, v in ratios.items():
        export_rows.append({"Chỉ tiêu": ratio_labels_vi.get(k, k), "Đầu kỳ": None, "Cuối kỳ": v})

    df_export = pd.DataFrame(export_rows)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv = df_export.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 Tải CSV", data=csv, file_name="bao_cao_tai_chinh.csv", mime="text/csv")
    with col_dl2:
        excel_buffer = pd.ExcelWriter("bao_cao_tai_chinh.xlsx", engine="openpyxl")
        df_export.to_excel(excel_buffer, index=False, sheet_name="Chỉ số tài chính")
        df_extracted.to_excel(excel_buffer, index=False, sheet_name="Chỉ tiêu BCĐKT")
        excel_buffer.close()
        with open("bao_cao_tai_chinh.xlsx", "rb") as f:
            st.download_button("📥 Tải Excel", data=f.read(), file_name="bao_cao_tai_chinh.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("👆 Nhập dữ liệu BCĐKT ở trên để bắt đầu phân tích.")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.markdown("🔧 Built with Streamlit + Plotly | 📊 Phân Tích BCĐKT v1.0")