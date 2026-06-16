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
    page_title="Phân Tích Tài Chính",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Phân Tích Báo Cáo Tài Chính")
st.markdown("Nhập dữ liệu BCĐKT & BCKQHĐKD → Tự động tính chỉ số tài chính → Biểu đồ trực quan")

# ============================================================
# HELPER FUNCTIONS
# ============================================================
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

def safe_div(a, b):
    if a is None or b is None or b == 0:
        return None
    return round(a / b, 4)

def parse_uploaded_file(uploaded):
    """Parse uploaded file (xlsx/xls/csv) into DataFrame."""
    if uploaded is None:
        return None
    try:
        if uploaded.name.lower().endswith(".csv"):
            return pd.read_csv(uploaded, header=None)
        else:
            return pd.read_excel(uploaded, header=None, engine="openpyxl" if uploaded.name.lower().endswith(".xlsx") else None)
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}\n\n💡 Thử lưu file dưới dạng .xlsx (Excel 2007+) rồi upload lại.")
        return None

def detect_columns(df, input_mode, col_start_label="Số đầu kỳ", col_end_label="Số cuối kỳ"):
    """Detect numeric and text columns from DataFrame."""
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    text_cols = [c for c in df.columns if c not in numeric_cols]

    if len(numeric_cols) >= 2:
        col_start = numeric_cols[0]
        col_end = numeric_cols[1]
    elif len(numeric_cols) == 1:
        col_start = col_end = numeric_cols[0]
    else:
        col_start = df.columns[-2] if len(df.columns) >= 2 else df.columns[0]
        col_end = df.columns[-1]

    if input_mode == "✏️ Nhập tay":
        col_start = col_start_label
        col_end = col_end_label
        text_col = "Tên chỉ tiêu"
    else:
        text_col = text_cols[0] if text_cols else df.columns[0]

    return col_start, col_end, text_col

def extract_items(df, key_items, col_start, col_end, text_col=None):
    """Extract key items from DataFrame based on keyword search."""
    extracted = {}
    for label, kws in key_items.items():
        val_start = find_value(df, kws, col_start)
        val_end = find_value(df, kws, col_end)
        extracted[label] = {"Đầu kỳ": val_start, "Cuối kỳ": val_end}
    return extracted

def format_metric(v, k, ratio_labels_vi):
    """Format a ratio value for display."""
    label = ratio_labels_vi.get(k, k)
    if v is None:
        return "N/A"
    if k in ["Đòn bẩy tài chính"]:
        return f"{v:.2f}x"
    if k in ["Tỷ lệ nợ/Tổng tài sản", "Tỷ lệ nợ/VCSH", "ROA", "ROE",
             "Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng",
             "Tỷ lệ nợ/VCSH (BCĐKT)", "Tỷ lệ nợ/Tổng TS (BCĐKT)", "ROA (kết hợp)", "ROE (kết hợp)"]:
        return f"{v*100:.1f}%"
    return f"{v:.2f}"

# ============================================================
# TAB SELECTION
# ============================================================
tab1, tab2, tab3 = st.tabs(["📋 BCĐKT", "📋 BCKQHĐKD", "🔄 Kết hợp"])

# ============================================================
# TAB 1: BCĐKT
# ============================================================
with tab1:
    st.header("Bảng Cân Đối Kế Toán")
    bccl_input_mode = st.radio("Chọn cách nhập dữ liệu BCĐKT:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True, key="bcdkt_mode")

    bcdkt_df = None

    if bccl_input_mode == "📂 Upload file Excel":
        uploaded_bcdkt = st.file_uploader("Chọn file BCĐKT (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="upload_bcdkt")
        bcdkt_df = parse_uploaded_file(uploaded_bcdkt)

    elif bccl_input_mode == "✏️ Nhập tay":
        default_bcdkt = """01,Tổng tài sản,50000,65000
02,Tài sản ngắn hạn,30000,40000
03,Tài sản dài hạn,20000,25000
04,Tiền và tương đương tiền,10000,15000
05,Phải thu ngắn hạn,8000,10000
06,Hàng tồn kho,12000,15000
10,Tổng nợ phải trả,30000,38000
11,Nợ ngắn hạn,18000,22000
12,Nợ dài hạn,12000,16000
30,Vốn chủ sở hữu,20000,27000
31,Vốn điều lệ,15000,15000
32,Lợi nhuận chưa phân phối,5000,12000"""
        text_bcdkt = st.text_area("Dữ liệu BCĐKT", value=default_bcdkt, height=250, key="text_bcdkt")
        try:
            bcdkt_df = pd.read_csv(StringIO(text_bcdkt), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
        except Exception:
            st.warning("Dữ liệu nhập tay chưa hợp lệ.")

    bcdkt_extracted = {}

    if bcdkt_df is not None and not bcdkt_df.empty:
        bcdkt_df.columns = [str(c) for c in bcdkt_df.columns]
        st.subheader("Dữ liệu gốc BCĐKT")
        st.dataframe(bcdkt_df, use_container_width=True)

        col_start_b, col_end_b, text_col_b = detect_columns(bcdkt_df, bccl_input_mode)

        bcdkt_items = {
            "Tổng tài sản": ["tổng tài sản", "tổng ts"],
            "Tài sản ngắn hạn": ["tài sản ngắn hạn", "ts ngắn hạn", "tài sản ngắn"],
            "Tài sản dài hạn": ["tài sản dài hạn", "ts dài hạn", "tài sản dài"],
            "Tiền và tương đương tiền": ["tiền và tương đương", "tiền"],
            "Phải thu ngắn hạn": ["phải thu ngắn hạn", "phải thu khách hàng", "phải thu"],
            "Hàng tồn kho": ["hàng tồn kho", "htk"],
            "Tổng nợ phải trả": ["tổng nợ phải trả", "tổng nợ", "nợ phải trả"],
            "Nợ ngắn hạn": ["nợ ngắn hạn", "nợ ngắn"],
            "Nợ dài hạn": ["nợ dài hạn", "nợ dài"],
            "Vốn chủ sở hữu": ["vốn chủ sở hữu", "vcsh", "vốn chủ"],
            "Lợi nhuận chưa phân phối": ["lợi nhuận chưa phân phối", "ln cpt"],
        }

        bcdkt_extracted = extract_items(bcdkt_df, bcdkt_items, col_start_b, col_end_b)

        df_bcdkt_ext = pd.DataFrame(bcdkt_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"})
        df_bcdkt_ext = df_bcdkt_ext.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

        if df_bcdkt_ext.empty:
            st.warning("⚠️ Không tự động trích được chỉ tiêu BCĐKT. Kiểm tra tên chỉ tiêu trong file.")
        else:
            st.subheader("Trích xuất chỉ tiêu BCĐKT")
            st.dataframe(df_bcdkt_ext.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

            # --- BCĐKT RATIOS ---
            st.subheader("Chỉ số thanh khoản & Đòn bẩy")

            ts_nh = bcdkt_extracted.get("Tài sản ngắn hạn", {}).get("Cuối kỳ")
            no_nh = bcdkt_extracted.get("Nợ ngắn hạn", {}).get("Cuối kỳ")
            tien = bcdkt_extracted.get("Tiền và tương đương tiền", {}).get("Cuối kỳ")
            htk = bcdkt_extracted.get("Hàng tồn kho", {}).get("Cuối kỳ")
            tong_ts = bcdkt_extracted.get("Tổng tài sản", {}).get("Cuối kỳ")
            tong_no = bcdkt_extracted.get("Tổng nợ phải trả", {}).get("Cuối kỳ")
            vcsh = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Cuối kỳ")
            no_dh = bcdkt_extracted.get("Nợ dài hạn", {}).get("Cuối kỳ")
            ts_dh = bcdkt_extracted.get("Tài sản dài hạn", {}).get("Cuối kỳ")

            bcdkt_ratios = {
                "Hệ số thanh toán hiện hành": safe_div(ts_nh, no_nh),
                "Hệ số thanh toán nhanh": safe_div((ts_nh or 0) - (htk or 0), no_nh),
                "Hệ số thanh toán tức thời": safe_div(tien, no_nh),
                "Tỷ lệ nợ/VCSH": safe_div(tong_no, vcsh),
                "Tỷ lệ nợ/Tổng tài sản": safe_div(tong_no, tong_ts),
                "Đòn bẩy tài chính": safe_div(tong_ts, vcsh),
            }

            bcdkt_ratio_labels = {
                "Hệ số thanh toán hiện hành": "Thanh khoản hiện hành",
                "Hệ số thanh toán nhanh": "Thanh khoản nhanh",
                "Hệ số thanh toán tức thời": "Thanh khoản tức thời",
                "Tỷ lệ nợ/VCSH": "Nợ / VCSH",
                "Tỷ lệ nợ/Tổng tài sản": "Nợ / Tổng TS",
                "Đòn bẩy tài chính": "Tổng TS / VCSH",
            }

            c1, c2, c3 = st.columns(3)
            for i, (k, v) in enumerate(bcdkt_ratios.items()):
                col = [c1, c2, c3][i % 3]
                with col:
                    st.metric(label=bcdkt_ratio_labels.get(k, k), value=format_metric(v, k, bcdkt_ratio_labels))

            # --- BCĐKT CHARTS ---
            st.subheader("Biểu đồ BCĐKT")
            chart_c1, chart_c2 = st.columns(2)

            with chart_c1:
                st.markdown("**🥧 Cơ cấu tài sản**")
                pie_vals, pie_labels = [], []
                for label, key in [("TS ngắn hạn", "Tài sản ngắn hạn"), ("TS dài hạn", "Tài sản dài hạn")]:
                    v = bcdkt_extracted.get(key, {}).get("Cuối kỳ")
                    if v:
                        pie_vals.append(v)
                        pie_labels.append(label)
                total_pie = sum(pie_vals)
                if tong_ts and total_pie < tong_ts:
                    pie_vals.append(tong_ts - total_pie)
                    pie_labels.append("Khác")
                if pie_vals:
                    fig_pie = px.pie(values=pie_vals, names=pie_labels, hole=0.4)
                    st.plotly_chart(fig_pie, use_container_width=True)

            with chart_c2:
                st.markdown("**📈 Nguồn vốn**")
                if tong_no and vcsh:
                    tong_no_dk = bcdkt_extracted.get("Tổng nợ phải trả", {}).get("Đầu kỳ")
                    vcsh_dk = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Đầu kỳ")
                    fig_stack = go.Figure()
                    fig_stack.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[tong_no_dk or 0, tong_no], name="Nợ phải trả", marker_color="#ef5545"))
                    fig_stack.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[vcsh_dk or 0, vcsh], name="VCSH", marker_color="#636efa"))
                    fig_stack.update_layout(barmode="stack", height=350, yaxis_title="Giá trị")
                    st.plotly_chart(fig_stack, use_container_width=True)

# ============================================================
# TAB 2: BCKQHĐKD
# ============================================================
with tab2:
    st.header("Báo Cáo Kết Quả Hoạt Động Kinh Doanh")
    bkq_mode = st.radio("Chọn cách nhập dữ liệu BCKQHĐKD:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True, key="bkq_mode")

    bkq_df = None

    if bkq_mode == "📂 Upload file Excel":
        uploaded_bkq = st.file_uploader("Chọn file BCKQHĐKD (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="upload_bkq")
        bkq_df = parse_uploaded_file(uploaded_bkq)

    elif bkq_mode == "✏️ Nhập tay":
        default_bkq = """01,Doanh thu thuần,120000,150000
02,Doanh thu hoạt động tài chính,5000,6000
03,Thu nhập khác,2000,3000
10,Chi phí vốn hàng bán,72000,90000
11,Chi phí bán hàng,12000,15000
12,Chi phí quản lý doanh nghiệp,10000,12000
13,Chi phí tài chính,4000,5000
14,Chi phí khác,1500,2000
20,Lợi nhuận gộp,48000,45000
21,Lợi nhuận thuần từ HĐKD,26000,25000
22,Lợi nhuận sau thuế,20000,22000
23,Thuế TNDN,6000,6600"""
        text_bkq = st.text_area("Dữ liệu BCKQHĐKD", value=default_bkq, height=300, key="text_bkq")
        try:
            bkq_df = pd.read_csv(StringIO(text_bkq), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
        except Exception:
            st.warning("Dữ liệu nhập tay chưa hợp lệ.")

    bkq_extracted = {}

    if bkq_df is not None and not bkq_df.empty:
        bkq_df.columns = [str(c) for c in bkq_df.columns]
        st.subheader("Dữ liệu gốc BCKQHĐKD")
        st.dataframe(bkq_df, use_container_width=True)

        col_start_kq, col_end_kq, text_col_kq = detect_columns(bkq_df, bkq_mode)

        bkq_items = {
            "Doanh thu thuần": ["doanh thu thuần", "doanh thu thuần về bán hàng", "dt thuần"],
            "Doanh thu hoạt động tài chính": ["doanh thu hoạt động tài chính", "dt tài chính", "doanh thu ht tài chính"],
            "Thu nhập khác": ["thu nhập khác", "thu nhập"],
            "Chi phí vốn hàng bán": ["chi phí vốn hàng bán", "giá vốn hàng bán", "giá vốn", "cp vốn hàng bán"],
            "Chi phí bán hàng": ["chi phí bán hàng", "cp bán hàng"],
            "Chi phí quản lý doanh nghiệp": ["chi phí quản lý doanh nghiệp", "cp quản lý dn", "chi phí quản lý"],
            "Chi phí tài chính": ["chi phí tài chính", "cp tài chính"],
            "Chi phí khác": ["chi phí khác", "cp khác"],
            "Lợi nhuận gộp": ["lợi nhuận gộp", "ln gộp", "lợi nhuận gộp về bán hàng"],
            "Lợi nhuận thuần từ HĐKD": ["lợi nhuận thuần từ hdkd", "ln thuần từ hdkd", "lợi nhuận từ hdkd"],
            "Lợi nhuận sau thuế": ["lợi nhuận sau thuế", "ln sau thuế", "lợi nhuận kế toán sau thuế"],
            "Thuế TNDN": ["thuế tndn", "thuế thu nhập doanh nghiệp", "thuế tndn hoãn lại"],
            "Tổng doanh thu": ["tổng doanh thu", "doanh thu"],
        }

        bkq_extracted = extract_items(bkq_df, bkq_items, col_start_kq, col_end_kq)

        # Also try "kỳ trước" / "kỳ này" column detection for income statements
        df_bkq_ext = pd.DataFrame(bkq_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"})
        df_bkq_ext = df_bkq_ext.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

        if df_bkq_ext.empty:
            st.warning("⚠️ Không tự động trích được chỉ tiêu BCKQHĐKD. Kiểm tra tên chỉ tiêu trong file.")
        else:
            st.subheader("Trích xuất chỉ tiêu BCKQHĐKD")
            st.dataframe(df_bkq_ext.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

            # --- BCKQHĐKD RATIOS ---
            st.subheader("Chỉ số sinh lời & Hiệu quả")

            dt_thuan = bkq_extracted.get("Doanh thu thuần", {}).get("Cuối kỳ")
            dt_thuan_dk = bkq_extracted.get("Doanh thu thuần", {}).get("Đầu kỳ")
            ln_gop = bkq_extracted.get("Lợi nhuận gộp", {}).get("Cuối kỳ")
            ln_gop_dk = bkq_extracted.get("Lợi nhuận gộp", {}).get("Đầu kỳ")
            cp_von = bkq_extracted.get("Chi phí vốn hàng bán", {}).get("Cuối kỳ")
            cp_ban = bkq_extracted.get("Chi phí bán hàng", {}).get("Cuối kỳ")
            cp_qly = bkq_extracted.get("Chi phí quản lý doanh nghiệp", {}).get("Cuối kỳ")
            cp_tc = bkq_extracted.get("Chi phí tài chính", {}).get("Cuối kỳ")
            ln_hdkd = bkq_extracted.get("Lợi nhuận thuần từ HĐKD", {}).get("Cuối kỳ")
            ln_st = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Cuối kỳ")
            ln_st_dk = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Đầu kỳ")
            dt_httc = bkq_extracted.get("Doanh thu hoạt động tài chính", {}).get("Cuối kỳ")

            bkq_ratios = {
                "Biên lợi nhuận gộp": safe_div(ln_gop, dt_thuan),
                "Biên lợi nhuận hoạt động": safe_div(ln_hdkd, dt_thuan),
                "Biên lợi nhuận ròng": safe_div(ln_st, dt_thuan),
                "Tỷ suất doanh thu / Vốn chủ sở hữu": safe_div(dt_thuan, bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Cuối kỳ")),
                "Tỷ lệ chi phí vốn hàng bán / Doanh thu": safe_div(cp_von, dt_thuan),
                "Tỷ lệ chi phí bán hàng / Doanh thu": safe_div(cp_ban, dt_thuan),
                "Tỷ lệ chi phí quản lý / Doanh thu": safe_div(cp_qly, dt_thuan),
                "Tăng trưởng doanh thu": safe_div((dt_thuan or 0) - (dt_thuan_dk or 0), dt_thuan_dk),
                "Tăng trưởng lợi nhuận sau thuế": safe_div((ln_st or 0) - (ln_st_dk or 0), ln_st_dk),
            }

            bkq_ratio_labels = {
                "Biên lợi nhuận gộp": "LN gộp / DT thuần",
                "Biên lợi nhuận hoạt động": "LN HĐKD / DT thuần",
                "Biên lợi nhuận ròng": "LN sau thuế / DT thuần",
                "Tỷ suất doanh thu / Vốn chủ sở hữu": "DT thuần / VCSH",
                "Tỷ lệ chi phí vốn hàng bán / Doanh thu": "CP vốn hàng bán / DT",
                "Tỷ lệ chi phí bán hàng / Doanh thu": "CP bán hàng / DT",
                "Tỷ lệ chi phí quản lý / Doanh thu": "CP quản lý / DT",
                "Tăng trưởng doanh thu": "(DT cuối kỳ - DT đầu kỳ) / DT đầu kỳ",
                "Tăng trưởng lợi nhuận sau thuế": "(LN cuối - LN đầu) / LN đầu",
            }

            c1, c2, c3 = st.columns(3)
            for i, (k, v) in enumerate(bkq_ratios.items()):
                col = [c1, c2, c3][i % 3]
                with col:
                    val_str = format_metric(v, k, bkq_ratio_labels)
                    st.metric(label=bkq_ratio_labels.get(k, k), value=val_str)

            # --- BCKQHĐKD CHARTS ---
            st.subheader("Biểu đồ BCKQHĐKD")
            chart_b1, chart_b2 = st.columns(2)

            with chart_b1:
                st.markdown("**📊 Cơ cấu chi phí**")
                cp_vals = []
                cp_labels = []
                for label, key in [("CP vốn hàng bán", "Chi phí vốn hàng bán"), ("CP bán hàng", "Chi phí bán hàng"),
                                   ("CP quản lý", "Chi phí quản lý doanh nghiệp"), ("CP tài chính", "Chi phí tài chính")]:
                    v = bkq_extracted.get(key, {}).get("Cuối kỳ")
                    if v:
                        cp_vals.append(v)
                        cp_labels.append(label)
                if cp_vals:
                    fig_cp = px.pie(values=cp_vals, names=cp_labels, hole=0.4, color_discrete_sequence=px.colors.sequential.Reds)
                    st.plotly_chart(fig_cp, use_container_width=True)

            with chart_b2:
                st.markdown("**📈 So sánh DT & LN**")
                fig_dtln = go.Figure()
                if dt_thuan or dt_thuan_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[dt_thuan_dk or 0, dt_thuan or 0], name="Doanh thu thuần", marker_color="#636efa"))
                if ln_st or ln_st_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[ln_st_dk or 0, ln_st or 0], name="LN sau thuế", marker_color="#2ca02c"))
                if ln_gop or ln_gop_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[ln_gop_dk or 0, ln_gop or 0], name="LN gộp", marker_color="#ff7f0e"))
                fig_dtln.update_layout(barmode="group", height=350, yaxis_title="Giá trị")
                st.plotly_chart(fig_dtln, use_container_width=True)

            # Biên lợi nhuận trend
            st.markdown("**📉 Biên lợi nhuận**")
            margin_vals = []
            margin_labels = []
            for k in ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng"]:
                v = bkq_ratios.get(k)
                if v is not None:
                    margin_vals.append(v * 100)
                    margin_labels.append(k.replace("Biên lợi nhuận ", "").title())
            if margin_vals:
                fig_margin = go.Figure(data=[go.Bar(x=margin_labels, y=margin_vals, marker_color=["#636efa", "#ff7f0e", "#2ca02c"])])
                fig_margin.update_layout(yaxis_title="%", height=350, yaxis=dict(ticksuffix="%"))
                st.plotly_chart(fig_margin, use_container_width=True)

# ============================================================
# TAB 3: COMBINED ANALYSIS
# ============================================================
with tab3:
    st.header("🔄 Phân Tích Kết Hợp BCĐKT + BCKQHĐKD")
    st.markdown("Kết hợp dữ liệu từ cả 2 báo cáo để tính các chỉ số tài chính tổng hợp.")

    has_bcdkt = bool(bcdkt_extracted)
    has_bkq = bool(bkq_extracted)

    if not has_bcdkt and not has_bkq:
        st.info("👆 Nhập dữ liệu BCĐKT và/hoặc BCKQHĐKD ở các tab trên để phân tích kết hợp.")
    else:
        # Get values from both
        tong_ts = bcdkt_extracted.get("Tổng tài sản", {}).get("Cuối kỳ")
        tong_ts_dk = bcdkt_extracted.get("Tổng tài sản", {}).get("Đầu kỳ")
        vcsh = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Cuối kỳ")
        vcsh_dk = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Đầu kỳ")
        tong_no = bcdkt_extracted.get("Tổng nợ phải trả", {}).get("Cuối kỳ")
        ts_nh = bcdkt_extracted.get("Tài sản ngắn hạn", {}).get("Cuối kỳ")
        no_nh = bcdkt_extracted.get("Nợ ngắn hạn", {}).get("Cuối kỳ")
        tien = bcdkt_extracted.get("Tiền và tương đương tiền", {}).get("Cuối kỳ")
        htk = bcdkt_extracted.get("Hàng tồn kho", {}).get("Cuối kỳ")
        pt_nh = bcdkt_extracted.get("Phải thu ngắn hạn", {}).get("Cuối kỳ")
        no_dh = bcdkt_extracted.get("Nợ dài hạn", {}).get("Cuối kỳ")
        ts_dh = bcdkt_extracted.get("Tài sản dài hạn", {}).get("Cuối kỳ")
        ts_nh_dk = bcdkt_extracted.get("Tài sản ngắn hạn", {}).get("Đầu kỳ")
        no_nh_dk = bcdkt_extracted.get("Nợ ngắn hạn", {}).get("Đầu kỳ")

        dt_thuan = bkq_extracted.get("Doanh thu thuần", {}).get("Cuối kỳ")
        dt_thuan_dk = bkq_extracted.get("Doanh thu thuần", {}).get("Đầu kỳ")
        ln_gop = bkq_extracted.get("Lợi nhuận gộp", {}).get("Cuối kỳ")
        ln_hdkd = bkq_extracted.get("Lợi nhuận thuần từ HĐKD", {}).get("Cuối kỳ")
        ln_st = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Cuối kỳ")
        ln_st_dk = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Đầu kỳ")
        cp_von = bkq_extracted.get("Chi phí vốn hàng bán", {}).get("Cuối kỳ")
        cp_ban = bkq_extracted.get("Chi phí bán hàng", {}).get("Cuối kỳ")
        cp_qly = bkq_extracted.get("Chi phí quản lý doanh nghiệp", {}).get("Cuối kỳ")

        # AVERAGE ASSETS (for ROA/ROE with income statement)
        ts_tb = safe_div((tong_ts or 0) + (tong_ts_dk or 0), 2) if tong_ts or tong_ts_dk else None
        vcsh_tb = safe_div((vcsh or 0) + (vcsh_dk or 0), 2) if vcsh or vcsh_dk else None

        combined_ratios = {}
        combined_labels = {}

        # Thanh khoản (from BCĐKT)
        if has_bcdkt:
            combined_ratios["Hệ số thanh toán hiện hành"] = safe_div(ts_nh, no_nh)
            combined_ratios["Hệ số thanh toán nhanh"] = safe_div((ts_nh or 0) - (htk or 0), no_nh)
            combined_ratios["Hệ số thanh toán tức thời"] = safe_div(tien, no_nh)
            combined_labels["Hệ số thanh toán hiện hành"] = "TS ngắn hạn / Nợ ngắn hạn"
            combined_labels["Hệ số thanh toán nhanh"] = "(TS ngắn hạn - HTK) / Nợ ngắn hạn"
            combined_labels["Hệ số thanh toán tức thời"] = "Tiền / Nợ ngắn hạn"

            combined_ratios["Tỷ lệ nợ/VCSH"] = safe_div(tong_no, vcsh)
            combined_ratios["Tỷ lệ nợ/Tổng tài sản"] = safe_div(tong_no, tong_ts)
            combined_ratios["Đòn bẩy tài chính"] = safe_div(tong_ts, vcsh)
            combined_labels["Tỷ lệ nợ/VCSH"] = "Nợ / VCSH"
            combined_labels["Tỷ lệ nợ/Tổng tài sản"] = "Nợ / Tổng TS"
            combined_labels["Đòn bẩy tài chính"] = "Tổng TS / VCSH"

        # Sinh lời (from BCKQHĐKD + BCĐKT)
        if has_bkq:
            combined_ratios["Biên lợi nhuận gộp"] = safe_div(ln_gop, dt_thuan)
            combined_ratios["Biên lợi nhuận hoạt động"] = safe_div(ln_hdkd, dt_thuan)
            combined_ratios["Biên lợi nhuận ròng"] = safe_div(ln_st, dt_thuan)
            combined_labels["Biên lợi nhuận gộp"] = "LN gộp / DT thuần"
            combined_labels["Biên lợi nhuận hoạt động"] = "LN HĐKD / DT thuần"
            combined_labels["Biên lợi nhuận ròng"] = "LN sau thuế / DT thuần"

        # Hiệu quả sử dụng vốn (needs both)
        if has_bcdkt and has_bkq:
            combined_ratios["ROA"] = safe_div(ln_st, ts_tb)
            combined_ratios["ROE"] = safe_div(ln_st, vcsh_tb)
            combined_labels["ROA"] = "LN sau thuế / TS trung bình"
            combined_labels["ROE"] = "LN sau thuế / VCSH trung bình"

            combined_ratios["Vòng quay tài sản"] = safe_div(dt_thuan, ts_tb)
            combined_labels["Vòng quay tài sản"] = "DT thuần / TS trung bình"

            combined_ratios["Vòng quay vốn chủ sở hữu"] = safe_div(dt_thuan, vcsh_tb)
            combined_labels["Vòng quay vốn chủ sở hữu"] = "DT thuần / VCSH trung bình"

            combined_ratios["Vòng quay hàng tồn kho"] = safe_div(cp_von, htk)
            combined_labels["Vòng quay hàng tồn kho"] = "CP vốn hàng bán / HTK"

            combined_ratios["Vòng quay khoản phải thu"] = safe_div(dt_thuan, pt_nh)
            combined_labels["Vòng quay khoản phải thu"] = "DT thuần / Phải thu ngắn hạn"

        # Hiệu quả chi phí
        if has_bkq:
            combined_ratios["Tỷ lệ CP vốn hàng bán / DT"] = safe_div(cp_von, dt_thuan)
            combined_ratios["Tỷ lệ CP bán hàng / DT"] = safe_div(cp_ban, dt_thuan)
            combined_ratios["Tỷ lệ CP quản lý / DT"] = safe_div(cp_qly, dt_thuan)
            combined_labels["Tỷ lệ CP vốn hàng bán / DT"] = "CP vốn / DT thuần"
            combined_labels["Tỷ lệ CP bán hàng / DT"] = "CP bán hàng / DT thuần"
            combined_labels["Tỷ lệ CP quản lý / DT"] = "CP quản lý / DT thuần"

        # Tăng trưởng
        if has_bkq and dt_thuan_dk:
            combined_ratios["Tăng trưởng doanh thu"] = safe_div((dt_thuan or 0) - (dt_thuan_dk or 0), dt_thuan_dk)
            combined_labels["Tăng trưởng doanh thu"] = "(DT cuối - DT đầu) / DT đầu"
        if has_bkq and ln_st_dk:
            combined_ratios["Tăng trưởng LN sau thuế"] = safe_div((ln_st or 0) - (ln_st_dk or 0), ln_st_dk)
            combined_labels["Tăng trưởng LN sau thuế"] = "(LN cuối - LN đầu) / LN đầu"

        # DISPLAY COMBINED RATIOS
        if combined_ratios:
            st.subheader("📋 Bảng tổng hợp chỉ số tài chính")

            # Group ratios
            thanh_khoan = [k for k in combined_ratios if k in ["Hệ số thanh toán hiện hành", "Hệ số thanh toán nhanh", "Hệ số thanh toán tức thời"]]
            don_bay = [k for k in combined_ratios if k in ["Tỷ lệ nợ/VCSH", "Tỷ lệ nợ/Tổng tài sản", "Đòn bẩy tài chính"]]
            sinh_loi = [k for k in combined_ratios if k in ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng", "ROA", "ROE"]]
            hieu_qua = [k for k in combined_ratios if k in ["Vòng quay tài sản", "Vòng quay vốn chủ sở hữu", "Vòng quay hàng tồn kho", "Vòng quay khoản phải thu"]]
            chi_phi = [k for k in combined_ratios if k.startswith("Tỷ lệ CP")]
            tang_truong = [k for k in combined_ratios if k.startswith("Tăng trưởng")]

            groups = [
                ("💧 Thanh khoản", thanh_khoan),
                ("🏗️ Đòn bẩy", don_bay),
                ("💰 Sinh lời", sinh_loi),
                ("⚡ Hiệu quả", hieu_qua),
                ("📊 Cơ cấu chi phí", chi_phi),
                ("📈 Tăng trưởng", tang_truong),
            ]

            for title, keys in groups:
                if not keys:
                    continue
                st.markdown(f"**{title}**")
                cols = st.columns(min(len(keys), 3))
                for i, k in enumerate(keys):
                    v = combined_ratios.get(k)
                    with cols[i % 3]:
                        val_str = format_metric(v, k, combined_labels)
                        st.metric(label=k, value=val_str)

            # COMBINED CHARTS
            st.subheader("📊 Biểu đồ tổng hợp")

            chart_g1, chart_g2 = st.columns(2)

            with chart_g1:
                # Du Pont decomposition
                if has_bcdkt and has_bkq:
                    st.markdown("**🔺 Phân tích Du Pont**")
                    margin = combined_ratios.get("Biên lợi nhuận ròng")
                    turnover = combined_ratios.get("Vòng quay tài sản")
                    leverage = combined_ratios.get("Đòn bẩy tài chính")
                    roe = combined_ratios.get("ROE")

                    fig_dupont = go.Figure()
                    labels = ["Biên LN ròng", "Vòng quay TS", "Đòn bẩy", "ROE"]
                    vals = [margin or 0, turnover or 0, leverage or 0, roe or 0]
                    # Normalize for display
                    vals_pct = [f"{(margin or 0)*100:.1f}%", f"{turnover or 0:.2f}x", f"{leverage or 0:.2f}x", f"{(roe or 0)*100:.1f}%"]
                    fig_dupont.add_trace(go.Bar(x=labels, y=vals, text=vals_pct, textposition="auto", marker_color=["#636efa", "#ff7f0e", "#2ca02c", "#ef5545"]))
                    fig_dupont.update_layout(height=350, yaxis_title="Giá trị")
                    st.plotly_chart(fig_dupont, use_container_width=True)

            with chart_g2:
                # Ratio radar
                st.markdown("**🕸️ Radar chỉ số tổng hợp**")
                radar_labels = []
                radar_vals = []
                for k, v in combined_ratios.items():
                    if v is not None:
                        radar_labels.append(k)
                        if k in ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng", "ROA", "ROE",
                                   "Tỷ lệ nợ/Tổng tài sản", "Tỷ lệ nợ/VCSH",
                                   "Tỷ lệ CP vốn hàng bán / DT", "Tỷ lệ CP bán hàng / DT", "Tỷ lệ CP quản lý / DT",
                                   "Tăng trưởng doanh thu", "Tăng trưởng LN sau thuế"]:
                            radar_vals.append(min(abs(v), 2.0))
                        elif k == "Đòn bẩy tài chính":
                            radar_vals.append(min(v / 5, 1.5))
                        elif k.startswith("Vòng quay"):
                            radar_vals.append(min(v / 5, 1.5))
                        else:
                            radar_vals.append(min(abs(v) / 3, 1.5))
                if radar_labels:
                    fig_radar = go.Figure(data=go.Scatterpolar(
                        r=radar_vals, theta=radar_labels, fill="toself", name="Cuối kỳ"
                    ))
                    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 1.5])), showlegend=False, height=400)
                    st.plotly_chart(fig_radar, use_container_width=True)

            # Growth comparison
            if has_bkq and (dt_thuan_dk or ln_st_dk):
                st.markdown("**📈 So sánh Đầu kỳ vs Cuối kỳ**")
                fig_growth = go.Figure()
                items_to_chart = {}
                for label, key in [("Doanh thu thuần", "Doanh thu thuần"), ("LN gộp", "Lợi nhuận gộp"), ("LN sau thuế", "Lợi nhuận sau thuế")]:
                    dk = bkq_extracted.get(key, {}).get("Đầu kỳ")
                    ck = bkq_extracted.get(key, {}).get("Cuối kỳ")
                    if dk or ck:
                        items_to_chart[label] = (dk, ck)
                # Add BCĐKT items if available
                if has_bcdkt:
                    for label, key in [("Tổng tài sản", "Tổng tài sản"), ("Vốn CSH", "Vốn chủ sở hữu"), ("Tổng nợ", "Tổng nợ phải trả")]:
                        dk = bcdkt_extracted.get(key, {}).get("Đầu kỳ")
                        ck = bcdkt_extracted.get(key, {}).get("Cuối kỳ")
                        if dk or ck:
                            items_to_chart[label] = (dk, ck)

                for label, (dk, ck) in items_to_chart.items():
                    fig_growth.add_trace(go.Bar(name=label, x=["Đầu kỳ", "Cuối kỳ"], y=[dk or 0, ck or 0]))
                fig_growth.update_layout(barmode="group", height=400, yaxis_title="Giá trị")
                st.plotly_chart(fig_growth, use_container_width=True)

        # EXPORT COMBINED
        st.header("📥 Xuất báo cáo tổng hợp")

        export_rows = []
        # BCĐKT items
        if has_bcdkt:
            for label in ["Tổng tài sản", "Tài sản ngắn hạn", "Tài sản dài hạn", "Tiền và tương đương tiền",
                         "Phải thu ngắn hạn", "Hàng tồn kho", "Tổng nợ phải trả", "Nợ ngắn hạn", "Nợ dài hạn",
                         "Vốn chủ sở hữu", "Lợi nhuận chưa phân phối"]:
                if label in bcdkt_extracted:
                    export_rows.append({"Chỉ tiêu": f"[BCĐKT] {label}", "Đầu kỳ": bcdkt_extracted[label].get("Đầu kỳ"), "Cuối kỳ": bcdkt_extracted[label].get("Cuối kỳ"), "Ghi chú": ""})
        # BCKQHĐKD items
        if has_bkq:
            for label in ["Doanh thu thuần", "Doanh thu hoạt động tài chính", "Chi phí vốn hàng bán",
                         "Chi phí bán hàng", "Chi phí quản lý doanh nghiệp", "Chi phí tài chính", "Chi phí khác",
                         "Lợi nhuận gộp", "Lợi nhuận thuần từ HĐKD", "Lợi nhuận sau thuế", "Thuế TNDN"]:
                if label in bkq_extracted:
                    export_rows.append({"Chỉ tiêu": f"[BCKQHĐKD] {label}", "Đầu kỳ": bkq_extracted[label].get("Đầu kỳ"), "Cuối kỳ": bkq_extracted[label].get("Cuối kỳ"), "Ghi chú": ""})
        # Ratios
        export_rows.append({"Chỉ tiêu": "--- CHỈ SỐ TÀI CHÍNH ---", "Đầu kỳ": "", "Cuối kỳ": "", "Ghi chú": ""})
        for k, v in combined_ratios.items():
            export_rows.append({"Chỉ tiêu": k, "Đầu kỳ": "", "Cuối kỳ": v, "Ghi chú": combined_labels.get(k, "")})

        df_export = pd.DataFrame(export_rows)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            csv = df_export.to_csv(index=False). encode("utf-8-sig")
            st.download_button("📥 Tải CSV tổng hợp", data=csv, file_name="bao_cao_tai_chinh_tong_hop.csv", mime="text/csv")
        with col_dl2:
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    df_export.to_excel(writer, index=False, sheet_name="Tổng hợp")
                    if has_bcdkt:
                        pd.DataFrame(bcdkt_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"}).to_excel(writer, index=False, sheet_name="BCĐKT")
                    if has_bkq:
                        pd.DataFrame(bkq_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"}).to_excel(writer, index=False, sheet_name="BCKQHĐKD")
                with open(tmp.name, "rb") as f:
                    st.download_button("📥 Tải Excel tổng hợp", data=f.read(), file_name="bao_cao_tai_chinh_tong_hop.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.markdown("🔧 Built with Streamlit + Plotly | 📊 Phân Tích Tài Chính v2.0")