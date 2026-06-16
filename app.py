import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import StringIO
import numpy as np
import tempfile

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Phân Tích Tài Chính",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Phân Tích Báo Cáo Tài Chính")
st.markdown("Nhập dữ liệu Bảng cân đối kế toán, Báo cáo kết quả hoạt động kinh doanh, Báo cáo lưu chuyển tiền tệ → Tự động tính chỉ số tài chính → Biểu đồ trực quan")

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def find_value(df, keywords, col_name):
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

def extract_items(df, key_items, col_start, col_end):
    extracted = {}
    for label, kws in key_items.items():
        val_start = find_value(df, kws, col_start)
        val_end = find_value(df, kws, col_end)
        extracted[label] = {"Đầu kỳ": val_start, "Cuối kỳ": val_end}
    return extracted

# ============================================================
# FORMULA DICTIONARY - hover tooltips for every ratio
# ============================================================
FORMULAS = {
    # BCĐKT
    "Hệ số thanh toán hiện hành": "Tài sản ngắn hạn / Nợ ngắn hạn",
    "Hệ số thanh toán nhanh": "(Tài sản ngắn hạn - Hàng tồn kho) / Nợ ngắn hạn",
    "Hệ số thanh toán tức thời": "Tiền và tương đương tiền / Nợ ngắn hạn",
    "Tỷ lệ nợ trên vốn chủ sở hữu": "Tổng nợ phải trả / Vốn chủ sở hữu",
    "Tỷ lệ nợ trên tổng tài sản": "Tổng nợ phải trả / Tổng tài sản",
    "Đòn bẩy tài chính": "Tổng tài sản / Vốn chủ sở hữu",
    # BCKQHĐKD
    "Biên lợi nhuận gộp": "Lợi nhuận gộp / Doanh thu thuần × 100%",
    "Biên lợi nhuận hoạt động": "Lợi nhuận thuần từ hoạt động kinh doanh / Doanh thu thuần × 100%",
    "Biên lợi nhuận ròng": "Lợi nhuận sau thuế / Doanh thu thuần × 100%",
    "Tỷ lệ chi phí vốn hàng bán trên doanh thu": "Chi phí vốn hàng bán / Doanh thu thuần × 100%",
    "Tỷ lệ chi phí bán hàng trên doanh thu": "Chi phí bán hàng / Doanh thu thuần × 100%",
    "Tỷ lệ chi phí quản lý trên doanh thu": "Chi phí quản lý doanh nghiệp / Doanh thu thuần × 100%",
    "Tăng trưởng doanh thu": "(Doanh thu cuối kỳ - Doanh thu đầu kỳ) / Doanh thu đầu kỳ × 100%",
    "Tăng trưởng lợi nhuận sau thuế": "(Lợi nhuận cuối kỳ - Lợi nhuận đầu kỳ) / Lợi nhuận đầu kỳ × 100%",
    # BCLCTT
    "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần": "Lưu chuyển tiền từ hoạt động kinh doanh / Doanh thu thuần",
    "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế": "Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế",
    "Tỷ lệ lưu chuyển tiền hoạt động đầu tư trên tổng lưu chuyển": "Lưu chuyển tiền từ hoạt động đầu tư / Lưu chuyển tiền thuần",
    "Tỷ lệ lưu chuyển tiền hoạt động tài chính trên tổng lưu chuyển": "Lưu chuyển tiền từ hoạt động tài chính / Lưu chuyển tiền thuần",
    # Combined
    "ROA": "Lợi nhuận sau thuế / Tổng tài sản trung bình × 100%\n(Tổng tài sản trung bình = (Đầu kỳ + Cuối kỳ) / 2)",
    "ROE": "Lợi nhuận sau thuế / Vốn chủ sở hữu trung bình × 100%\n(Vốn chủ sở hữu trung bình = (Đầu kỳ + Cuối kỳ) / 2)",
    "Vòng quay tài sản": "Doanh thu thuần / Tổng tài sản trung bình",
    "Vòng quay vốn chủ sở hữu": "Doanh thu thuần / Vốn chủ sở hữu trung bình",
    "Vòng quay hàng tồn kho": "Chi phí vốn hàng bán / Hàng tồn kho",
    "Vòng quay khoản phải thu": "Doanh thu thuần / Phải thu ngắn hạn",
    "Dòng tiền tự do": "Lưu chuyển tiền từ hoạt động kinh doanh + Lưu chuyển tiền từ hoạt động đầu tư",
}

def fmt_pct(v, k):
    """Format ratio for display - all labels are now full Vietnamese."""
    if v is None:
        return "N/A"
    pct_keys = ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng",
                "ROA", "ROE", "Tỷ lệ nợ trên vốn chủ sở hữu", "Tỷ lệ nợ trên tổng tài sản",
                "Tỷ lệ chi phí vốn hàng bán trên doanh thu", "Tỷ lệ chi phí bán hàng trên doanh thu",
                "Tỷ lệ chi phí quản lý trên doanh thu",
                "Tăng trưởng doanh thu", "Tăng trưởng lợi nhuận sau thuế",
                "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần",
                "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế",
                "Tỷ lệ lưu chuyển tiền hoạt động đầu tư trên tổng lưu chuyển",
                "Tỷ lệ lưu chuyển tiền hoạt động tài chính trên tổng lưu chuyển"]
    if k == "Đòn bẩy tài chính":
        return f"{v:.2f}x"
    if k in pct_keys:
        return f"{v*100:.1f}%"
    return f"{v:.2f}"

# ============================================================
# TAB SELECTION
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs(["📋 Bảng cân đối kế toán", "📋 Báo cáo kết quả KDKD", "💲 Báo cáo lưu chuyển tiền tệ", "🔄 Kết hợp"])

# ============================================================
# TAB 1: BCĐKT
# ============================================================
with tab1:
    st.header("Bảng Cân Đối Kế Toán")
    bccl_input_mode = st.radio("Chọn cách nhập dữ liệu:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True, key="bcdkt_mode")
    bcdkt_df = None

    if bccl_input_mode == "📂 Upload file Excel":
        uploaded_bcdkt = st.file_uploader("Chọn file Bảng cân đối kế toán (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="upload_bcdkt")
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
        text_bcdkt = st.text_area("Dữ liệu Bảng cân đối kế toán", value=default_bcdkt, height=250, key="text_bcdkt")
        try:
            bcdkt_df = pd.read_csv(StringIO(text_bcdkt), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
        except Exception:
            st.warning("Dữ liệu nhập tay chưa hợp lệ.")

    bcdkt_extracted = {}

    if bcdkt_df is not None and not bcdkt_df.empty:
        bcdkt_df.columns = [str(c) for c in bcdkt_df.columns]
        st.subheader("Dữ liệu gốc")
        st.dataframe(bcdkt_df, use_container_width=True)

        col_start_b, col_end_b, _ = detect_columns(bcdkt_df, bccl_input_mode)

        bcdkt_items = {
            "Tổng tài sản": ["tổng tài sản", "tổng ts", "tổng ts hợp nhất"],
            "Tài sản ngắn hạn": ["tài sản ngắn hạn", "ts ngắn hạn", "tài sản ngắn", "tsngắn hạn"],
            "Tài sản dài hạn": ["tài sản dài hạn", "ts dài hạn", "tài sản dài", "tsdài hạn"],
            "Tiền và tương đương tiền": ["tiền và tương đương", "tiền", "tiền và tương đương tiền", "tiền mặt", "tiền gửi ngân hàng"],
            "Phải thu ngắn hạn": ["phải thu ngắn hạn", "phải thu khách hàng", "phải thu", "phải thu ngắn hạn khác", "khoản phải thu ngắn hạn"],
            "Hàng tồn kho": ["hàng tồn kho", "htk", "hàng tồn kho ròng"],
            "Tổng nợ phải trả": ["tổng nợ phải trả", "tổng nợ", "nợ phải trả", "tổng nợ phải trả hợp nhất"],
            "Nợ ngắn hạn": ["nợ ngắn hạn", "nợ ngắn", "nợ và phải trả ngắn hạn"],
            "Nợ dài hạn": ["nợ dài hạn", "nợ dài", "nợ và phải trả dài hạn"],
            "Vốn chủ sở hữu": ["vốn chủ sở hữu", "vcsh", "vốn chủ", "vốn chủ sở hữu hợp nhất"],
            "Lợi nhuận chưa phân phối": ["lợi nhuận chưa phân phối", "ln cpt", "lợi nhuận sau thuế chưa phân phối"],
        }

        bcdkt_extracted = extract_items(bcdkt_df, bcdkt_items, col_start_b, col_end_b)

        df_bcdkt_ext = pd.DataFrame(bcdkt_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"})
        df_bcdkt_ext = df_bcdkt_ext.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

        if df_bcdkt_ext.empty:
            st.warning("⚠️ Không tự động trích được chỉ tiêu. Kiểm tra tên chỉ tiêu trong file.")
        else:
            st.subheader("Trích xuất chỉ tiêu chính")
            st.dataframe(df_bcdkt_ext.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

            st.subheader("Chỉ số thanh khoản và Đòn bẩy")
            ts_nh = bcdkt_extracted.get("Tài sản ngắn hạn", {}).get("Cuối kỳ")
            no_nh = bcdkt_extracted.get("Nợ ngắn hạn", {}).get("Cuối kỳ")
            tien = bcdkt_extracted.get("Tiền và tương đương tiền", {}).get("Cuối kỳ")
            htk = bcdkt_extracted.get("Hàng tồn kho", {}).get("Cuối kỳ")
            tong_ts = bcdkt_extracted.get("Tổng tài sản", {}).get("Cuối kỳ")
            tong_no = bcdkt_extracted.get("Tổng nợ phải trả", {}).get("Cuối kỳ")
            vcsh = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Cuối kỳ")
            no_dh = bcdkt_extracted.get("Nợ dài hạn", {}).get("Cuối kỳ")

            bcdkt_ratios = {
                "Hệ số thanh toán hiện hành": safe_div(ts_nh, no_nh),
                "Hệ số thanh toán nhanh": safe_div((ts_nh or 0) - (htk or 0), no_nh),
                "Hệ số thanh toán tức thời": safe_div(tien, no_nh),
                "Tỷ lệ nợ trên vốn chủ sở hữu": safe_div(tong_no, vcsh),
                "Tỷ lệ nợ trên tổng tài sản": safe_div(tong_no, tong_ts),
                "Đòn bẩy tài chính": safe_div(tong_ts, vcsh),
            }

            c1, c2, c3 = st.columns(3)
            for i, (k, v) in enumerate(bcdkt_ratios.items()):
                with [c1, c2, c3][i % 3]:
                    st.metric(label=k, value=fmt_pct(v, k), help=FORMULAS.get(k, ""))

            st.subheader("Biểu đồ")
            ch1, ch2 = st.columns(2)

            with ch1:
                st.markdown("**🥧 Cơ cấu tài sản (Cuối kỳ)**")
                pie_vals, pie_labels = [], []
                for label, key in [("Tài sản ngắn hạn", "Tài sản ngắn hạn"), ("Tài sản dài hạn", "Tài sản dài hạn")]:
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

            with ch2:
                st.markdown("**📈 Cơ cấu nguồn vốn (Đầu kỳ vs Cuối kỳ)**")
                tong_no_dk = bcdkt_extracted.get("Tổng nợ phải trả", {}).get("Đầu kỳ")
                vcsh_dk = bcdkt_extracted.get("Vốn chủ sở hữu", {}).get("Đầu kỳ")
                if tong_no and vcsh:
                    fig_stack = go.Figure()
                    fig_stack.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[tong_no_dk or 0, tong_no], name="Nợ phải trả", marker_color="#ef5545"))
                    fig_stack.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[vcsh_dk or 0, vcsh], name="Vốn chủ sở hữu", marker_color="#636efa"))
                    fig_stack.update_layout(barmode="stack", height=350, yaxis_title="Giá trị")
                    st.plotly_chart(fig_stack, use_container_width=True)

            # Additional charts for BCĐKT
            ch3, ch4 = st.columns(2)

            with ch3:
                st.markdown("**📊 So sánh các chỉ tiêu (Đầu kỳ vs Cuối kỳ)**")
                fig_bcdkt_bar = go.Figure()
                for label in ["Tổng tài sản", "Tài sản ngắn hạn", "Tài sản dài hạn", "Tổng nợ phải trả", "Nợ ngắn hạn", "Nợ dài hạn", "Vốn chủ sở hữu"]:
                    if label in bcdkt_extracted:
                        dk = bcdkt_extracted[label].get("Đầu kỳ")
                        ck = bcdkt_extracted[label].get("Cuối kỳ")
                        fig_bcdkt_bar.add_trace(go.Bar(name=label, x=["Đầu kỳ", "Cuối kỳ"], y=[dk or 0, ck or 0]))
                fig_bcdkt_bar.update_layout(barmode="group", height=350, yaxis_title="Giá trị")
                st.plotly_chart(fig_bcdkt_bar, use_container_width=True)

            with ch4:
                st.markdown("**🥧 Cơ cấu nguồn vốn (Cuối kỳ)**")
                nv_pie_vals, nv_pie_labels = [], []
                if tong_no:
                    nv_pie_vals.append(tong_no)
                    nv_pie_labels.append("Nợ phải trả")
                if vcsh:
                    nv_pie_vals.append(vcsh)
                    nv_pie_labels.append("Vốn chủ sở hữu")
                if nv_pie_vals:
                    fig_nv_pie = px.pie(values=nv_pie_vals, names=nv_pie_labels, hole=0.4, color_discrete_sequence=["#ef5545", "#636efa"])
                    st.plotly_chart(fig_nv_pie, use_container_width=True)

            # Thanh khoản gauge
            st.markdown("**💧 Biểu đồ thanh khoản**")
            tk_names = []
            tk_vals = []
            for k in ["Hệ số thanh toán hiện hành", "Hệ số thanh toán nhanh", "Hệ số thanh toán tức thời"]:
                v = bcdkt_ratios.get(k)
                if v is not None:
                    tk_names.append(k.replace("Hệ số thanh toán ", "").title())
                    tk_vals.append(round(v, 2))
            if tk_vals:
                fig_tk = go.Figure(data=[go.Bar(x=tk_names, y=tk_vals, marker_color=["#2ca02c", "#ff7f0e", "#636efa"])])
                fig_tk.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="Mức an toàn = 1.0")
                fig_tk.update_layout(height=350, yaxis_title="Hệ số")
                st.plotly_chart(fig_tk, use_container_width=True)

# ============================================================
# TAB 2: BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH
# ============================================================
with tab2:
    st.header("Báo Cáo Kết Quả Hoạt Động Kinh Doanh")
    bkq_mode = st.radio("Chọn cách nhập dữ liệu:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True, key="bkq_mode")
    bkq_df = None

    if bkq_mode == "📂 Upload file Excel":
        uploaded_bkq = st.file_uploader("Chọn file Báo cáo kết quả hoạt động kinh doanh (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="upload_bkq")
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
21,Lợi nhuận thuần từ hoạt động kinh doanh,26000,25000
22,Lợi nhuận sau thuế,20000,22000
23,Thuế thu nhập doanh nghiệp,6000,6600"""
        text_bkq = st.text_area("Dữ liệu Báo cáo kết quả hoạt động kinh doanh", value=default_bkq, height=300, key="text_bkq")
        try:
            bkq_df = pd.read_csv(StringIO(text_bkq), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
        except Exception:
            st.warning("Dữ liệu nhập tay chưa hợp lệ.")

    bkq_extracted = {}

    if bkq_df is not None and not bkq_df.empty:
        bkq_df.columns = [str(c) for c in bkq_df.columns]
        st.subheader("Dữ liệu gốc")
        st.dataframe(bkq_df, use_container_width=True)

        col_start_kq, col_end_kq, _ = detect_columns(bkq_df, bkq_mode)

        bkq_items = {
            "Doanh thu thuần": ["doanh thu thuần", "doanh thu thuần về bán hàng", "dt thuần", "doanh thu thuần về bán hàng và cung cấp dịch vụ", "doanh thu"],
            "Doanh thu hoạt động tài chính": ["doanh thu hoạt động tài chính", "dt tài chính", "doanh thu ht tài chính", "doanh thu từ hoạt động tài chính"],
            "Thu nhập khác": ["thu nhập khác", "thu nhập", "thu nhập khác từ hoạt động tài chính"],
            "Chi phí vốn hàng bán": ["chi phí vốn hàng bán", "giá vốn hàng bán", "giá vốn", "cp vốn hàng bán", "giá vốn hàng bán và cung cấp dịch vụ"],
            "Chi phí bán hàng": ["chi phí bán hàng", "cp bán hàng", "chi phí bán hàng và cung cấp dịch vụ"],
            "Chi phí quản lý doanh nghiệp": ["chi phí quản lý doanh nghiệp", "cp quản lý dn", "chi phí quản lý", "chi phí quản lý doanh nghiệp và cung cấp dịch vụ"],
            "Chi phí tài chính": ["chi phí tài chính", "cp tài chính", "chi phí từ hoạt động tài chính"],
            "Chi phí khác": ["chi phí khác", "cp khác", "chi phí khác từ hoạt động tài chính"],
            "Lợi nhuận gộp": ["lợi nhuận gộp", "ln gộp", "lợi nhuận gộp về bán hàng", "lợi nhuận gộp về bán hàng và cung cấp dịch vụ"],
            "Lợi nhuận thuần từ hoạt động kinh doanh": ["lợi nhuận thuần từ hoạt động kinh doanh", "lợi nhuận thuần từ hdkd", "ln thuần từ hdkd", "lợi nhuận từ hdkd", "lợi nhuận thuần từ hoạt động kinh doanh và cung cấp dịch vụ", "lợi nhuận từ hoạt động kinh doanh"],
            "Lợi nhuận sau thuế": ["lợi nhuận sau thuế", "ln sau thuế", "lợi nhuận kế toán sau thuế", "lợi nhuận sau thuế thu nhập doanh nghiệp"],
            "Thuế thu nhập doanh nghiệp": ["thuế tndn", "thuế thu nhập doanh nghiệp", "thuế tndn hoãn lại", "chi phí thuế thu nhập doanh nghiệp"],
        }

        bkq_extracted = extract_items(bkq_df, bkq_items, col_start_kq, col_end_kq)
        df_bkq_ext = pd.DataFrame(bkq_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"})
        df_bkq_ext = df_bkq_ext.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

        if df_bkq_ext.empty:
            st.warning("⚠️ Không tự động trích được chỉ tiêu. Kiểm tra tên chỉ tiêu trong file.")
        else:
            st.subheader("Trích xuất chỉ tiêu chính")
            st.dataframe(df_bkq_ext.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

            st.subheader("Chỉ số sinh lời và Hiệu quả")
            dt_thuan = bkq_extracted.get("Doanh thu thuần", {}).get("Cuối kỳ")
            dt_thuan_dk = bkq_extracted.get("Doanh thu thuần", {}).get("Đầu kỳ")
            ln_gop = bkq_extracted.get("Lợi nhuận gộp", {}).get("Cuối kỳ")
            ln_hdkd = bkq_extracted.get("Lợi nhuận thuần từ hoạt động kinh doanh", {}).get("Cuối kỳ")
            ln_st = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Cuối kỳ")
            ln_st_dk = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Đầu kỳ")
            cp_von = bkq_extracted.get("Chi phí vốn hàng bán", {}).get("Cuối kỳ")
            cp_ban = bkq_extracted.get("Chi phí bán hàng", {}).get("Cuối kỳ")
            cp_qly = bkq_extracted.get("Chi phí quản lý doanh nghiệp", {}).get("Cuối kỳ")

            bkq_ratios = {
                "Biên lợi nhuận gộp": safe_div(ln_gop, dt_thuan),
                "Biên lợi nhuận hoạt động": safe_div(ln_hdkd, dt_thuan),
                "Biên lợi nhuận ròng": safe_div(ln_st, dt_thuan),
                "Tỷ lệ chi phí vốn hàng bán trên doanh thu": safe_div(cp_von, dt_thuan),
                "Tỷ lệ chi phí bán hàng trên doanh thu": safe_div(cp_ban, dt_thuan),
                "Tỷ lệ chi phí quản lý trên doanh thu": safe_div(cp_qly, dt_thuan),
                "Tăng trưởng doanh thu": safe_div((dt_thuan or 0) - (dt_thuan_dk or 0), dt_thuan_dk) if dt_thuan_dk else None,
                "Tăng trưởng lợi nhuận sau thuế": safe_div((ln_st or 0) - (ln_st_dk or 0), ln_st_dk) if ln_st_dk else None,
            }

            c1, c2, c3 = st.columns(3)
            for i, (k, v) in enumerate(bkq_ratios.items()):
                with [c1, c2, c3][i % 3]:
                    st.metric(label=k, value=fmt_pct(v, k), help=FORMULAS.get(k, ""))

            st.subheader("Biểu đồ")
            ch1, ch2 = st.columns(2)

            with ch1:
                st.markdown("**📊 Cơ cấu chi phí**")
                cp_vals, cp_labels = [], []
                for label, key in [("Chi phí vốn hàng bán", "Chi phí vốn hàng bán"), ("Chi phí bán hàng", "Chi phí bán hàng"),
                                   ("Chi phí quản lý", "Chi phí quản lý doanh nghiệp"), ("Chi phí tài chính", "Chi phí tài chính")]:
                    v = bkq_extracted.get(key, {}).get("Cuối kỳ")
                    if v:
                        cp_vals.append(v)
                        cp_labels.append(label)
                if cp_vals:
                    fig_cp = px.pie(values=cp_vals, names=cp_labels, hole=0.4, color_discrete_sequence=px.colors.sequential.Reds)
                    st.plotly_chart(fig_cp, use_container_width=True)

            with ch2:
                st.markdown("**📈 So sánh Doanh thu và Lợi nhuận**")
                fig_dtln = go.Figure()
                ln_gop_dk = bkq_extracted.get("Lợi nhuận gộp", {}).get("Đầu kỳ")
                if dt_thuan or dt_thuan_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[dt_thuan_dk or 0, dt_thuan or 0], name="Doanh thu thuần", marker_color="#636efa"))
                if ln_st or ln_st_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[ln_st_dk or 0, ln_st or 0], name="Lợi nhuận sau thuế", marker_color="#2ca02c"))
                if ln_gop or ln_gop_dk:
                    fig_dtln.add_trace(go.Bar(x=["Đầu kỳ", "Cuối kỳ"], y=[ln_gop_dk or 0, ln_gop or 0], name="Lợi nhuận gộp", marker_color="#ff7f0e"))
                fig_dtln.update_layout(barmode="group", height=350, yaxis_title="Giá trị")
                st.plotly_chart(fig_dtln, use_container_width=True)

            # Additional chart: CP/DT breakdown bar
            ch3, ch4 = st.columns(2)

            with ch3:
                st.markdown("**📊 Tỷ trọng chi phí trên doanh thu thuần**")
                cp_on_dt_names = []
                cp_on_dt_vals = []
                cp_on_dt_colors = []
                for label, key, color in [("Chi phí vốn hàng bán", "Tỷ lệ chi phí vốn hàng bán trên doanh thu", "#d62728"),
                                          ("Chi phí bán hàng", "Tỷ lệ chi phí bán hàng trên doanh thu", "#ff7f0e"),
                                          ("Chi phí quản lý", "Tỷ lệ chi phí quản lý trên doanh thu", "#636efa")]:
                    v = bkq_ratios.get(key)
                    if v is not None:
                        cp_on_dt_names.append(label)
                        cp_on_dt_vals.append(round(v * 100, 1))
                        cp_on_dt_colors.append(color)
                if cp_on_dt_vals:
                    fig_cpdt = go.Figure(data=[go.Bar(x=cp_on_dt_names, y=cp_on_dt_vals, marker_color=cp_on_dt_colors)])
                    fig_cpdt.update_layout(height=350, yaxis_title="% trên doanh thu", yaxis=dict(ticksuffix="%"))
                    st.plotly_chart(fig_cpdt, use_container_width=True)

            with ch4:
                st.markdown("**🥧 Cơ cấu doanh thu (Lợi nhuận gộp vs Chi phí vốn)**")
                dt_pie_vals = []
                dt_pie_labels = []
                ln_gop_val = bkq_extracted.get("Lợi nhuận gộp", {}).get("Cuối kỳ")
                cp_von_val = bkq_extracted.get("Chi phí vốn hàng bán", {}).get("Cuối kỳ")
                if ln_gop_val:
                    dt_pie_vals.append(ln_gop_val)
                    dt_pie_labels.append("Lợi nhuận gộp")
                if cp_von_val:
                    dt_pie_vals.append(cp_von_val)
                    dt_pie_labels.append("Chi phí vốn hàng bán")
                if dt_pie_vals:
                    fig_dt_pie = px.pie(values=dt_pie_vals, names=dt_pie_labels, hole=0.4, color_discrete_sequence=["#2ca02c", "#d62728"])
                    st.plotly_chart(fig_dt_pie, use_container_width=True)

            # Tăng trưởng comparison
            if dt_thuan_dk or ln_st_dk:
                st.markdown("**📈 Tăng trưởng Đầu kỳ vs Cuối kỳ**")
                fig_growth_bkq = go.Figure()
                items_bkq = [("Doanh thu thuần", "Doanh thu thuần"), ("Lợi nhuận gộp", "Lợi nhuận gộp"),
                             ("Lợi nhuận từ hoạt động kinh doanh", "Lợi nhuận thuần từ hoạt động kinh doanh"), ("Lợi nhuận sau thuế", "Lợi nhuận sau thuế")]
                for label, key in items_bkq:
                    dk = bkq_extracted.get(key, {}).get("Đầu kỳ")
                    ck = bkq_extracted.get(key, {}).get("Cuối kỳ")
                    if dk or ck:
                        fig_growth_bkq.add_trace(go.Bar(name=label, x=["Đầu kỳ", "Cuối kỳ"], y=[dk or 0, ck or 0]))
                fig_growth_bkq.update_layout(barmode="group", height=350, yaxis_title="Giá trị")
                st.plotly_chart(fig_growth_bkq, use_container_width=True)

# ============================================================
# TAB 3: BÁO CÁO LƯU CHUYỂN TIỀN TỆ
# ============================================================
with tab3:
    st.header("Báo Cáo Lưu Chuyển Tiền Tệ")
    cf_mode = st.radio("Chọn cách nhập dữ liệu:", ["📂 Upload file Excel", "✏️ Nhập tay"], horizontal=True, key="cf_mode")
    cf_df = None

    if cf_mode == "📂 Upload file Excel":
        uploaded_cf = st.file_uploader("Chọn file Báo cáo lưu chuyển tiền tệ (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="upload_cf")
        cf_df = parse_uploaded_file(uploaded_cf)
    elif cf_mode == "✏️ Nhập tay":
        default_cf = """01,Lưu chuyển tiền từ hoạt động kinh doanh,15000,20000
02,Tiền thu từ bán hàng,80000,100000
03,Tiền trả cho người bán,55000,65000
04,Tiền trả cho người lao động,10000,12000
05,Tiền trả cho các chi phí khác,5000,3000
10,Lưu chuyển tiền từ hoạt động đầu tư,-10000,-8000
11,Tiền thu thanh lý tài sản cố định,2000,3000
12,Tiền chi mua sắm tài sản cố định,-12000,-11000
20,Lưu chuyển tiền từ hoạt động tài chính,5000,4000
21,Tiền thu từ vay vốn,10000,8000
22,Tiền trả nợ vay,-5000,-4000
30,Lưu chuyển tiền thuần,10000,16000
31,Tiền đầu kỳ,20000,30000
32,Tiền cuối kỳ,30000,46000"""
        text_cf = st.text_area("Dữ liệu Báo cáo lưu chuyển tiền tệ", value=default_cf, height=300, key="text_cf")
        try:
            cf_df = pd.read_csv(StringIO(text_cf), header=None, names=["Mã", "Tên chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"])
        except Exception:
            st.warning("Dữ liệu nhập tay chưa hợp lệ.")

    cf_extracted = {}

    if cf_df is not None and not cf_df.empty:
        cf_df.columns = [str(c) for c in cf_df.columns]
        st.subheader("Dữ liệu gốc")
        st.dataframe(cf_df, use_container_width=True)

        col_start_cf, col_end_cf, _ = detect_columns(cf_df, cf_mode)

        cf_items = {
            "Lưu chuyển tiền từ hoạt động kinh doanh": ["lưu chuyển tiền từ hoạt động kinh doanh", "lưu chuyển tiền từ hđkd", "lưu chuyển từ hoạt động kinh doanh", "lctt từ hoạt động kinh doanh", "lctt từ hđkd", "lưu chuyển từ hđkd", "lct hđkd", "lct hoạt động kinh doanh", "lưu chuyển tiền thuần từ hđkd", "dòng tiền từ hoạt động kinh doanh", "dòng tiền hđkd", "lưu chuyển tiền lũy kế từ hoạt động kinh doanh"],
            "Tiền thu từ bán hàng": ["tiền thu từ bán hàng", "thu từ bán hàng", "tiền thu bán hàng", "thu tiền từ bán hàng"],
            "Tiền trả cho người bán": ["tiền trả cho người bán", "trả cho người bán", "chi cho người bán", "tiền thanh toán cho người bán"],
            "Tiền trả cho người lao động": ["tiền trả cho người lao động", "trả người lao động", "chi cho người lao động", "trả cho nlđ", "tiền lương", "chi phí nhân sự", "chi trả cho người lao động"],
            "Tiền trả chi phí khác": ["tiền trả cho các chi phí khác", "chi phí khác", "tiền trả chi phí khác", "thuế và các khoản khác"],
            "Lưu chuyển tiền từ hoạt động đầu tư": ["lưu chuyển tiền từ hoạt động đầu tư", "lưu chuyển tiền từ hđ đầu tư", "lưu chuyển từ hoạt động đầu tư", "lctt từ hoạt động đầu tư", "lctt từ hđ đầu tư", "lct hđ đầu tư", "lct từ hđ đầu tư", "lct hoạt động đầu tư", "dòng tiền từ hoạt động đầu tư", "dòng tiền hđ đầu tư", "lưu chuyển tiền lũy kế từ hoạt động đầu tư"],
            "Tiền thu thanh lý tài sản cố định": ["thanh lý tài sản cố định", "thanh lý tscđ", "thu thanh lý", "thu bán tscđ", "tiền thu thanh lý tscđ", "tiền thu từ thanh lý", "thu tiền từ bán tscđ"],
            "Tiền chi mua sắm tài sản cố định": ["mua sắm tài sản cố định", "mua sắm tscđ", "chi mua sắm tscđ", "tiền chi mua sắm", "đầu tư tscđ", "mua sắm csxd", "tiền chi cho hoạt động đầu tư", "chi mua sắm tài sản cố định", "đầu tư tài sản cố định"],
            "Lưu chuyển tiền từ hoạt động tài chính": ["lưu chuyển tiền từ hoạt động tài chính", "lưu chuyển tiền từ hđ tài chính", "lưu chuyển từ hoạt động tài chính", "lctt từ hoạt động tài chính", "lctt từ hđ tài chính", "lct hđ tài chính", "lct hoạt động tài chính", "dòng tiền từ hoạt động tài chính", "dòng tiền hđ tài chính", "lưu chuyển tiền lũy kế từ hoạt động tài chính"],
            "Tiền thu từ vay vốn": ["tiền thu từ vay vốn", "thu từ vay", "vay vốn", "tiền vay nhận được"],
            "Tiền trả nợ vay": ["tiền trả nợ vay", "trả nợ vay", "trả nợ", "thanh toán nợ vay", "chi trả nợ vay"],
            "Lưu chuyển tiền thuần": ["lưu chuyển tiền thuần", "lct thuần", "lưu chuyển thuần", "tăng/giảm tiền", "lưu chuyển tiền thuần trong kỳ", "tăng giảm tiền"],
            "Tiền đầu kỳ": ["tiền đầu kỳ", "số dư đầu kỳ", "tiền đầu năm", "tiền và tương đương tiền đầu kỳ", "số dư tiền đầu kỳ"],
            "Tiền cuối kỳ": ["tiền cuối kỳ", "số dư cuối kỳ", "tiền cuối năm", "tiền và tương đương tiền cuối kỳ", "số dư tiền cuối kỳ"],
        }

        cf_extracted = extract_items(cf_df, cf_items, col_start_cf, col_end_cf)
        df_cf_ext = pd.DataFrame(cf_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"})
        df_cf_ext = df_cf_ext.dropna(subset=["Đầu kỳ", "Cuối kỳ"], how="all")

        if df_cf_ext.empty:
            st.warning("⚠️ Không tự động trích được chỉ tiêu. Kiểm tra tên chỉ tiêu trong file.")
        else:
            st.subheader("Trích xuất chỉ tiêu chính")
            st.dataframe(df_cf_ext.style.format({"Đầu kỳ": "{:,.0f}", "Cuối kỳ": "{:,.0f}"}), use_container_width=True)

            # --- BCLCTT ANALYSIS ---
            st.subheader("Phân tích lưu chuyển tiền tệ")

            lct_hdkd = cf_extracted.get("Lưu chuyển tiền từ hoạt động kinh doanh", {}).get("Cuối kỳ")
            lct_hd_dt = cf_extracted.get("Lưu chuyển tiền từ hoạt động đầu tư", {}).get("Cuối kỳ")
            lct_hd_tc = cf_extracted.get("Lưu chuyển tiền từ hoạt động tài chính", {}).get("Cuối kỳ")
            lct_thuan = cf_extracted.get("Lưu chuyển tiền thuần", {}).get("Cuối kỳ")
            tien_ck = cf_extracted.get("Tiền cuối kỳ", {}).get("Cuối kỳ")

            dt_thuan_cf = bkq_extracted.get("Doanh thu thuần", {}).get("Cuối kỳ") if bkq_extracted else None
            ln_st_cf = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Cuối kỳ") if bkq_extracted else None

            cf_ratios = {}
            cf_ratio_labels = {}
            cf_ratios["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần"] = safe_div(lct_hdkd, dt_thuan_cf)
            cf_ratio_labels["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần"] = "Lưu chuyển tiền từ hoạt động kinh doanh / Doanh thu thuần"
            cf_ratios["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế"] = safe_div(lct_hdkd, ln_st_cf)
            cf_ratio_labels["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế"] = "Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế"
            cf_ratios["Tỷ lệ lưu chuyển tiền hoạt động đầu tư trên tổng lưu chuyển"] = safe_div(lct_hd_dt, lct_thuan) if lct_thuan and lct_thuan != 0 else None
            cf_ratio_labels["Tỷ lệ lưu chuyển tiền hoạt động đầu tư trên tổng lưu chuyển"] = "Lưu chuyển tiền từ hoạt động đầu tư / Lưu chuyển thuần"
            cf_ratios["Tỷ lệ lưu chuyển tiền hoạt động tài chính trên tổng lưu chuyển"] = safe_div(lct_hd_tc, lct_thuan) if lct_thuan and lct_thuan != 0 else None
            cf_ratio_labels["Tỷ lệ lưu chuyển tiền hoạt động tài chính trên tổng lưu chuyển"] = "Lưu chuyển tiền từ hoạt động tài chính / Lưu chuyển thuần"

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("💰 Lưu chuyển tiền từ hoạt động kinh doanh", value=f"{lct_hdkd:,.0f}" if lct_hdkd else "N/A")
                st.metric("📈 Lưu chuyển tiền từ hoạt động đầu tư", value=f"{lct_hd_dt:,.0f}" if lct_hd_dt else "N/A")
            with c2:
                st.metric("🏦 Lưu chuyển tiền từ hoạt động tài chính", value=f"{lct_hd_tc:,.0f}" if lct_hd_tc else "N/A")
                st.metric("💵 Lưu chuyển tiền thuần", value=f"{lct_thuan:,.0f}" if lct_thuan else "N/A")
            with c3:
                for k, v in cf_ratios.items():
                    st.metric(label=cf_ratio_labels.get(k, k), value=fmt_pct(v, k), help=FORMULAS.get(k, ""))

            # Quality assessment
            if lct_hdkd and ln_st_cf and ln_st_cf != 0:
                quality = lct_hdkd / ln_st_cf
                st.subheader("🔍 Chất lượng lợi nhuận")
                if quality > 1:
                    st.success(f"✅ Tỷ lệ Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế = **{quality:.2f}x** — Doanh nghiệp tạo tiền tốt, lợi nhuận chất lượng cao.")
                elif quality > 0.7:
                    st.warning(f"⚠️ Tỷ lệ Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế = **{quality:.2f}x** — Khá, nhưng cần theo dõi chất lượng lợi nhuận.")
                else:
                    st.error(f"🔴 Tỷ lệ Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế = **{quality:.2f}x** — Lợi nhuận chưa chuyển thành tiền, rủi ro chất lượng lợi nhuận.")

            # BCLCTT CHARTS
            st.subheader("Biểu đồ")
            ch1, ch2 = st.columns(2)

            with ch1:
                st.markdown("**📊 Cơ cấu dòng tiền (Cuối kỳ)**")
                cf_bar_labels = []
                cf_bar_vals = []
                cf_bar_colors = []
                for label, key, color in [("Hoạt động kinh doanh", "Lưu chuyển tiền từ hoạt động kinh doanh", "#2ca02c"),
                                          ("Hoạt động đầu tư", "Lưu chuyển tiền từ hoạt động đầu tư", "#636efa"),
                                          ("Hoạt động tài chính", "Lưu chuyển tiền từ hoạt động tài chính", "#ff7f0e")]:
                    v = cf_extracted.get(key, {}).get("Cuối kỳ")
                    if v is not None:
                        cf_bar_labels.append(label)
                        cf_bar_vals.append(v)
                        cf_bar_colors.append(color)
                if cf_bar_vals:
                    fig_cf_pie = go.Figure(data=[go.Bar(x=cf_bar_labels, y=cf_bar_vals, marker_color=cf_bar_colors,
                                                        text=[f"{v:,.0f}" for v in cf_bar_vals], textposition="auto")])
                    fig_cf_pie.update_layout(height=350, yaxis_title="Giá trị")
                    st.plotly_chart(fig_cf_pie, use_container_width=True)

            with ch2:
                st.markdown("**📈 Dòng tiền theo hoạt động**")
                fig_cf_bar = go.Figure()
                for label, key in [("Hoạt động kinh doanh", "Lưu chuyển tiền từ hoạt động kinh doanh"),
                                    ("Hoạt động đầu tư", "Lưu chuyển tiền từ hoạt động đầu tư"),
                                    ("Hoạt động tài chính", "Lưu chuyển tiền từ hoạt động tài chính"),
                                    ("Lưu chuyển thuần", "Lưu chuyển tiền thuần")]:
                    dk = cf_extracted.get(key, {}).get("Đầu kỳ")
                    ck = cf_extracted.get(key, {}).get("Cuối kỳ")
                    if dk is not None or ck is not None:
                        fig_cf_bar.add_trace(go.Bar(name=label, x=["Đầu kỳ", "Cuối kỳ"], y=[dk or 0, ck or 0]))
                fig_cf_bar.update_layout(barmode="group", height=350, yaxis_title="Giá trị")
                st.plotly_chart(fig_cf_bar, use_container_width=True)

            # Chart: Cash inflows (Tiền thu)
            ch5, ch6 = st.columns(2)
            with ch5:
                st.markdown("**💰 Tiền thu**")
                thu_items = [
                    ("Thu từ bán hàng", "Tiền thu từ bán hàng"),
                    ("Thu thanh lý TSCĐ", "Tiền thu thanh lý tài sản cố định"),
                    ("Thu từ vay vốn", "Tiền thu từ vay vốn"),
                ]
                thu_labels, thu_dk, thu_ck = [], [], []
                for lbl, key in thu_items:
                    dk_v = cf_extracted.get(key, {}).get("Đầu kỳ")
                    ck_v = cf_extracted.get(key, {}).get("Cuối kỳ")
                    if dk_v is not None or ck_v is not None:
                        thu_labels.append(lbl)
                        thu_dk.append(dk_v or 0)
                        thu_ck.append(ck_v or 0)
                if thu_labels:
                    fig_thu = go.Figure()
                    fig_thu.add_trace(go.Bar(name="Đầu kỳ", x=thu_labels, y=thu_dk, marker_color="#636efa"))
                    fig_thu.add_trace(go.Bar(name="Cuối kỳ", x=thu_labels, y=thu_ck, marker_color="#2ca02c"))
                    fig_thu.update_layout(barmode="group", height=350, yaxis_title="Giá trị", legend=dict(orientation="h", yanchor="bottom", y=1.02))
                    st.plotly_chart(fig_thu, use_container_width=True)
                else:
                    st.info("Chưa có dữ liệu tiền thu.")

            with ch6:
                st.markdown("**💸 Tiền trả & Chi**")
                chi_items = [
                    ("Trả người bán", "Tiền trả cho người bán"),
                    ("Trả người lao động", "Tiền trả cho người lao động"),
                    ("Trả chi phí khác", "Tiền trả chi phí khác"),
                    ("Mua sắm TSCĐ", "Tiền chi mua sắm tài sản cố định"),
                    ("Trả nợ vay", "Tiền trả nợ vay"),
                ]
                chi_labels, chi_dk, chi_ck = [], [], []
                for lbl, key in chi_items:
                    dk_v = cf_extracted.get(key, {}).get("Đầu kỳ")
                    ck_v = cf_extracted.get(key, {}).get("Cuối kỳ")
                    if dk_v is not None or ck_v is not None:
                        chi_labels.append(lbl)
                        chi_dk.append(dk_v or 0)
                        chi_ck.append(ck_v or 0)
                if chi_labels:
                    fig_chi = go.Figure()
                    fig_chi.add_trace(go.Bar(name="Đầu kỳ", x=chi_labels, y=chi_dk, marker_color="#636efa"))
                    fig_chi.add_trace(go.Bar(name="Cuối kỳ", x=chi_labels, y=chi_ck, marker_color="#ef5545"))
                    fig_chi.update_layout(barmode="group", height=350, yaxis_title="Giá trị", legend=dict(orientation="h", yanchor="bottom", y=1.02))
                    st.plotly_chart(fig_chi, use_container_width=True)
                else:
                    st.info("Chưa có dữ liệu tiền trả & chi.")

# ============================================================
# TAB 4: COMBINED ANALYSIS
# ============================================================
with tab4:
    st.header("🔄 Phân Tích Kết Hợp 3 Báo Cáo")
    st.markdown("Kết hợp dữ liệu từ Bảng cân đối kế toán, Báo cáo kết quả hoạt động kinh doanh, Báo cáo lưu chuyển tiền tệ để tính các chỉ số tài chính tổng hợp.")

    has_bcdkt = bool(bcdkt_extracted)
    has_bkq = bool(bkq_extracted)
    has_cf = bool(cf_extracted)

    if not has_bcdkt and not has_bkq and not has_cf:
        st.info("👆 Nhập dữ liệu ở các tab trên để phân tích kết hợp.")
    else:
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

        dt_thuan = bkq_extracted.get("Doanh thu thuần", {}).get("Cuối kỳ") if has_bkq else None
        dt_thuan_dk = bkq_extracted.get("Doanh thu thuần", {}).get("Đầu kỳ") if has_bkq else None
        ln_gop = bkq_extracted.get("Lợi nhuận gộp", {}).get("Cuối kỳ") if has_bkq else None
        ln_hdkd = bkq_extracted.get("Lợi nhuận thuần từ hoạt động kinh doanh", {}).get("Cuối kỳ") if has_bkq else None
        ln_st = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Cuối kỳ") if has_bkq else None
        ln_st_dk = bkq_extracted.get("Lợi nhuận sau thuế", {}).get("Đầu kỳ") if has_bkq else None
        cp_von = bkq_extracted.get("Chi phí vốn hàng bán", {}).get("Cuối kỳ") if has_bkq else None
        cp_ban = bkq_extracted.get("Chi phí bán hàng", {}).get("Cuối kỳ") if has_bkq else None
        cp_qly = bkq_extracted.get("Chi phí quản lý doanh nghiệp", {}).get("Cuối kỳ") if has_bkq else None

        lct_hdkd = cf_extracted.get("Lưu chuyển tiền từ hoạt động kinh doanh", {}).get("Cuối kỳ") if has_cf else None
        lct_hd_dt = cf_extracted.get("Lưu chuyển tiền từ hoạt động đầu tư", {}).get("Cuối kỳ") if has_cf else None

        ts_tb = safe_div((tong_ts or 0) + (tong_ts_dk or 0), 2) if tong_ts or tong_ts_dk else None
        vcsh_tb = safe_div((vcsh or 0) + (vcsh_dk or 0), 2) if vcsh or vcsh_dk else None

        combined_ratios = {}
        combined_labels = {}

        # Thanh khoản
        if has_bcdkt:
            combined_ratios["Hệ số thanh toán hiện hành"] = safe_div(ts_nh, no_nh)
            combined_ratios["Hệ số thanh toán nhanh"] = safe_div((ts_nh or 0) - (htk or 0), no_nh)
            combined_ratios["Hệ số thanh toán tức thời"] = safe_div(tien, no_nh)
            combined_labels["Hệ số thanh toán hiện hành"] = "Tài sản ngắn hạn / Nợ ngắn hạn"
            combined_labels["Hệ số thanh toán nhanh"] = "(Tài sản ngắn hạn - Hàng tồn kho) / Nợ ngắn hạn"
            combined_labels["Hệ số thanh toán tức thời"] = "Tiền / Nợ ngắn hạn"

            combined_ratios["Tỷ lệ nợ trên vốn chủ sở hữu"] = safe_div(tong_no, vcsh)
            combined_ratios["Tỷ lệ nợ trên tổng tài sản"] = safe_div(tong_no, tong_ts)
            combined_ratios["Đòn bẩy tài chính"] = safe_div(tong_ts, vcsh)
            combined_labels["Tỷ lệ nợ trên vốn chủ sở hữu"] = "Tổng nợ / Vốn chủ sở hữu"
            combined_labels["Tỷ lệ nợ trên tổng tài sản"] = "Tổng nợ / Tổng tài sản"
            combined_labels["Đòn bẩy tài chính"] = "Tổng tài sản / Vốn chủ sở hữu"

        # Sinh lời
        if has_bkq:
            combined_ratios["Biên lợi nhuận gộp"] = safe_div(ln_gop, dt_thuan)
            combined_ratios["Biên lợi nhuận hoạt động"] = safe_div(ln_hdkd, dt_thuan)
            combined_ratios["Biên lợi nhuận ròng"] = safe_div(ln_st, dt_thuan)
            combined_labels["Biên lợi nhuận gộp"] = "Lợi nhuận gộp / Doanh thu thuần"
            combined_labels["Biên lợi nhuận hoạt động"] = "Lợi nhuận từ hoạt động kinh doanh / Doanh thu thuần"
            combined_labels["Biên lợi nhuận ròng"] = "Lợi nhuận sau thuế / Doanh thu thuần"

        # Hiệu quả sử dụng vốn
        if has_bcdkt and has_bkq:
            combined_ratios["ROA"] = safe_div(ln_st, ts_tb)
            combined_ratios["ROE"] = safe_div(ln_st, vcsh_tb)
            combined_labels["ROA"] = "Lợi nhuận sau thuế / Tổng tài sản trung bình"
            combined_labels["ROE"] = "Lợi nhuận sau thuế / Vốn chủ sở hữu trung bình"

            combined_ratios["Vòng quay tài sản"] = safe_div(dt_thuan, ts_tb)
            combined_labels["Vòng quay tài sản"] = "Doanh thu thuần / Tổng tài sản trung bình"

            combined_ratios["Vòng quay vốn chủ sở hữu"] = safe_div(dt_thuan, vcsh_tb)
            combined_labels["Vòng quay vốn chủ sở hữu"] = "Doanh thu thuần / Vốn chủ sở hữu trung bình"

            combined_ratios["Vòng quay hàng tồn kho"] = safe_div(cp_von, htk)
            combined_labels["Vòng quay hàng tồn kho"] = "Chi phí vốn hàng bán / Hàng tồn kho"

            combined_ratios["Vòng quay khoản phải thu"] = safe_div(dt_thuan, pt_nh)
            combined_labels["Vòng quay khoản phải thu"] = "Doanh thu thuần / Phải thu ngắn hạn"

        # Hiệu quả chi phí
        if has_bkq:
            combined_ratios["Tỷ lệ chi phí vốn hàng bán trên doanh thu"] = safe_div(cp_von, dt_thuan)
            combined_ratios["Tỷ lệ chi phí bán hàng trên doanh thu"] = safe_div(cp_ban, dt_thuan)
            combined_ratios["Tỷ lệ chi phí quản lý trên doanh thu"] = safe_div(cp_qly, dt_thuan)
            combined_labels["Tỷ lệ chi phí vốn hàng bán trên doanh thu"] = "Chi phí vốn hàng bán / Doanh thu thuần"
            combined_labels["Tỷ lệ chi phí bán hàng trên doanh thu"] = "Chi phí bán hàng / Doanh thu thuần"
            combined_labels["Tỷ lệ chi phí quản lý trên doanh thu"] = "Chi phí quản lý / Doanh thu thuần"

        # Tăng trưởng
        if has_bkq and dt_thuan_dk:
            combined_ratios["Tăng trưởng doanh thu"] = safe_div((dt_thuan or 0) - (dt_thuan_dk or 0), dt_thuan_dk)
            combined_labels["Tăng trưởng doanh thu"] = "(Doanh thu cuối kỳ - Doanh thu đầu kỳ) / Doanh thu đầu kỳ"
        if has_bkq and ln_st_dk:
            combined_ratios["Tăng trưởng lợi nhuận sau thuế"] = safe_div((ln_st or 0) - (ln_st_dk or 0), ln_st_dk)
            combined_labels["Tăng trưởng lợi nhuận sau thuế"] = "(Lợi nhuận cuối - Lợi nhuận đầu) / Lợi nhuận đầu"

        # Dòng tiền
        if has_cf:
            combined_ratios["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần"] = safe_div(lct_hdkd, dt_thuan)
            combined_labels["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần"] = "Lưu chuyển tiền từ hoạt động kinh doanh / Doanh thu thuần"
            combined_ratios["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế"] = safe_div(lct_hdkd, ln_st)
            combined_labels["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế"] = "Lưu chuyển tiền từ hoạt động kinh doanh / Lợi nhuận sau thuế"
            fcf_val = None
            if lct_hdkd is not None or lct_hd_dt is not None:
                fcf_val = (lct_hdkd or 0) + (lct_hd_dt or 0)
            combined_ratios["Dòng tiền tự do"] = fcf_val
            combined_labels["Dòng tiền tự do"] = "Lưu chuyển tiền từ hoạt động kinh doanh + Lưu chuyển tiền từ hoạt động đầu tư"

        # DISPLAY
        if combined_ratios:
            st.subheader("📋 Bảng tổng hợp chỉ số tài chính")

            thanh_khoan = [k for k in combined_ratios if k in ["Hệ số thanh toán hiện hành", "Hệ số thanh toán nhanh", "Hệ số thanh toán tức thời"]]
            don_bay = [k for k in combined_ratios if k in ["Tỷ lệ nợ trên vốn chủ sở hữu", "Tỷ lệ nợ trên tổng tài sản", "Đòn bẩy tài chính"]]
            sinh_loi = [k for k in combined_ratios if k in ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng", "ROA", "ROE"]]
            hieu_qua = [k for k in combined_ratios if k in ["Vòng quay tài sản", "Vòng quay vốn chủ sở hữu", "Vòng quay hàng tồn kho", "Vòng quay khoản phải thu"]]
            chi_phi = [k for k in combined_ratios if k.startswith("Tỷ lệ chi phí")]
            tang_truong = [k for k in combined_ratios if k.startswith("Tăng trưởng")]
            dong_tien = [k for k in combined_ratios if k in ["Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần", "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế", "Dòng tiền tự do"]]

            groups = [
                ("💧 Thanh khoản", thanh_khoan),
                ("🏗️ Đòn bẩy", don_bay),
                ("💰 Sinh lời", sinh_loi),
                ("⚡ Hiệu quả", hieu_qua),
                ("📊 Cơ cấu chi phí", chi_phi),
                ("📈 Tăng trưởng", tang_truong),
                ("💵 Dòng tiền", dong_tien),
            ]

            for title, keys in groups:
                if not keys:
                    continue
                st.markdown(f"**{title}**")
                cols = st.columns(min(len(keys), 3))
                for i, k in enumerate(keys):
                    v = combined_ratios.get(k)
                    with cols[i % 3]:
                        val_str = fmt_pct(v, k)
                        st.metric(label=k, value=val_str, help=FORMULAS.get(k, ""))

            # CHARTS
            st.subheader("📊 Biểu đồ tổng hợp")
            ch1, ch2 = st.columns(2)

            with ch1:
                if has_bcdkt and has_bkq:
                    st.markdown("**🔺 Phân tích Du Pont**")
                    margin = combined_ratios.get("Biên lợi nhuận ròng")
                    turnover = combined_ratios.get("Vòng quay tài sản")
                    leverage = combined_ratios.get("Đòn bẩy tài chính")
                    roe = combined_ratios.get("ROE")
                    fig_dupont = go.Figure()
                    labels = ["Biên lợi nhuận ròng", "Vòng quay tài sản", "Đòn bẩy tài chính", "ROE"]
                    vals = [margin or 0, turnover or 0, leverage or 0, roe or 0]
                    vals_pct = [f"{(margin or 0)*100:.1f}%", f"{turnover or 0:.2f}x", f"{leverage or 0:.2f}x", f"{(roe or 0)*100:.1f}%"]
                    fig_dupont.add_trace(go.Bar(x=labels, y=vals, text=vals_pct, textposition="auto", marker_color=["#636efa", "#ff7f0e", "#2ca02c", "#ef5545"]))
                    fig_dupont.update_layout(height=350, yaxis_title="Giá trị")
                    st.plotly_chart(fig_dupont, use_container_width=True)

            with ch2:
                st.markdown("**🕸️ Radar chỉ số tổng hợp**")
                radar_labels, radar_vals = [], []
                for k, v in combined_ratios.items():
                    if v is not None:
                        radar_labels.append(k)
                        if k in ["Biên lợi nhuận gộp", "Biên lợi nhuận hoạt động", "Biên lợi nhuận ròng", "ROA", "ROE",
                                   "Tỷ lệ nợ trên tổng tài sản", "Tỷ lệ nợ trên vốn chủ sở hữu",
                                   "Tỷ lệ chi phí vốn hàng bán trên doanh thu", "Tỷ lệ chi phí bán hàng trên doanh thu", "Tỷ lệ chi phí quản lý trên doanh thu",
                                   "Tăng trưởng doanh thu", "Tăng trưởng lợi nhuận sau thuế",
                                   "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên doanh thu thuần", "Tỷ lệ lưu chuyển tiền hoạt động kinh doanh trên lợi nhuận sau thuế"]:
                            radar_vals.append(min(abs(v), 2.0))
                        elif k == "Đòn bẩy tài chính":
                            radar_vals.append(min(v / 5, 1.5))
                        elif k.startswith("Vòng quay"):
                            radar_vals.append(min(v / 5, 1.5))
                        else:
                            radar_vals.append(min(abs(v) / 3, 1.5))
                if radar_labels:
                    fig_radar = go.Figure(data=go.Scatterpolar(r=radar_vals, theta=radar_labels, fill="toself", name="Cuối kỳ"))
                    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 1.5])), showlegend=False, height=400)
                    st.plotly_chart(fig_radar, use_container_width=True)

            # Growth comparison
            st.markdown("**📈 So sánh Đầu kỳ và Cuối kỳ**")
            fig_growth = go.Figure()
            items_to_chart = {}
            if has_bkq:
                for label, key in [("Doanh thu thuần", "Doanh thu thuần"), ("Lợi nhuận gộp", "Lợi nhuận gộp"), ("Lợi nhuận sau thuế", "Lợi nhuận sau thuế")]:
                    dk = bkq_extracted.get(key, {}).get("Đầu kỳ")
                    ck = bkq_extracted.get(key, {}).get("Cuối kỳ")
                    if dk or ck:
                        items_to_chart[label] = (dk, ck)
            if has_bcdkt:
                for label, key in [("Tổng tài sản", "Tổng tài sản"), ("Vốn chủ sở hữu", "Vốn chủ sở hữu"), ("Tổng nợ phải trả", "Tổng nợ phải trả")]:
                    dk = bcdkt_extracted.get(key, {}).get("Đầu kỳ")
                    ck = bcdkt_extracted.get(key, {}).get("Cuối kỳ")
                    if dk or ck:
                        items_to_chart[label] = (dk, ck)
            if has_cf:
                for label, key in [("Lưu chuyển tiền từ hoạt động kinh doanh", "Lưu chuyển tiền từ hoạt động kinh doanh"),
                                    ("Lưu chuyển tiền từ hoạt động đầu tư", "Lưu chuyển tiền từ hoạt động đầu tư"),
                                    ("Lưu chuyển tiền từ hoạt động tài chính", "Lưu chuyển tiền từ hoạt động tài chính"),
                                    ("Lưu chuyển tiền thuần", "Lưu chuyển tiền thuần")]:
                    dk = cf_extracted.get(key, {}).get("Đầu kỳ")
                    ck = cf_extracted.get(key, {}).get("Cuối kỳ")
                    if dk or ck:
                        items_to_chart[label] = (dk, ck)

            for label, (dk, ck) in items_to_chart.items():
                fig_growth.add_trace(go.Bar(name=label, x=["Đầu kỳ", "Cuối kỳ"], y=[dk or 0, ck or 0]))
            fig_growth.update_layout(barmode="group", height=400, yaxis_title="Giá trị")
            st.plotly_chart(fig_growth, use_container_width=True)

        # EXPORT
        st.header("📥 Xuất báo cáo tổng hợp")

        export_rows = []
        if has_bcdkt:
            for label in ["Tổng tài sản", "Tài sản ngắn hạn", "Tài sản dài hạn", "Tiền và tương đương tiền",
                         "Phải thu ngắn hạn", "Hàng tồn kho", "Tổng nợ phải trả", "Nợ ngắn hạn", "Nợ dài hạn",
                         "Vốn chủ sở hữu", "Lợi nhuận chưa phân phối"]:
                if label in bcdkt_extracted:
                    export_rows.append({"Chỉ tiêu": f"[Bảng cân đối KT] {label}", "Đầu kỳ": bcdkt_extracted[label].get("Đầu kỳ"), "Cuối kỳ": bcdkt_extracted[label].get("Cuối kỳ"), "Ghi chú": ""})
        if has_bkq:
            for label in ["Doanh thu thuần", "Doanh thu hoạt động tài chính", "Chi phí vốn hàng bán",
                         "Chi phí bán hàng", "Chi phí quản lý doanh nghiệp", "Chi phí tài chính", "Chi phí khác",
                         "Lợi nhuận gộp", "Lợi nhuận thuần từ hoạt động kinh doanh", "Lợi nhuận sau thuế", "Thuế thu nhập doanh nghiệp"]:
                if label in bkq_extracted:
                    export_rows.append({"Chỉ tiêu": f"[Báo cáo kết quả hoạt động kinh doanh] {label}", "Đầu kỳ": bkq_extracted[label].get("Đầu kỳ"), "Cuối kỳ": bkq_extracted[label].get("Cuối kỳ"), "Ghi chú": ""})
        if has_cf:
            for label in ["Lưu chuyển tiền từ hoạt động kinh doanh", "Lưu chuyển tiền từ hoạt động đầu tư", "Lưu chuyển tiền từ hoạt động tài chính",
                         "Lưu chuyển tiền thuần", "Tiền đầu kỳ", "Tiền cuối kỳ"]:
                if label in cf_extracted:
                    export_rows.append({"Chỉ tiêu": f"[Báo cáo lưu chuyển tiền tệ] {label}", "Đầu kỳ": cf_extracted[label].get("Đầu kỳ"), "Cuối kỳ": cf_extracted[label].get("Cuối kỳ"), "Ghi chú": ""})

        export_rows.append({"Chỉ tiêu": "--- CHỈ SỐ TÀI CHÍNH ---", "Đầu kỳ": "", "Cuối kỳ": "", "Ghi chú": ""})
        for k, v in combined_ratios.items():
            export_rows.append({"Chỉ tiêu": k, "Đầu kỳ": "", "Cuối kỳ": v, "Ghi chú": combined_labels.get(k, "")})

        df_export = pd.DataFrame(export_rows)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            csv = df_export.to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 Tải CSV tổng hợp", data=csv, file_name="bao_cao_tai_chinh_tong_hop.csv", mime="text/csv")
        with col_dl2:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    df_export.to_excel(writer, index=False, sheet_name="Tổng hợp")
                    if has_bcdkt:
                        pd.DataFrame(bcdkt_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"}).to_excel(writer, index=False, sheet_name="Bảng cân đối KT")
                    if has_bkq:
                        pd.DataFrame(bkq_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"}).to_excel(writer, index=False, sheet_name="Báo cáo kết quả hoạt động kinh doanh")
                    if has_cf:
                        pd.DataFrame(cf_extracted).T.reset_index().rename(columns={"index": "Chỉ tiêu"}).to_excel(writer, index=False, sheet_name="Báo cáo lưu chuyển tiền tệ")
                with open(tmp.name, "rb") as f:
                    st.download_button("📥 Tải Excel tổng hợp", data=f.read(), file_name="bao_cao_tai_chinh_tong_hop.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.markdown("🔧 Built with Streamlit + Plotly | 📊 Phân Tích Tài Chính v3.0")