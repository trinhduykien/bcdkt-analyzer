import pandas as pd

# Sample BCĐKT data
data = {
    "Mã": ["01", "02", "03", "04", "05", "06", "10", "11", "12", "13", "30", "31", "32"],
    "Tên chỉ tiêu": [
        "Tổng tài sản",
        "Tài sản ngắn hạn",
        "Tiền và tương đương tiền",
        "Phải thu ngắn hạn",
        "Hàng tồn kho",
        "Tài sản dài hạn",
        "Tổng nợ phải trả",
        "Nợ ngắn hạn",
        "Nợ dài hạn",
        "Nợ khác",
        "Vốn chủ sở hữu",
        "Vốn điều lệ",
        "Lợi nhuận chưa phân phối",
    ],
    "Số đầu kỳ": [50000, 30000, 10000, 8000, 12000, 20000, 30000, 18000, 12000, None, 20000, 15000, 5000],
    "Số cuối kỳ": [65000, 40000, 15000, 10000, 15000, 25000, 38000, 22000, 16000, None, 27000, 15000, 12000],
}

df = pd.DataFrame(data)
df.to_excel("D:\\WORK\\bcdkt_analyzer\\sample_bcdkt.xlsx", index=False, sheet_name="BCDKT")
print("Sample file created: sample_bcdkt.xlsx")