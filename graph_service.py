"""
backend/graph_service.py — Matplotlib Graph Engine
Tạo đồ thị / biểu đồ chuẩn xác cho đề thi, trả về PNG (base64 hoặc file).

Hỗ trợ các loại:
  1. do_thi_ham_so     — Đồ thị hàm số (bậc 1/2/3, sin, cos, tan, log, exp...)
  2. bieu_do_cot       — Biểu đồ cột (bar chart)
  3. bieu_do_duong     — Biểu đồ đường (line chart)
  4. bieu_do_tron      — Biểu đồ tròn (pie chart)
  5. histogram         — Histogram + đường cong chuẩn
  6. hinh_hoc_oxy      — Hình học trên hệ trục Oxy (điểm, đoạn thẳng, đa giác, đường tròn)
  7. do_thi_vat_ly     — Đồ thị vật lý (v-t, s-t, U-I...)
"""

import matplotlib
matplotlib.use('Agg')  # Non-interactive backend — tránh lỗi GUI trên server

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import io
import base64
import os
import uuid
import math

# ══════════════════════════════════════════════════════════════════════
#  CẤU HÌNH CHUNG CHO TẤT CẢ ĐỒ THỊ
# ══════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    'font.family': 'serif',
    'font.serif': ['Times New Roman', 'DejaVu Serif'],
    'font.size': 11,
    'axes.labelsize': 12,
    'axes.titlesize': 13,
    'figure.dpi': 150,
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.linestyle': '--',
})

# Thư mục lưu ảnh tạm
GRAPH_OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'temp_graphs')
os.makedirs(GRAPH_OUTPUT_DIR, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════
#  1. ĐỒ THỊ HÀM SỐ
# ══════════════════════════════════════════════════════════════════════
def ve_do_thi_ham_so(data: dict) -> dict:
    """
    Vẽ đồ thị hàm số trên mặt phẳng Oxy.
    
    data = {
        "hamSo": "x**2 - 4*x + 3",         # Biểu thức Python (dùng numpy)
        "xRange": [-2, 6],                   # Khoảng x
        "yRange": [-2, 10],                  # Khoảng y (tùy chọn, auto nếu không có)
        "tieuDe": "y = x² - 4x + 3",        # Tiêu đề (tùy chọn)
        "nhanX": "x", "nhanY": "y",          # Nhãn trục
        "diemDacBiet": [                      # Các điểm đánh dấu (tùy chọn)
            {"x": 1, "y": 0, "nhan": "A(1;0)"},
            {"x": 3, "y": 0, "nhan": "B(3;0)"},
            {"x": 2, "y": -1, "nhan": "Đỉnh(2;-1)"}
        ],
        "duongPhu": [                         # Các đường phụ (tùy chọn)
            {"hamSo": "0", "mau": "gray", "nhan": "y = 0"}
        ]
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(6, 4.5))
    
    x_range = data.get('xRange', [-5, 5])
    x = np.linspace(x_range[0], x_range[1], 500)
    
    # Vẽ hàm số chính
    ham_so = data.get('hamSo', 'x')
    try:
        y = eval(ham_so, {"__builtins__": {}, "x": x, "np": np,
                          "sin": np.sin, "cos": np.cos, "tan": np.tan,
                          "log": np.log, "exp": np.exp, "sqrt": np.sqrt,
                          "abs": np.abs, "pi": np.pi, "e": np.e})
    except Exception as e:
        ax.text(0.5, 0.5, f"Lỗi biểu thức: {ham_so}\n{e}",
                transform=ax.transAxes, ha='center', va='center',
                fontsize=12, color='red')
        return _save_and_return(fig)
    
    # Giới hạn y để tránh vẽ vô cực
    y_finite = np.where(np.isfinite(y), y, np.nan)
    ax.plot(x, y_finite, 'b-', linewidth=2, label=data.get('tieuDe', f'y = {ham_so}'))
    
    # Vẽ các đường phụ
    for dp in data.get('duongPhu', []):
        try:
            y_phu = eval(dp['hamSo'], {"__builtins__": {}, "x": x, "np": np,
                                        "sin": np.sin, "cos": np.cos, "tan": np.tan,
                                        "log": np.log, "exp": np.exp, "sqrt": np.sqrt,
                                        "abs": np.abs, "pi": np.pi, "e": np.e})
            y_phu_f = np.where(np.isfinite(y_phu), y_phu, np.nan)
            ax.plot(x, y_phu_f, color=dp.get('mau', 'gray'), linewidth=1,
                    linestyle='--', label=dp.get('nhan', ''))
        except:
            pass
    
    # Vẽ trục Ox, Oy
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.axvline(x=0, color='black', linewidth=0.8)
    
    # Đánh dấu điểm đặc biệt
    for diem in data.get('diemDacBiet', []):
        px, py = diem.get('x', 0), diem.get('y', 0)
        ax.plot(px, py, 'ro', markersize=6)
        nhan = diem.get('nhan', f'({px};{py})')
        ax.annotate(nhan, (px, py), textcoords="offset points",
                    xytext=(8, 8), fontsize=9, color='red',
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='yellow', alpha=0.7))
    
    # Cấu hình trục
    if data.get('yRange'):
        ax.set_ylim(data['yRange'])
    else:
        y_valid = y_finite[np.isfinite(y_finite)]
        if len(y_valid) > 0:
            y_min, y_max = np.min(y_valid), np.max(y_valid)
            margin = (y_max - y_min) * 0.15 or 1
            ax.set_ylim(y_min - margin, y_max + margin)
    
    ax.set_xlim(x_range)
    ax.set_xlabel(data.get('nhanX', 'x'))
    ax.set_ylabel(data.get('nhanY', 'y'))
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    ax.legend(loc='best', fontsize=9)
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  2. BIỂU ĐỒ CỘT
# ══════════════════════════════════════════════════════════════════════
def ve_bieu_do_cot(data: dict) -> dict:
    """
    data = {
        "tieuDe": "Dân số các nước ĐNÁ 2023",
        "nhanX": "Quốc gia", "nhanY": "Dân số (triệu người)",
        "nhan": ["VN", "Thái Lan", "Indonesia", "Philippines"],
        "giaTri": [100, 72, 275, 114],
        "mau": ["#3b82f6", "#10b981", "#f59e0b", "#ef4444"],  # tùy chọn
        "hienGiaTri": true  # hiện số trên cột
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(6, 4.5))
    
    nhan = data.get('nhan', [])
    gia_tri = data.get('giaTri', [])
    mau = data.get('mau', None)
    
    if not mau or len(mau) < len(nhan):
        mau = plt.cm.Set2(np.linspace(0, 1, len(nhan)))
    
    bars = ax.bar(nhan, gia_tri, color=mau, edgecolor='white', linewidth=0.8)
    
    # Hiện giá trị trên đầu cột
    if data.get('hienGiaTri', True):
        for bar, val in zip(bars, gia_tri):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(gia_tri) * 0.02,
                    str(val), ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    ax.set_xlabel(data.get('nhanX', ''))
    ax.set_ylabel(data.get('nhanY', ''))
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    ax.set_ylim(0, max(gia_tri) * 1.2)
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  3. BIỂU ĐỒ ĐƯỜNG
# ══════════════════════════════════════════════════════════════════════
def ve_bieu_do_duong(data: dict) -> dict:
    """
    data = {
        "tieuDe": "GDP Việt Nam 2018-2023",
        "nhanX": "Năm", "nhanY": "GDP (tỷ USD)",
        "nhan": ["2018", "2019", "2020", "2021", "2022", "2023"],
        "chuoiDuLieu": [
            {"ten": "Việt Nam", "giaTri": [245, 262, 271, 366, 409, 430]},
            {"ten": "Thái Lan", "giaTri": [506, 544, 500, 506, 495, 512]}
        ]
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(6, 4.5))
    
    nhan = data.get('nhan', [])
    colors = ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6']
    markers = ['o', 's', '^', 'D', 'v']
    
    for i, chuoi in enumerate(data.get('chuoiDuLieu', [])):
        color = colors[i % len(colors)]
        marker = markers[i % len(markers)]
        ax.plot(nhan, chuoi['giaTri'], color=color, marker=marker,
                linewidth=2, markersize=6, label=chuoi.get('ten', f'Chuỗi {i+1}'))
    
    ax.set_xlabel(data.get('nhanX', ''))
    ax.set_ylabel(data.get('nhanY', ''))
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    ax.legend(loc='best')
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  4. BIỂU ĐỒ TRÒN
# ══════════════════════════════════════════════════════════════════════
def ve_bieu_do_tron(data: dict) -> dict:
    """
    data = {
        "tieuDe": "Cơ cấu kinh tế Việt Nam 2023",
        "nhan": ["Nông nghiệp", "Công nghiệp", "Dịch vụ"],
        "giaTri": [12, 38, 50],
        "mau": ["#10b981", "#3b82f6", "#f59e0b"]  # tùy chọn
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(5.5, 4.5))
    
    nhan = data.get('nhan', [])
    gia_tri = data.get('giaTri', [])
    mau = data.get('mau', None)
    
    if not mau or len(mau) < len(nhan):
        mau = plt.cm.Set2(np.linspace(0, 1, len(nhan)))
    
    # Explode miếng lớn nhất
    explode = [0.05 if v == max(gia_tri) else 0 for v in gia_tri]
    
    wedges, texts, autotexts = ax.pie(
        gia_tri, labels=nhan, autopct='%1.1f%%', startangle=90,
        colors=mau, explode=explode, shadow=False,
        textprops={'fontsize': 10}
    )
    for autotext in autotexts:
        autotext.set_fontweight('bold')
    
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold', pad=15)
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  5. HISTOGRAM
# ══════════════════════════════════════════════════════════════════════
def ve_histogram(data: dict) -> dict:
    """
    data = {
        "tieuDe": "Phân bố điểm thi lớp 12A",
        "nhanX": "Điểm", "nhanY": "Số học sinh",
        "duLieu": [3, 4, 5, 5, 6, 6, 6, 7, 7, 7, 7, 8, 8, 8, 9, 9, 10],
        "soCot": 8,                   # Số bin (tùy chọn)
        "veDuongCong": true            # Vẽ đường cong chuẩn (tùy chọn)
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(6, 4.5))
    
    du_lieu = data.get('duLieu', [])
    so_cot = data.get('soCot', 10)
    
    n, bins, p = ax.hist(du_lieu, bins=so_cot, color='#3b82f6',
                          edgecolor='white', alpha=0.8, density=False)
    
    if data.get('veDuongCong', False) and len(du_lieu) > 5:
        from scipy import stats
        mu, sigma = np.mean(du_lieu), np.std(du_lieu)
        x_curve = np.linspace(min(du_lieu), max(du_lieu), 100)
        y_curve = stats.norm.pdf(x_curve, mu, sigma) * len(du_lieu) * (bins[1] - bins[0])
        ax.plot(x_curve, y_curve, 'r-', linewidth=2, label=f'μ={mu:.1f}, σ={sigma:.1f}')
        ax.legend()
    
    ax.set_xlabel(data.get('nhanX', ''))
    ax.set_ylabel(data.get('nhanY', 'Tần số'))
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  6. HÌNH HỌC TRÊN OXY
# ══════════════════════════════════════════════════════════════════════
def ve_hinh_hoc_oxy(data: dict) -> dict:
    """
    data = {
        "tieuDe": "Tam giác ABC",
        "xRange": [-1, 6], "yRange": [-1, 5],
        "diem": [
            {"x": 0, "y": 0, "nhan": "A"},
            {"x": 5, "y": 0, "nhan": "B"},
            {"x": 2, "y": 4, "nhan": "C"}
        ],
        "doanThang": [
            {"x1": 0, "y1": 0, "x2": 5, "y2": 0},
            {"x1": 5, "y1": 0, "x2": 2, "y2": 4},
            {"x1": 2, "y1": 4, "x2": 0, "y2": 0}
        ],
        "duongTron": [
            {"tamX": 2, "tamY": 1, "banKinh": 3, "nhan": "(O;3)"}
        ]
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(5.5, 5))
    
    x_range = data.get('xRange', [-5, 5])
    y_range = data.get('yRange', [-5, 5])
    
    # Vẽ trục
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.axvline(x=0, color='black', linewidth=0.8)
    
    # Vẽ đoạn thẳng
    for dt in data.get('doanThang', []):
        ax.plot([dt['x1'], dt['x2']], [dt['y1'], dt['y2']],
                '-', linewidth=2, color=dt.get('mau', '#3b82f6'))
    
    # Vẽ đường tròn
    for dc in data.get('duongTron', []):
        circle = plt.Circle((dc['tamX'], dc['tamY']), dc['banKinh'],
                             fill=False, color=dc.get('mau', '#3b82f6'),
                             linewidth=2)
        ax.add_patch(circle)
        if dc.get('nhan'):
            ax.annotate(dc['nhan'], (dc['tamX'], dc['tamY']),
                        textcoords="offset points", xytext=(5, 5),
                        fontsize=9, color='blue')
    
    # Vẽ điểm
    for diem in data.get('diem', []):
        ax.plot(diem['x'], diem['y'], 'ko', markersize=5)
        nhan = diem.get('nhan', f"({diem['x']};{diem['y']})")
        offset_x = diem.get('offsetX', 8)
        offset_y = diem.get('offsetY', 8)
        ax.annotate(nhan, (diem['x'], diem['y']),
                    textcoords="offset points", xytext=(offset_x, offset_y),
                    fontsize=10, fontweight='bold', color='black')
    
    ax.set_xlim(x_range)
    ax.set_ylim(y_range)
    ax.set_aspect('equal')
    ax.set_xlabel('x')
    ax.set_ylabel('y')
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  7. ĐỒ THỊ VẬT LÝ (v-t, s-t, U-I, P-V...)
# ══════════════════════════════════════════════════════════════════════
def ve_do_thi_vat_ly(data: dict) -> dict:
    """
    data = {
        "tieuDe": "Đồ thị vận tốc - thời gian",
        "nhanX": "t (s)", "nhanY": "v (m/s)",
        "doanThang": [
            {"x1": 0, "y1": 0, "x2": 5, "y2": 20},    # Gia tốc đều
            {"x1": 5, "y1": 20, "x2": 10, "y2": 20},   # Chuyển động đều
            {"x1": 10, "y1": 20, "x2": 15, "y2": 0}    # Giảm tốc
        ],
        "diemDacBiet": [
            {"x": 5, "y": 20, "nhan": "A(5;20)"}
        ]
    }
    """
    fig, ax = plt.subplots(1, 1, figsize=(6, 4.5))
    
    # Vẽ các đoạn thẳng (đặc trưng của đồ thị vật lý)
    for dt in data.get('doanThang', []):
        ax.plot([dt['x1'], dt['x2']], [dt['y1'], dt['y2']],
                'b-', linewidth=2, color=dt.get('mau', '#3b82f6'))
    
    # Vẽ hàm số liên tục (nếu có)
    if data.get('hamSo'):
        x_range = data.get('xRange', [0, 10])
        x = np.linspace(x_range[0], x_range[1], 300)
        try:
            y = eval(data['hamSo'], {"__builtins__": {}, "x": x, "np": np,
                                      "sin": np.sin, "cos": np.cos,
                                      "sqrt": np.sqrt, "pi": np.pi})
            y_f = np.where(np.isfinite(y), y, np.nan)
            ax.plot(x, y_f, 'b-', linewidth=2)
        except:
            pass
    
    # Trục
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.axvline(x=0, color='black', linewidth=0.8)
    
    # Điểm đặc biệt
    for diem in data.get('diemDacBiet', []):
        ax.plot(diem['x'], diem['y'], 'ro', markersize=6)
        nhan = diem.get('nhan', f"({diem['x']};{diem['y']})")
        ax.annotate(nhan, (diem['x'], diem['y']),
                    textcoords="offset points", xytext=(8, 8),
                    fontsize=9, color='red',
                    bbox=dict(boxstyle='round,pad=0.2', facecolor='yellow', alpha=0.7))
    
    ax.set_xlabel(data.get('nhanX', ''))
    ax.set_ylabel(data.get('nhanY', ''))
    if data.get('tieuDe'):
        ax.set_title(data['tieuDe'], fontweight='bold')
    
    fig.tight_layout()
    return _save_and_return(fig)


# ══════════════════════════════════════════════════════════════════════
#  DISPATCHER — Gọi đúng hàm vẽ theo loại
# ══════════════════════════════════════════════════════════════════════
GRAPH_HANDLERS = {
    'do_thi_ham_so':  ve_do_thi_ham_so,
    'bieu_do_cot':    ve_bieu_do_cot,
    'bieu_do_duong':  ve_bieu_do_duong,
    'bieu_do_tron':   ve_bieu_do_tron,
    'histogram':      ve_histogram,
    'hinh_hoc_oxy':   ve_hinh_hoc_oxy,
    'do_thi_vat_ly':  ve_do_thi_vat_ly,
}

def generate_graph(data: dict) -> dict:
    """
    Entry point chính. Nhận JSON metadata, trả về dict chứa:
    {
        "success": True/False,
        "base64": "data:image/png;base64,...",
        "filePath": "/path/to/graph.png",
        "error": "..."  (nếu lỗi)
    }
    """
    loai = data.get('loai', '')
    handler = GRAPH_HANDLERS.get(loai)
    
    if not handler:
        return {
            "success": False,
            "error": f"Loại đồ thị không hỗ trợ: '{loai}'. "
                     f"Các loại hỗ trợ: {list(GRAPH_HANDLERS.keys())}"
        }
    
    try:
        result = handler(data)
        return result
    except Exception as e:
        return {"success": False, "error": f"Lỗi khi vẽ đồ thị: {str(e)}"}


# ══════════════════════════════════════════════════════════════════════
#  HELPER — Lưu figure → PNG → base64 + file
# ══════════════════════════════════════════════════════════════════════
def _save_and_return(fig) -> dict:
    """Lưu matplotlib figure thành PNG, trả về base64 và file path."""
    try:
        # Tạo base64
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.1)
        buf.seek(0)
        b64_str = base64.b64encode(buf.read()).decode('utf-8')
        buf.close()
        
        # Lưu file
        filename = f"graph_{uuid.uuid4().hex[:12]}.png"
        filepath = os.path.join(GRAPH_OUTPUT_DIR, filename)
        fig.savefig(filepath, format='png', bbox_inches='tight', pad_inches=0.1)
        
        plt.close(fig)
        
        return {
            "success": True,
            "base64": f"data:image/png;base64,{b64_str}",
            "filePath": filepath,
            "fileName": filename
        }
    except Exception as e:
        plt.close(fig)
        return {"success": False, "error": f"Lỗi lưu ảnh: {str(e)}"}
