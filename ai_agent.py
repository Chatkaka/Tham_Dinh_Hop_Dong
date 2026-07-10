import os
import re
import json
from pydantic import BaseModel, Field
from typing import Optional

# Check if google-genai is installed
try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

class TaskUpdateInfo(BaseModel):
    ma_ngan_sach: str = Field(description="Mã ngân sách hoặc STT công việc được nhắc đến, ví dụ: '28.11.2.1' hoặc 'TD.BĐS.28.11.2.1'")
    trang_thai: str = Field(description="Trạng thái của công việc: 'Todo', 'In-Progress', 'Done', hoặc 'Delayed'")
    tien_do: float = Field(description="Tiến độ công việc dưới dạng phần trăm (0 đến 100)")
    dieu_kien_ghi_nhan: str = Field(description="Điều kiện ghi nhận kết quả hoặc mô tả ngắn gọn nội dung công việc đã hoàn thành")

def get_gemini_client(api_key: str = None):
    # Use provided api_key, otherwise fallback to environment variable
    key = api_key or os.environ.get("GEMINI_API_KEY")
    if not key:
        return None
    if not HAS_GENAI:
        return None
    try:
        return genai.Client(api_key=key)
    except Exception as e:
        print(f"Không thể khởi tạo Gemini Client: {e}")
        return None

def parse_natural_language_update(text: str, api_key: str = None) -> dict:
    """
    Sử dụng Gemini Pro (hoặc regex fallback) để bóc tách văn bản cập nhật tiến độ công việc.
    """
    client = get_gemini_client(api_key)
    if client:
        try:
            prompt = (
                f"Hãy phân tích câu thông báo sau đây của quản lý dự án để bóc tách thông tin tiến độ công việc:\n"
                f"Câu thông báo: \"{text}\"\n\n"
                f"Hãy trích xuất mã công việc (ma_ngan_sach hoặc STT), xác định trạng thái (Todo, In-Progress, Done, Delayed), "
                f"tiến độ (0-100), và điều kiện ghi nhận kết quả."
            )
            
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=TaskUpdateInfo,
                    temperature=0.1
                ),
            )
            
            # Phản hồi dạng JSON tuân thủ schema
            data = json.loads(response.text)
            return {
                "ma_ngan_sach": data.get("ma_ngan_sach", ""),
                "trang_thai": data.get("trang_thai", "Done"),
                "tien_do": float(data.get("tien_do", 100.0)),
                "dieu_kien_ghi_nhan": data.get("dieu_kien_ghi_nhan", text),
                "method": "Gemini AI"
            }
        except Exception as e:
            print(f"Lỗi khi gọi Gemini API: {e}. Sử dụng Regex fallback...")
    
    # Regex fallback nếu không có API key hoặc lỗi
    return regex_parse_fallback(text)

def regex_parse_fallback(text: str) -> dict:
    """
    Trích xuất bằng Regex cho mục đích thử nghiệm offline
    """
    # Tìm chuỗi số dạng WBS hoặc STT (ví dụ: 28.11.2.1 hoặc TD.BĐS.28...)
    match = re.search(r'([A-Za-z\.]+)?\d+(\.\d+)+', text)
    ma_ngan_sach = match.group(0) if match else "1.1"

    # Đoán trạng thái và tiến độ từ từ khóa
    trang_thai = "Done"
    tien_do = 100.0
    
    lower_text = text.lower()
    if "chưa" in lower_text or "chậm" in lower_text:
        trang_thai = "Delayed"
        tien_do = 30.0
    elif "đang" in lower_text or "bắt đầu" in lower_text:
        trang_thai = "In-Progress"
        tien_do = 50.0
    elif "hoàn thành" in lower_text or "đã xong" in lower_text or "đã duyệt" in lower_text or "phê duyệt" in lower_text:
        trang_thai = "Done"
        tien_do = 100.0

    # Điều kiện ghi nhận là mô tả từ văn bản thô
    dieu_kien_ghi_nhan = text.strip()

    return {
        "ma_ngan_sach": ma_ngan_sach,
        "trang_thai": trang_thai,
        "tien_do": tien_do,
        "dieu_kien_ghi_nhan": dieu_kien_ghi_nhan,
        "method": "Regex Fallback (Offline)"
    }

def evaluate_financial_risk(project_total_budget: float, total_spent: float, spending_details: list, api_key: str = None) -> str:
    """
    Đánh giá rủi ro tài chính và tiến độ dựa trên số liệu ngân sách và thực chi.
    """
    client = get_gemini_client(api_key)
    ratio = (total_spent / project_total_budget * 100) if project_total_budget > 0 else 0.0
    
    if client:
        try:
            details_str = json.dumps(spending_details[:10], ensure_ascii=False)
            prompt = (
                f"Bạn là Giám đốc Tài chính (CFO) của dự án Bất động sản Ven Sông Vinh.\n"
                f"Hãy đánh giá rủi ro tài chính của dự án dựa trên số liệu sau:\n"
                f"- Tổng ngân sách dự án: {project_total_budget:,.2f} Trđ\n"
                f"- Lũy kế thực chi: {total_spent:,.2f} Trđ (Tỷ lệ giải ngân: {ratio:.2f}%)\n"
                f"- Chi tiết các khoản chi lớn/vượt ngân sách gần đây: {details_str}\n\n"
                f"Hãy viết một báo cáo đánh giá rủi ro tài chính và tiến độ ngắn gọn (dưới 200 từ), "
                f"chỉ ra các mối nguy hại tiềm ẩn (nếu có) và đưa ra 2 khuyến nghị hành động nhanh cho CEO."
            )
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt
            )
            return response.text
        except Exception as e:
            print(f"Lỗi khi gọi Gemini API đánh giá rủi ro: {e}")

    # Fallback báo cáo tài chính ngoại tuyến
    if ratio > 90:
        status_text = "CẢNH BÁO ĐỎ: Ngân sách đã giải ngân đạt mức báo động (>90%)."
        recs = "1. Dừng ngay việc duyệt tất cả các khoản chi phát sinh ngoài ngân sách.\n2. Tổ chức họp khẩn cấp với các Ban quản lý để rà soát tổng thể mức đầu tư."
    elif ratio > 75:
        status_text = "CẢNH BÁO VÀNG: Tỷ lệ giải ngân khá cao, cần kiểm soát chặt chẽ."
        recs = "1. Yêu cầu báo cáo chi tiết các hạng mục chuẩn bị thi công.\n2. Tối ưu hóa lại chi phí nhân công và vật tư xây dựng."
    else:
        status_text = "AN TOÀN: Tỷ lệ giải ngân nằm trong tầm kiểm soát."
        recs = "1. Tiếp tục bám sát tiến độ giải ngân theo kế hoạch quý.\n2. Phê duyệt nhanh các hạng mục thi công đúng tiến độ."

    report = (
        f"**BÁO CÁO ĐÁNH GIÁ RỦI RO TÀI CHÍNH DỰ ÁN (OFFLINE)**\n\n"
        f"- **Tổng ngân sách:** {project_total_budget:,.2f} Trđ\n"
        f"- **Lũy kế thực chi:** {total_spent:,.2f} Trđ ({ratio:.2f}%)\n"
        f"- **Đánh giá chung:** {status_text}\n\n"
        f"**Khuyến nghị hành động:**\n{recs}"
    )
    return report

class ExcelMappingInfo(BaseModel):
    ma_ngan_sach_idx: Optional[int] = Field(None, description="Index of column containing WBS code or budget code (e.g. 'TD.BĐS.1.2')")
    stt_idx: Optional[int] = Field(None, description="Index of column containing STT hierarchy number (e.g. '1.1.2')")
    ten_cong_viec_idx: Optional[int] = Field(None, description="Index of column containing task name or description")
    phong_ban_idx: Optional[int] = Field(None, description="Index of column containing department responsible")
    co_quan_idx: Optional[int] = Field(None, description="Index of column containing resolving agency")
    ho_so_dau_ra_idx: Optional[int] = Field(None, description="Index of column containing output files or deliverables")
    dieu_kien_ghi_nhan_idx: Optional[int] = Field(None, description="Index of column containing result recognition conditions")
    thoi_han_hoan_thanh_idx: Optional[int] = Field(None, description="Index of column containing deadline date or month")
    tien_do_idx: Optional[int] = Field(None, description="Index of column containing task progress percentage")
    trang_thai_idx: Optional[int] = Field(None, description="Index of column containing status (Todo, Done, In-Progress)")
    ngan_sach_idx: Optional[int] = Field(None, description="Index of column containing budget total in Trđ")

def extract_excel_mapping_with_ai(rows_data: list, api_key: str = None) -> dict:
    """
    Sử dụng Gemini Pro để bóc tách cột Excel mẫu thành các chỉ số trường dữ liệu phù hợp.
    Có fallback offline bằng heuristic nếu không có API key.
    """
    client = get_gemini_client(api_key)
    if client:
        try:
            prompt = (
                f"Bạn là chuyên gia phân tích dữ liệu dự án.\n"
                f"Dưới đây là một số dòng dữ liệu mẫu trích xuất từ file Excel (dòng đầu tiên có thể là tiêu đề hoặc chứa tiêu đề):\n"
                f"{json.dumps(rows_data, ensure_ascii=False)}\n\n"
                f"Hãy phân tích cấu trúc cột của bảng này và xác định chỉ số cột (0-based index) tương ứng với các trường dữ liệu sau. "
                f"Nếu trường nào không có trong bảng hoặc bạn không chắc chắn, hãy trả về null cho trường đó."
            )
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ExcelMappingInfo,
                    temperature=0.1
                ),
            )
            return json.loads(response.text)
        except Exception as e:
            print(f"Lỗi gọi Gemini bóc tách cột Excel: {e}. Chuyển sang Heuristic fallback...")
            
    # Heuristic fallback offline
    mapping = {
        "ma_ngan_sach_idx": None,
        "stt_idx": None,
        "ten_cong_viec_idx": None,
        "phong_ban_idx": None,
        "co_quan_idx": None,
        "ho_so_dau_ra_idx": None,
        "dieu_kien_ghi_nhan_idx": None,
        "thoi_han_hoan_thanh_idx": None,
        "tien_do_idx": None,
        "trang_thai_idx": None,
        "ngan_sach_idx": None
    }
    
    # Duyệt qua tối đa 5 dòng đầu để tìm dòng tiêu đề
    header_row = None
    for r in rows_data[:5]:
        if any(isinstance(val, str) and any(kw in val.lower() for kw in ["stt", "wbs", "tên công việc", "nội dung", "ngân sách", "phòng ban"]) for val in r if val):
            header_row = [str(x).lower().strip() if x is not None else "" for x in r]
            break
            
    if not header_row and len(rows_data) > 0:
        header_row = [str(x).lower().strip() if x is not None else "" for x in rows_data[0]]
        
    if header_row:
        for idx, col in enumerate(header_row):
            if "wbs" in col or "mã ngân sách" in col or "mã công việc" in col:
                mapping["ma_ngan_sach_idx"] = idx
            elif "stt" in col or "số thứ tự" in col:
                mapping["stt_idx"] = idx
            elif "tên" in col or "nội dung" in col or "công việc" in col:
                if mapping["ten_cong_viec_idx"] is None:
                    mapping["ten_cong_viec_idx"] = idx
            elif "phòng" in col or "phụ trách" in col or "ban" in col:
                mapping["phong_ban_idx"] = idx
            elif "cơ quan" in col or "giải quyết" in col:
                mapping["co_quan_idx"] = idx
            elif "hồ sơ" in col or "kết quả" in col or "đầu ra" in col:
                mapping["ho_so_dau_ra_idx"] = idx
            elif "điều kiện" in col or "ghi nhận" in col:
                mapping["dieu_kien_ghi_nhan_idx"] = idx
            elif "thời hạn" in col or "ngày hoàn thành" in col or "deadline" in col:
                mapping["thoi_han_hoan_thanh_idx"] = idx
            elif "tiến độ" in col or "tiến trình" in col or "phần trăm" in col:
                mapping["tien_do_idx"] = idx
            elif "trạng thái" in col or "status" in col:
                mapping["trang_thai_idx"] = idx
            elif "ngân sách" in col or "dự toán" in col or "budget" in col or "tổng" in col:
                if mapping["ngan_sach_idx"] is None:
                    mapping["ngan_sach_idx"] = idx

    # Đảm bảo các cột tối thiểu có giá trị mặc định nếu heuristic không tìm thấy
    if mapping["stt_idx"] is None: mapping["stt_idx"] = 2
    if mapping["ma_ngan_sach_idx"] is None: mapping["ma_ngan_sach_idx"] = 1
    if mapping["ten_cong_viec_idx"] is None: mapping["ten_cong_viec_idx"] = 3
    if mapping["phong_ban_idx"] is None: mapping["phong_ban_idx"] = 4
    if mapping["co_quan_idx"] is None: mapping["co_quan_idx"] = 6
    if mapping["ho_so_dau_ra_idx"] is None: mapping["ho_so_dau_ra_idx"] = 7
    if mapping["dieu_kien_ghi_nhan_idx"] is None: mapping["dieu_kien_ghi_nhan_idx"] = 11
    
    return mapping

def generate_generic_text(prompt: str, api_key: str = None) -> str:
    client = get_gemini_client(api_key)
    if client:
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt
            )
            return response.text
        except Exception as e:
            return f"Lỗi gọi Gemini: {e}"
    return "Gemini API Client không khả dụng (thiếu API Key hoặc lỗi cài đặt thư viện)."
