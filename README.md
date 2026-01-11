# Excel Metadata Tool

Công cụ tự động tạo file CSV barcode từ Excel nguồn và template, kèm giao diện desktop (Tkinter).

## Tính năng
- Lấy thông tin run từ ô H14, tạo header barcode và xuất CSV theo từng file nguồn.
- Dùng workbook template để tra cứu (VLOOKUP) dữ liệu primer và ghi kết quả từ dòng 4.
- Chế độ thư mục xử lý toàn bộ file Excel trong folder, gom và báo lỗi chi tiết theo từng file.
- Giao diện hiển thị tiến trình, thời gian chạy, gợi ý nhanh tên file/thư mục đã chọn; tạo thư mục đích theo giá trị H14.
- Chặn ghi đè: nếu CSV đích đã tồn tại sẽ báo lỗi, yêu cầu xóa/đổi tên trước khi chạy lại.

## Yêu cầu
- Python 3.8+
- Thư viện: `pandas`, `openpyxl`
- Windows (dùng `os.startfile` và hỗ trợ icon `.ico` cho UI)


## Chạy giao diện (UI)
```bash
python Fontend.py
```
- Chọn Source File / Source Folder, Template và Output.
- Nhấn **Convert File** hoặc **Convert Folder**.
- Kết quả: CSV nằm trong `output/<H14_da_lam_sach>/barcode_<Runname>.csv`.

## Dùng mã nguồn trực tiếp (không UI)
```python
from backend import process_excel
process_excel(source_path, template_path, output_path)

from backend_funtion_convert_folder import process_folder
process_folder(source_folder, template_path, output_path)
```

## Xử lý lỗi
- H14 trống hoặc sai định dạng: báo lỗi rõ ràng.
- Chế độ folder: gom lỗi và liệt kê file nào hỏng.
- File CSV đã tồn tại: dừng và yêu cầu xoá/đổi tên trước khi chạy lại.

## Ghi chú
- Tên thư mục đích lấy đúng giá trị H14 (được làm sạch ký tự cấm của Windows).
- Run name được chuẩn hoá bắt đầu bằng `R`, dùng cho tên file CSV.
