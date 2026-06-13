# FoodMood

FoodMood là ứng dụng web phân tích cảm xúc đánh giá nhà hàng tiếng Việt.

## Truy cập app

Mở link sau để dùng trực tiếp:

https://foodmood-cgl8.onrender.com/

Lưu ý: Render bản miễn phí có thể mất khoảng 30-60 giây để khởi động nếu lâu không ai truy cập.

## Cách sử dụng

1. Vào link app.
2. Chọn trang `Đánh giá`.
3. Nhập một câu review hoặc tải file Excel/CSV.
4. Xem kết quả phân loại: `Hài lòng`, `Trung lập`, `Không hài lòng`.

## Upload file

App hỗ trợ các định dạng:

```text
.xlsx, .xls, .csv
File cần có một cột nội dung review với một trong các tên:
sentence, text, feedback, comment, review, content
Chạy local
cd sentiment_app
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
Sau đó mở:
http://127.0.0.1:5000
