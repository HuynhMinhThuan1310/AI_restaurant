# Vietnamese Sentiment App

Ứng dụng Flask dùng bộ phân tích đánh giá nhà hàng mới đã train:

```text
C:\AI\results\tfidf_restaurant_review_3class_final_overall_model.joblib
```

Bộ phân tích chia đánh giá nhà hàng/quán ăn thành 3 nhóm hiển thị: `Không hài lòng`, `Trung lập`, `Hài lòng`.

Chỉ số kiểm tra chính: Độ chính xác `0.8406`, Độ cân bằng `0.8038`, Điểm tổng thể `0.8417`.

## Chạy local

```powershell
Set-Location -LiteralPath "C:\AI\sentiment_app"
python -m venv ".venv"
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

Nếu chạy trong VS Code, hãy chọn interpreter:

```text
C:\AI\sentiment_app\.venv\Scripts\python.exe
```

Hoặc chạy trực tiếp bằng:

```powershell
& "C:\AI\sentiment_app\.venv\Scripts\python.exe" "C:\AI\sentiment_app\app.py"
```

Mở trình duyệt tại:

```text
http://127.0.0.1:5000
```

## Input file

App hỗ trợ `.xlsx`, `.xls`, `.csv`. File cần có một cột văn bản với tên một trong các tên sau:

```text
sentence, text, feedback, comment, review, content
```

File kết quả sẽ nằm trong thư mục `outputs/` và có thêm cột:

```text
sentiment_result
```
