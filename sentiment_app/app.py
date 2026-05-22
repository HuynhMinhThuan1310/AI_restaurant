import re
import unicodedata
from datetime import datetime
from pathlib import Path

import joblib
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
MODEL_CANDIDATES = [
    BASE_DIR.parent / "results" / "tfidf_restaurant_review_3class_final_overall_model.joblib",
    BASE_DIR / "tfidf_restaurant_review_3class_final_overall_model.joblib",
]
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}
TEXT_CANDIDATES = ["sentence", "text", "feedback", "comment", "review", "content"]
SENTIMENT_LABELS = ["negative", "neutral", "positive"]
LABEL_DISPLAY = {
    "negative": "Không hài lòng",
    "neutral": "Trung lập",
    "positive": "Hài lòng",
}
ASPECT_CONFIG = [
    {
        "key": "food",
        "title": "Món ăn",
        "keywords": [
            "món", "đồ ăn", "thức ăn", "nước dùng", "nước chấm", "đồ uống", "món chính", "khẩu vị",
            "hải sản", "cơm", "bún", "phở", "lẩu", "mì", "miến", "gà", "bò", "cá", "thịt", "rau",
            "bánh", "tráng miệng", "topping", "gia vị", "sốt", "menu",
        ],
        "positive_terms": [
            "ngon", "rất ngon", "ngon miệng", "vừa miệng", "vừa ăn", "tươi", "tươi ngon", "đậm đà",
            "thơm", "giòn", "mềm", "nóng hổi", "hấp dẫn", "đẹp mắt", "sạch sẽ", "chất lượng",
            "đầy đặn", "nhiều topping", "nhiều thịt", "khá ngon", "ổn áp", "xuất sắc", "tuyệt vời",
            "không chê được", "hài lòng",
        ],
        "neutral_terms": [
            "tạm", "tạm ổn", "tạm được", "bình thường", "trung bình", "ổn", "được", "chấp nhận được",
            "không đặc sắc", "không quá đặc biệt", "không có gì đặc biệt", "vừa phải", "ăn được",
        ],
        "negative_terms": [
            "dở", "tệ", "chán", "không ngon", "quá tệ", "nguội", "nguội lạnh", "nhạt", "nhạt nhẽo",
            "mặn", "mặn chát", "ngọt gắt", "cay quá", "dầu mỡ", "nhiều dầu", "tanh", "hôi", "khô",
            "bở", "sống", "cháy", "cũ", "ít", "ít thịt", "ít topping", "không tươi", "không sạch",
            "không hợp khẩu vị", "không đáng ăn", "thất vọng",
        ],
    },
    {
        "key": "service",
        "title": "Nhân viên/phục vụ",
        "keywords": [
            "nhân viên", "bạn nhân viên", "phục vụ", "thái độ", "order", "gọi món", "xử lý",
            "chu đáo", "nhiệt tình", "dễ thương", "thu ngân", "bảo vệ", "quản lý", "chủ quán",
            "staff", "waiter", "waitress", "tư vấn", "hỗ trợ",
        ],
        "positive_terms": [
            "nhiệt tình", "dễ thương", "chu đáo", "nhanh", "nhanh nhẹn", "thân thiện", "niềm nở",
            "vui vẻ", "lịch sự", "lễ phép", "tận tâm", "chuyên nghiệp", "tốt", "hỗ trợ tốt",
            "tư vấn kỹ", "xử lý nhanh", "không chậm", "hài lòng",
        ],
        "neutral_terms": [
            "bình thường", "ổn", "tạm", "tạm ổn", "được", "không vấn đề", "đúng quy trình",
            "không quá nổi bật", "chấp nhận được",
        ],
        "negative_terms": [
            "thái độ", "thái độ kém", "khó chịu", "cau có", "cọc", "bất lịch sự", "không thân thiện",
            "không nhiệt tình", "hời hợt", "thờ ơ", "lơ khách", "không quan tâm", "chậm", "chậm chạp",
            "lâu", "phục vụ lâu", "phục vụ chậm", "sai order", "order sai", "quên món", "xử lý chậm",
            "thiếu chuyên nghiệp", "tệ", "quát", "không xin lỗi", "bỏ mặc khách",
        ],
    },
    {
        "key": "space",
        "title": "Không gian/địa điểm",
        "keywords": [
            "không gian", "quán", "địa điểm", "vị trí", "chỗ ngồi", "bàn ghế", "nhà vệ sinh",
            "gửi xe", "view", "decor", "trang trí", "âm nhạc", "ánh sáng", "máy lạnh", "mùi", "sạch",
            "đẹp", "đông", "ồn",
        ],
        "positive_terms": [
            "đẹp", "sạch", "sạch sẽ", "thoáng", "thoáng mát", "rộng", "rộng rãi", "ấm cúng",
            "dễ chịu", "yên tĩnh", "view đẹp", "decor đẹp", "mát", "phù hợp", "gọn gàng", "hài lòng",
        ],
        "neutral_terms": ["ổn", "bình thường", "tạm", "tạm ổn", "được", "vừa đủ", "không quá rộng"],
        "negative_terms": [
            "ồn", "ồn ào", "bẩn", "không sạch", "chật", "chật chội", "nóng", "bí", "ngột ngạt",
            "đông", "quá đông", "khó chịu", "hôi", "mùi khó chịu", "nhà vệ sinh bẩn", "thiếu chỗ ngồi",
            "khó gửi xe", "bàn ghế bẩn", "không thoáng",
        ],
    },
    {
        "key": "price",
        "title": "Giá cả",
        "keywords": [
            "giá", "giá cả", "giá tiền", "đắt", "rẻ", "hợp lý", "cao", "mắc", "phí", "chi phí",
            "hóa đơn", "bill", "voucher", "khuyến mãi", "combo", "phần ăn", "suất", "size",
        ],
        "positive_terms": [
            "hợp lý", "rẻ", "giá tốt", "đáng tiền", "đáng giá", "phù hợp", "vừa túi tiền", "nhiều so với giá",
            "không đắt", "khuyến mãi tốt", "hài lòng",
        ],
        "neutral_terms": ["tạm", "bình thường", "ổn", "tạm ổn", "vừa phải", "chấp nhận được", "không quá đắt"],
        "negative_terms": [
            "đắt", "mắc", "cao", "quá cao", "hơi đắt", "khá đắt", "không đáng", "không đáng tiền",
            "ít", "phần ít", "giá chát", "phụ thu", "tính phí", "bill sai", "không rõ giá",
        ],
    },
    {
        "key": "wait",
        "title": "Thời gian chờ",
        "keywords": [
            "chờ", "đợi", "lâu", "nhanh", "lên món", "món lên", "ra món", "đợi món", "lên đồ",
            "phục vụ lâu", "phục vụ nhanh", "xếp hàng", "thời gian", "chờ bàn", "mang món",
        ],
        "positive_terms": [
            "nhanh", "rất nhanh", "lên nhanh", "ra nhanh", "lên món nhanh", "phục vụ nhanh", "không phải chờ",
            "không chờ lâu", "không lâu", "đúng giờ", "mang món nhanh", "order nhanh",
        ],
        "neutral_terms": ["ổn", "tạm", "tạm ổn", "bình thường", "vừa phải", "chấp nhận được"],
        "negative_terms": [
            "lâu", "rất lâu", "quá lâu", "chờ lâu", "đợi lâu", "đợi món lâu", "chậm", "lên chậm",
            "ra chậm", "phục vụ lâu", "phục vụ chậm", "món lên lâu", "xếp hàng lâu", "chờ bàn lâu",
            "order lâu", "quên món",
        ],
    },
]
MODEL_TITLE = "Bộ phân tích đánh giá nhà hàng"
MODEL_METRICS = {
    "accuracy": 0.8406,
    "macro_f1": 0.8038,
    "weighted_f1": 0.8417,
}

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


def find_model_path():
    for path in MODEL_CANDIDATES:
        if path.exists():
            return path
    checked = ", ".join(str(path) for path in MODEL_CANDIDATES)
    raise FileNotFoundError(f"Không tìm thấy bộ phân tích đánh giá nhà hàng. Đã kiểm tra: {checked}")

def restore_sklearn_classes(estimator):
    """Handle RidgeClassifier joblib files saved by a different sklearn version."""
    if hasattr(estimator, "steps"):
        for _, step in estimator.steps:
            restore_sklearn_classes(step)
        return estimator

    if hasattr(estimator, "estimators_"):
        for child in estimator.estimators_:
            restore_sklearn_classes(child)

    if not hasattr(estimator, "classes_"):
        label_binarizer = getattr(estimator, "_label_binarizer", None)
        if label_binarizer is not None and hasattr(label_binarizer, "classes_"):
            estimator.classes_ = label_binarizer.classes_

    return estimator


MODEL_PATH = find_model_path()
model = restore_sklearn_classes(joblib.load(MODEL_PATH))


def clean_text(text):
    text = unicodedata.normalize("NFC", str(text)).lower()
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def detect_text_column(columns):
    normalized_columns = {str(col).strip().lower(): col for col in columns}
    for candidate in TEXT_CANDIDATES:
        if candidate in normalized_columns:
            return normalized_columns[candidate]
    return None


def read_table(file_path):
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(file_path)
    if suffix == ".csv":
        for encoding in ["utf-8-sig", "utf-8", "cp1258", "latin1"]:
            try:
                return pd.read_csv(file_path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(file_path)
    raise ValueError("Chỉ hỗ trợ file .xlsx, .xls hoặc .csv")


def predict_texts(texts):
    cleaned = [clean_text(text) for text in texts]
    try:
        return model.predict(cleaned)
    except AttributeError as exc:
        if "classes_" not in str(exc):
            raise
        restore_sklearn_classes(model)
        return model.predict(cleaned)


def display_label(label):
    return LABEL_DISPLAY.get(str(label), str(label))


def split_review_segments(text):
    text = str(text).strip()
    if not text:
        return []
    text = re.sub(r"\b(nhưng|tuy nhiên|còn|bù lại|ngoài ra|trong khi|đồng thời)\b", ".", text, flags=re.IGNORECASE)
    parts = re.split(r"[.!?;\n]+", text)
    return [part.strip(" ,") for part in parts if part.strip(" ,")]


def contains_any(text, terms):
    return any(contains_term(text, term) for term in terms)


def contains_term(text, term):
    return re.search(rf"(?<!\w){re.escape(term)}(?!\w)", text) is not None


def collect_aspect_terms(cleaned_segment, aspect):
    found = {"positive": [], "neutral": [], "negative": []}
    for label in found:
        for term in aspect[f"{label}_terms"]:
            if contains_term(cleaned_segment, term) and term not in found[label]:
                found[label].append(term)
    return remove_embedded_terms(found)


def remove_embedded_terms(found):
    matched_terms = [(label, term) for label, terms in found.items() for term in terms]
    for label, terms in found.items():
        found[label] = [
            term
            for term in terms
            if not any(
                other_label != label
                and term != other_term
                and term in other_term
                and len(other_term) > len(term)
                for other_label, other_term in matched_terms
            )
        ]
    return found


def infer_aspect_sentiment(model_label, found_terms):
    scores = {label: len(found_terms[label]) for label in SENTIMENT_LABELS}
    if not any(scores.values()):
        return model_label if model_label in SENTIMENT_LABELS else "neutral"

    if scores["negative"] > scores["positive"] and scores["negative"] >= scores["neutral"]:
        return "negative"
    if scores["positive"] > scores["negative"] and scores["positive"] >= scores["neutral"]:
        return "positive"
    if scores["neutral"] > scores["positive"] and scores["neutral"] > scores["negative"]:
        return "neutral"
    return model_label if model_label in SENTIMENT_LABELS else "neutral"


def top_terms(counter, limit=4):
    return [term for term, _ in sorted(counter.items(), key=lambda item: (-item[1], item[0]))[:limit]]


def build_aspect_analysis(texts):
    records = []
    for text in texts:
        for segment in split_review_segments(text):
            cleaned_segment = clean_text(segment)
            for aspect in ASPECT_CONFIG:
                if contains_any(cleaned_segment, aspect["keywords"]):
                    records.append(
                        {
                            "aspect": aspect,
                            "segment": segment,
                            "cleaned_segment": cleaned_segment,
                        }
                    )

    stats = {
        aspect["key"]: {
            "config": aspect,
            "total": 0,
            "sentiments": {label: 0 for label in SENTIMENT_LABELS},
            "terms": {label: {} for label in SENTIMENT_LABELS},
            "examples": [],
        }
        for aspect in ASPECT_CONFIG
    }

    if records:
        segment_predictions = predict_texts([record["segment"] for record in records])
        for record, prediction in zip(records, segment_predictions):
            found_terms = collect_aspect_terms(record["cleaned_segment"], record["aspect"])
            label = infer_aspect_sentiment(str(prediction), found_terms)
            aspect_key = record["aspect"]["key"]
            aspect_stats = stats[aspect_key]
            aspect_stats["total"] += 1
            if label in aspect_stats["sentiments"]:
                aspect_stats["sentiments"][label] += 1

            for term_label, terms in found_terms.items():
                for term in terms:
                    aspect_stats["terms"][term_label][term] = aspect_stats["terms"][term_label].get(term, 0) + 1

            if len(aspect_stats["examples"]) < 2:
                aspect_stats["examples"].append(record["segment"][:110])

    analysis = []
    for aspect in ASPECT_CONFIG:
        aspect_stats = stats[aspect["key"]]
        total = aspect_stats["total"]
        if total:
            dominant_label = max(SENTIMENT_LABELS, key=lambda label: aspect_stats["sentiments"][label])
            if dominant_label == "positive":
                status_label = "Nghiêng về tốt"
            elif dominant_label == "negative":
                status_label = "Cần chú ý"
            else:
                status_label = "Khá trung lập"
        else:
            dominant_label = "empty"
            status_label = "Chưa thấy nhắc rõ"

        analysis.append(
            {
                "key": aspect["key"],
                "title": aspect["title"],
                "status": dominant_label,
                "status_label": status_label,
                "total": total,
                "sentiments": aspect_stats["sentiments"],
                "positive_terms": top_terms(aspect_stats["terms"]["positive"]),
                "neutral_terms": top_terms(aspect_stats["terms"]["neutral"]),
                "negative_terms": top_terms(aspect_stats["terms"]["negative"]),
                "examples": aspect_stats["examples"],
            }
        )

    return analysis


def build_prediction_summary(predictions):
    counts = pd.Series(predictions).value_counts().reindex(SENTIMENT_LABELS, fill_value=0)
    total = int(counts.sum())
    if total == 0:
        return None

    percentages = (counts / total * 100).round(2)
    rows = []
    for label in SENTIMENT_LABELS:
        rows.append(
            {
                "label": label,
                "display_label": display_label(label),
                "count": int(counts[label]),
                "percent": float(percentages[label]),
            }
        )

    positive_rate = float(percentages["positive"])
    neutral_rate = float(percentages["neutral"])
    negative_rate = float(percentages["negative"])
    dominant_label = str(counts.idxmax())

    if positive_rate >= 55 and positive_rate > negative_rate:
        conclusion_title = "Đánh giá khá tích cực"
        conclusion_detail = (
            "Phần lớn đánh giá đang nghiêng về hướng tốt. Có thể xem đây là tín hiệu tích cực "
            "về trải nghiệm ăn uống, món ăn hoặc dịch vụ hiện tại."
        )
        conclusion_level = "positive"
    elif negative_rate >= 40 and negative_rate >= positive_rate:
        conclusion_title = "Nhiều đánh giá chưa tốt"
        conclusion_detail = (
            "Tệp này có khá nhiều ý kiến chưa tốt. Nên xem lại các đánh giá thuộc nhóm Không hài lòng "
            "để biết khách hàng chưa hài lòng ở món ăn, giá cả, không gian hay phục vụ."
        )
        conclusion_level = "negative"
    elif neutral_rate >= 45:
        conclusion_title = "Nhiều đánh giá trung lập"
        conclusion_detail = (
            "Nhiều đánh giá chưa thể hiện rõ khen hay chê, hoặc có cả điểm tốt và điểm chưa tốt. "
            "Nhìn chung, trải nghiệm khách hàng đang ở mức trung bình/ổn định."
        )
        conclusion_level = "neutral"
    elif positive_rate > negative_rate:
        conclusion_title = "Nghiêng nhẹ về tích cực"
        conclusion_detail = (
            "Số đánh giá tích cực cao hơn đánh giá chưa tốt. Kết quả nhìn chung khá ổn, nhưng vẫn nên xem thêm các ý kiến còn lại."
        )
        conclusion_level = "positive"
    elif negative_rate > positive_rate:
        conclusion_title = "Nghiêng nhẹ về chưa hài lòng"
        conclusion_detail = (
            "Số đánh giá chưa tốt đang nhiều hơn đánh giá tích cực. Có thể dùng danh sách kết quả bên dưới để xem nhanh các ý kiến cần chú ý."
        )
        conclusion_level = "negative"
    else:
        conclusion_title = "Phân bổ cảm xúc cân bằng"
        conclusion_detail = (
            "Các nhóm ý kiến không chênh lệch nhiều. Nên xem thêm từng phản hồi để có nhận định đầy đủ hơn."
        )
        conclusion_level = "neutral"

    return {
        "total": total,
        "rows": rows,
        "dominant_label": dominant_label,
        "positive_rate": positive_rate,
        "neutral_rate": neutral_rate,
        "negative_rate": negative_rate,
        "conclusion_title": conclusion_title,
        "conclusion_detail": conclusion_detail,
        "conclusion_level": conclusion_level,
    }


@app.route("/")
def index():
    return render_template(
        "index.html",
        active_page="home",
        model_title=MODEL_TITLE,
        model_metrics=MODEL_METRICS,
    )


@app.route("/predict", methods=["GET", "POST"])
def predict():
    single_prediction = None
    single_text = ""
    table_preview = None
    download_filename = None
    prediction_summary = None
    aspect_analysis = None
    error = None

    if request.method == "POST":
        action = request.form.get("action")

        try:
            if action == "predict_text":
                single_text = request.form.get("text", "").strip()
                if not single_text:
                    raise ValueError("Hãy nhập nội dung đánh giá cần phân tích.")
                single_prediction = predict_texts([single_text])[0]

            elif action == "predict_file":
                upload = request.files.get("file")
                if upload is None or upload.filename == "":
                    raise ValueError("Hãy chọn file Excel hoặc CSV.")
                if not allowed_file(upload.filename):
                    raise ValueError("Chỉ hỗ trợ file .xlsx, .xls hoặc .csv.")

                original_name = secure_filename(upload.filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                input_path = UPLOAD_DIR / f"{timestamp}_{original_name}"
                upload.save(input_path)

                df = read_table(input_path)
                text_col = detect_text_column(df.columns)
                if text_col is None:
                    raise ValueError(
                        "Không tìm thấy cột văn bản. Hãy đặt tên cột là một trong: "
                        + ", ".join(TEXT_CANDIDATES)
                    )

                raw_predictions = predict_texts(df[text_col].fillna(""))
                df["sentiment_result"] = [display_label(label) for label in raw_predictions]
                prediction_summary = build_prediction_summary(raw_predictions)
                aspect_analysis = build_aspect_analysis(df[text_col].fillna("").tolist())
                output_name = f"predictions_{timestamp}_{Path(original_name).stem}.xlsx"
                output_path = OUTPUT_DIR / output_name
                df.to_excel(output_path, index=False)

                preview_df = df[[text_col, "sentiment_result"]].head(20).rename(
                    columns={text_col: "Nội dung", "sentiment_result": "Kết quả cảm xúc"}
                )
                table_preview = preview_df.to_dict(orient="records")
                download_filename = output_name

        except Exception as exc:
            error = str(exc)

    return render_template(
        "predict.html",
        single_text=single_text,
        single_prediction=single_prediction,
        single_prediction_display=display_label(single_prediction) if single_prediction else None,
        table_preview=table_preview,
        download_filename=download_filename,
        prediction_summary=prediction_summary,
        aspect_analysis=aspect_analysis,
        error=error,
        text_candidates=TEXT_CANDIDATES,
        active_page="predict",
        model_title=MODEL_TITLE,
    )


@app.route("/download/<path:filename>")
def download(filename):
    safe_name = secure_filename(filename)
    file_path = OUTPUT_DIR / safe_name
    if not file_path.exists():
        return "File không tồn tại", 404
    return send_file(file_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
