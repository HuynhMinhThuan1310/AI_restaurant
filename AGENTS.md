# AGENTS.md

## Workspace Shape
- Small report/notebook/Flask-demo workspace; there is no root package manifest, lockfile, CI, formatter, linter, or test runner.
- Main training notebook is `train_restaurant_ml_kaggle_final_overall.ipynb`; final artifacts live in `results/`.
- Flask app entrypoint is `sentiment_app/app.py`; templates are `sentiment_app/templates/`, styling is `sentiment_app/static/style.css`.
- Word report is `bao_cao_tfidf_linear_svm_3class_fixed_image.docx`; `_rewrite_report_fixed_image.py` can regenerate it but may overwrite manual Word edits.

## Current Model And Data
- Current task is 3-class Vietnamese restaurant-review sentiment classification, not student feedback, phone reviews, Book1, or ViSFD.
- Dataset/artifact name: `vietnamese-restaurant-review-sentiment-dataset`.
- Raw labels map as `0 -> negative`, `1 -> positive`, `2 -> neutral`; app displays them as `Không hài lòng`, `Hài lòng`, `Trung lập`.
- Final model path used by the app: `results/tfidf_restaurant_review_3class_final_overall_model.joblib`.
- Final approach: classical ML, TF-IDF word/character features + LinearSVC; avoid changing to deep learning unless explicitly requested.
- Final test metrics in `results/final_overall_model_info.json`: Accuracy `0.8406`, Macro F1 `0.8038`, Weighted F1 `0.8417`.

## Flask App Gotchas
- Run with the app venv, not global Python or VS Code Run Code: `C:\AI\sentiment_app\.venv\Scripts\python.exe C:\AI\sentiment_app\app.py`.
- `sentiment_app/requirements.txt` pins `scikit-learn==1.6.1`; keep this because the joblib model was saved for that version.
- Browser URL is `http://127.0.0.1:5000`; `127.0.0.1:5500` is Live Server and will not run Flask routes.
- Uploads support `.xlsx`, `.xls`, `.csv`; text column detection checks `sentence`, `text`, `feedback`, `comment`, `review`, `content`.
- File predictions write the Vietnamese display label to `sentiment_result`; downloaded files are written under `sentiment_app/outputs/`.
- `restore_sklearn_classes()` in `app.py` is compatibility glue for joblib/sklearn class metadata; keep it unless a regenerated model is verified without it.

## Report Workflow
- Do not regenerate the Word report unless the user asks; the user may manually edit the cover and section text in Word.
- If regeneration is explicitly requested, run from `C:\AI`: `python -X utf8 "_rewrite_report_fixed_image.py"`.
- If Word has the report open, the script may save to `bao_cao_tfidf_linear_svm_3class_fixed_image_updated.docx`; close Word to overwrite the main file.
- Report language should avoid code-style parameter names for non-technical readers; explain TF-IDF, LinearSVC, and metrics in Vietnamese.

## Focused Verification
- App syntax from `C:\AI`: `C:\AI\sentiment_app\.venv\Scripts\python.exe -m py_compile "sentiment_app\app.py"`.
- Route smoke test from `C:\AI\sentiment_app`: `C:\AI\sentiment_app\.venv\Scripts\python.exe -c "import app; c=app.app.test_client(); assert c.get('/').status_code==200; assert c.get('/predict').status_code==200"`.
- Model smoke test from `C:\AI\sentiment_app`: `C:\AI\sentiment_app\.venv\Scripts\python.exe -c "import app; print(app.display_label(app.predict_texts(['Quán ăn rất ngon, nhân viên nhiệt tình.'])[0]))"`.
- Report smoke check after regeneration: ensure it no longer mentions student feedback, Book1, ViSFD, phone reviews, or Teachable Machine.

## UI Context
- Current product copy is about evaluating restaurant quality from customer reviews; avoid reverting to generic “tổng hợp cảm xúc” wording.
- Top navigation is `Trang chủ`, `Tính năng`, `Cách dùng`, `Đánh giá`; `/predict` is the evaluation page.
- Homepage hero currently uses an external Unsplash restaurant image with overlay cards; if offline reliability matters, replace with a local asset instead of removing the visual.
