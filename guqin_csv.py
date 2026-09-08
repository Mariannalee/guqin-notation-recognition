"""使用 guqin_labeler.py 產生的圖片與 labels.csv 訓練相同格式的古琴辨識模型。"""  # 說明本程式負責 CSV 資料的模型訓練。
import argparse  # 匯入命令列參數工具以接受資料夾與模型位置。
import csv  # 匯入 CSV 工具以讀取標註程式保存的答案。
import json  # 匯入 JSON 工具以清楚輸出訓練結果。
from pathlib import Path  # 匯入跨平台路徑工具以支援 Windows 與 macOS。
import joblib  # 匯入模型保存工具以產生 guqin_model.joblib。
import numpy as np  # 匯入數值工具以組合所有圖片特徵。
from sklearn.ensemble import ExtraTreesClassifier  # 匯入小型影像資料適用的隨機樹分類器。
from sklearn.model_selection import cross_val_score  # 匯入交叉驗證工具以估算獨立資料表現。
from guqin import RANDOM_SEED, image_feature  # 沿用原辨識程式完全相同的影像特徵與亂數設定。

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}  # 定義訓練程式接受的圖片格式。


def load_csv_samples(folder: Path) -> tuple[list[Path], list[dict[str, str]], list[str]]:  # 定義讀取標註圖片與 CSV 答案的函式。
    csv_path = folder / "labels.csv"  # 設定標註檔必須位於圖片資料夾內。
    if not csv_path.exists():  # 檢查標註程式是否尚未產生 labels.csv。
        raise FileNotFoundError(f"找不到標籤檔：{csv_path}")  # 顯示缺少標籤檔的明確錯誤訊息。
    image_files: list[Path] = []  # 建立存放有效圖片完整路徑的清單。
    rows: list[dict[str, str]] = []  # 建立存放每張圖片答案的清單。
    with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:  # 用相容 Excel 的 UTF-8 編碼開啟標籤檔。
        reader = csv.DictReader(csv_file)  # 將 CSV 每一列轉成以欄位名稱索引的字典。
        fieldnames = reader.fieldnames or []  # 取得 CSV 內實際存在的欄位名稱。
        if "image" not in fieldnames:  # 檢查 CSV 是否有圖片路徑欄位。
            raise ValueError("labels.csv 必須包含 image 欄位。")  # 缺少圖片欄時停止並提示正確格式。
        headers = [field for field in fieldnames if field != "image"]  # 將 image 以外的欄位全部視為預測目標。
        if not headers:  # 檢查是否完全沒有可訓練的標籤欄位。
            raise ValueError("labels.csv 必須至少包含一個標籤欄位。")  # 提醒使用者補上模型輸出欄位。
        for line_number, row in enumerate(reader, start=2):  # 從 CSV 第二列開始逐筆檢查標註資料。
            image_name = str(row.get("image", "")).strip()  # 取得並清理圖片檔名。
            answers = {header: str(row.get(header, "")).strip() for header in headers}  # 取得並清理所有標籤答案。
            if not image_name or any(not answer for answer in answers.values()):  # 檢查這筆資料是否尚未完成標註。
                continue  # 自動略過未完成的圖片，讓使用者能先用已完成部分訓練。
            image_path = folder / image_name  # 將 CSV 圖片檔名組合成完整路徑。
            if not image_path.is_file() or image_path.suffix.lower() not in IMAGE_SUFFIXES:  # 檢查圖片是否存在且格式受到支援。
                print(f"警告：略過第 {line_number} 列，找不到圖片 {image_name}")  # 顯示被略過資料的列號與圖片名稱。
                continue  # 略過無法讀取的圖片而繼續檢查其他資料。
            image_files.append(image_path)  # 將有效圖片加入訓練路徑清單。
            rows.append(answers)  # 保留 CSV 內的欄名與答案供各欄位分別訓練。
    if not image_files:  # 檢查是否沒有任何完整且有效的標註資料。
        raise ValueError("labels.csv 內沒有可供訓練的完整資料。")  # 提醒使用者先完成至少一筆標註。
    return image_files, rows, headers  # 回傳圖片、答案與模型輸出欄位名稱。


def train_csv(folder: Path, model_path: Path) -> None:  # 定義使用圖片資料夾及 CSV 訓練模型的主要函式。
    image_files, rows, headers = load_csv_samples(folder.resolve())  # 讀取並驗證所有已完成標註的資料。
    features = np.vstack([image_feature(path) for path in image_files])  # 將每張圖片轉成與原模型一致的特徵矩陣。
    models: dict[str, ExtraTreesClassifier] = {}  # 建立保存 String 與 hui 分類器的字典。
    scores: dict[str, float | None] = {}  # 建立保存各欄位交叉驗證準確率的字典。
    for header in headers:  # 分別處理 String 與 hui 兩個預測目標。
        labels = np.array([row[header] for row in rows])  # 取出目前預測目標的全部正確答案。
        model = ExtraTreesClassifier(n_estimators=500, random_state=RANDOM_SEED, n_jobs=-1, max_features="sqrt")  # 建立與 Excel 訓練流程相同的多核心分類器。
        class_counts = np.unique(labels, return_counts=True)[1]  # 計算每個類別各有多少張訓練圖片。
        folds = min(5, int(class_counts.min())) if len(class_counts) > 1 else 0  # 依樣本最少類別決定安全的驗證折數。
        scores[header] = float(cross_val_score(model, features, labels, cv=folds, n_jobs=-1).mean()) if folds >= 2 else None  # 樣本足夠時估算模型對未見圖片的準確率。
        model.fit(features, labels)  # 使用所有已標註資料訓練最後的分類器。
        models[header] = model  # 將訓練完成的分類器保存到對應欄位。
    package = {"models": models, "headers": headers, "image_size": 64, "samples": len(rows), "scores": scores, "source": str(folder)}  # 組合預測所需模型與訓練摘要。
    model_path.parent.mkdir(parents=True, exist_ok=True)  # 建立模型檔所在的上層資料夾。
    joblib.dump(package, model_path)  # 將模型保存成原 guqin.py 可以直接使用的格式。
    print(json.dumps({"status": "CSV 資料訓練完成", "model": str(model_path), "samples": len(rows), "validation_accuracy": scores}, ensure_ascii=False, indent=2))  # 輸出模型位置、資料量與驗證結果。


def build_parser() -> argparse.ArgumentParser:  # 定義建立命令列參數的函式。
    parser = argparse.ArgumentParser(description="使用圖片資料夾內的 labels.csv 訓練古琴辨識模型。")  # 建立具有用途說明的參數解析器。
    parser.add_argument("folder", nargs="?", type=Path, default=Path("dataset/images"), help="含有圖片與 labels.csv 的資料夾。")  # 加入可省略的標註資料夾位置。
    parser.add_argument("--model", type=Path, default=Path("guqin_model.joblib"), help="要輸出的模型檔位置。")  # 加入可自訂的模型輸出位置。
    return parser  # 回傳設定完成的參數解析器。


def main() -> None:  # 定義程式的命令列入口函式。
    arguments = build_parser().parse_args()  # 解析使用者提供的資料夾與模型位置。
    train_csv(arguments.folder, arguments.model)  # 使用 CSV 標註資料執行完整模型訓練。


if __name__ == "__main__":  # 檢查本檔案是否由 Python 直接執行。
    main()  # 啟動 CSV 圖片模型訓練流程。
