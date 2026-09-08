"""從含有儲存格內圖片的 Excel 訓練古琴影像辨識模型，或用模型預測新圖片。"""  # 說明本程式的用途。
import argparse  # 匯入命令列參數解析工具。
import json  # 匯入 JSON 工具以輸出容易閱讀與串接的預測結果。
import posixpath  # 匯入 ZIP 內部路徑的跨平台處理工具。
import re  # 匯入正規表示式以取得儲存格欄名。
import shutil  # 匯入檔案複製工具以匯出訓練圖片。
import tempfile  # 匯入暫存目錄工具以安全處理 Excel 內的圖片。
import xml.etree.ElementTree as ET  # 匯入 XML 解析器以讀取 xlsx 內部資料。
import zipfile  # 匯入 ZIP 工具，因為 xlsx 實際上是 ZIP 封裝格式。
from pathlib import Path  # 匯入路徑物件以安全處理 Windows 與其他平台路徑。
import joblib  # 匯入模型序列化工具以儲存及載入訓練成果。
import numpy as np  # 匯入數值陣列工具以計算影像特徵。
from PIL import Image, ImageOps  # 匯入影像讀取、旋轉修正及灰階處理工具。
from sklearn.ensemble import ExtraTreesClassifier  # 匯入適合小型資料集的隨機樹分類器。
from sklearn.model_selection import cross_val_score  # 匯入交叉驗證工具以估計模型表現。

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"  # 定義 Excel 工作表 XML 命名空間。
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"  # 定義 xlsx 關係檔命名空間。
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"  # 定義 Office 關係屬性的命名空間。
RICH_REL_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"  # 定義儲存格內圖片關係命名空間。
RICH_DATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"  # 定義 Excel 豐富資料命名空間。
IMAGE_SIZE = 64  # 設定每張圖片統一縮放後的寬度與高度。
RANDOM_SEED = 42  # 設定固定亂數種子以讓訓練結果可以重現。


def xml_root(book: zipfile.ZipFile, name: str) -> ET.Element:  # 定義從 xlsx 讀取指定 XML 根節點的函式。
    return ET.fromstring(book.read(name))  # 讀取 ZIP 內的 XML 位元組並解析後回傳。


def shared_strings(book: zipfile.ZipFile) -> list[str]:  # 定義讀取 Excel 共用字串表的函式。
    if "xl/sharedStrings.xml" not in book.namelist():  # 檢查活頁簿是否沒有共用字串表。
        return []  # 沒有共用字串時回傳空清單。
    root = xml_root(book, "xl/sharedStrings.xml")  # 解析共用字串 XML。
    return ["".join(node.text or "" for node in item.iter(f"{{{MAIN_NS}}}t")) for item in root.findall(f"{{{MAIN_NS}}}si")]  # 合併每個字串項目的所有文字片段。


def cell_value(cell: ET.Element, strings: list[str]) -> str | None:  # 定義把 Excel 儲存格節點轉成 Python 值的函式。
    value_node = cell.find(f"{{{MAIN_NS}}}v")  # 尋找儲存格的值節點。
    if value_node is None or value_node.text is None:  # 檢查儲存格是否為空。
        return None  # 空儲存格回傳空值。
    raw_value = value_node.text  # 取得尚未轉換的儲存格文字。
    if cell.get("t") == "s":  # 檢查儲存格是否引用共用字串。
        return strings[int(raw_value)]  # 按索引取回真正的共用字串。
    return raw_value  # 數字或其他類型保留原始文字以免標籤精度改變。


def worksheet_rows(book: zipfile.ZipFile) -> tuple[list[str], list[dict[str, str | None]], list[int | None]]:  # 定義讀取標題、資料列及圖片中繼索引的函式。
    strings = shared_strings(book)  # 讀取活頁簿的共用字串。
    root = xml_root(book, "xl/worksheets/sheet1.xml")  # 解析第一張工作表 XML。
    rows = root.findall(f".//{{{MAIN_NS}}}row")  # 找出工作表內的所有資料列。
    if not rows:  # 檢查工作表是否完全沒有資料。
        raise ValueError("Excel 工作表沒有資料。")  # 用明確訊息中止以避免訓練空模型。
    header_cells = rows[0].findall(f"{{{MAIN_NS}}}c")  # 取得第一列的所有標題儲存格。
    header_map = {re.match(r"[A-Z]+", cell.get("r", ""))[0]: cell_value(cell, strings) for cell in header_cells}  # 建立欄名字母到標題文字的對照。
    image_column = next((column for column, title in header_map.items() if str(title).strip().lower() in {"標題", "image", "圖片"}), "A")  # 找到圖片欄，找不到時使用 A 欄。
    label_columns = [column for column in header_map if column != image_column]  # 將圖片欄以外的所有欄位視為要預測的標籤。
    headers = [str(header_map[column]) for column in label_columns]  # 依工作表順序建立模型輸出欄名。
    data_rows: list[dict[str, str | None]] = []  # 建立存放每列標籤資料的清單。
    metadata_indexes: list[int | None] = []  # 建立存放每列圖片中繼索引的清單。
    for row in rows[1:]:  # 從第二列開始逐列處理訓練資料。
        cells = {re.match(r"[A-Z]+", cell.get("r", ""))[0]: cell for cell in row.findall(f"{{{MAIN_NS}}}c")}  # 建立本列欄名字母到儲存格的對照。
        image_cell = cells.get(image_column)  # 取得本列的圖片儲存格。
        metadata_indexes.append(int(image_cell.get("vm")) - 1 if image_cell is not None and image_cell.get("vm") else None)  # 將從一開始的 vm 索引轉為從零開始的索引。
        data_rows.append({str(header_map[column]): cell_value(cells[column], strings) if column in cells else None for column in label_columns})  # 讀取本列每一個目標欄位的值。
    return headers, data_rows, metadata_indexes  # 回傳標題、標籤列與圖片中繼索引。


def image_paths_by_metadata(book: zipfile.ZipFile) -> dict[int, str]:  # 定義把圖片中繼索引解析成 xlsx 內部圖片路徑的函式。
    metadata_root = xml_root(book, "xl/metadata.xml")  # 解析儲存格豐富資料的中繼資訊。
    rich_indexes = [int(node.get("i", "-1")) for node in metadata_root.findall(f".//{{{RICH_DATA_NS}}}rvb")]  # 依 vm 順序讀出豐富資料索引。
    rich_root = xml_root(book, "xl/richData/rdrichvalue.xml")  # 解析每筆豐富資料的內容。
    rich_values = rich_root.findall(f"{{{RICH_DATA_NS}}}rv")  # 取得所有儲存格內圖片記錄。
    relation_root = xml_root(book, "xl/richData/richValueRel.xml")  # 解析豐富資料到關係 ID 的順序表。
    relation_ids = [node.get(f"{{{REL_NS}}}id") for node in relation_root.findall(f"{{{RICH_REL_NS}}}rel")]  # 依順序取得每張圖片的關係 ID。
    package_root = xml_root(book, "xl/richData/_rels/richValueRel.xml.rels")  # 解析關係 ID 到實際圖片檔的對照表。
    targets = {node.get("Id"): node.get("Target") for node in package_root.findall(f"{{{PKG_NS}}}Relationship")}  # 建立關係 ID 到相對圖片路徑的字典。
    result: dict[int, str] = {}  # 建立圖片中繼索引到 ZIP 內部路徑的結果字典。
    for metadata_index, rich_index in enumerate(rich_indexes):  # 逐一處理每個 Excel 圖片中繼索引。
        values = rich_values[rich_index].findall(f"{{{RICH_DATA_NS}}}v")  # 取得這筆圖片豐富資料內的所有值。
        relation_index = int(values[0].text)  # 第一個值是從零開始的圖片關係順序索引。
        relation_id = relation_ids[relation_index]  # 將關係順序索引轉成實際關係 ID。
        target = targets[relation_id]  # 取得關係 ID 所指向的圖片相對路徑。
        result[metadata_index] = posixpath.normpath(posixpath.join("xl/richData", target))  # 將相對路徑正規化為 ZIP 內完整路徑。
    return result  # 回傳所有圖片中繼索引與路徑的對照。


def load_excel_samples(excel_path: Path, extraction_dir: Path) -> tuple[list[Path], list[dict[str, str]], list[str]]:  # 定義從 Excel 配對並匯出圖片與標籤的函式。
    with zipfile.ZipFile(excel_path) as book:  # 以唯讀方式開啟 xlsx ZIP 封裝。
        headers, rows, metadata_indexes = worksheet_rows(book)  # 讀取標籤資料與每列圖片索引。
        image_map = image_paths_by_metadata(book)  # 建立圖片索引到壓縮檔內路徑的對照。
        image_files: list[Path] = []  # 建立存放匯出圖片路徑的清單。
        valid_rows: list[dict[str, str]] = []  # 建立存放標籤完整資料列的清單。
        extraction_dir.mkdir(parents=True, exist_ok=True)  # 建立圖片匯出目錄並容許目錄已存在。
        for row_number, (row, metadata_index) in enumerate(zip(rows, metadata_indexes), start=2):  # 逐列配對 Excel 標籤和圖片。
            if metadata_index is None or metadata_index not in image_map:  # 檢查本列是否缺少可解析的圖片。
                continue  # 缺少圖片時略過這列以避免錯誤標記。
            if any(row.get(header) in {None, ""} for header in headers):  # 檢查本列是否有任何目標標籤空白。
                continue  # 標籤不完整時略過這列。
            internal_path = image_map[metadata_index]  # 取得圖片在 xlsx ZIP 內的實際路徑。
            suffix = Path(internal_path).suffix.lower() or ".png"  # 保留原圖副檔名，沒有副檔名時採用 PNG。
            output_path = extraction_dir / f"row_{row_number:04d}{suffix}"  # 用 Excel 列號建立可追溯的圖片檔名。
            with book.open(internal_path) as source, output_path.open("wb") as target:  # 同時開啟壓縮檔內來源與輸出檔案。
                shutil.copyfileobj(source, target)  # 將圖片位元組完整複製到輸出檔。
            image_files.append(output_path)  # 把成功匯出的圖片路徑加入訓練清單。
            valid_rows.append({header: str(row[header]) for header in headers})  # 把標籤統一轉成字串後加入訓練清單。
    if not image_files:  # 檢查是否完全沒有成功配對的圖片。
        raise ValueError("找不到可與標籤配對的儲存格內圖片。")  # 提示資料格式問題並停止訓練。
    return image_files, valid_rows, headers  # 回傳圖片路徑、標籤資料及輸出欄名。


def image_feature(image_path: Path) -> np.ndarray:  # 定義把單張圖片轉成固定長度特徵向量的函式。
    with Image.open(image_path) as opened_image:  # 開啟圖片並確保使用完畢後自動關閉檔案。
        image = ImageOps.exif_transpose(opened_image).convert("L")  # 按 EXIF 修正方向並轉成灰階影像。
        image = ImageOps.autocontrast(image)  # 自動拉伸明暗範圍以降低不同拍攝亮度的影響。
        image = ImageOps.fit(image, (IMAGE_SIZE, IMAGE_SIZE), method=Image.Resampling.LANCZOS)  # 等比例裁切並縮放到固定尺寸。
        pixels = np.asarray(image, dtype=np.float32) / 255.0  # 將像素轉成零到一的浮點數陣列。
    gradient_y, gradient_x = np.gradient(pixels)  # 計算每個像素在水平與垂直方向的亮度梯度。
    magnitude = np.hypot(gradient_x, gradient_y)  # 計算每個像素的邊緣強度。
    angle = (np.arctan2(gradient_y, gradient_x) + np.pi) % np.pi  # 將邊緣方向正規化到零至一百八十度。
    bin_index = np.minimum((angle * 9 / np.pi).astype(int), 8)  # 將每個邊緣方向分配到九個方向區間。
    hog_parts: list[np.ndarray] = []  # 建立存放各區塊方向直方圖的清單。
    for top in range(0, IMAGE_SIZE, 8):  # 每八個像素由上到下切分影像。
        for left in range(0, IMAGE_SIZE, 8):  # 每八個像素由左到右切分影像。
            histogram = np.bincount(bin_index[top:top + 8, left:left + 8].ravel(), weights=magnitude[top:top + 8, left:left + 8].ravel(), minlength=9)  # 統計區塊內九個方向的加權邊緣量。
            hog_parts.append(histogram / (np.linalg.norm(histogram) + 1e-7))  # 正規化區塊直方圖並加入特徵清單。
    small_pixels = np.asarray(Image.fromarray((pixels * 255).astype(np.uint8)).resize((16, 16), Image.Resampling.BILINEAR), dtype=np.float32).ravel() / 255.0  # 加入低解析度外觀特徵以保留整體形狀。
    return np.concatenate([small_pixels, *hog_parts]).astype(np.float32)  # 合併外觀與邊緣特徵並回傳固定長度向量。


def train(excel_path: Path, model_path: Path, export_dir: Path | None) -> None:  # 定義由 Excel 訓練並儲存模型的主要函式。
    temporary = tempfile.TemporaryDirectory(prefix="guqin_") if export_dir is None else None  # 未指定匯出目錄時建立自動清理的暫存目錄。
    extraction_dir = Path(temporary.name) if temporary is not None else export_dir  # 決定圖片要匯出到暫存或使用者指定目錄。
    image_files, rows, headers = load_excel_samples(excel_path, extraction_dir)  # 從 Excel 取出已正確配對的圖片與標籤。
    features = np.vstack([image_feature(path) for path in image_files])  # 將所有圖片轉成機器學習可用的特徵矩陣。
    models: dict[str, ExtraTreesClassifier] = {}  # 建立存放每個輸出欄位分類器的字典。
    scores: dict[str, float | None] = {}  # 建立存放各欄位交叉驗證準確率的字典。
    for header in headers:  # 逐一針對 String、hui 等每個欄位建立獨立模型。
        labels = np.array([row[header] for row in rows])  # 取出目前欄位的所有正確答案。
        model = ExtraTreesClassifier(n_estimators=500, random_state=RANDOM_SEED, n_jobs=-1, max_features="sqrt")  # 建立具多核心支援的隨機樹分類器並避開新版套件的類別權重相容問題。
        class_counts = np.unique(labels, return_counts=True)[1]  # 統計目前欄位各類別的樣本數。
        folds = min(5, int(class_counts.min())) if len(class_counts) > 1 else 0  # 依最少類別樣本數決定可安全執行的分層驗證折數。
        scores[header] = float(cross_val_score(model, features, labels, cv=folds, n_jobs=-1).mean()) if folds >= 2 else None  # 樣本足夠時估算交叉驗證平均準確率。
        model.fit(features, labels)  # 使用全部資料訓練最後要保存的分類器。
        models[header] = model  # 將訓練完成的分類器按欄位名稱保存。
    package = {"models": models, "headers": headers, "image_size": IMAGE_SIZE, "samples": len(rows), "scores": scores}  # 將模型與必要中繼資料組成單一模型套件。
    model_path.parent.mkdir(parents=True, exist_ok=True)  # 建立模型輸出檔的上層目錄。
    joblib.dump(package, model_path)  # 將完整模型套件寫入磁碟。
    if temporary is not None:  # 檢查本次是否使用自動建立的暫存目錄。
        temporary.cleanup()  # 清除暫存的 Excel 圖片以節省空間。
    print(json.dumps({"status": "訓練完成", "model": str(model_path), "samples": len(rows), "validation_accuracy": scores}, ensure_ascii=False, indent=2))  # 輸出訓練摘要與各欄位驗證準確率。


def predict(model_path: Path, image_path: Path) -> None:  # 定義載入模型並預測一張新圖片的函式。
    package = joblib.load(model_path)  # 從磁碟載入先前訓練完成的模型套件。
    feature = image_feature(image_path).reshape(1, -1)  # 將新圖片轉成單筆二維特徵矩陣。
    result: dict[str, dict[str, str | float]] = {}  # 建立存放每個欄位預測答案與信心值的字典。
    for header in package["headers"]:  # 依序預測模型內的 String、hui 等所有欄位。
        model = package["models"][header]  # 取得目前欄位所對應的分類器。
        probabilities = model.predict_proba(feature)[0]  # 計算新圖片屬於各類別的機率。
        best_index = int(np.argmax(probabilities))  # 找出機率最高的類別索引。
        result[header] = {"value": str(model.classes_[best_index]), "confidence": round(float(probabilities[best_index]), 4)}  # 保存預測值與四位小數信心值。
    print(json.dumps({"image": str(image_path), "prediction": result}, ensure_ascii=False, indent=2))  # 以 JSON 格式輸出可讀且可供其他程式使用的結果。


def build_parser() -> argparse.ArgumentParser:  # 定義建立命令列介面的函式。
    parser = argparse.ArgumentParser(description="用 Excel 內的古琴圖片訓練模型，並辨識新圖片的 String 與 hui。")  # 建立主命令列解析器及說明。
    subparsers = parser.add_subparsers(dest="command", required=True)  # 建立必填的訓練或預測子命令。
    train_parser = subparsers.add_parser("train", help="從 Excel 圖片及旁邊欄位訓練模型。")  # 建立訓練子命令。
    train_parser.add_argument("--excel", type=Path, default=Path("古琴資料.xlsx"), help="訓練資料 Excel 路徑。")  # 加入可省略且預設為現有檔案的 Excel 參數。
    train_parser.add_argument("--model", type=Path, default=Path("guqin_model.joblib"), help="模型輸出路徑。")  # 加入模型輸出檔參數。
    train_parser.add_argument("--export-images", type=Path, default=None, help="可選：保留從 Excel 匯出的圖片目錄。")  # 加入選擇性保留訓練圖片的參數。
    predict_parser = subparsers.add_parser("predict", help="使用已訓練模型辨識一張新圖片。")  # 建立預測子命令。
    predict_parser.add_argument("image", type=Path, help="要辨識的新圖片路徑。")  # 加入必填的新圖片位置參數。
    predict_parser.add_argument("--model", type=Path, default=Path("guqin_model.joblib"), help="已訓練模型路徑。")  # 加入可省略的模型位置參數。
    return parser  # 回傳設定完成的命令列解析器。


def main() -> None:  # 定義程式命令列入口函式。
    arguments = build_parser().parse_args()  # 解析使用者輸入的所有命令列參數。
    if arguments.command == "train":  # 檢查使用者是否選擇訓練模式。
        train(arguments.excel, arguments.model, arguments.export_images)  # 執行 Excel 圖片取出、特徵計算與模型訓練。
    else:  # 使用者選擇的不是訓練模式時即為預測模式。
        predict(arguments.model, arguments.image)  # 載入模型並辨識指定的新圖片。


if __name__ == "__main__":  # 檢查本檔案是否由 Python 直接執行。
    main()  # 啟動命令列主流程。
