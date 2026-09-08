"""標註 PDF 切出的古琴單字圖片，支援 Windows 與 macOS。"""  # 說明這個程式的用途。
import argparse  # 匯入命令列參數工具以接受圖片資料夾位置。
import csv  # 匯入 CSV 工具以讀寫每張圖片的標註內容。
import json  # 匯入 JSON 工具以保存關閉程式後仍可恢復的圖片進度。
import tkinter as tk  # 匯入 Python 內建的跨平台視窗工具。
from pathlib import Path  # 匯入跨平台路徑工具以支援 Windows 與 macOS。
from tkinter import filedialog, messagebox, ttk  # 匯入資料夾選擇、提示訊息及現代化介面元件。
from PIL import Image, ImageOps, ImageTk  # 匯入 Pillow 以顯示 PNG、JPG 等圖片。

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}  # 定義可載入的圖片副檔名。
LABEL_FIELDS = ["String1", "String2", "Hui1", "Hui2", "Other"]  # 定義使用者指定的五個標註欄位。
STRING_OPTIONS = ["", "1", "2", "3", "4", "5", "6", "7", "1&3", "Nondefine"]  # 定義兩個 String 欄位的常用選項。
HUI_OPTIONS = ["", "0", "7", "8", "8.5", "8.9", "9", "10", "11", "13", "4down", "5up", "6up", "8up", "9up", "Nondefine"]  # 定義兩個 Hui 欄位的常用選項。
OTHER_OPTIONS = ["", "Nondefine"]  # 定義 Other 欄位的初始選項並保留自由輸入能力。


class GuqinLabeler:  # 定義管理整個圖片標註視窗的類別。
    def __init__(self, root: tk.Tk, folder: Path) -> None:  # 定義標註視窗的初始化流程。
        self.root = root  # 保存主視窗物件供其他方法使用。
        self.folder = folder.resolve()  # 將圖片根資料夾轉成完整路徑。
        self.csv_path = self.folder / "labels.csv"  # 將標籤檔固定放在圖片根資料夾內。
        self.state_path = self.folder / ".guqin_labeler_state.json"  # 設定保存目前圖片位置的隱藏進度檔路徑。
        self.images: list[Path] = []  # 建立存放所有待標註圖片的清單。
        self.labels: dict[str, dict[str, str]] = {}  # 建立圖片相對路徑到五個答案的對照表。
        self.index = 0  # 將目前顯示位置設定為第一張圖片。
        self.rotation = 0  # 將目前圖片的預覽旋轉角度設定為零。
        self.photo: ImageTk.PhotoImage | None = None  # 保留畫面圖片參照以避免被自動回收。
        self.values = {field: tk.StringVar() for field in LABEL_FIELDS}  # 為五個輸入欄位建立各自的同步文字變數。
        self.input_boxes: list[ttk.Combobox] = []  # 建立依填寫順序保存五個輸入框的清單。
        self.status_value = tk.StringVar()  # 建立與畫面狀態列同步的文字變數。
        self.root.title("古琴單字圖片標註工具")  # 設定主視窗標題。
        self.root.geometry("980x800")  # 設定適合五個欄位的初始視窗大小。
        self.root.minsize(700, 620)  # 限制最小視窗尺寸以免控制項互相遮擋。
        self.build_interface()  # 建立圖片、五個輸入欄位與操作按鈕。
        self.load_folder(self.folder)  # 遞迴載入啟動時指定資料夾內的切割圖片。
        self.root.bind("<Left>", lambda event: self.previous_image())  # 綁定左方向鍵以切換上一張圖片。
        self.root.bind("<Right>", lambda event: self.save_and_next())  # 綁定右方向鍵以儲存並切換下一張圖片。
        self.root.bind("<Control-s>", lambda event: self.save_current())  # 綁定 Windows 的 Ctrl+S 儲存快捷鍵。
        self.root.bind("<Command-s>", lambda event: self.save_current())  # 綁定 macOS 的 Command+S 儲存快捷鍵。
        self.root.bind("<Escape>", lambda event: self.skip_image())  # 綁定 Windows 與 macOS 都可使用的 Esc 跳過快捷鍵。
        self.root.bind("<Return>", self.focus_next_field)  # 綁定 Enter 以依序前往下一個待填欄位或完成目前圖片。

    def build_interface(self) -> None:  # 定義建立標註畫面的函式。
        toolbar = ttk.Frame(self.root, padding=8)  # 建立最上方的資料夾工具列。
        toolbar.pack(fill="x")  # 讓工具列填滿視窗寬度。
        ttk.Button(toolbar, text="選擇裁切圖片根資料夾", command=self.choose_folder).pack(side="left")  # 建立選擇 pdf_crops 資料夾的按鈕。
        self.folder_label = ttk.Label(toolbar, text=str(self.folder))  # 建立顯示目前圖片根資料夾的標籤。
        self.folder_label.pack(side="left", padx=12, fill="x", expand=True)  # 顯示路徑並使用剩餘寬度。
        self.image_label = ttk.Label(self.root, anchor="center")  # 建立顯示目前單字圖片的主要區域。
        self.image_label.pack(fill="both", expand=True, padx=12, pady=6)  # 讓圖片區域隨視窗大小自動伸縮。
        editor = ttk.LabelFrame(self.root, text="標註內容", padding=10)  # 建立包含五個答案欄位的群組區域。
        editor.pack(fill="x", padx=10, pady=5)  # 讓答案輸入區填滿視窗寬度。
        for column, field in enumerate(LABEL_FIELDS):  # 依序建立 String1、String2、Hui1、Hui2 與 Other 欄位。
            ttk.Label(editor, text=f"{field}：").grid(row=0, column=column, padx=5, pady=(0, 4), sticky="w")  # 在輸入框上方顯示欄位名稱。
            options = STRING_OPTIONS if field.startswith("String") else HUI_OPTIONS if field.startswith("Hui") else OTHER_OPTIONS  # 依欄位類型選擇適合的下拉選項。
            box = ttk.Combobox(editor, textvariable=self.values[field], values=options, state="normal", width=15)  # 建立可選擇也可自由輸入答案的下拉欄位。
            box.grid(row=1, column=column, padx=5, pady=(0, 4), sticky="ew")  # 將輸入框放在對應欄名下方。
            self.input_boxes.append(box)  # 按畫面順序保存輸入框以支援 Enter 快速移動。
            editor.columnconfigure(column, weight=1)  # 允許每個輸入欄隨視窗寬度平均伸縮。
        ttk.Label(editor, text="String1、String2、Hui1、Hui2、Other 全部必填。", foreground="#555555").grid(row=2, column=0, columnspan=5, pady=(5, 0), sticky="w")  # 顯示五個欄位全部必須填寫的操作提示。
        controls = ttk.Frame(self.root, padding=8)  # 建立圖片切換與操作按鈕區域。
        controls.pack(fill="x")  # 讓操作區填滿視窗寬度。
        ttk.Button(controls, text="← 上一張", command=self.previous_image).pack(side="left", padx=4)  # 建立切換上一張圖片的按鈕。
        ttk.Button(controls, text="向左轉", command=lambda: self.rotate_preview(90)).pack(side="left", padx=4)  # 建立只將預覽向左旋轉的按鈕。
        ttk.Button(controls, text="向右轉", command=lambda: self.rotate_preview(-90)).pack(side="left", padx=4)  # 建立只將預覽向右旋轉的按鈕。
        ttk.Button(controls, text="跳過此張", command=self.skip_image).pack(side="left", padx=4)  # 建立不儲存答案並直接切換下一張的按鈕。
        ttk.Button(controls, text="跳到未標註", command=self.jump_to_unlabeled).pack(side="left", padx=4)  # 建立快速尋找下一張未標註圖片的按鈕。
        ttk.Button(controls, text="只儲存", command=self.save_current).pack(side="right", padx=4)  # 建立只儲存目前答案而不換圖的按鈕。
        ttk.Button(controls, text="儲存並下一張 →", command=self.save_and_next).pack(side="right", padx=4)  # 建立儲存答案並切換下一張的主要按鈕。
        ttk.Label(self.root, textvariable=self.status_value, anchor="center", padding=8).pack(fill="x")  # 建立顯示相對路徑與完成進度的狀態列。

    def choose_folder(self) -> None:  # 定義讓使用者選擇裁切圖片根資料夾的函式。
        selected = filedialog.askdirectory(initialdir=self.folder, title="選擇 pdf_crops 輸出資料夾")  # 開啟 Windows 與 macOS 原生資料夾選擇視窗。
        if selected:  # 檢查使用者是否選擇資料夾而不是取消。
            self.load_folder(Path(selected))  # 遞迴載入新資料夾內的 page_XXXX 單字圖片。

    def relative_name(self, image: Path) -> str:  # 定義取得圖片相對於資料根目錄之可攜路徑的函式。
        return image.relative_to(self.folder).as_posix()  # 使用跨平台統一的斜線格式回傳相對路徑。

    def is_character_image(self, path: Path) -> bool:  # 定義判斷檔案是否為待標註單字圖的函式。
        lowered = path.name.lower()  # 將檔名轉成小寫以進行不分大小寫比較。
        excluded = lowered.startswith("preview_") or lowered.startswith("binary_")  # 判斷是否為裁切程式產生的整頁檢查圖。
        return path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES and not excluded  # 只接受支援格式且不是預覽或二值圖的檔案。

    def load_folder(self, folder: Path) -> None:  # 定義遞迴掃描裁切圖片及讀取既有標籤的函式。
        self.folder = folder.resolve()  # 保存目前裁切圖片根資料夾的完整路徑。
        self.csv_path = self.folder / "labels.csv"  # 將標籤檔位置更新到目前根資料夾。
        self.state_path = self.folder / ".guqin_labeler_state.json"  # 將進度檔位置更新到目前根資料夾。
        self.folder.mkdir(parents=True, exist_ok=True)  # 在資料夾不存在時建立它。
        self.images = sorted([path for path in self.folder.rglob("*") if self.is_character_image(path)], key=lambda path: self.relative_name(path).lower())  # 遞迴找出 page_XXXX 子資料夾內所有單字圖片並依路徑排序。
        self.labels = self.read_labels()  # 從 labels.csv 讀取先前已完成的標註。
        self.index = self.read_progress_index()  # 優先回到上次關閉前停留的圖片位置。
        self.rotation = 0  # 切換資料夾時重設預覽旋轉角度。
        self.folder_label.config(text=str(self.folder))  # 更新畫面上顯示的資料夾路徑。
        self.show_current()  # 顯示目前圖片及其既有答案。

    def read_labels(self) -> dict[str, dict[str, str]]:  # 定義從 CSV 讀取既有五欄標註的函式。
        result: dict[str, dict[str, str]] = {}  # 建立暫存全部標籤的結果字典。
        if not self.csv_path.exists():  # 檢查根資料夾內是否尚未建立 labels.csv。
            return result  # 第一次使用時直接回傳空字典。
        with self.csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:  # 用 Excel 相容的 UTF-8 編碼開啟 CSV。
            for row in csv.DictReader(csv_file):  # 逐列讀取圖片路徑與五個答案。
                image_name = str(row.get("image", "")).strip().replace("\\", "/")  # 取得並統一 Windows 或 macOS 的路徑分隔符號。
                if image_name:  # 檢查本列是否包含有效圖片路徑。
                    result[image_name] = {field: str(row.get(field, "")).strip() for field in LABEL_FIELDS}  # 保存本列五個欄位的答案。
        return result  # 回傳從 CSV 讀到的全部標籤。

    def read_progress_index(self) -> int:  # 定義從進度檔恢復上次停留圖片位置的函式。
        if not self.images:  # 檢查目前資料夾是否完全沒有圖片。
            return 0  # 沒有圖片時回傳安全的零位置。
        if self.state_path.exists():  # 檢查這個圖片資料夾是否已有先前保存的進度。
            try:  # 開始捕捉進度檔可能損壞或格式不正確的情況。
                state = json.loads(self.state_path.read_text(encoding="utf-8"))  # 讀取並解析上次保存的 JSON 進度資料。
                saved_name = str(state.get("current_image", "")).replace("\\", "/")  # 取得並統一路徑分隔符號的上次圖片名稱。
                names = [self.relative_name(image) for image in self.images]  # 建立目前所有圖片相對路徑的順序清單。
                if saved_name in names:  # 檢查上次圖片是否仍存在於目前資料夾。
                    return names.index(saved_name)  # 回傳上次停留圖片在目前清單中的位置。
            except (OSError, ValueError, TypeError):  # 捕捉讀檔失敗、JSON 損壞或資料類型錯誤。
                pass  # 忽略無法使用的舊進度並改用未標註圖片位置。
        return self.first_unlabeled_index()  # 沒有有效進度時從第一張未完成圖片開始。

    def write_progress(self) -> None:  # 定義立即保存目前停留圖片位置的函式。
        if not self.images:  # 檢查目前是否沒有可記錄位置的圖片。
            return  # 沒有圖片時不建立無意義的進度檔。
        state = {"current_image": self.relative_name(self.images[self.index]), "position": self.index + 1, "total": len(self.images)}  # 組合目前圖片相對路徑與方便查看的進度數字。
        temporary_path = self.state_path.with_suffix(".json.tmp")  # 建立暫存進度檔以避免中途關閉造成檔案損壞。
        temporary_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")  # 將目前進度以容易閱讀的 JSON 格式寫入暫存檔。
        temporary_path.replace(self.state_path)  # 完整寫入後安全替換正式進度檔。

    def write_labels(self) -> None:  # 定義將目前所有標註安全寫入 CSV 的函式。
        temporary_path = self.csv_path.with_suffix(".csv.tmp")  # 建立暫存檔以避免寫入途中損壞正式 CSV。
        with temporary_path.open("w", encoding="utf-8-sig", newline="") as csv_file:  # 用 Excel 可直接開啟的 UTF-8 編碼建立暫存 CSV。
            writer = csv.DictWriter(csv_file, fieldnames=["image", *LABEL_FIELDS])  # 建立包含圖片路徑及五個答案的 CSV 寫入器。
            writer.writeheader()  # 寫入 image、String1、String2、Hui1、Hui2、Other 標題列。
            for image in self.images:  # 依圖片顯示順序逐張輸出標籤。
                name = self.relative_name(image)  # 取得目前圖片可跨平台搬移的相對路徑。
                label = self.labels.get(name, {field: "" for field in LABEL_FIELDS})  # 取得既有答案或建立全部空白答案。
                writer.writerow({"image": name, **{field: label.get(field, "") for field in LABEL_FIELDS}})  # 將相對路徑與五個答案寫成一列。
        temporary_path.replace(self.csv_path)  # 完成後以暫存檔安全替換正式 CSV。

    def first_unlabeled_index(self, start: int = 0) -> int:  # 定義尋找下一張尚未填完五個答案之圖片位置的函式。
        if not self.images:  # 檢查目前是否完全沒有圖片。
            return 0  # 沒有圖片時回傳安全的零位置。
        for offset in range(len(self.images)):  # 最多循環檢查全部圖片一次。
            index = (start + offset) % len(self.images)  # 計算包含結尾循環的候選位置。
            name = self.relative_name(self.images[index])  # 取得候選圖片的相對路徑。
            if any(not self.labels.get(name, {}).get(field) for field in LABEL_FIELDS):  # 檢查五個必填答案是否有任何一個仍為空白。
                return index  # 回傳第一張未標註圖片的位置。
        return start % len(self.images)  # 全部完成時保留目前或指定的起始位置。

    def show_current(self) -> None:  # 定義顯示目前圖片與既有五個答案的函式。
        if not self.images:  # 檢查目前資料夾是否沒有任何單字圖片。
            self.image_label.config(image="", text="找不到單字圖片\n請選擇 pdf_cropper.py 的輸出資料夾")  # 顯示正確資料夾選擇提示。
            for value in self.values.values():  # 逐一處理五個答案變數。
                value.set("")  # 清除沒有圖片時的所有答案欄位。
            self.status_value.set("共 0 張單字圖片")  # 更新狀態列顯示目前沒有圖片。
            return  # 結束本次畫面更新。
        self.index %= len(self.images)  # 保證目前位置永遠落在有效圖片範圍內。
        image_path = self.images[self.index]  # 取得目前要顯示的單字圖片路徑。
        with Image.open(image_path) as source:  # 開啟目前圖片並在處理後自動關閉檔案。
            preview = ImageOps.exif_transpose(source).convert("RGB")  # 修正圖片方向並轉成一致的 RGB 格式。
            preview = preview.rotate(self.rotation, expand=True)  # 依使用者設定旋轉預覽而不修改原檔。
            preview.thumbnail((700, 500), Image.Resampling.NEAREST)  # 放大或縮小顯示區使用的單字預覽。
            if preview.width < 350 and preview.height < 350:  # 檢查單字圖是否因固定畫布而顯示得太小。
                factor = min(350 / preview.width, 350 / preview.height)  # 計算清楚顯示單字的整體放大倍率。
                preview = preview.resize((round(preview.width * factor), round(preview.height * factor)), Image.Resampling.NEAREST)  # 使用不模糊筆畫的最近鄰方法放大預覽。
            self.photo = ImageTk.PhotoImage(preview.copy())  # 將 Pillow 圖片轉成 Tkinter 可顯示的物件。
        self.image_label.config(image=self.photo, text="")  # 在主要區域顯示目前單字圖片。
        name = self.relative_name(image_path)  # 取得目前圖片的跨平台相對路徑。
        label = self.labels.get(name, {field: "" for field in LABEL_FIELDS})  # 取得已保存答案或建立空白答案。
        for field in LABEL_FIELDS:  # 逐一更新五個畫面輸入欄位。
            self.values[field].set(label.get(field, ""))  # 將目前圖片的既有答案放入對應輸入框。
        completed = sum(1 for image in self.images if all(self.labels.get(self.relative_name(image), {}).get(field) for field in LABEL_FIELDS))  # 計算五個必填答案都已填寫的圖片數量。
        self.status_value.set(f"{self.index + 1} / {len(self.images)}　{name}　已完成 {completed} 張")  # 顯示目前位置、相對路徑與完成進度。
        self.write_progress()  # 每次顯示圖片就立即保存位置以支援關閉後接續標註。

    def rotate_preview(self, degrees: int) -> None:  # 定義只旋轉預覽而不改動原圖的函式。
        self.rotation = (self.rotation + degrees) % 360  # 更新並正規化預覽旋轉角度。
        self.show_current()  # 重新顯示套用新角度的圖片。

    def focus_next_field(self, event: tk.Event | None = None) -> str:  # 定義按 Enter 後移動到下一個輸入欄或下一張圖片的函式。
        focused = self.root.focus_get()  # 取得目前正在接受鍵盤輸入的畫面元件。
        if focused in self.input_boxes:  # 檢查目前游標是否位於五個標註輸入框之一。
            current_index = self.input_boxes.index(focused)  # 取得目前輸入框在五欄填寫順序中的位置。
            if current_index < len(self.input_boxes) - 1:  # 檢查目前欄位後方是否還有下一個待填欄位。
                self.input_boxes[current_index + 1].focus_set()  # 將鍵盤游標移到下一個輸入欄位。
                self.input_boxes[current_index + 1].selection_range(0, tk.END)  # 反白下一欄既有內容以方便直接覆寫。
            else:  # 目前游標位於最後一個 Other 欄位。
                old_index = self.index  # 保存目前圖片位置以判斷是否成功儲存並切換。
                self.save_and_next()  # 驗證五欄、儲存目前答案並嘗試切換下一張圖片。
                if self.index != old_index or len(self.images) == 1:  # 檢查儲存流程是否成功完成。
                    self.input_boxes[0].focus_set()  # 將鍵盤游標移回下一張圖片的 String1 欄位。
                    self.input_boxes[0].selection_range(0, tk.END)  # 反白 String1 既有內容以方便快速輸入。
        elif self.input_boxes:  # 目前游標不在輸入框但畫面已建立五個欄位。
            self.input_boxes[0].focus_set()  # 第一次按 Enter 時將游標移到 String1 欄位。
            self.input_boxes[0].selection_range(0, tk.END)  # 反白 String1 內容以準備輸入。
        return "break"  # 阻止 Enter 同時觸發 Combobox 或其他元件的預設動作。

    def save_current(self, show_message: bool = True) -> bool:  # 定義驗證並儲存目前圖片答案的函式。
        if not self.images:  # 檢查目前是否沒有可標註圖片。
            return False  # 沒有圖片時回報未儲存。
        answers = {field: self.values[field].get().strip() for field in LABEL_FIELDS}  # 取得並清理五個輸入框的答案。
        missing_fields = [field for field in LABEL_FIELDS if not answers[field]]  # 找出五個必填答案中仍然空白的欄位。
        if missing_fields:  # 檢查是否至少有一個必填欄位尚未填寫。
            messagebox.showwarning("答案未填完整", f"以下欄位全部必填：{', '.join(missing_fields)}")  # 列出所有缺少答案的欄位供使用者補填。
            return False  # 回報本次沒有成功儲存。
        name = self.relative_name(self.images[self.index])  # 取得目前圖片的相對路徑。
        self.labels[name] = answers  # 更新目前圖片在記憶體中的五個答案。
        self.write_labels()  # 立即把全部標籤寫入 CSV 以避免意外遺失。
        self.show_current()  # 更新畫面上的已完成數量。
        if show_message:  # 檢查這次操作是否需要顯示成功狀態。
            summary = "、".join(f"{field}={answer}" for field, answer in answers.items() if answer)  # 將所有非空答案組合成簡短摘要。
            self.status_value.set(f"已儲存：{name}　{summary}")  # 在狀態列顯示剛保存的圖片與答案。
        return True  # 回報本次答案已成功儲存。

    def save_and_next(self) -> None:  # 定義儲存目前答案並切換下一張的函式。
        if self.save_current(show_message=False):  # 先驗證並保存目前答案。
            self.index = (self.index + 1) % len(self.images)  # 將位置移到下一張並在結尾循環。
            self.rotation = 0  # 切換圖片時重設預覽旋轉角度。
            self.show_current()  # 顯示下一張圖片及既有答案。

    def previous_image(self) -> None:  # 定義切換上一張圖片的函式。
        if self.images:  # 檢查目前確實有圖片可以切換。
            self.index = (self.index - 1) % len(self.images)  # 將位置移到上一張並在開頭循環。
            self.rotation = 0  # 切換圖片時重設預覽旋轉角度。
            self.show_current()  # 顯示上一張圖片及既有答案。

    def skip_image(self) -> None:  # 定義不儲存目前答案並直接跳過這張圖片的函式。
        if self.images:  # 檢查目前確實有圖片可以跳過。
            self.index = (self.index + 1) % len(self.images)  # 將位置移到下一張並在結尾循環。
            self.rotation = 0  # 跳過圖片時重設預覽旋轉角度。
            self.show_current()  # 顯示下一張圖片且保留被跳過圖片的未標註狀態。

    def jump_to_unlabeled(self) -> None:  # 定義跳到下一張未標註圖片的函式。
        if self.images:  # 檢查目前有圖片可以搜尋。
            self.index = self.first_unlabeled_index(self.index + 1)  # 從下一張開始尋找尚未填完五個答案的圖片。
            self.rotation = 0  # 跳轉圖片時重設預覽旋轉角度。
            self.show_current()  # 顯示找到的未標註圖片。


def build_parser() -> argparse.ArgumentParser:  # 定義建立命令列參數的函式。
    parser = argparse.ArgumentParser(description="標註 PDF 裁切出的古琴單字圖片。")  # 建立具有用途說明的命令列解析器。
    parser.add_argument("folder", nargs="?", type=Path, default=Path("pdf_crops_samples"), help="pdf_cropper.py 的輸出資料夾。")  # 加入可省略並預設為目前實測輸出的資料夾參數。
    return parser  # 回傳設定完成的參數解析器。


def main() -> None:  # 定義程式的主要啟動函式。
    arguments = build_parser().parse_args()  # 解析使用者提供的裁切圖片根資料夾。
    root = tk.Tk()  # 建立 Windows 與 macOS 都能使用的 Tkinter 主視窗。
    GuqinLabeler(root, arguments.folder)  # 建立並初始化古琴單字標註工具。
    root.mainloop()  # 啟動視窗事件循環直到使用者關閉程式。


if __name__ == "__main__":  # 檢查本檔案是否由 Python 直接執行。
    main()  # 啟動古琴單字圖片標註視窗。
