# Sum The Time Clock

重建版打卡單辨識工具。此專案由既有 PyInstaller 成品做靜態抽取後，重新整理成可修改的 Python / PySide6 原始碼專案。

## 功能

- 選擇或拖拉打卡單照片進入視窗。
- 支援 `.jpg`、`.jpeg`、`.png`、`.bmp`、`.webp`。
- 依 OCR 結果填入每日上班、下班時間。
- 依班段狀態用顏色標示 `OK`、空白、需確認、失敗、工時異常。
- 從 `config.json` 讀取班段驗證門檻。
- 記錄使用者手動修正到 `logs/correction_history.jsonl`。
- 同一 OCR 原始文字累積至少 2 次相同修正後，下次自動套用。
- 匯出到 `範本/時數表.xlsx` 格式的 Excel。

## 安裝

```bat
py -3.12 -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## 執行

```bat
python main.py
```

## 測試

```bat
python -m compileall .
python test_static_logic.py
python test_gui_logic.py
```

## 打包

```bat
pyinstaller 打卡單辨識工具.spec --clean --noconfirm
```

打包成品會出現在 `dist/打卡單辨識工具/`。若要避免覆蓋其他成品，可使用：

```bat
pyinstaller 打卡單辨識工具.spec --clean --noconfirm --distpath dist_new
```

## 注意

此 repo 不包含原本打包後的 `.exe` 或 `_internal/`。執行時會讀取 exe 同層的 `config.json`；若沒有，才讀取 PyInstaller `_internal` 中的 bundled 設定。
