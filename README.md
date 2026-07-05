# Sum The Time Clock

Sum The Time Clock is a Python / PySide6 desktop tool for reading attendance sheet photos, reviewing recognized clock-in and clock-out times, and exporting the result to an Excel timesheet.

## Highlights

- Drag and drop attendance photos into the app.
- Supports `.jpg`, `.jpeg`, `.png`, `.bmp`, and `.webp` images.
- Handles one photo or a first-half / second-half pair.
- Shows loaded file names before recognition.
- Shows an OCR progress bar while images are being processed.
- Uses a sidebar workflow layout inspired by the generated UI design.
- Highlights blank, review-needed, failed, short-shift, and abnormal-shift rows.
- Reads validation rules from `config.json`.
- Records manual table edits in `logs/correction_history.jsonl`.
- Applies repeated correction-history matches automatically.
- Exports to the bundled Excel template at `templates/timesheet.xlsx`.
- Includes a language selector for English, Traditional Chinese, and Japanese.

## Project Structure

```text
main.py
models.py
utils.py
roi_config.py
image_preprocess.py
ocr_engine.py
excel_writer.py
drag_drop_widgets.py
correction_history.py
config.json
requirements.txt
Sum-The-Time-clock.spec
assets/
  drop-image.svg
templates/
  timesheet.xlsx
logs/
  .gitkeep
output/
  .gitkeep
```

The repository intentionally excludes the old packaged executable, `_internal/`, virtual environments, reverse-engineering cache, and generated build output.

## Install

```bat
py -3.12 -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Python 3.11 also works with the reconstructed source.

## Run

```bat
python main.py
```

## Test

```bat
python -m compileall .
python test_static_logic.py
python test_gui_logic.py
```

## Build

```bat
pyinstaller Sum-The-Time-clock.spec --clean --noconfirm
```

To keep generated files outside the source tree:

```bat
pyinstaller Sum-The-Time-clock.spec --clean --noconfirm --distpath ..\release_new --workpath ..\release_build
```

The packaged app will be created at:

```text
..\release_new\Sum-The-Time-clock\
```

## Configuration

Validation thresholds are read from `config.json`:

```json
{
  "validation": {
    "min_shift_minutes": 120,
    "max_shift_minutes": 960,
    "warn_short_shift_minutes": 240,
    "ok_confidence_threshold": 0.88
  }
}
```

When packaged with PyInstaller, the app first looks for `config.json` next to the executable so users can edit settings after deployment. If that file is missing, it falls back to the bundled copy inside PyInstaller data files.
