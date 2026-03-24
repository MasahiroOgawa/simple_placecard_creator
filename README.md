# Simple Place Card Creator (席札作成ツール)

Create printable wedding place cards (席札) from a guest list file. Outputs editable LibreOffice Impress (.pptx) and PDF files with 8 foldable cards per A4 portrait page (2 columns x 4 rows).

## Setup

```bash
uv sync
```

Requires LibreOffice installed for PDF conversion.

## Usage

### 1. Extract guest names from xlsx/csv

```bash
uv run python3 extract_guests.py data/ゲスト一覧.xlsx -o output/guests.csv
```

Also supports CSV input (auto-detects Japanese encodings):

```bash
uv run python3 extract_guests.py data/ゲスト一覧.csv -o output/guests.csv
```

Extracts all attending guests (ご出席) including 連名 (joint names) and outputs a CSV with display names and 様 suffix.

### 2. Create place cards

```bash
uv run python3 create_placecards.py output/guests.csv -o output/
```

Accepts Excel (.xlsx), CSV (from step 1), or a plain text file (one name per line). Output defaults to `output/placecards.pptx`.

#### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o`, `--output` | `output/` | Output directory |
| `--welcome` | `welcome` | Text above the name |
| `--date` | `April 19, 2026` | Text below the name |
| `--font` | `Noto Serif CJK JP` | Font family |
| `--no-pdf` | off | Skip PDF conversion |

## Output

- **A4 portrait**, 8 foldable cards per page (2x4)
- Separator lines between cards for cutting
- Each card: welcome text (Great Vibes script font), guest name with 様, date
- Font size auto-adjusts based on name length

## Special case handling

`extract_guests.py` includes handlers for data quirks in the guest list:

- **連名 with first-name only**: Family name inherited from main guest (e.g. 河野 family)
- **連名 with swapped fields**: Family name in name column, given name in furigana column (e.g. 小川 family)
- **CamelCase English names**: Auto-split (e.g. `McDermottEthan` → `McDermott Ethan`)
- **Slash-separated names**: Reordered (e.g. `Murphy/Josiah` → `Josiah Murphy`)

To add new special cases, register a handler function in the `SPECIAL_CASES` dict keyed by guest ID.
