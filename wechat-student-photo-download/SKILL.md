---
name: wechat-student-photo-download
description: Use when Codex needs to collect student photos from a WeChat mini program class list, extract 学号 and 姓名 into a clean roster, batch-download photos named as 学号姓名.jpg, or fall back to copying WeChat cache images when direct download is incomplete.
---

# WeChat Student Photo Download

Extract a clean student roster from mini-program text or OCR output, then batch-download named student photos with the bundled scripts. Prefer direct download by 学号 first; use cache copying only as fallback.

## Workflow

1. Build or clean a roster file containing `学号 姓名`.
2. Download photos into a new output folder.
3. Check the generated CSV report for failures or suspicious files.
4. Copy WeChat cache JPEGs only if direct download is incomplete.
5. Delete raw `f_*.jpg` cache copies after finishing.

## Prepare The Roster

Use `scripts/normalize_student_list.py` when the source text comes from WeChat list pages, screenshots after OCR, or mixed clipboard text.

Expected final line format:

```text
124232024001 卓振绅
124232024003 何女晏
```

Run:

```powershell
python .\scripts\normalize_student_list.py --input .\raw.txt --output .\students.txt
```

The script keeps valid student IDs, removes duplicates by 学号, and writes UTF-8 text ready for batch download.

## Download Photos

Run:

```powershell
python .\scripts\photo_batch_tools.py --list .\students.txt --out .\23数媒 --download
```

The script downloads from:

```text
https://zhjw.fjcuc.cn/photo/{student_id}.jpg
```

Output naming rule:

```text
学号姓名.jpg
```

Supported roster formats for `--list`: `.txt`, `.csv`, `.json`.

## Use Cache As Fallback

If some students fail to download, copy likely original photos from WeChat cache:

```powershell
python .\scripts\photo_batch_tools.py --out .\23数媒 --copy-cache "$env:APPDATA\Tencent\xwechat\radium\web\profiles\webview_404ee7fc58206f51bad67b919e89c585\Cache\Cache_Data"
```

This only copies `480x640` JPEG files named like `f_*`. Treat these files as temporary fallback material, not final deliverables.

After you finish matching or replacing them, clean the raw copies:

```powershell
python .\scripts\photo_batch_tools.py --out .\23数媒 --delete-cache-copies
```

## Check Reports

Review `_photo_download_report.csv` in the output folder.

Important statuses:
- `downloaded`: usable result.
- `failed`: request or image validation failed.
- `bad_content` / `bad_content_9664`: server returned suspicious content; verify manually.

## Operational Notes

- Start each class in a fresh output folder.
- Keep the roster file as the naming source of truth.
- If the mini program exposes text directly, prefer copying text over screenshot OCR.
- If OCR output is noisy, save it first and normalize with the script instead of hand-editing dozens of lines.
- Do not rely on cache filenames for student identity; they are only temporary raw images.

## Resources

- `scripts/normalize_student_list.py`: clean raw WeChat text or OCR output into `学号 姓名`.
- `scripts/photo_batch_tools.py`: download by 学号, copy cache JPEGs, and clean cache copies.
