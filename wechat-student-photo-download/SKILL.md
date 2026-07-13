---
name: wechat-student-photo-download
description: Use when Codex needs to collect student photos from a WeChat mini program class list, extract 学号 and 姓名 into a clean roster, batch-download photos named as 学号姓名.jpg, or fall back to copying WeChat cache images when direct download is incomplete.
---

# WeChat Student Photo Download

将小程序文本、OCR 结果或整理后的名单规范成可用花名册，然后批量下载学生照片。优先按学号直连下载，微信缓存图片仅作为补救手段。

## Workflow

1. 准备或清洗名单文件，格式为 `学号 姓名`。
2. 下载照片到新的输出文件夹。
3. 检查生成的报表：同时保留 `_photo_download_report.csv`，并生成“输出文件夹同名 `.xlsx`”。
4. 若直连下载不完整，再复制微信缓存 JPEG。
5. 完成替换后，删除临时 `f_*.jpg` 缓存副本。

## Prepare The Roster

当源数据来自小程序页面文本、截图 OCR 或混杂剪贴板内容时，先使用 `scripts/normalize_student_list.py` 清洗。

最终行格式：

```text
124232024001 卓振维
124232024003 何女晶
```

运行：

```powershell
python .\scripts\normalize_student_list.py --input .\raw.txt --output .\students.txt
```

脚本会保留合法学号，按学号去重，并输出 UTF-8 编码文本供后续下载使用。

## Download Photos

运行：

```powershell
python .\scripts\photo_batch_tools.py --list .\students.txt --out .\23数媒 --download
```

下载地址模板：

```text
https://zhjw.fjcuc.cn/photo/{student_id}.jpg
```

照片命名规则：

```text
学号姓名.jpg
```

`--list` 支持 `.txt`、`.csv`、`.json`。

## Use Cache As Fallback

若部分学生下载失败，可从微信缓存复制可能的原始照片：

```powershell
python .\scripts\photo_batch_tools.py --out .\23数媒 --copy-cache "$env:APPDATA\Tencent\xwechat\radium\web\profiles\webview_404ee7fc58206f51bad67b919e89c585\Cache\Cache_Data"
```

该步骤仅复制文件名为 `f_*` 且分辨率为 `480x640` 的 JPEG。它们只是临时补救素材，不应作为最终交付命名依据。

完成匹配或替换后，清理这些原始缓存副本：

```powershell
python .\scripts\photo_batch_tools.py --out .\23数媒 --delete-cache-copies
```

## Check Reports

下载完成后，输出目录中会同时生成两份主报表：

- `_photo_download_report.csv`：完整字段，便于复核下载地址、状态和字节数
- `输出文件夹同名.xlsx`：仅保留前三列，便于交付和核对名单

例如输出目录为 `23数媒`，则会生成：

- `23数媒\_photo_download_report.csv`
- `23数媒\23数媒.xlsx`

XLSX 工作表仅包含以下中文表头：

- `序号`
- `学号`
- `姓名`

CSV 仍保留完整字段：

- `seq`
- `student_id`
- `name`
- `url`
- `status`
- `bytes`

常见状态：

- `downloaded`：下载正常。
- `failed`：请求失败或图片校验失败。
- `bad_content` / `bad_content_9664`：返回内容可疑，需要人工复核。

## Operational Notes

- 每个班级使用独立输出目录。
- 名单文件是最终命名基准。
- 小程序如果能直接复制文本，优先复制文本，不要先截图再 OCR。
- OCR 噪声较大时，先落盘再用脚本清洗，不要手工逐行修。
- 不要用缓存文件名识别学生身份，缓存图仅用于补救。

## Resources

- `scripts/normalize_student_list.py`：将原始小程序文本或 OCR 结果清洗成 `学号 姓名`。
- `scripts/photo_batch_tools.py`：按学号下载照片、复制缓存 JPEG、清理缓存副本，并生成 CSV + XLSX 报表。
