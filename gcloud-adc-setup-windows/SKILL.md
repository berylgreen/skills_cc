---
name: gcloud-adc-setup-windows
description: Use when需要在 Windows 上安装 Google Cloud SDK 并完成 ADC 认证，且可能遇到 winget source 冲突、PATH 未生效或 quota project 未配置问题。
---

# gcloud ADC Setup (Windows)

## Overview
用于在 Windows 环境稳定完成 Google Cloud SDK 安装、账号登录、ADC 初始化与 quota project 配置，并提供可复现的验证命令。

## When to Use
- 新机器首次配置 `gcloud`
- `gcloud: command not found`
- `winget` 安装时出现 `msstore` 证书/源冲突
- 已登录但 ADC/配额项目未配置导致调用报错

## Quick Reference

### 1) 安装 Cloud SDK
```bash
winget install --id Google.CloudSDK -e --accept-package-agreements --accept-source-agreements
```

若报错（如 `Failed when searching source: msstore` / 证书不匹配），改用：
```bash
winget install --id Google.CloudSDK -e --source winget --accept-package-agreements --accept-source-agreements
```

### 2) 定位可执行文件（PATH 未生效时）
```bash
GCLOUD_CMD="/c/Program Files (x86)/Google/Cloud SDK/google-cloud-sdk/bin/gcloud.cmd"
"$GCLOUD_CMD" --version
```

### 3) 写入 User PATH（永久）
```bash
powershell.exe -NoProfile -Command '$target="C:\Program Files (x86)\Google\Cloud SDK\google-cloud-sdk\bin"; $userPath=[Environment]::GetEnvironmentVariable("Path","User"); if ([string]::IsNullOrEmpty($userPath)) { $userPath="" }; $parts=$userPath -split ";" | Where-Object { $_ -ne "" }; if ($parts -notcontains $target) { $newPath = if ([string]::IsNullOrEmpty($userPath)) { $target } else { "$userPath;$target" }; [Environment]::SetEnvironmentVariable("Path",$newPath,"User") }'
```
> 新开终端后可直接用 `gcloud`。

### 4) 登录并更新 ADC（浏览器授权）
```bash
"$GCLOUD_CMD" auth login --update-adc
"$GCLOUD_CMD" auth application-default login
```

### 5) 配置 quota project
```bash
"$GCLOUD_CMD" projects list --format="table(projectId,name,projectNumber)"
"$GCLOUD_CMD" auth application-default set-quota-project <PROJECT_ID>
```

### 6) 验证
```bash
"$GCLOUD_CMD" --version
"$GCLOUD_CMD" auth list
"$GCLOUD_CMD" auth application-default print-access-token > /dev/null && echo "ADC token OK"
"$GCLOUD_CMD" auth application-default set-quota-project <PROJECT_ID> --quiet
```

## Common Mistakes
- **`gcloud` 找不到**：先用 `GCLOUD_CMD` 全路径执行，再修 PATH。
- **winget 源冲突**：显式加 `--source winget`。
- **有 ADC 但仍报配额/API 错误**：通常是未设置 quota project。
- **验证时泄露 token**：建议重定向到 `/dev/null`，仅检查命令成功。
