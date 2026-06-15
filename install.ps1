#Requires -Version 5.1
<#
.SYNOPSIS
    TKC仕訳帳ビューアー (VSCode拡張) を最新リリースからインストールする。

.DESCRIPTION
    GitHub Releases から最新の .vsix をダウンロードして
    `code --install-extension` で VSCode に入れる。

.EXAMPLE
    irm https://raw.githubusercontent.com/REPLACE_OWNER/tkc-journal-viewer/main/install.ps1 | iex
#>

$ErrorActionPreference = "Stop"

# ↓ GitHub に push したらここを書き換える
$Owner = "ryobang"
$Repo  = "tkc-journal-viewer"

# 1. code コマンドの存在チェック
if (-not (Get-Command code -ErrorAction SilentlyContinue)) {
    Write-Host "[ERR] VSCodeの 'code' コマンドが見つからない。" -ForegroundColor Red
    Write-Host "      VSCodeで Ctrl+Shift+P → 'Shell Command: Install code command in PATH' を実行してね。"
    exit 1
}

# 2. 最新リリースの.vsixを取得
$apiUrl = "https://api.github.com/repos/$Owner/$Repo/releases/latest"
Write-Host "[1/3] 最新リリース確認: $apiUrl"
$release = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "tkc-installer" }
$asset = $release.assets | Where-Object { $_.name -like "*.vsix" } | Select-Object -First 1
if (-not $asset) {
    Write-Host "[ERR] リリースに .vsix が見つからない。" -ForegroundColor Red
    exit 1
}

# 3. ダウンロード
$dest = Join-Path $env:TEMP $asset.name
Write-Host "[2/3] ダウンロード: $($asset.name)"
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $dest -UseBasicParsing

# 4. インストール
Write-Host "[3/3] VSCodeにインストール中..."
& code --install-extension $dest --force | Out-Host

Write-Host ""
Write-Host "[OK] $($release.tag_name) インストール完了。VSCodeを再起動してね。" -ForegroundColor Green
