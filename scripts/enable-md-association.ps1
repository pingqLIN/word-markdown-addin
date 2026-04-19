param(
  [string]$WordPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
  [switch]$Undo
)

$extension = ".md"
$fileType = "Word.MarkdownDocument"

if ($Undo) {
  Write-Host "移除本機 .md 使用者關聯（HKCU 範圍）..."
  Remove-Item -Path "HKCU:\Software\Classes\$extension" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path "HKCU:\Software\Classes\$fileType" -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host "已清除關聯。請確認有無其他企業原則覆蓋。"
  return
}

if (-not (Test-Path $WordPath)) {
  throw "找不到 Word 可執行檔：$WordPath"
}

$shell = New-Item -Path "HKCU:\Software\Classes\$extension" -Force | Out-Null
New-Item -Path "HKCU:\Software\Classes\$extension" -Name "OpenWithProgids" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$extension" -Name "(default)" -Value $fileType
New-Item -Path "HKCU:\Software\Classes\$fileType\shell\open\command" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$fileType\shell\open\command" -Name "(default)" -Value "`"$WordPath`" /n /q /mFileOpen `"%1`""
New-Item -Path "HKCU:\Software\Classes\$fileType\DefaultIcon" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$fileType\DefaultIcon" -Name "(default)" -Value "$WordPath,0"

Write-Host "完成：.md 已關聯到 Word。"
Write-Host "注意：僅在 Word 已 sideload Word Markdown Companion 後才能使用加值集的 Markdown 匯入匯出流程。"
