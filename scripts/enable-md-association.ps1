param(
  [string]$WordPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
  [switch]$Undo
)

$extension = ".md"
$fileType = "Word.MarkdownDocument"
$launcherPath = Join-Path $PSScriptRoot "open-markdown-in-word.js"
$wordApplicationProgId = "Applications\WINWORD.EXE"
$settingsRoot = "HKCU:\Software\WordMarkdownCompanion"
$settingsBackupPath = Join-Path $settingsRoot "Backup"

function Get-UserChoiceProgId {
  try {
    return (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\UserChoice" -ErrorAction Stop).ProgId
  } catch {
    return $null
  }
}

function Get-OpenCommandValue([string]$progId) {
  if (-not $progId) {
    return $null
  }

  try {
    return (Get-Item "HKCU:\Software\Classes\$progId\shell\open\command" -ErrorAction Stop).GetValue("")
  } catch {
    return $null
  }
}

function Set-DefaultRegistryValue([string]$path, [string]$value) {
  New-Item -Path $path -Force | Out-Null
  Set-Item -Path $path -Value $value
}

function Set-OpenCommandValue([string]$progId, [string]$commandValue) {
  if (-not $progId) {
    return
  }

  Set-DefaultRegistryValue -path "HKCU:\Software\Classes\$progId\shell\open\command" -value $commandValue
}

if ($Undo) {
  Write-Host "Removing per-user .md association from HKCU..."
  $backup = Get-ItemProperty -Path $settingsBackupPath -ErrorAction SilentlyContinue
if ($backup -and $backup.UserChoiceProgId -and $backup.UserChoiceCommand) {
    Set-OpenCommandValue -progId $backup.UserChoiceProgId -commandValue $backup.UserChoiceCommand
  }
  if ($backup -and $backup.WordApplicationCommand) {
    Set-OpenCommandValue -progId $wordApplicationProgId -commandValue $backup.WordApplicationCommand
  }

  Remove-Item -Path "HKCU:\Software\Classes\$extension" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path "HKCU:\Software\Classes\$fileType" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path $settingsRoot -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host "Association removed. Check whether any enterprise policy restores it."
  return
}

if (-not (Test-Path $WordPath)) {
  throw "Word executable not found: $WordPath"
}

if (-not (Test-Path $launcherPath)) {
  throw "Markdown launcher not found: $launcherPath"
}

$nodePath = (Get-Command node -ErrorAction Stop).Source
$launcherCommand = "`"$nodePath`" `"$launcherPath`" `"%1`""
$userChoiceProgId = Get-UserChoiceProgId
$existingUserChoiceCommand = Get-OpenCommandValue -progId $userChoiceProgId

New-Item -Path $settingsBackupPath -Force | Out-Null
if ($userChoiceProgId) {
  Set-ItemProperty -Path $settingsBackupPath -Name "UserChoiceProgId" -Value $userChoiceProgId
}
if ($existingUserChoiceCommand) {
  Set-ItemProperty -Path $settingsBackupPath -Name "UserChoiceCommand" -Value $existingUserChoiceCommand
}
$existingWordApplicationCommand = Get-OpenCommandValue -progId $wordApplicationProgId
if ($existingWordApplicationCommand) {
  Set-ItemProperty -Path $settingsBackupPath -Name "WordApplicationCommand" -Value $existingWordApplicationCommand
}

$shell = New-Item -Path "HKCU:\Software\Classes\$extension" -Force | Out-Null
New-Item -Path "HKCU:\Software\Classes\$extension" -Name "OpenWithProgids" -Force | Out-Null
Set-Item -Path "HKCU:\Software\Classes\$extension" -Value $fileType
Set-DefaultRegistryValue -path "HKCU:\Software\Classes\$fileType\shell\open\command" -value $launcherCommand
Set-DefaultRegistryValue -path "HKCU:\Software\Classes\$fileType\DefaultIcon" -value "$WordPath,0"

if ($userChoiceProgId) {
  Set-OpenCommandValue -progId $userChoiceProgId -commandValue $launcherCommand
}
Set-OpenCommandValue -progId $wordApplicationProgId -commandValue $launcherCommand

Write-Host "Done: .md is now associated with the Word Markdown Companion launcher."
Write-Host "Double-clicking a .md file now stages its content first, then opens a blank Word window instead of letting Word open the markdown file directly as plain text."
