param(
  [string]$ProjectId = "3713261754210573091",
  [string]$ListPath = ".\lab_posts_latest.json",
  [string]$RulesPath = ".\preconstruction_validation_rules.yaml",
  [string]$OutputPath = ".\preconstruction_validation_report_latest.txt"
)

$ErrorActionPreference = "Stop"

function Get-PythonCommand {
  if (Test-Path 'C:\Python314\python.exe') {
    return @('C:\Python314\python.exe')
  }

  $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
  if ($pythonCmd -and $pythonCmd.Source) {
    return @($pythonCmd.Source)
  }

  $pyCmd = Get-Command py -ErrorAction SilentlyContinue
  if ($pyCmd -and $pyCmd.Source) {
    return @($pyCmd.Source, '-3')
  }

  throw "Python is not installed. Run install_jk_deadline_bot_server.ps1 first."
}

$Token = $env:DOORAY_API_TOKEN
if ([string]::IsNullOrWhiteSpace($Token)) {
  throw "DOORAY_API_TOKEN is not set."
}

$headers = @{ Authorization = "dooray-api $Token" }
$postsUrl = "https://api.gov-dooray.com/project/v1/projects/$ProjectId/posts?page=0&size=100"
Invoke-WebRequest -Headers $headers -Uri $postsUrl -OutFile $ListPath | Out-Null

Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
$pythonCommand = Get-PythonCommand
$pythonExe = $pythonCommand[0]
$pythonArgs = @()
if ($pythonCommand.Count -gt 1) {
  $pythonArgs = $pythonCommand[1..($pythonCommand.Count - 1)]
}
& $pythonExe @pythonArgs .\validate_preconstruction_meetings.py $Token $ListPath $RulesPath $OutputPath
if ($LASTEXITCODE -ne 0) {
  throw "validate_preconstruction_meetings.py failed with exit code $LASTEXITCODE"
}

Write-Output "list_path=$ListPath"
Write-Output "report_path=$OutputPath"
Write-Output "json_path=$([System.IO.Path]::ChangeExtension($OutputPath, '.json'))"
