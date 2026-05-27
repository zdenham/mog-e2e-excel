param(
  [Parameter(Mandatory = $true)]
  [string] $Path,

  [int] $TimeoutSeconds = 120
)

$ErrorActionPreference = "Stop"

function Write-CheckResult {
  param(
    [Parameter(Mandatory = $true)]
    [string] $Status,

    [Parameter(Mandatory = $true)]
    [string] $Message
  )

  [pscustomobject]@{
    status = $Status
    message = $Message
  } | ConvertTo-Json -Compress
}

function Stop-ProcessTree {
  param([int] $ProcessId)

  try {
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
  } catch {
  }
}

function Get-DialogText {
  param([int] $ExcelProcessId)

  try {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes -ErrorAction Stop
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
      $ExcelProcessId
    )
    $windows = $root.FindAll(
      [System.Windows.Automation.TreeScope]::Children,
      $processCondition
    )
    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($window in $windows) {
      if ($window.Current.Name) {
        $parts.Add($window.Current.Name)
      }

      $textCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Text
      )
      $texts = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $textCondition
      )
      foreach ($text in $texts) {
        if ($text.Current.Name) {
          $parts.Add($text.Current.Name)
        }
      }

      $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button
      )
      $buttons = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $buttonCondition
      )
      foreach ($button in $buttons) {
        if ($button.Current.Name) {
          $parts.Add($button.Current.Name)
        }
      }
    }

    return ($parts -join " ")
  } catch {
    return ""
  }
}

function Dismiss-ExcelDialogs {
  param([int] $ExcelProcessId)

  try {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes -ErrorAction Stop
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
      $ExcelProcessId
    )
    $windows = $root.FindAll(
      [System.Windows.Automation.TreeScope]::Children,
      $processCondition
    )
    $buttonNames = @("Don't Save", "No", "Cancel", "OK", "Close")

    foreach ($window in $windows) {
      foreach ($buttonName in $buttonNames) {
        $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
          [System.Windows.Automation.AutomationElement]::NameProperty,
          $buttonName
        )
        $button = $window.FindFirst(
          [System.Windows.Automation.TreeScope]::Descendants,
          $nameCondition
        )
        if ($null -ne $button) {
          $invoke = $button.GetCurrentPattern(
            [System.Windows.Automation.InvokePattern]::Pattern
          )
          $invoke.Invoke()
          Start-Sleep -Milliseconds 250
        }
      }
    }
  } catch {
  }
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
  Write-CheckResult -Status "unsupported" -Message "Excel COM validation requires Windows."
  exit 0
}

try {
  $fullPath = (Resolve-Path -LiteralPath $Path).Path
} catch {
  Write-CheckResult -Status "error" -Message "Workbook is not readable: $Path"
  exit 1
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("mog-excel-com-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$workerPath = Join-Path $tempDir "worker.ps1"
$resultPath = Join-Path $tempDir "result.json"
$pidPath = Join-Path $tempDir "excel.pid"

function Quote-Argument {
  param([string] $Value)
  return '"' + ($Value -replace '"', '\"') + '"'
}

@'
param(
  [Parameter(Mandatory = $true)]
  [string] $WorkbookPath,

  [Parameter(Mandatory = $true)]
  [string] $ResultPath,

  [Parameter(Mandatory = $true)]
  [string] $PidPath
)

$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeMethods {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Write-WorkerResult {
  param([string] $Status, [string] $Message)
  [pscustomobject]@{
    status = $Status
    message = $Message
  } | ConvertTo-Json -Compress | Set-Content -LiteralPath $ResultPath -Encoding UTF8
}

$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $true
  $excel.AskToUpdateLinks = $false
  try {
    $excel.AutomationSecurity = 3
  } catch {
  }

  $excelPid = 0
  [void][NativeMethods]::GetWindowThreadProcessId([intptr]$excel.Hwnd, [ref]$excelPid)
  Set-Content -LiteralPath $PidPath -Value $excelPid -Encoding ASCII

  $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true)
  Start-Sleep -Seconds 2
  $workbook.Close($false)
  Write-WorkerResult -Status "ok" -Message "OK"
} catch {
  $message = $_.Exception.Message
  if ($message -match "(?i)corrupt|repair|recovered|unreadable|found a problem") {
    Write-WorkerResult -Status "corrupt" -Message $message
  } else {
    Write-WorkerResult -Status "error" -Message $message
  }
} finally {
  try {
    if ($null -ne $workbook) {
      $workbook.Close($false)
    }
  } catch {
  }
  try {
    if ($null -ne $excel) {
      $excel.Quit()
    }
  } catch {
  }
  if ($null -ne $workbook) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }
  if ($null -ne $excel) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
}
'@ | Set-Content -LiteralPath $workerPath -Encoding UTF8

$powershell = (Get-Command powershell.exe).Source
$workerArguments = @(
  "-NoProfile",
  "-STA",
  "-ExecutionPolicy",
  "Bypass",
  "-File",
  $workerPath,
  "-WorkbookPath",
  $fullPath,
  "-ResultPath",
  $resultPath,
  "-PidPath",
  $pidPath
) | ForEach-Object { Quote-Argument $_ }

$worker = Start-Process `
  -FilePath $powershell `
  -ArgumentList ($workerArguments -join " ") `
  -PassThru

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$excelPid = $null
$dialogPattern = "(?i)found a problem|corrupt|repair|recovered|unreadable content"

while ((Get-Date) -lt $deadline) {
  if ((Test-Path -LiteralPath $pidPath) -and $null -eq $excelPid) {
    $pidText = Get-Content -LiteralPath $pidPath -Raw
    if ($pidText -match "\d+") {
      $excelPid = [int]$Matches[0]
    }
  }

  if ($null -ne $excelPid) {
    $dialogText = Get-DialogText -ExcelProcessId $excelPid
    if ($dialogText -match $dialogPattern) {
      Dismiss-ExcelDialogs -ExcelProcessId $excelPid
      Stop-ProcessTree -ProcessId $worker.Id
      Stop-ProcessTree -ProcessId $excelPid
      Write-CheckResult -Status "corrupt" -Message ("CORRUPT_DIALOG: " + $dialogText)
      exit 1
    }
  }

  if (Test-Path -LiteralPath $resultPath) {
    $json = Get-Content -LiteralPath $resultPath -Raw
    Write-Output $json
    $parsed = $json | ConvertFrom-Json
    if ($parsed.status -eq "ok") {
      exit 0
    }
    exit 1
  }

  if ($worker.HasExited) {
    break
  }

  Start-Sleep -Milliseconds 500
}

if ($null -ne $excelPid) {
  Dismiss-ExcelDialogs -ExcelProcessId $excelPid
  Stop-ProcessTree -ProcessId $excelPid
}
Stop-ProcessTree -ProcessId $worker.Id
Write-CheckResult -Status "error" -Message "Excel COM validation timed out."
exit 1
