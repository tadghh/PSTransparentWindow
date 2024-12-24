# First, include the enhanced WindowTransparency class with additional API methods
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class WindowTransparency2 {
    [DllImport("user32.dll")]
    public static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(System.Drawing.Point point);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out System.Drawing.Point lpPoint);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    // Added APIs for window enumeration
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
}
'@

# Configuration file path
$configPath = Join-Path $env:APPDATA 'WindowTransparency\config.json'

function Initialize-TransparencyConfig {
  if (-not (Test-Path (Split-Path $configPath -Parent))) {
    New-Item -ItemType Directory -Path (Split-Path $configPath -Parent) -Force | Out-Null
  }
  if (-not (Test-Path $configPath)) {
    @{
      Processes = @{}
      Windows   = @{}
    } | ConvertTo-Json | Set-Content $configPath
  }
}

function Get-TransparencyConfig {
  if (Test-Path $configPath) {
    Get-Content $configPath | ConvertFrom-Json -AsHashtable
  }
  else {
    @{
      Processes = @{}
      Windows   = @{}
    }
  }
}
function Wait-ForMouseClick {
  $VK_LBUTTON = 0x01
  while ($true) {
    $state = [WindowTransparency2]::GetAsyncKeyState($VK_LBUTTON)
    if ($state -band 0x8000) {
      Start-Sleep -Milliseconds 100  # Small delay to ensure click is complete
      return
    }
    Start-Sleep -Milliseconds 10
  }
}
function Get-WindowUnderCursor {
  $point = New-Object System.Drawing.Point
  [void][WindowTransparency2]::GetCursorPos([ref]$point)
  $hwnd = [WindowTransparency2]::WindowFromPoint($point)

  # Get window title
  $title = New-Object System.Text.StringBuilder 256
  [WindowTransparency2]::GetWindowText($hwnd, $title, 256)

  # Get window class
  $className = New-Object System.Text.StringBuilder 256
  [WindowTransparency2]::GetClassName($hwnd, $className, 256)

  # Get process ID
  $processId = 0
  [void][WindowTransparency2]::GetWindowThreadProcessId($hwnd, [ref]$processId)

  $process = Get-Process -Id $processId -ErrorAction SilentlyContinue

  return @{
    Handle      = $hwnd
    Title       = $title.ToString()
    ClassName   = $className.ToString()
    ProcessName = $process.ProcessName
    ProcessId   = $processId
  }
}

function Get-WindowsByProcessName {
  param(
      [Parameter(Mandatory=$true)]
      [string]$ProcessName
  )

  # First, ensure we have access to the Win32 API if not already defined
  if (-not ('WindowTransparency2' -as [type])) {
      Add-Type @"
          using System;
          using System.Runtime.InteropServices;
          public class WindowTransparency2 {
              [DllImport("user32.dll")]
              public static extern bool GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);
              [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
              public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
              [DllImport("user32.dll")]
              public static extern uint GetWindowThreadProcessId(IntPtr hWnd, ref int processId);
          }
"@
  }

  Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
  Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
  ForEach-Object {
      $hwnd = $_.MainWindowHandle

      # Get window title
      $title = New-Object System.Text.StringBuilder 256
      [WindowTransparency2]::GetWindowText($hwnd, $title, 256)

      # Get window class
      $className = New-Object System.Text.StringBuilder 256
      [WindowTransparency2]::GetClassName($hwnd, $className, 256)

      # Get process ID
      $processId = 0
      [void][WindowTransparency2]::GetWindowThreadProcessId($hwnd, [ref]$processId)

      @{
          Handle      = $hwnd
          Title       = $title.ToString()
          ClassName   = $className.ToString()
          ProcessName = $_.ProcessName
          ProcessId   = $processId
      }
  }
}
function Set-TransparencyConfig {
  param (
    [string]$ProcessName,
    [int]$Transparency,
    [string]$WindowClass = $null
  )

  $ProcessName = $ProcessName.ToLower()
  $config = Get-TransparencyConfig
  if ($WindowClass) {
    if (-not $config.Windows) { $config.Windows = @{} }
    $config.Windows["$ProcessName|$WindowClass"] = @{
      Transparency = $Transparency
      WindowClass  = $WindowClass
    }
  }
  # else {

    # If an existing entry in the config exists that contains WindowClass it will override the process (else)
    # $config.Processes[$ProcessName] = $Transparency
  # }
  $config | ConvertTo-Json | Set-Content $configPath
}

function Get-WindowsOfProcess {
  param (
    [string]$ProcessName
  )

  $windows = New-Object System.Collections.ArrayList
  $processIds = (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Id

  if (-not $processIds) { return $windows }

  $enumWindowsCallback = {
    param(
      [IntPtr]$hwnd,
      [IntPtr]$lParam
    )

    $processId = 0
    [void][WindowTransparency2]::GetWindowThreadProcessId($hwnd, [ref]$processId)

    if ($processIds -contains $processId -and [WindowTransparency2]::IsWindowVisible($hwnd)) {
      $className = New-Object System.Text.StringBuilder 256
      [WindowTransparency2]::GetClassName($hwnd, $className, 256)

      $title = New-Object System.Text.StringBuilder 256
      [WindowTransparency2]::GetWindowText($hwnd, $title, 256)

      [void]$windows.Add(@{
          Handle    = $hwnd
          Title     = $title.ToString()
          ClassName = $className.ToString()
          ProcessId = $processId
        })
    }
    return $true
  }

  $enumWindowsDelegate = [WindowTransparency2+EnumWindowsProc]$enumWindowsCallback
  [WindowTransparency2]::EnumWindows($enumWindowsDelegate, [IntPtr]::Zero)

  return $windows
}

function Apply-WindowTransparency {
  param (
    [IntPtr]$WindowHandle,
    [int]$Transparency
  )

  $WS_EX_LAYERED = 0x80000
  $LWA_ALPHA = 0x2
  $GWL_EXSTYLE = -20

  if ($WindowHandle -and $WindowHandle -ne [IntPtr]::Zero) {
    $transparencyValue = [Math]::Round(($Transparency / 100) * 255)
    $style = [WindowTransparency2]::GetWindowLong($WindowHandle, $GWL_EXSTYLE)
    $style = $style -bor $WS_EX_LAYERED
    [void][WindowTransparency2]::SetWindowLong($WindowHandle, $GWL_EXSTYLE, $style)
    [void][WindowTransparency2]::SetLayeredWindowAttributes($WindowHandle, 0, $transparencyValue, $LWA_ALPHA)
    return $true
  }
  return $false
}

function Start-TransparencyMonitor {

  $job = Start-Job -ScriptBlock {
    param($configPath)

    $processedWindows = @{}
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class WindowTransparency2 {
    [DllImport("user32.dll")]
    public static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern IntPtr WindowFromPoint(System.Drawing.Point point);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out System.Drawing.Point lpPoint);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    // Added APIs for window enumeration
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
}
'@

    while ($true) {
      $config = Get-Content $configPath | ConvertFrom-Json -AsHashtable

      # Process regular process rules
      foreach ($processEntry in $config.Processes.GetEnumerator()) {
        $processName = $processEntry.Key
        $transparency = $processEntry.Value

        Get-Process -Name $processName -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
        ForEach-Object {
          $hwnd = $_.MainWindowHandle
          if (-not $processedWindows.ContainsKey($hwnd)) {
            $style = [WindowTransparency2]::GetWindowLong($hwnd, -20)
            $style = $style -bor 0x80000
            [void][WindowTransparency2]::SetWindowLong($hwnd, -20, $style)
            [void][WindowTransparency2]::SetLayeredWindowAttributes($hwnd, 0, [Math]::Round(($transparency / 100) * 255), 2)
            $processedWindows[$hwnd] = $true
          }
        }
      }

      # Process window class specific rules
      if ($config.Windows) {
        foreach ($windowEntry in $config.Windows.GetEnumerator()) {
          $processName, $targetClass = $windowEntry.Key -split '\|'
          $transparency = $windowEntry.Value.Transparency

          $enumCallback = {
            param([IntPtr]$hwnd, [IntPtr]$lParam)

            $className = New-Object System.Text.StringBuilder 256
            [WindowTransparency2]::GetClassName($hwnd, $className, 256)

            if ($className.ToString() -eq $targetClass -and [WindowTransparency2]::IsWindowVisible($hwnd)) {
              $processId = 0
              [WindowTransparency2]::GetWindowThreadProcessId($hwnd, [ref]$processId)

              try {
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if ($process.ProcessName -eq $processName) {
                  $style = [WindowTransparency2]::GetWindowLong($hwnd, -20)
                  $style = $style -bor 0x80000
                  [void][WindowTransparency2]::SetWindowLong($hwnd, -20, $style)
                  [void][WindowTransparency2]::SetLayeredWindowAttributes($hwnd, 0, [Math]::Round(($transparency / 100) * 255), 2)
                }
              }
              catch {}
            }
            return $true
          }

          $enumDelegate = [WindowTransparency2+EnumWindowsProc]$enumCallback
          [WindowTransparency2]::EnumWindows($enumDelegate, [IntPtr]::Zero)
        }
      }

      Start-Sleep -Milliseconds 500
    }
  } -ArgumentList $configPath

  return $job
}
function Add-TransparencyRule {
  param (
    [Parameter(Mandatory = $false)]
    [string]$ProcessName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Transparency,

    [Parameter(Mandatory = $false)]
    [string]$WindowClass
  )

  Initialize-TransparencyConfig

  # If no process name provided, use cursor selection
  if (-not $ProcessName) {
    Write-Host 'Click on the window you want to make transparent...' -ForegroundColor Cyan
    Wait-ForMouseClick
    $selectedWindow = Get-WindowUnderCursor
    $ProcessName = $selectedWindow.ProcessName
    $WindowClass = $selectedWindow.ClassName
    Write-Host "Selected window: $($selectedWindow.Title)" -ForegroundColor Green
    Write-Host "Process: $ProcessName" -ForegroundColor Green
    Write-Host "Window Class: $WindowClass" -ForegroundColor Green
  }

  # If no transparency provided, prompt for it
  if (-not $PSBoundParameters.ContainsKey('Transparency')) {
    do {
      $input = Read-Host 'Enter transparency level (1-100)'
      if ($input -match '^\d+$') {
        $Transparency = [int]$input
        if ($Transparency -lt 1 -or $Transparency -gt 100) {
          Write-Host 'Please enter a number between 1 and 100' -ForegroundColor Yellow
        }
      }
      else {
        Write-Host 'Please enter a valid number' -ForegroundColor Yellow
      }
    } while ($Transparency -lt 1 -or $Transparency -gt 100)
  }

  if ($WindowClass) {
    Write-Host "Adding transparency rule for $ProcessName windows with class $WindowClass" -ForegroundColor Green
    Set-TransparencyConfig -ProcessName $ProcessName -Transparency $Transparency -WindowClass $WindowClass

    # Apply immediately to existing windows
    $windows = Get-WindowsOfProcess -ProcessName $ProcessName
    $windows | Where-Object { $_.ClassName -eq $WindowClass } | ForEach-Object {
      Apply-WindowTransparency -WindowHandle $_.Handle -Transparency $Transparency
    }
  }
  else {
    Write-Host "Adding transparency rule for $ProcessName" -ForegroundColor Green
    $processInfo = Get-WindowsByProcessName -ProcessName $ProcessName
    Set-TransparencyConfig -ProcessName $ProcessName -Transparency $Transparency -WindowClass $processInfo.ClassName

    # Apply immediately to existing windows
    $windows = Get-WindowsOfProcess -ProcessName $ProcessName
    $windows | Where-Object { $_.ClassName -eq $WindowClass } | ForEach-Object {
      Apply-WindowTransparency -WindowHandle $_.Handle -Transparency $Transparency
    }
  }
}

# Initialize and start the monitor
Initialize-TransparencyConfig
$monitorJob = Start-TransparencyMonitor

# Example usage:
# For regular windows:
# Add-TransparencyRule -ProcessName "notepad" -Transparency 70

# For Explorer windows (including "This PC"):
# Add-TransparencyRule -ProcessName "explorer" -WindowClass "CabinetWClass" -Transparency 70
