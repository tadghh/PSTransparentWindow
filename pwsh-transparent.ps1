# Add the required Windows API types
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class WindowTransparency {
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
}
"@

function Wait-ForMouseClick {
    $VK_LBUTTON = 0x01
    while ($true) {
        $state = [WindowTransparency]::GetAsyncKeyState($VK_LBUTTON)
        if ($state -band 0x8000) {
            Start-Sleep -Milliseconds 100  # Small delay to ensure click is complete
            return
        }
        Start-Sleep -Milliseconds 10
    }
}

function Get-WindowUnderCursor {
    $point = New-Object System.Drawing.Point
    [WindowTransparency]::GetCursorPos([ref]$point)
    $hwnd = [WindowTransparency]::WindowFromPoint($point)

    # Get window title
    $stringBuilder = New-Object System.Text.StringBuilder 256
    [WindowTransparency]::GetWindowText($hwnd, $stringBuilder, 256)

    return @{
        Handle = $hwnd
        Title = $stringBuilder.ToString()
    }
}

function Set-WindowTransparency {
    param (
        [Parameter(Mandatory=$false)]
        [string]$WindowTitle,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1,100)]
        [int]$Transparency
    )

    # Constants for Windows API
    $WS_EX_LAYERED = 0x80000
    $LWA_ALPHA = 0x2
    $GWL_EXSTYLE = -20

    if ($WindowTitle) {
        # Find window by title
        $hwnd = (Get-Process | Where-Object {$_.MainWindowTitle -like "*$WindowTitle*"} | Select-Object -First 1).MainWindowHandle
        if (-not $hwnd) {
            Write-Host "No window found with title containing: $WindowTitle" -ForegroundColor Red
            return
        }
    } else {
        # Use click selection
        Write-Host "Click on the window you want to make transparent..." -ForegroundColor Cyan
        Wait-ForMouseClick
        $selectedWindow = Get-WindowUnderCursor
        $hwnd = $selectedWindow.Handle
        $WindowTitle = $selectedWindow.Title
    }

    Write-Host "Selected window: $WindowTitle" -ForegroundColor Green

    # Only prompt for transparency if it wasn't provided as a parameter
    if (-not $PSBoundParameters.ContainsKey('Transparency')) {
        do {
            $input = Read-Host "Enter transparency level (1-100)"
            if ($input -match '^\d+$') {
                $Transparency = [int]$input
                if ($Transparency -lt 1 -or $Transparency -gt 100) {
                    Write-Host "Please enter a number between 1 and 100" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Please enter a valid number" -ForegroundColor Yellow
            }
        } while ($Transparency -lt 1 -or $Transparency -gt 100)
    }

    # Convert 1-100 range to 0-255
    $transparencyValue = [Math]::Round(($Transparency / 100) * 255)

    if ($hwnd -and $hwnd -ne [IntPtr]::Zero) {
        # Get current window style
        $style = [WindowTransparency]::GetWindowLong($hwnd, $GWL_EXSTYLE)

        # Add layered window style
        $style = $style -bor $WS_EX_LAYERED
        [void][WindowTransparency]::SetWindowLong($hwnd, $GWL_EXSTYLE, $style)

        # Set the transparency
        [void][WindowTransparency]::SetLayeredWindowAttributes($hwnd, 0, $transparencyValue, $LWA_ALPHA)

        Write-Host "Successfully set transparency for window: $WindowTitle" -ForegroundColor Green
        Write-Host "Transparency: $Transparency% (Value: $transparencyValue)" -ForegroundColor Green
    } else {
        Write-Host "No valid window selected or found" -ForegroundColor Red
    }
}

# Example usage:
# Set-WindowTransparency                                # Click selection, will prompt for transparency
# Set-WindowTransparency -Transparency 70              # Click selection with preset transparency
# Set-WindowTransparency -WindowTitle "Notepad"        # Title selection, will prompt for transparency
# Set-WindowTransparency -WindowTitle "Notepad" -Transparency 70  # Title selection with preset transparency