$code = @"
    [DllImport("user32.dll", SetLastError=true)]
    public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
"@
 
$win = Add-Type -MemberDefinition $code -Name SwitchWindow -Namespace SwitchWindow2 -PassThru
##Process Name of Skype StandAlone
$ProcessName = "Skype"
##Command to start Skype APP W10
$command = "start Skype:"

Function Get-SkypeHandle {
    Get-Process $ProcessName| 
        Select MainWindowHandle, Name
}

Function Is-SkypeApp {
    if((get-process $ProcessName -ErrorAction SilentlyContinue) -eq $Null){ 
        return $false
    }
    else{ 
        return $true
    }
}
 
Function Swap-Window {
    $proc = Get-SkypeHandle
    Write-Host "Swapping to process: $($proc.Name) - handle: $($proc.MainWindowHandle)..."
    $win::SwitchToThisWindow([System.IntPtr]$proc.MainWindowHandle, $true)
}

if(Is-SkypeApp){
    Swap-Window
}
else{
    iex $command
}
