$code = @"
    [DllImport("user32.dll", SetLastError=true)]
    public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
"@
 
$win = Add-Type -MemberDefinition $code -Name SwitchWindow -Namespace SwitchWindow2 -PassThru
##Process Name of Skype StandAlone
$ProcessName = "Skype"
##Command to start Skype APP W10
$command = "start Skype:"

# Creating a WScript.Shell onject  
 $keyBoardObject = New-Object -ComObject WScript.Shell  
# CAPSLOCK / SCROL LOCK /  Num Lock key Status   
# Getting the status of Keys using .Net Object , this result of the ISkeyLock is in Boolean 
# Capsock 
 $capsLockKeySatus = [System.Windows.Forms.Control]::IsKeyLocked('CapsLock') 

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

if ( $capsLockKeySatus -eq $true )  
{ 
    Write-Host "Capslock key is on"  
         
    # if you want to ON the CapsLock {IF CAPSLOCK IS ALREADY OFF }, then please unComment the below Command. 
        $keyBoardObject.SendKeys("{CAPSLOCK}") 
}  

if(Is-SkypeApp){
    Swap-Window
}
else{
    iex $command
}
