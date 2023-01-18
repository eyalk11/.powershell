#Write-Host "started"
using namespace System.Management.Automation
    New-Alias z Get-Help -ErrorAction SilentlyContinue
function AddWrapper([parameter(mandatory=$true, position=0)][string]$For,[parameter(mandatory=$true, position=1)][string]$To) 
{
    $paramDictionary = [RuntimeDefinedParameterDictionary]::new()
    $paramset= $(Get-Command $For).Parameters.Values | %{[System.Management.Automation.RuntimeDefinedParameter]::new($_.Name,$_.ParameterType,$_.Attributes)}
    $paramsetlet= $(Get-Command empt).Parameters.Keys 
    $paramsetlet+= $(Get-Command $To).ScriptBlock.Ast.Body.ParamBlock.Parameters.Name | %{ $_.VariablePath.UserPath }
    $paramset | %{ if ( -not ($paramsetlet -contains $_.Name) ) {$paramDictionary.Add($_.Name,$_)}}
    return $paramDictionary
}
function GetRestOfParams()
{
    #if dontincludecommon provide source function else dst function
    Param([parameter(mandatory=$true, position=1)][hashtable]$params, 
    [parameter(mandatory=$true, position=0)][string]$dstsource,
    [parameter(mandatory=$false, position=2)][switch][bool]$dontincludecommon=$true)
    $dstorgparams=$(Get-Command $dstsource).Parameters.Keys
    $z= $params
    if ( -not $dontincludecommon)
    {
    $z.Keys | %{ if ( -not ($dstorgparams -contains $_) ) {$z.Remove($_)} } | Out-Null
    }
    else 
    { 
        $dyn= $(Get-Command $dstsource).Parameters.Values | Where-Object -Property IsDynamic -Eq $false
        $dyn | %{ $z.Remove($_.Name) } | Out-Null
    }
    return $z
}

function Empt
{
    [CmdletBinding()]
        Param([parameter(mandatory=$true, position=0)][string]$aaaa)
    1
}

function Let
{
    [CmdletBinding()]
Param([parameter(mandatory=$true, position=0)][string]$Option,[parameter(mandatory=$false, position=0)][string]$OptionB)


DynamicParam {
    AddWrapper -For Get -To $MyInvocation.MyCommand.Name
}
Begin { 
   $params = GetRestOfParams Let $PSBoundParameters -dontincludecommon
}
Process {
    Get @params -OptionB ( $OptionB + "1"
            )}
}

function Get
{
    [CmdLetBinding()]
        Param([parameter(mandatory=$false, position=0)][string]$OptionA,
[parameter(mandatory=$false, position=1)][string]$OptionB)
Write-Host "opta",$OptionA
Write-Host "optb",$OptionB
}
Function Term($Proc,$cmd)
{
(Get-Process -Name $Proc) | Where-Object CommandLine -like $cmd | %{Get-CimInstance Win32_Process -Filter ("ProcessId = {0}" -f ($_.Id)) } | %{ Invoke-CimMethod -InputObject $_ -MethodName Terminate }
}

Function KillAllPyCharm()
{
    Term python *pydevd*
    Term python *ibsrv*
    Term cmd *ibsrv*
}
Function EditInNeo($ar)
{
Write-Host nvr --remote $ar --servername  $(Get-Content C:\temp\listen.txt)
nvr --remote $ar --servername  $(Get-Content C:\temp\listen.txt)
if ($LASTEXITCODE -eq 1) {&"C:\Users\ekarni\Neovim\bin\nvim-qt.exe" $ar }
Show-Window nvim-qt
}


Function ResetNeo($a)
{
    DelProcess nvim-qt
    if ($a){
    Start-Process "C:\Users\ekarni\Neovim\bin\nvim-qt.exe" -ArgumentList ($a)
    }
    else{ Start-Process "C:\Users\ekarni\Neovim\bin\nvim-qt.exe"}

 #ps | Where-Object -Property ProcessName  -Like "*goneovim*"| %{Write-Host $_.Id ,$_.ProcessName ;$_.Kill()}
 #C:\Users\ekarni\Downloads\Goneovim-v0.4.12-win64\goneovim.exe
}

Function DelProcess($name)
{
ps | Where-Object -Property ProcessName  -Like "*$name*"| %{Write-Host $_.Id ,$_.ProcessName ;$_.Kill()}
}
function TranslatePath($fil)
{
	wsl bash -c "wslpath -w '$fil'"
}
function RunBash($fil)
{
wsl bash -c "source /home/ekarni/.bash_profile; $fil" 
}

function Show-Window {
  param(
    [Parameter(Mandatory)]
    [string] $ProcessName
  )

  # As a courtesy, strip '.exe' from the name, if present.
  $ProcessName = $ProcessName -replace '\.exe$'

  # Get the PID of the first instance of a process with the given name
  # that has a non-empty window title.
  # NOTE: If multiple instances have visible windows, it is undefined
  #       which one is returned.
  $hWnd = (Get-Process -ErrorAction Ignore $ProcessName).Where({ $_.MainWindowTitle }, 'First').MainWindowHandle

  if (-not $hWnd) { Throw "No $ProcessName process with a non-empty window title found." }

  $type = Add-Type -PassThru -NameSpace Util -Name SetFgWin -MemberDefinition @'
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool IsIconic(IntPtr hWnd);    // Is the window minimized?
'@ 

  # Note: 
  #  * This can still fail, because the window could have bee closed since
  #    the title was obtained.
  #  * If the target window is currently minimized, it gets the *focus*, but its
  #    *not restored*.
  $null = $type::SetForegroundWindow($hWnd)
  # If the window is minimized, restore it.
  # Note: We don't call ShowWindow() *unconditionally*, because doing so would
  #       restore a currently *maximized* window instead of activating it in its current state.
  if ($type::IsIconic($hwnd)) {
    $type::ShowWindow($hwnd, 9) # SW_RESTORE
  }

}
