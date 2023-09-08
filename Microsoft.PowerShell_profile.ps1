
using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms

. $PSScriptRoot\secret.ps1
#Write-Host "started"

New-Alias z Get-Help -ErrorAction SilentlyContinue
# Remove the default cd alias
Remove-Alias cd
# Create a new cd function
function MyCD {
    Set-Location @args
    $curtime =$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $dict = @{
        Id = "30"
        CommandLine = "cd $(Get-Location)"
        ExecutionStatus = "Completed"
        StartExecutionTime = $curtime
        EndExecutionTime = $curtime
        Duration = "00:00:00.0389011"
    }
    $historyObject = New-Object -TypeName PSObject -Property $dict
    Add-History -InputObject $historyObject
}
# Set cd to use the new function
Set-Alias cd MyCD
function SimpHistEx
{
     $va=$(SimpHist)
     [System.Windows.Forms.SendKeys]::SendWait($va)


}
function SimpHist 
{
    $historyLocation = $(Get-PSReadLineOption).HistorySavePath
    $all = Get-Content $historyLocation
    return $($all | Sort-Object -Unique | FZF)
}
# Function to get history of saved locations
function StupidHist {
 $historyLocation = $(Get-PSReadLineOption).HistorySavePath
 $all = Get-Content $historyLocation | select-string -Pattern "^cd .:" | %{ echo ($_ -replace "^cd (.*)","`$1") } | Sort-Object -Unique 
 return $all | Where-Object { Test-Path $($_) }
}
# Function to change to the last visited location
function CdLast {
    $location = StupidHist | FZF
    if ($location) {
        Set-Location $location
    }
}
# Create an alias for CdLast
Set-Alias q CdLast
function ConVM
{
$Username = "User"
$Password = ConvertTo-SecureString "Password1" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
$Session = New-PSSession -VMName win10 -Credential $Credential
return $Session 
}

function ClearShada
{
    rm C:\Users\ekarni\AppData\Local\nvim-data\shada\*
    ResetNeo
}
function Which($arg)
{
    python -c "import shutil; print(shutil.which('$arg'))"
}
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
function OtherPython($a)
{
    Invoke-expression "C:\users\ekarni\AppData\Local\Programs\Python\Python39\python.exe $a"
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
Function Get-LockingProcess {

[cmdletbinding()]
Param(
    [Parameter(Position=0, Mandatory=$True,
    HelpMessage="What is the path or filename? You can enter a partial name without wildcards")]
    [Alias("name")]
    [ValidateNotNullorEmpty()]
    [string]$Path
)

# Define the path to Handle.exe
# //$Handle = "G:\Sysinternals\handle.exe"
$Handle = "C:\SysinternalsSuite\handle.exe"

# //[regex]$matchPattern = "(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\b(\d+)\b)\s+type:\s+(?<Type>\w+)\s+\w+:\s+(?<Path>.*)"
# //[regex]$matchPattern = "(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\d+)\s+type:\s+(?<Type>\w+)\s+\w+:\s+(?<Path>.*)"
# (?m) for multiline matching.
# It must be . (not \.) for user group.
[regex]$matchPattern = "(?m)^(?<Name>\w+\.\w+)\s+pid:\s+(?<PID>\d+)\s+type:\s+(?<Type>\w+)\s+(?<User>.+)\s+\w+:\s+(?<Path>.*)$"

# skip processing banner
$data = &$handle -u $path -nobanner
# join output for multi-line matching
$data = $data -join "`n"
$MyMatches = $matchPattern.Matches( $data )

# //if ($MyMatches.value) {
if ($MyMatches.count) {

    $MyMatches | foreach {
        [pscustomobject]@{
            FullName = $_.groups["Name"].value
            Name = $_.groups["Name"].value.split(".")[0]
            ID = $_.groups["PID"].value
            Type = $_.groups["Type"].value
            User = $_.groups["User"].value.trim()
            Path = $_.groups["Path"].value
            toString = "pid: $($_.groups["PID"].value), user: $($_.groups["User"].value), image: $($_.groups["Name"].value)"
        } #hashtable
    } #foreach
} #if data
else {
    Write-Warning "No matching handles found"
}
} #end function
function copy-foldertovirtualmachine {
    param(
            [parameter (mandatory = $true, valuefrompipeline = $true)]
            [string]$VMName,
            [string]$FromFolder = '.\'
         )
        foreach ($File in (Get-ChildItem $Folder -recurse | ? Mode -ne 'd-----')){

              $relativePath = $item.FullName.Substring($Root.Length)
            Copy-VMFile -VM (Get-VM $VMName) -SourcePath $file.fullname -DestinationPath $file.fullname -FileSource Host -CreateFullPath -Force}
}

function NewVMDrive
{
$Username = "user"
$Password = ConvertTo-SecureString "Password1" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
New-PSDrive -Name "V" -PSProvider "FileSystem" -Root "\\192.168.10.2\c$" -Credential $cred -Persist 
}
function GetGitStash
{
     git stash list | ss mychanges | %{ $_ -replace ":.*$"} | %{ git diff $_^1 $_}
}
function SquashCommits([int]$count) {
$commitHashes = git log --pretty=format:%h -n $count    

$commands= ( 0..$($count-2) ) |  %{   "sed -i 's/^pick $($commitHashes[$_])/squash $($commitHashes[$_])/' `$file"    }
$st= $commands -join "`n"

$st="func() { 
local file=`$1
$st 
}; func"
$env:GIT_SEQUENCE_EDITOR=$st
try{
        git rebase -i HEAD~$count
    }finally
    {
        Remove-Item Env:\GIT_SEQUENCE_EDITOR
    }
}

function RemoveCommit([string]$commit) {
$commitid=git log --pretty="%h" --grep=$commit


$st= "sed -i 's/^pick $($commitid)/drop $($commitid)/' `$file"
$st= $commands -join "`n"

$st="func() {
local file=`$1
$st
}; func"
$env:GIT_SEQUENCE_EDITOR=$st
try{
        git rebase -i HEAD~$count
    }finally
    {
        Remove-Item Env:\GIT_SEQUENCE_EDITOR
    }
}


