
#using namespace System.Management.Automation
#try{ 
#Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
#} catch {}
#Set-PSReadlineOption -AddToHistoryHandler
#  $x | where { $y=cat $_ | ss  -Pattern "\bcall" | ss "\bput" ; $y.Length -ge 1 }
#Write-Host "started"
function cc 
{
    claude --continue
}
New-Alias ss Select-String
New-Alias z Get-Help -ErrorAction SilentlyContinue
New-Alias m Get-Member
New-Alias P pwsh
New-Alias gitp GitPullKeepLocal
New-Alias cl claude

# Remove the default cd alias
# Create a new cd function
#
#
function ProfileCommands {
<#
.SYNOPSIS
Lists all functions in the common module with their synopsis and description.
#>
    #$ErrorActionPreference=SilentlyContinue
    get-command -module common | %{ get-help $_  } | Format-Table Name, SYNOPSIS ,DESCRIPTION
}

function Checkout-FileWithDifferentName {
<#
.SYNOPSIS
Checks out a file from a git branch and saves it under a new name.
.DESCRIPTION
Stashes local changes to the file, checks it out from the specified branch, renames it to the new name, then restores the stash.
.PARAMETER FilePath
Path to the file to check out.
.PARAMETER NewFileName
New name/path for the checked-out file.
.PARAMETER Branch
Git branch to check out the file from. Defaults to 'main'.
#>
    param (
        [string]$FilePath,
        [string]$NewFileName,
        [string]$Branch = "main"
    )
    # Check if the file exists in the current directory
    if (-Not (Test-Path $FilePath)) {
        Write-Error "File '$FilePath' does not exist in the current directory."
        return
    }
    # Get the directory and file name from the file path
    $directory = Split-Path $FilePath
    $fileName = Split-Path $FilePath -Leaf
    # Change to the directory containing the file
    Push-Location $directory
    try {
        # Stash any local changes to the file
        git stash push $fileName
        # Checkout the file from the specified branch
        git checkout $Branch -- $fileName
        # Rename the checked-out file
        mv  $fileName  $NewFileName
        # Restore the stashed changes
        git stash pop
    }
    catch {
        Write-Error "An error occurred: $_"
    }
    finally {
        # Return to the original directory
        Pop-Location
    }
}


function ConvertPSObjectToHashtable
{
<#
.SYNOPSIS
Recursively converts a PSObject to a hashtable.
.DESCRIPTION
Handles nested objects and arrays. Useful for working with JSON data (from ConvertFrom-Json) that needs to be mutable or key-accessible as a hashtable.
.PARAMETER InputObject
The PSObject, array, or scalar value to convert. Accepts pipeline input.
#>
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process
    {
        if ($null -eq $InputObject)
        { return $null 
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            $collection = @(
                foreach ($object in $InputObject)
                { ConvertPSObjectToHashtable $object 
                }
            )

            Write-Output -NoEnumerate $collection
        } elseif ($InputObject -is [psobject])
        {
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties)
            {
                $hash[$property.Name] = (ConvertPSObjectToHashtable $property.Value).PSObject.BaseObject
            }

            $hash
        } else
        {
            $InputObject
        }
    }
}

$global:jsonFile = Join-Path -Path $env:USERPROFILE -ChildPath ('cmdLines.json' )
# Shared shell-state file consumed by window_switcher: map PID -> {title, cwd, time, processid, command}.
# Concurrent shells coordinate via a named mutex.
$global:wsStateFile = 'C:\temp\wt_state.json'

$ExecutionContext.InvokeCommand.PostCommandLookupAction = {
try{
    $cmdLine = $MyInvocation.Line
    if ($args[1].CommandOrigin -ne 'Runspace' -or $cmdLine -match 'PostCommandLookupAction|^prompt$')
    { return
    }

    $currentDir = (Get-Location).Path

    if (!(Test-Path -Path $global:jsonFile) -or (Get-Item $global:jsonFile).Length -eq 0)
    {
        @{ $currentDir = @($cmdLine) } | ConvertTo-Json | Set-Content -Path $global:jsonFile
    } else
    {
        $existingCmdLines = Get-Content -Path $global:jsonFile -Raw | ConvertFrom-Json
        $existingCmdLines = ConvertPSObjectToHashtable $existingCmdLines
        if ($null -eq $existingCmdLines) { $existingCmdLines = @{} }

        if (!$existingCmdLines.ContainsKey($currentDir))
        {
            $existingCmdLines.Add($currentDir, @($cmdLine))
        } else
        {
            if (!$existingCmdLines[$currentDir].Contains($cmdLine))
            {
                $existingCmdLines[$currentDir] += $cmdLine
            }
        }
        $existingCmdLines | ConvertTo-Json | Set-Content -Path $global:jsonFile
    }

    # window_switcher state export ---
    $entry = [ordered]@{
        title     = $Host.UI.RawUI.WindowTitle
        cwd       = $currentDir
        time      = [DateTimeOffset]::Now.ToUnixTimeSeconds()
        processid = $PID
        command   = $cmdLine
    }
    $mutex = New-Object System.Threading.Mutex($false, 'Global\WindowSwitcherStateMutex')
    try {
        [void]$mutex.WaitOne(1500)
        $state = @{}
        if (Test-Path -LiteralPath $global:wsStateFile) {
            try {
                $raw = Get-Content -Raw -LiteralPath $global:wsStateFile -ErrorAction Stop
                if ($raw) {
                    $obj = $raw | ConvertFrom-Json -ErrorAction Stop
                    $state = ConvertPSObjectToHashtable $obj
                    if ($null -eq $state) { $state = @{} }
                }
            } catch { $state = @{} }
        }
        foreach ($k in @($state.Keys)) {
            if (-not (Get-Process -Id ([int]$k) -ErrorAction SilentlyContinue)) {
                $state.Remove($k)
            }
        }
        $state["$PID"] = $entry
        $state | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $global:wsStateFile -Encoding UTF8
    } finally {
        [void]$mutex.ReleaseMutex()
        $mutex.Dispose()
    }
    }catch {
Write-Debug "error in PostCommandLookupAction: $_"
    }
}
$parameters = @{
    Key = 'Alt+q'
    BriefDescription = 'Go to last dir'
    LongDescription = 'Go to last dir'
    ScriptBlock = {
        param($key, $arg)   # The arguments are ignored in this example
        CdLast 
    }
}
Set-PSReadLineKeyHandler @parameters
$parameters = @{
    Key = 'Alt+e'
    BriefDescription = 'Execute from last same direrctory'
    LongDescription = 'Execute from last commands typed in same direrctory'
    ScriptBlock = {
        param($key, $arg)   # The arguments are ignored in this example
        [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert( $(GrepOnCurDir) )
        #[Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()

    }
}
Set-PSReadLineKeyHandler @parameters
$parameters = @{
    Key = 'Alt+h'
    BriefDescription = 'Grep from last same direrctory'
    LongDescription = 'Grep from last commands typed globally'
    ScriptBlock = {
        param($key, $arg)   # The arguments are ignored in this example
        [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert( $(SimpHist) )
    }
}
Set-PSReadLineKeyHandler @parameters
$parameters = @{
    Key = 'Alt+c'
    BriefDescription = 'Open claude in last dir'
    LongDescription = 'Pick a previously visited directory via fzf and open claude there'
    ScriptBlock = {
        param($key, $arg)
        [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert('ClaudeLast')
        [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
    }
}
Set-PSReadLineKeyHandler @parameters



function GrepOnCurDir()
{
<#
.SYNOPSIS
Interactively select a command previously run in the current directory.
.DESCRIPTION
Reads the per-directory command history JSON file and presents matching commands for the current directory via fzf.
#>
    $currentDir = (Get-Location).Path
    if (!(Test-Path -Path $global:jsonFile) -or (Get-Item $global:jsonFile).Length -eq 0) { return '' }
    $existingCmdLines = Get-Content -Path $global:jsonFile -Raw | ConvertFrom-Json
    $existingCmdLines = ConvertPSObjectToHashtable $existingCmdLines
    if ($null -eq $existingCmdLines -or -not $existingCmdLines.ContainsKey($currentDir)) { return '' }
    $existingCmdLines[$currentDir] | fzf
}
function MyCD
{
<#
.SYNOPSIS
Custom cd that records the destination in PSReadLine history.
.DESCRIPTION
Wraps Set-Location and appends a 'cd <path>' entry to the PSReadLine history file so directory navigation is searchable via SimpHist/CdLast.
#>
    try{
        Set-Location @args
    }catch{ 
        Write-Error "asdas"
        Write-Error $_.Exception.InnerException.Message
        return
    }
    #$curtime =$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    #$dict = @{
    #Id = "30"
    #CommandLine = "cd $(Get-Location)"
    #ExecutionStatus = "Completed"
    #StartExecutionTime = $curtime
    #EndExecutionTime = $curtime
    #Duration = "00:00:00.0389011"
    #}
    #$historyObject = New-Object -TypeName PSObject -Property $dict
    #Add-History -InputObject $historyObject
    try{ 
    $historyLocation = $(Get-PSReadLineOption).HistorySavePath

    Add-Content -Path $historyLocation -Value "cd $(Get-Location)"
    }catch {
        Write-Debug "error in MyCD: $_"
    }
}
# Set cd to use the new function
if ($PSVersionTable.PSVersion -like "7*")
    {
Set-Alias cd MyCD
    }
function SimpHistEx
{
<#
.SYNOPSIS
Fuzzy-search command history and immediately execute the selected entry.
.DESCRIPTION
Calls SimpHist to pick a history entry via fzf, inserts it on the command line, and executes it.
#>
    $va=$(SimpHist)
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert( $va )
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()

    #[System.Windows.Forms.SendKeys]::SendWait($va)


}
function SimpHist
{
<#
.SYNOPSIS
Fuzzy-search the full command history and return the selected entry.
.DESCRIPTION
Reads the PSReadLine history file, deduplicates entries, and presents them via fzf for interactive selection.
#>
    $historyLocation = $(Get-PSReadLineOption).HistorySavePath
    $all = Get-Content $historyLocation
    return $($all | Sort-Object -Unique | FZF)
}
# Function to get history of saved locations
function StupidHist
{
<#
.SYNOPSIS
Returns a list of previously visited directories from command history.
.DESCRIPTION
Extracts 'cd' entries from the PSReadLine history file and filters to only those that currently exist on disk.
#>
    $historyLocation = $(Get-PSReadLineOption).HistorySavePath
    $all = Get-Content $historyLocation | select-string -Pattern "^cd .:" | %{ echo ($_ -replace "^cd (.*)","`$1") } | Sort-Object -Unique 
    return $all | Where-Object { Test-Path $($_) }
}
# Function to change to the last visited location
function CdLast
{
<#
.SYNOPSIS
Jump to a previously visited directory using fzf.
.DESCRIPTION
Presents the list of known visited directories (from StupidHist) via fzf and navigates to the selected one. Also aliased as 'q'.
#>
    $location = StupidHist | FZF
    if ($location)
    {
        Set-Location $location
    }
}
# Create an alias for CdLast
Set-Alias q CdLast
function ClaudeLast
{
<#
.SYNOPSIS
Pick a previously visited directory via fzf and open claude there.
.DESCRIPTION
Uses StupidHist to list visited directories, presents them via fzf, prompts for confirmation, then runs claude.exe in the selected directory.
#>
    $location = StupidHist | FZF
    if (-not $location) { return }


    cd  $location
    try { claude.exe @args } finally {  }
}
function ConVM
{
<#
.SYNOPSIS
Opens a PSSession to the local 'win10' Hyper-V virtual machine.
.DESCRIPTION
Creates a PowerShell remoting session to the VM named 'win10' using default credentials and returns the session object.
#>
    $Username = "User"
    $Password = ConvertTo-SecureString "Password1" -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
    $Session = New-PSSession -VMName win10 -Credential $Credential
    return $Session 
}

function ClearShada
{
<#
.SYNOPSIS
Clears neovim shada (shared data) files and restarts neovim-qt.
.DESCRIPTION
Deletes all files in the nvim-data shada directory, then calls ResetNeo to restart the editor.
#>
    rm ~\AppData\Local\nvim-data\shada\*
    ResetNeo
}
function Which($arg)
{
<#
.SYNOPSIS
Finds the full path of an executable, similar to Unix 'which'.
.PARAMETER arg
The executable name to look up.
#>
    python -c "import shutil; print(shutil.which('$arg'))"
}
Function Get-ProcessCwd {
<#
.SYNOPSIS
Reads the current working directory of a process by reading its PEB. 64-bit only.
#>
    param([Parameter(Mandatory)][int]$Id)
    if (-not ('Util.PebReader' -as [type])) {
        Add-Type -NameSpace Util -Name PebReader -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError=true)]
public static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
[DllImport("kernel32.dll", SetLastError=true)]
public static extern bool CloseHandle(IntPtr h);
[DllImport("kernel32.dll", SetLastError=true)]
public static extern bool ReadProcessMemory(IntPtr h, IntPtr addr, byte[] buf, IntPtr size, out IntPtr read);
[DllImport("ntdll.dll")]
public static extern int NtQueryInformationProcess(IntPtr h, int infoClass, IntPtr buf, int len, out int ret);
'@
    }
    $h = [Util.PebReader]::OpenProcess(0x1010, $false, $Id)
    if ($h -eq [IntPtr]::Zero) { return $null }
    try {
        $pbi = [System.Runtime.InteropServices.Marshal]::AllocHGlobal(48)
        try {
            $rl = 0
            if ([Util.PebReader]::NtQueryInformationProcess($h, 0, $pbi, 48, [ref]$rl) -ne 0) { return $null }
            $peb = [System.Runtime.InteropServices.Marshal]::ReadIntPtr($pbi, 8)
        } finally { [System.Runtime.InteropServices.Marshal]::FreeHGlobal($pbi) }
        $buf = New-Object byte[] 8; $r = [IntPtr]::Zero
        if (-not [Util.PebReader]::ReadProcessMemory($h, [IntPtr]([long]$peb + 0x20), $buf, [IntPtr]8, [ref]$r)) { return $null }
        $procParams = [IntPtr][BitConverter]::ToInt64($buf, 0)
        $us = New-Object byte[] 16
        if (-not [Util.PebReader]::ReadProcessMemory($h, [IntPtr]([long]$procParams + 0x38), $us, [IntPtr]16, [ref]$r)) { return $null }
        $len = [BitConverter]::ToUInt16($us, 0)
        if ($len -eq 0) { return '' }
        $bufPtr = [IntPtr][BitConverter]::ToInt64($us, 8)
        $str = New-Object byte[] $len
        if (-not [Util.PebReader]::ReadProcessMemory($h, $bufPtr, $str, [IntPtr]$len, [ref]$r)) { return $null }
        ([System.Text.Encoding]::Unicode.GetString($str)).TrimEnd('\')
    } finally { [Util.PebReader]::CloseHandle($h) | Out-Null }
}

Function LookFor {
<#
.SYNOPSIS
Find running processes by name, command line, or window title.
.DESCRIPTION
Filters Get-Process results using wildcard patterns. Useful for locating specific process instances by their command line or window title.
.PARAMETER Proc
Process name pattern (wildcard). Defaults to '*' (all processes).
.PARAMETER cmd
Command line pattern (wildcard). Defaults to '*'.
.PARAMETER Title
Main window title pattern (wildcard). Defaults to '*'.
.PARAMETER ShowTable
If specified, returns the raw Process objects instead of the default summary (Id, Name, PrivateMB, CPU(s), MainWindowTitle, CommandLine, Cwd).
.PARAMETER Tree
If specified, also includes descendant processes (recursively) of each match, indented under their parent.
.PARAMETER ParentId
If specified (non-zero), only include processes whose parent process id equals this value.
#>
    param(
        [Parameter(Position=0)]
        [string]$Proc = "*",

        [Parameter(Position=1)]
        [string]$cmd = "*",

        [Parameter()]
        [string]$Title = "*",

        [Parameter()]
        [switch]$ShowTable,

        [Parameter()]
        [switch]$Tree,

        [Parameter()]
        [switch]$Object,

        [Parameter()]
        [int]$ParentId = 0
    )

    $allCim = Get-CimInstance Win32_Process
    $ppidMap = @{}
    $pnameMap = @{}
    $byParent = @{}
    foreach ($p in $allCim) {
        $ppidMap[[int]$p.ProcessId] = [int]$p.ParentProcessId
        $pnameMap[[int]$p.ProcessId] = [string]$p.Name
        if (-not $byParent.ContainsKey([int]$p.ParentProcessId)) {
            $byParent[[int]$p.ParentProcessId] = @()
        }
        $byParent[[int]$p.ParentProcessId] += $p
    }

    $processes = Get-Process | Where-Object {
        $_.Name -like $Proc -and
        $_.CommandLine -like $cmd -and
        ($_.MainWindowTitle -like $Title) -and
        ($ParentId -eq 0 -or $ppidMap[[int]$_.Id] -eq $ParentId)
    }

    if ($Tree) {
        $allProc = @{}
        Get-Process | ForEach-Object { $allProc[[int]$_.Id] = $_ }

        $rows = New-Object System.Collections.Generic.List[object]
        $seen = @{}
        $emit = {
            param($pid_, $depth)
            if ($seen.ContainsKey($pid_)) { return }
            $seen[$pid_] = $true
            $proc = $allProc[$pid_]
            if ($proc) {
                $rows.Add([pscustomobject]@{
                    Id              = $proc.Id
                    PPID            = $ppidMap[[int]$proc.Id]
                    PName           = $pnameMap[[int]$ppidMap[[int]$proc.Id]]
                    Name            = ('  ' * $depth) + $proc.Name
                    PrivateMB       = [math]::Round($proc.PrivateMemorySize64 / 1MB, 1)
                    'CPU(s)'        = if ($proc.CPU) { [math]::Round($proc.CPU, 1) } else { 0 }
                    MainWindowTitle = $proc.MainWindowTitle
                    CommandLine     = $proc.CommandLine
                    Cwd             = Get-ProcessCwd -Id $proc.Id
                })
            }
            if ($byParent.ContainsKey($pid_)) {
                foreach ($child in $byParent[$pid_]) {
                    & $emit ([int]$child.ProcessId) ($depth + 1)
                }
            }
        }
        foreach ($p in $processes) { & $emit ([int]$p.Id) 0 }
        $rows
    }
    elseif ($ShowTable) {
        $processes
    } else {
        $processes | Select-Object Id, `
            @{Name='PPID'; Expression={ $ppidMap[[int]$_.Id] }}, `
            @{Name='PName'; Expression={ $pnameMap[[int]$ppidMap[[int]$_.Id]] }}, `
            Name, `
            @{Name='PrivateMB'; Expression={ [math]::Round($_.PrivateMemorySize64 / 1MB, 1) }}, `
            @{Name='CPU(s)'; Expression={ if ($_.CPU) { [math]::Round($_.CPU, 1) } else { 0 } }}, `
            MainWindowTitle, CommandLine, `
            @{Name='Cwd'; Expression={ Get-ProcessCwd -Id $_.Id }}
    }
}


Function Term($Proc,$cmd="*")
{
<#
.SYNOPSIS
Terminates processes matching a name and optional command line pattern.
.PARAMETER Proc
Wildcard pattern for the process name.
.PARAMETER cmd
Wildcard pattern for the process command line. Defaults to '*'.
#>
(Get-Process) | Where { $_.name -like $Proc}    | Where-Object CommandLine -like $cmd | ForEach-Object{Get-CimInstance Win32_Process -Filter ("ProcessId = {0}" -f ($_.Id)) } | %{ Invoke-CimMethod -InputObject $_ -MethodName Terminate }
}

Function KillAllPyCharm()
{
<#
.SYNOPSIS
Kills all PyCharm-related Python and cmd processes.
.DESCRIPTION
Terminates python processes associated with the PyCharm debugger (pydevd) and ibsrv, as well as any cmd processes running ibsrv.
#>
    Term python *pydevd*
    Term python *ibsrv*
    Term cmd *ibsrv*
}
Function EditInNeo($ar, $line)
{
<#
.SYNOPSIS
Opens a file in the running neovim instance (or starts nvim-qt if not running).
.DESCRIPTION
Uses nvr (neovim-remote) to open the file in an existing neovim server. If nvr fails, falls back to launching nvim-qt directly. Brings the neovim-qt window to the foreground.
.PARAMETER ar
File path to open.
.PARAMETER line
Optional line number to jump to.
#>
    Write-Host $ar $line
    $fileArg = $ar
    if ($line) {
        $lineArg = "+$line"
        Write-Host nvr --remote $lineArg $ar --servername  $(Get-Content C:\temp\listen.txt)
        #cmd /S "pause"
        nvr --remote $lineArg $ar --servername  $(Get-Content C:\temp\listen.txt)
        if ($LASTEXITCODE -eq 1)
        {&$qtpath $lineArg $ar
        }
    } else {
        Write-Host nvr --remote $ar --servername  $(Get-Content C:\temp\listen.txt)
        nvr --remote $ar --servername  $(Get-Content C:\temp\listen.txt)
        if ($LASTEXITCODE -eq 1)
        {&$qtpath $ar
        }
    }
    Show-Window nvim-qt
}



Function DelProcess($name)
{
<#
.SYNOPSIS
Kills all processes whose name contains the given string.
.PARAMETER name
Substring to match against process names.
#>
    ps | Where-Object -Property ProcessName  -Like "*$name*"| %{Write-Host $_.Id ,$_.ProcessName ;$_.Kill()}
}
function TranslatePath($fil)
{
<#
.SYNOPSIS
Converts a WSL/Linux path to a Windows path using wslpath.
.PARAMETER fil
The WSL path to convert.
#>
    wsl bash -c "wslpath -w '$fil'"
}
function RunInBash
{
<#
.SYNOPSIS
Sources ~/.bash_profile in WSL and runs the given command.
.DESCRIPTION
Joins all arguments into a single bash command line, sources /home/ekarni/.bash_profile, then executes the command in WSL bash.
#>
    $cmd = $args -join ' '
$c= @" 
wsl.exe bash -c `"source /home/ekarni/.bash_profile && $cmd`"
"@
Write-Host $c
    iex $c
}
function Show-Window
{
<#
.SYNOPSIS
Brings a process window to the foreground, restoring it if minimized.
.DESCRIPTION
Uses Win32 API (SetForegroundWindow / ShowWindow) to activate the main window of the named process. Strips the .exe extension automatically.
.PARAMETER ProcessName
Name of the process whose window should be brought to focus.
#>
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

    if (-not $hWnd)
    { Throw "No $ProcessName process with a non-empty window title found." 
    }

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
    if ($type::IsIconic($hwnd))
    {
        $type::ShowWindow($hwnd, 9) # SW_RESTORE
    }

}
Function Get-LockingProcess
{
<#
.SYNOPSIS
Finds processes that have a lock on a given file or path.
.DESCRIPTION
Uses Sysinternals handle.exe to enumerate open handles and returns process info (name, PID, type, user, path) for any process locking the specified path.
.PARAMETER Path
The file path or partial name to check for locking processes.
#>
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
    $Handle = "C:\SysinternalsSuite\handle64.exe"

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
    if ($MyMatches.count)
    {

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
    else
    {
        Write-Warning "No matching handles found"
    }
} #end function
function copy-foldertovirtualmachine
{
<#
.SYNOPSIS
Copies all files in a folder to a Hyper-V virtual machine.
.PARAMETER VMName
Name of the Hyper-V VM to copy files to.
.PARAMETER FromFolder
Source folder path. Defaults to the current directory.
#>
    param(
        [parameter (mandatory = $true, valuefrompipeline = $true)]
        [string]$VMName,
        [string]$FromFolder = '.\'
    )
    foreach ($File in (Get-ChildItem $Folder -recurse | ? Mode -ne 'd-----'))
    {

        $relativePath = $item.FullName.Substring($Root.Length)
        Copy-VMFile -VM (Get-VM $VMName) -SourcePath $file.fullname -DestinationPath $file.fullname -FileSource Host -CreateFullPath -Force
    }
}

function NewVMDrive
{
<#
.SYNOPSIS
Maps the VM's C: drive as a persistent network drive (V:).
.DESCRIPTION
Creates a PSDrive pointing to \\192.168.10.2\c$ using stored credentials, making the VM's file system directly accessible from the host.
#>
    $Username = "user"
    $Password = ConvertTo-SecureString "Password1" -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)
    New-PSDrive -Name "V" -PSProvider "FileSystem" -Root "\\192.168.10.2\c$" -Credential $cred -Persist 
}
function GetGitStash
{
<#
.SYNOPSIS
Shows the diff of stash entries named 'mychanges'.
.DESCRIPTION
Lists all stash entries, filters by the name 'mychanges', and shows the diff of each matching stash against its parent commit.
#>
    git stash list | ss mychanges | %{ $_ -replace ":.*$"} | %{ git diff $_^1 $_}
}
function  CheckCommit ($n,$line)
{
<#
.SYNOPSIS
Searches the last N commits for a specific string.
.PARAMETER n
Number of recent commits to search.
.PARAMETER line
String pattern to search for in each commit's diff.
#>
    $commits= git log --pretty=format:%h -n $n
    $commits | %{ git show $_ | select-string $line} 
}

function RemoveCommit([string]$commit)
{
<#
.SYNOPSIS
Removes a commit from history using interactive rebase (drop).
.DESCRIPTION
Looks up the commit hash by its message, then sets GIT_SEQUENCE_EDITOR to a sed command that marks it as 'drop' in the rebase todo list.
.PARAMETER commit
Commit message (or part of it) to search for and remove.
#>
    $commitid=git log --pretty="%h" --grep=$commit


    $st= "sed -i 's/^pick $($commitid)/drop $($commitid)/' `$file"
    $st= $commands -join "`n"

    $st="func() {
local file=`$1
$st
}; func"
    $env:GIT_SEQUENCE_EDITOR=$st
    try
    {
        git rebase -i HEAD~$count
    } finally
    {
        Remove-Item Env:\GIT_SEQUENCE_EDITOR
    }
}
function ExtractFromLastStash($file)
{
<#
.SYNOPSIS
Gets the diff of a specific file from the most recent stash.
.PARAMETER file
Path of the file to extract the diff for.
#>
    $x=git diff stash@`{0`}^1 stash@`{0`} -- $file 
    return $x
}

function Checkout-FileFromStash {
<#
.SYNOPSIS
Interactively checks out a single file from a chosen stash entry.
.DESCRIPTION
Lists available stashes, prompts the user to pick one by index (if not provided), then runs 'git checkout stash@{N} -- FilePath'.
.PARAMETER FilePath
The file path to restore from the stash.
.PARAMETER StashIndex
Index of the stash to use (e.g. 0 for stash@{0}). If omitted, the user is prompted.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $false)]
        [int]$StashIndex
    )
    # Get the list of stashes
    $stashes = git stash list
    if ($stashes.Count -eq 0) {
        Write-Host "No stashes found."
        return
    }
    # Display the list of stashes
    Write-Host "Available stashes:"
    $stashes | ForEach-Object { Write-Host $_ }
    # If StashIndex is not provided, prompt the user to select a stash
    if (-not $PSBoundParameters.ContainsKey('StashIndex')) {
        $StashIndex = Read-Host "Enter the index of the stash you want to use (e.g., 0 for stash@{0})"
        # Validate the user's input
        if (-not $StashIndex -match '^\d+$') {
            Write-Error "Invalid input. Please enter a valid stash index."
            return
        }
    }
    # Checkout the specified file from the selected stash
    try {
        git checkout stash@{$StashIndex} -- $FilePath
        Write-Host "File '$FilePath' has been checked out from stash@{$StashIndex}."
    } catch {
        Write-Error "An error occurred while checking out the file: $_"
    }
}
function StashAll($name)
{
<#
.SYNOPSIS
Creates a named stash without modifying the working tree (silent stash).
.DESCRIPTION
Runs 'git stash create' to create a stash object and immediately stores it with a name via 'git stash store', leaving the working tree unchanged.
.PARAMETER name
The name/message to assign to the stash entry.
#>
    git stash store $(git stash create) -m $name
}
#function GitPullKeepLocal ()
#{
    #param ( 
        #[parameter()][switch]$keeplocalinconflict =$null,
        #[parameter()][switch]$dontkeepstash=$false 

        #) 

    #$commit_hash=$(git rev-parse HEAD)
    #git stash save | Out-Null
    #git pull --rebase 
    #$conflicts = $(git diff --name-only --diff-filter=U)
    #$changes = $(git diff --name-only $commit_hash)
    #if ($conflicts)
    #{
        #Write-Host "There are merge conflicts. Please run git pull. Aborting"
        ##abort the pull
        #git rebase --abort
        

        ## Exit or throw an error here, if you want to stop the script
    #} else
    #{

        ## Checkout files from the stash
        #git checkout stash -- . | Out-Null
        #git reset | Out-Null
        #$localch= $(git diff --name-only)
        #$int = $localch | ?{ $changes -contains $_  } 
        #if ($int)
        #{
            #echo "Following files are in both: $int " 

        
            #if ($(-not ($keeplocalinconflict)))
            #{
                #$userInput = Read-Host -Prompt "Do you want to keep local changes in case of conflict? (y/n/merge)"
                #if ($userInput -eq "y") {
                    #$keeplocalinconflict = $true
                #} else {
                    #echo "reseting to remote"
                    #git checkout -- $int
                #}
                #if ($userInput -eq "merge")
                #{
                    #git stash apply
                #}
            #}
        #}

        #if ($dontkeepstash) 
        #{
            #git stash drop
        #}
        ## Drop the stash
    #}
#}

function RestartWsl()
{
<#
.SYNOPSIS
Restarts the WSL (Windows Subsystem for Linux) service.
.DESCRIPTION
Restarts the LxssManager service, which forces a full WSL restart.
#>
    Get-Service LxssManager | Restart-Service

}
function UpdateVim($typ)
{
<#
.SYNOPSIS
Downloads and installs a specified neovim release.
.DESCRIPTION
Downloads nvim-win64.zip for the given release tag from GitHub, backs up the current Neovim installation to nvim-temp, and extracts the new version.
.PARAMETER typ
Release tag to download (e.g. 'nightly', 'v0.9.0').
#>
    cd ~ 
    Write-Host "usage: new-version-zip-filename (ie nightly)"
    Remove-Item -Path nvim-win64.zip -ErrorAction SilentlyContinue
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile("https://github.com/neovim/neovim/releases/download/$typ/nvim-win64.zip", "C:\Users\ekarni\nvim-win64.zip")
    if (Test-Path -Path nvim-temp)
    {
        Write-Host "moving temp to last temp"
        Remove-Item -Path ./neovim-lasttemp -Recurse -Force -ErrorAction SilentlyContinue
        Move-Item -Path nvim-temp -Destination nvim-lasttemp
    }
    #Move-Item -Path nvim-temp -Destination nvim-lasttemp     -ErrorAction SilentlyContinue
    Move-Item -Path ./Neovim -Destination nvim-temp
    Expand-Archive -Path nvim-win64.zip -DestinationPath ./Neovim -Force


}

function Add-ToPath {
<#
.SYNOPSIS
Permanently adds a directory to the machine-level system PATH.
.PARAMETER PathToAdd
The directory path to append to PATH. No-op if already present.
#>
    param (
        [string]$PathToAdd
    )
    # Check if the path already exists in the PATH variable
    $currentPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
    if ($currentPath -like "*$PathToAdd;*") {
        Write-Host "The path '$PathToAdd' is already in the system PATH."
        return
    }
    # Add the new path to the existing PATH variable
    $newPath = $currentPath + ";" + $PathToAdd
    # Update the system PATH variable
    [System.Environment]::SetEnvironmentVariable("Path", $newPath, [System.EnvironmentVariableTarget]::Machine)
    Write-Host "The path '$PathToAdd' has been added to the system PATH."
}
function FindGitFile ($x)
{
<#
.SYNOPSIS
Shows the full git history for a file, following renames.
.PARAMETER x
File path to trace through git history.
#>
    git log --follow -- $x
}

function IIF($condition, $truePart, $falsePart) {
<#
.SYNOPSIS
Ternary-style conditional expression (inline if).
.PARAMETER condition
Boolean condition to evaluate.
.PARAMETER truePart
Value returned when condition is true.
.PARAMETER falsePart
Value returned when condition is false.
#>
    if ($condition) {
        return $truePart
    } else {
        return $falsePart
    }
}

# PowerShell functions for chess analyzer operations


function Ext2 {
<#
.SYNOPSIS
Runs a script block in a new PowerShell window (using 'start' with base64-encoded command).
.DESCRIPTION
Encodes the script block as a base64 EncodedCommand and launches a new pwsh process via 'start'. Supports NoExit and custom working directory.
.PARAMETER ScriptBlock
The script block to execute in the new window.
.PARAMETER WorkingDirectory
Working directory for the new process. Defaults to the current directory.
.PARAMETER NoExit
Keep the new window open after the script completes.
.PARAMETER ArgumentList
Additional arguments to pass to the new PowerShell process.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory=$false)]
        [string]$WorkingDirectory = $PWD.Path,
        
        [Parameter(Mandatory=$false)]
        [switch]$NoExit,
        
        [Parameter(Mandatory=$false)]
        [string[]]$ArgumentList
    )
    
    # Convert the scriptblock to a string
    $scriptString = $ScriptBlock.ToString()
    
    # Convert the script string to base64
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptString)
    $encodedCommand = [Convert]::ToBase64String($bytes)
    
    # Build the PowerShell arguments
    $pwshArgs = @("pwsh","-EncodedCommand", $encodedCommand)
    
    # Add NoExit if specified
    if ($NoExit) {
        $pwshArgs = @("-NoExit") + $pwshArgs
    }
    
    # Add any additional arguments
    if ($ArgumentList) {
        $pwshArgs += $ArgumentList
    }
    
    # Start the new PowerShell process
    $startInfo = @{
        FilePath = "start"
        ArgumentList = $pwshArgs
        WorkingDirectory = $WorkingDirectory
        UseNewEnvironment = $false
    }
    
    Start-Process @startInfo
}
function DebugIt {
<#
.SYNOPSIS
Sets up SSH tunnel and kubectl port-forward for remote debugging.
.DESCRIPTION
Opens an SSH tunnel to the remote server with local port forwarding, then sets up kubectl port-forward from a running pod to localhost for debugger connectivity.
#>
    # SSH with port forwarding
    Write-Host "Setting up SSH tunnel for debugging..." -ForegroundColor Cyan
    Start-Process powershell -ArgumentList "-Command ssh ubuntu@54.228.92.153 -L1234:localhost:8888"
    
    # Get running pod and set up port forwarding
    Write-Host "Setting up Kubernetes port forwarding..." -ForegroundColor Cyan
    $RUN_POD = kubectl get pods | Where-Object { $_ -match "Running" } | ForEach-Object { ($_ -split '\s+')[0] } | Select-Object -First 1
    if ($RUN_POD) {
        Write-Host "Found running pod: $RUN_POD" -ForegroundColor Green
        kubectl port-forward $RUN_POD 8888:1234 --address localhost
    } else {
        Write-Host "No running pods found!" -ForegroundColor Red
    }
}

function Select-Zip {
<#
.SYNOPSIS
Zips two sequences together, similar to Python's zip() or enumerate().
.DESCRIPTION
Uses LINQ Enumerable.Zip to pair elements from two sequences. When only one sequence is provided, it zips with an infinite index (enumerate behavior).
.PARAMETER First
The first sequence to zip.
.PARAMETER Second
The second sequence to zip. Defaults to an integer counter (enumerate mode).
.PARAMETER ResultSelector
A script block that defines how to combine elements. Defaults to returning an array of both elements.
#>
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    $First,

    [Parameter(Position=1)]
    $Second = (0..([int]::MaxValue)),

    [Parameter(Position=2)]
    $ResultSelector = { ,$args }
)

# If Second is not explicitly provided, default to zipping with indices (like Python enumerate)
if ($PSBoundParameters.Count -eq 1) {
    $Second = 0..([int]::MaxValue)
}

[System.Linq.Enumerable]::Zip($First, $Second, [Func[Object, Object, Object[]]]$ResultSelector)
} 

function Tuple-Zip {
<#
.SYNOPSIS
Zips two arrays into an array of typed System.Tuple objects.
.DESCRIPTION
Uses Select-Zip to pair elements and wraps each pair in a System.Tuple, scaling the second array's values by 100.
.PARAMETER Array1
The first array.
.PARAMETER Array2
The second array (values multiplied by 100 in the resulting tuples).
#>
    param (
        [Parameter(Mandatory=$true)]
        [array]$Array1,
        
        [Parameter(Mandatory=$true)]
        [array]$Array2
    )
    $zipped= select-Zip $Array1 $Array2 
    return @($zipped.ForEach({ [System.Tuple]::Create($_[0], $_[1]*100) })) 
}

function Find {
 <# 
 .SYNOPSIS 
 
 Does find files similar to find in bash 
 .PARAMETER name 
     Filters name inside like (*a*) is accepted
 .PARAMETER norecurse 
         Don't do it recursively 
 .PARAMETER justname 
     Return the name and not full path
 .PARAMETER nohidden
     Exclude hidden files and directories
 .PARAMETER exec
  Accept a script block to execute and pass it $_ param
 #> 
 param(
     [Parameter(Mandatory=$true)]
     [string]$path ,
 
     [Parameter(Mandatory=$false)]
     [string]$fullpath ="*",
 
     [Parameter(Mandatory=$false)]
     [string]$name ="*",
 
     [Parameter(Mandatory=$false)]
     [ValidateSet('All', 'File', 'Directory')]
     [string]$type = 'All',
 
     [switch]$norecurse = $false ,
     [switch]$justname = $false ,
     [switch]$retitem = $false ,
     [switch]$nohidden = $false ,
     [switch]$noignoreerr = $false ,
 
 [Parameter(Mandatory=$false)]
     [scriptblock]$exec=$null 
 
 )
     $act= IIF $noignoreerr Continue SilentlyContinue
 
     Get-ChildItem -Path $path -Recurse:$(-not $norecurse) -Force:$(-not $nohidden) -ErrorAction $act |
     Where-Object { $_.Name -like $name } | Where-Object { $_.FullName -like $fullpath } |  ForEach-Object { 
         $x=$_
         $item = switch ($type) {
             'All' { $x }
             'File' { if (!$x.PSIsContainer) { $x } }
             'Directory' { if ($x.PSIsContainer) { $x } }
         }
         
         if ($item) {
             if ($retitem) {
                 $item
             } else {
                 if ($exec) {
                     $exec.InvokeWithContext($null, [psvariable]::new('_', $item))
                 }
                 
                 if ($justname) {
                     $item.Name
                 } else {
                     $item.FullName
                 }
             }
         }
     }
 
 }
 
 function FilesInCommit($cmt) {
 <#
 .SYNOPSIS
 Lists files changed in a specific git commit.
 .PARAMETER cmt
 Commit hash or reference.
 #>
     return $(git diff-tree --no-commit-id --name-only $cmt -r)
 }
 function RunPreCommitOnCommit($cmt) {
 <#
 .SYNOPSIS
 Runs pre-commit hooks against the files changed in a specific commit.
 .PARAMETER cmt
 Commit hash or reference.
 #>
     pre-commit run  --files $(git diff-tree --no-commit-id --name-only $cmt -r)
 }
 function AddOnModified($ext) {
 <#
 .SYNOPSIS
 Stages all modified tracked files, optionally filtered by extension.
 .PARAMETER ext
 Optional file extension pattern to filter (e.g. '.py').
 #>
 
     if ($ext) {
 
             git add $(git ls-files --modified | select-string $ext)
     }
     else {
     git add $(git ls-files --modified)
     }
 }
 
 function GitChangesForFile($fil) {
 <#
 .SYNOPSIS
 Shows all diffs for a file across the full git history (including reflog).
 .PARAMETER fil
 File path to inspect.
 #>
 
     git log --reflog --follow --format=%h -- $fil | %{ Write-host " change $_"; git --no-pager diff $_ -- $fil } 
 }
 function GitLogForFile($fil) {
 <#
 .SYNOPSIS
 Shows commit history for a file or directory across all branches and reflog.
 .PARAMETER fil
 File or directory path to inspect.
 #>
     git log --all --first-parent --remotes --reflog --author-date-order -- $fil
 }
 function StagedFiles() {
     <#
     .SYNOPSIS
     Lists files currently staged for the next commit.
     #>
         git diff --name-only --cached
 }
 function GitPullAdvanced ()
 {
 <#
 .SYNOPSIS
 Git pull that stashes local changes, pulls, then reapplies them with conflict resolution prompts.
 .PARAMETER keeplocalinconflict
 Keep local version when the same file changed both locally and remotely.
 .PARAMETER dontkeepstash
 Drop the stash after reapplying (don't keep it).
 .PARAMETER checkout
 Checkout the branch instead of pulling.
 .PARAMETER useremotefiles
 Silently prefer remote files on conflict instead of prompting.
 .PARAMETER branch
 Remote branch to pull/checkout.
 .PARAMETER repository
 Remote repository name (required when branch is specified).
 #>
     param (
         [parameter()][switch]$keeplocalinconflict =$null,
         [parameter()][switch]$dontkeepstash=$false,
         [parameter()][switch]$checkout=$false,
         [parameter()][switch]$useremotefiles=$false,
         [parameter(Position=0)][string]$branch,
         [parameter()][string]$repository
     )
     function DoPull($x)
     {
         if ($branch )
         {
             if (-not $repository) {Write-Error "no repo";return }
             git pull --rebase $repository $branch @x
         }
         else 
         {
             git pull @x
         }
 
     }
 try { 
     $oldrend= $PSStyle.OutputRendering 
     $PSStyle.OutputRendering = 1
 
 
     $commit_hash=$(git rev-parse HEAD)
     git stash save | Out-Null
     $outcheckout=""
     git fetch 
     if ($checkout)
     {
         $outcheckout=git checkout $branch 2>&1 
         echo $outcheckout
     }else {
         DoPull
     }
     if ($outcheckout -like "*The following untracked working tree files would be overwritten*") {
         $y= $outcheckout | where { $_  -like "`t*" }  | %{ $($_ | Out-string) -replace "`t","" } | %{ $_ -replace "`r`n","" } | %{ $_ -replace "`n","" }
         echo "adding to stash $y"
         git stash pop
         $y | %{ git add $_ }
         git stash  | Out-Null
         git checkout $branch 2>&1
 
     }
     if (-not $?)
     {
         Write-Host "unsucessful $LASTEXITCODE" #for now
         $conflicts = $(git diff --name-only --diff-filter=U  )
         if ((-not $conflicts))
         {
             $z= askyn "countinue" 
             if (-not $z){return}
         }
     }
 
     $conflicts = $(git diff --name-only --diff-filter=U  )
     $changes = $(git diff --name-only $commit_hash )
     if ($conflicts)
     {
         if ($checkout) {
             Write-Host "wtf "
                 return
         }
         Write-Host "There are merge conflicts. Please run git pull. Aborting"
     $userInput = Read-Host -Prompt @"
Do you want to resolve  conflict using ours/theirs/no? 
no just cancels (type exactly)
"@ 
 
         if ($userInput -eq "no" ){
             return 
         }
         git rebase --abort
         DoPull @("-X","$userInput")
         if (-not $?){Write-Host "unsucessful pull";return}
     }
 
     # Checkout files from the stash
     git checkout stash -- . | Out-Null
     git reset | Out-Null
     $localch= $(git diff --name-only)
     $int = $localch | ?{ $changes -contains $_  } 
     if ($int)
     {
         Write-Host "Following files are different in local branch: $int " 
         if ($(-not ($keeplocalinconflict)))
         {
             if (-not $useremotefiles) 
             {
             $userInput = Read-Host -Prompt "Do you want to keep local changes in case of conflict? (y/n/apply)"
             } else{ $userInput='n'} 
             if ($userInput -eq "y") {
                 $keeplocalinconflict = $true
             } else {
                 Write-Host "reseting to remote"
                     if ($branch)
                     {
                         Write-Host "git checkout $branch @int "
                         git checkout $branch -- @int 
                     }
                     else 
                     {
                         git checkout -- @int
                     }
             } 
             if ($userInput -eq "apply")
             {
                 git stash apply
             }
         }
     }
 
     if ($dontkeepstash) 
     {
         git stash drop
     }
 }
 finally 
 {$PSStyle.OutputRendering=$oldrend}
 # Drop the stash
 }
 
 function SquashCommits([int]$count) {
 <#
 .SYNOPSIS
 Squashes the last N commits into one via interactive rebase.
 .PARAMETER count
 Number of recent commits to squash together.
 #>
 $commitHashes = git log --pretty=format:%h -n $count
 
 $commands= ( 0..$($count-2) ) |  %{   "sed -i 's/^pick $($commitHashes[$_])/squash $($commitHashes[$_])/' `$file"    }
 $st= $commands -join "`n"
 
 $st="func() {
 local file=`$1
 $st
 }; func"
 Write-Host $st
 $env:GIT_SEQUENCE_EDITOR=$st
 try{
         git rebase -i HEAD~$count
     }finally
     {
         Remove-Item Env:\GIT_SEQUENCE_EDITOR
     }
 }
  function Find-GitFileFromReflog {
 <#
 .SYNOPSIS
 Searches git reflog for exact file paths based on partial filename matches
 
 .DESCRIPTION
 This function searches through the git reflog (all log entries) to find exact file paths
 that match a partial filename. It uses git log --all --name-only to examine all commits
 and their changed files, then filters results based on the partial filename provided.
 
 .PARAMETER PartialFilename
 Partial filename to search for (e.g., "config", "*.ps1", "test")
 
 .PARAMETER MaxResults
 Maximum number of results to return (default: 50)
 
 .PARAMETER IncludeDeleted
 Switch to include files that have been deleted
 
 .EXAMPLE
 Find-GitFileFromReflog "config"
 Finds all file paths containing "config" in their name
 
 .EXAMPLE
 Find-GitFileFromReflog "*.ps1" -MaxResults 10
 Finds up to 10 PowerShell files
 
 .EXAMPLE
 Find-GitFileFromReflog "test" -IncludeDeleted
 Finds all files with "test" in name, including deleted ones
 
 .OUTPUTS
 Array of unique file paths that match the partial filename
 #>
     [CmdletBinding()]
     param(
         [Parameter(Mandatory=$true, Position=0)]
         [string]$PartialFilename,
         
         [Parameter(Mandatory=$false)]
         [int]$MaxResults = 50,
         
         [Parameter(Mandatory=$false)]
         [switch]$IncludeDeleted
     )
     
     try {
         # Check if we're in a git repository
         $gitRoot = git rev-parse --show-toplevel 2>$null
         if ($LASTEXITCODE -ne 0) {
             Write-Error "Not in a git repository"
             return
         }
         
         Write-Verbose "Searching git reflog for files matching: $PartialFilename"
         
         # Get all files from git log --all (includes reflog entries)
         $gitCommand = "git log --all --name-only --pretty=format:"
         if ($IncludeDeleted) {
             $gitCommand += " --diff-filter=ACDMRT"
         }
         
         $allFiles = Invoke-Expression $gitCommand | Where-Object { $_ -ne "" }
         
         # Convert partial filename to regex pattern for flexible matching
         $pattern = $PartialFilename -replace '\*', '.*' -replace '\?', '.'
         
         # Filter files that match the pattern
         $matchingFiles = $allFiles | Where-Object { 
             $fileName = Split-Path $_ -Leaf
             $fileName -match $pattern -or $_ -match $pattern
         } | Sort-Object -Unique
         
         # Limit results if specified
         if ($MaxResults -gt 0) {
             $matchingFiles = $matchingFiles | Select-Object -First $MaxResults
         }
         
         if ($matchingFiles.Count -eq 0) {
             Write-Warning "No files found matching pattern: $PartialFilename"
             return
         }
         
         Write-Host "Found $($matchingFiles.Count) unique file(s) matching '$PartialFilename':" -ForegroundColor Green
         
         # Return results with additional metadata
         $results = @()
         foreach ($file in $matchingFiles) {
             # Check if file currently exists
             $exists = Test-Path (Join-Path $gitRoot $file)
             
             # Get last commit that touched this file
             $lastCommit = git log -1 --pretty=format:"%h %ai %s" -- $file 2>$null
              git log --reflog --all --follow --format=%h --
              $reflog =  git log --reflog --all --follow --format=%h -- $file 
             
             $result = [PSCustomObject]@{
                 Path = $file
                 FullPath = Join-Path $gitRoot $file
                 Exists = $exists
                 LastCommit = $lastCommit
                 RefLog = $reflog
                 FileName = Split-Path $file -Leaf
             }
             $results += $result
         }
         
         return $results
     }
     catch {
         Write-Error "Error searching git reflog: $_"
     }
 }
 
 function UnfilterList ($ls)
 {
<#
.SYNOPSIS
Pipeline filter: passes through items that do NOT match any pattern in the list.
.PARAMETER ls
Array of regex patterns to exclude. Pipeline items matching any pattern are dropped.
#>
     Process {
         $el=$_;
         If (-not $( $ls | where { $el -imatch $_}) )
         {
             return $_ 
         }
     }
 
 }
 
 function FilterList ($ls)
 {
<#
.SYNOPSIS
Pipeline filter: passes through items that match at least one pattern in the list.
.PARAMETER ls
Array of regex patterns. Only pipeline items matching at least one are kept.
#>
     Process {
         $el=$_;
         If ($( $ls | where { $el -imatch $_}) )
         {
             return $_ 
         }
     }
 
 }
 function ExtendedPSList($file)
 {
     <#
     .SYNOPSIS
     Returns extended process info (owner, command line, creation date) via WMI.
     #>
         return  Get-WmiObject Win32_Process | Select Name,ProcessId,ParentProcessId,@{Name="UserName";Expression={$_.GetOwner().User}}, CommandLine, @{Name='CreationDate'; Expression={ [System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate)}}
 }
 function Ext
 {<#
 .SYNOPSIS
 Runs a script block in a new window using the current PowerShell host executable.
 .PARAMETER ScriptBlock
 The script block to execute.
 #>
         [CmdletBinding()]
         param (
                 [parameter(Position=0)]
                 [ScriptBlock]$ScriptBlock
 
               )
 
             $cmd=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptBlock.ToString()))
             Start-Process -FilePath  "$([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)" -ArgumentList ("-EncodedCommand",$cmd)
 }
 function ExtPwsh
 {<#
 .SYNOPSIS
 Runs a script block in a new pwsh window.
 .PARAMETER ScriptBlock
 The script block to execute.
 #>
         [CmdletBinding()]
         param (
                 [parameter(Position=0)]
                 [ScriptBlock]$ScriptBlock
 
               )
 
             $cmd=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptBlock.ToString()))
             Start-Process "pwsh" -ArgumentList ("-EncodedCommand",$cmd)
 }
 Function DoGridView ($v)
 {
 <#
 .SYNOPSIS
 Displays pipeline data in Out-GridView from a separate pwsh process.
 .PARAMETER v
 The data to display.
 #>
     $v | Export-Clixml -Path c:\temp\outgridview.tmp
         ExtPwsh { try{  $v = Import-Clixml -Path c:\temp\outgridview.tmp 
             $v | Out-GridView
         } catch {} 
 
         Read-Host -Prompt "Press Enter to continue"
         }
 
 }
 function IntroduceGitTreeAlias
 {
 <#
 .SYNOPSIS
 Adds a 'git tree' alias that shows a decorated one-line graph of all branches.
 #>
     git config --global alias.tree "log --oneline --decorate --all --graph"
 }

function Get-MD5 {
    <#
    .SYNOPSIS
    Calculate MD5 hash for a string or file

    .DESCRIPTION
    Computes the MD5 hash of a string or file and returns it as a hexadecimal string

    .PARAMETER String
    The string to hash

    .PARAMETER Path
    Path to the file to hash

    .EXAMPLE
    Get-MD5 -String "hello world"

    .EXAMPLE
    Get-MD5 -Path "C:\file.txt"
    #>
    [CmdletBinding(DefaultParameterSetName='String')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='String', Position=0)]
        [string]$String,

        [Parameter(Mandatory=$true, ParameterSetName='File')]
        [string]$Path
    )

    if ($PSCmdlet.ParameterSetName -eq 'File') {
        (Get-FileHash -Algorithm MD5 -Path $Path).Hash
    }
    else {
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $hash = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))
        [System.BitConverter]::ToString($hash) -replace '-', ''
    }
}
Function Format-ErrorWithStackTrace($ErrorRecord)
{
<#
.SYNOPSIS
Formats an error record as a detailed string with enhanced stack trace and source lines.
.PARAMETER ErrorRecord
The error record to format (e.g. from $Error[0] or a catch block's $_).
#>
    # Helper function to get source line from file
    function Get-SourceLine($scriptPath, $lineNumber) {
        if ($scriptPath -and $lineNumber -and (Test-Path $scriptPath)) {
            try {
                $lines = Get-Content $scriptPath -ErrorAction SilentlyContinue
                if ($lines -and $lineNumber -le $lines.Count) {
                    return $lines[$lineNumber - 1].Trim()
                }
            } catch {
                # Silently continue if we can't get source line
            }
        }
        return $null
    }

    # Helper function to enhance stack trace with source lines
    function Add-SourceLinesToStackTrace($stackTrace) {
        if (-not $stackTrace) { return "" }

        $enhancedTrace = ""
        $lines = $stackTrace -split "`n"

        foreach ($line in $lines) {
            $enhancedTrace += $line + "`n"

            # Parse stack trace line to extract file path and line number
            # Format: "at <ScriptBlock>, <path>: line <number>"
            if ($line -match 'at .+, (.+):\s*line\s+(\d+)') {
                $scriptPath = $matches[1]
                $lineNum = [int]$matches[2]

                $sourceLine = Get-SourceLine $scriptPath $lineNum
                if ($sourceLine) {
                    $enhancedTrace += "   SOURCE: $sourceLine`n"
                }
            }
        }

        return $enhancedTrace
    }

$errorInfo = @"
ERROR DETAILS:
==============
Message: $($ErrorRecord.Exception.Message)
Exception Type:
$($ErrorRecord.Exception.GetType().FullName)

STACK TRACE:
============
$(Add-SourceLinesToStackTrace $ErrorRecord.ScriptStackTrace)

EXCEPTION STACK TRACE:
=====================
$($ErrorRecord.Exception.StackTrace)

"@

# Add inner exception if it exists
        if ($ErrorRecord.Exception.InnerException) {
            $errorInfo += @"
INNER EXCEPTION:
================
Message:
$($ErrorRecord.Exception.InnerException.Message)
Type: $($ErrorRecord.Exception.InnerException.GetType().FullName)
Stack Trace:
$($ErrorRecord.Exception.InnerException.StackTrace)
"@
        }

# Add position info
$errorInfo += @"
POSITION:
=========
Line:
$($ErrorRecord.InvocationInfo.ScriptLineNumber)
Offset: $($ErrorRecord.InvocationInfo.OffsetInLine)
Script: $($ErrorRecord.InvocationInfo.ScriptName)
Command: $($ErrorRecord.InvocationInfo.MyCommand)

"@

    # Add source line at error position
    $sourceLine = Get-SourceLine $ErrorRecord.InvocationInfo.ScriptName $ErrorRecord.InvocationInfo.ScriptLineNumber
    if ($sourceLine) {
        $errorInfo += @"
SOURCE LINE:
============
$sourceLine

"@
    }

    return $errorInfo
} 
function FindFileRg {
    <#
    .SYNOPSIS
    Find all files and search them with ripgrep

    .DESCRIPTION
    Recursively finds files that match pattern (rg)

    .PARAMETER SearchString
    The string/pattern to search for using ripgrep

    .PARAMETER Path
    The starting directory path (default: current directory)

    .PARAMETER FilePattern
    Optional file pattern to filter files (e.g., "*.ps1", "*.txt")

    .EXAMPLE
    FindRg "function"
    Searches for "function" in all files in current directory

    .EXAMPLE
    FindRg "TODO" -Path "C:\Projects" -FilePattern "*.cs"
    Searches for "TODO" in all .cs files under C:\Projects
    #>
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$SearchString,

        [Parameter(Position=1)]
        [string]$Path = ".",

        [Parameter()]
        [string]$FilePattern = "*"
    )

    # Get all files using Find function
    $files = @(Find -path $Path -name $FilePattern -type File)

    if ($files.Count -gt 0) {
        # Pass files list to ripgrep as arguments
        echo $files | rg $SearchString -

    } else {
        Write-Warning "No files found in path: $Path"
    }
}

function Get-Histogram {
    [CmdletBinding(DefaultParameterSetName='BucketCount')]
    Param(
        [Parameter(Mandatory, ValueFromPipeline, Position=1)]
        [ValidateNotNullOrEmpty()]
        [array]
        $InputObject
        ,
        [Parameter(Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Property
        ,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [float]
        $Minimum
        ,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [float]
        $Maximum
        ,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Alias('Width')]
        [float]
        $BucketWidth = 1
        ,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Alias('Count')]
        [float]
        $BucketCount
        ,
        [Parameter()]
        [switch]
        $Visualize
        ,
        [Parameter()]
        [ValidateRange(1, 200)]
        [int]
        $BarWidth = 73
        ,
        [Parameter()]
        [switch]
        $Weighted
    )

    Begin {
        Write-Verbose ('[{0}] Initializing' -f $MyInvocation.MyCommand)

        $Buckets = @{}
        $Data = @()
    }

    Process {
        Write-Verbose ('[{0}] Processing {1} items' -f $MyInvocation.MyCommand, $InputObject.Length)

        $InputObject | ForEach-Object {
            if ($Weighted) {
                # Expect data in format (probability, value) or [probability, value]
                if ($_ -is [array] -and $_.Count -eq 2) {
                    $Data += [PSCustomObject]@{
                        Weight = $_[0]
                        Value = $_[1]
                    }
                } elseif ($_.GetType().ToString() -like 'System.Tuple*') {
                    $Data += [PSCustomObject]@{
                        Weight = $_.Item1
                        Value = $_.Item2
                    }
                } else {
                    Write-Host $_
                    throw ('Weighted data must be in format (probability, value) or [probability, value]')
                }
            } else {
                if ($Property) {
                    if (-Not ($_ | Select-Object -ExpandProperty $Property -ErrorAction SilentlyContinue)) {
                        throw ('Input object does not contain a property called <{0}>.' -f $Property)
                    }
                }
                $Data += $_
            }
        }
    }

    End {
        Write-Verbose ('[{0}] Building histogram' -f $MyInvocation.MyCommand)

        Write-Debug ('[{0}] Retrieving measurements from upstream cmdlet.' -f $MyInvocation.MyCommand)
        if ($Weighted) {
            $Stats = $Data | Microsoft.PowerShell.Utility\Measure-Object -Minimum -Maximum -Property Value
        } elseif ($Property) {
            $Stats = $Data | Microsoft.PowerShell.Utility\Measure-Object -Minimum -Maximum -Property $Property
        } else {
            $Stats = $Data | Microsoft.PowerShell.Utility\Measure-Object -Minimum -Maximum
        }

        if (-Not $PSBoundParameters.ContainsKey('Minimum')) {
            $Minimum = $Stats.Minimum
            Write-Debug ('[{0}] Minimum value not specified. Using smallest value ({1}) from input data.' -f $MyInvocation.MyCommand, $Minimum)
        }
        if (-Not $PSBoundParameters.ContainsKey('Maximum')) {
            $Maximum = $Stats.Maximum
            Write-Debug ('[{0}] Maximum value not specified. Using largest value ({1}) from input data.' -f $MyInvocation.MyCommand, $Maximum)
        }
        if (-Not $PSBoundParameters.ContainsKey('BucketCount')) {
            $BucketCount = [math]::Ceiling(($Maximum - $Minimum) / $BucketWidth)
            Write-Debug ('[{0}] Bucket count not specified. Calculated {1} buckets from width of {2}.' -f $MyInvocation.MyCommand, $BucketCount, $BucketWidth)
        }
        if ($BucketCount -gt 100) {
            Write-Warning ('[{0}] Generating {1} buckets' -f $MyInvocation.MyCommand, $BucketCount)
        }

        Write-Debug ('[{0}] Building buckets using: Minimum=<{1}> Maximum=<{2}> BucketWidth=<{3}> BucketCount=<{4}>' -f $MyInvocation.MyCommand, $Minimum, $Maximum, $BucketWidth, $BucketCount)
        $OverallCount = 0
        $Buckets = 1..$BucketCount | ForEach-Object {
            [pscustomobject]@{
                Index         = $_
                lowerBound    = $Minimum + ($_ - 1) * $BucketWidth
                upperBound    = $Minimum +  $_      * $BucketWidth
                Count         = 0
                RelativeCount = 0
                Group         = @()
                PSTypeName    = 'HistogramBucket'
            }
        }

        Write-Debug ('[{0}] Building histogram' -f $MyInvocation.MyCommand)
        $Data | ForEach-Object {
            if ($Weighted) {
                $Value = $_.Value
                $Weight = $_.Weight
            } elseif ($Property) {
                $Value = $_.$Property
                $Weight = 1
            } else {
                $Value = $_
                $Weight = 1
            }

            if ($Value -ge $Minimum -and $Value -le $Maximum) {
                $BucketIndex = [math]::Floor(($Value - $Minimum) / $BucketWidth)
                if ($BucketIndex -lt $Buckets.Length) {
                    $Buckets[$BucketIndex].Count += $Weight
                    $Buckets[$BucketIndex].Group += $_
                    $OverallCount += $Weight
                }
            }
        }

        Write-Debug ('[{0}] Adding relative count' -f $MyInvocation.MyCommand)
        $Buckets | ForEach-Object {
            if ($OverallCount -gt 0) {
                $_.RelativeCount = $_.Count / $OverallCount
            } else {
                $_.RelativeCount = 0
            }
        }

        if ($Visualize) {
            Write-Debug ('[{0}] Generating visualization' -f $MyInvocation.MyCommand)

            $MaxCount = ($Buckets | Measure-Object -Property Count -Maximum).Maximum

            $Buckets | Where-Object { $_.Count -gt 0 } | ForEach-Object {
                # Format the bucket range/label
                $Label = if ($Property) {
                    "[{0:N1}-{1:N1}]" -f $_.lowerBound, $_.upperBound
                } else {
                    "[{0:N1}-{1:N1}]" -f $_.lowerBound, $_.upperBound
                }

                # Calculate percentage
                $Percentage = if ($OverallCount -gt 0) {
                    [int](100 * $_.Count / $OverallCount)
                } else {
                    0
                }

                # Calculate bar length based on proportion to max count
                $BarLength = if ($MaxCount -gt 0) {
                    [int]($BarWidth * $_.Count / $MaxCount)
                } else {
                    0
                }

                # Create the bar
                $Bar = ("*" * $BarLength).PadRight($BarWidth)

                # Format and display the line
                $DataPointCount = $_.Group.Length
                $Line = "{0} {1}% {2} [n={3}, wt={4:N2}]" -f `
                    $Label.PadLeft(20), `
                    $Percentage.ToString().PadLeft(3), `
                    $Bar, `
                    $DataPointCount, `
                    $_.Count

                Write-Host $Line -ForegroundColor Green
            }
        }

        Write-Debug ('[{0}] Returning histogram' -f $MyInvocation.MyCommand)
        $Buckets
    }
}
function SetPrimary()
{
<#
.SYNOPSIS
Sets the DELL U2419H monitor as the primary display.
#>
    $di=  $(Get-DisplayInfo | ? {$_.DisplayName -eq  "DELL U2419H"} ).DisplayId
    Set-DisplayPrimary -DisplayId $di
}
function SetExtendedDisplayPort()
{
<#
.SYNOPSIS
Sets extended display mode and makes the DisplayPort monitor the primary display.
.DESCRIPTION
Finds all DisplayPort-connected monitors via Get-DisplayInfo, enables them as extended displays, and sets the first one as the primary display.
#>
    $dpDisplays = Get-DisplayInfo | Where-Object { $_.ConnectionType -eq 'DisplayPort' }
    if (-not $dpDisplays) {
        Write-Warning "No DisplayPort monitor found."
        return
    }
    $dpIds = @($dpDisplays | ForEach-Object { $_.DisplayId })
    # Switch to "Extend these displays" mode
    DisplaySwitch.exe /extend
    # Enable as extended (not cloned)
    Enable-Display -DisplayId $dpIds
    # Set the first DisplayPort display as primary
    Set-DisplayPrimary -DisplayId $dpIds[0]
    Write-Host "DisplayPort monitor '$($dpDisplays[0].DisplayName)' (ID $($dpIds[0])) set as extended primary display." -ForegroundColor Green
}
function RunWTAdmin()
{
<#
.SYNOPSIS
Launches Windows Terminal with administrator privileges.
#>
     Start-process -Verb RunAs "~\AppData\Local\Microsoft\WindowsApps\wt.exe"
}
function CloseClaudeDesktop() {
<#
.SYNOPSIS
Closes the Claude desktop app (WindowsApps package).
#>
    Get-process Claude | where {$_.Path -like "*WindowsApp*" }  | %{ echo $_ } | Stop-Process
    }

function New-TrayIcon {
<#
.SYNOPSIS
Creates a system tray icon with a configurable context menu.
.DESCRIPTION
Creates a Windows system tray (NotifyIcon) with a right-click context menu.
Each menu item can run a configurable command (scriptblock or string).
Runs the message loop in a STA runspace so the calling shell stays interactive.
Returns a hashtable with 'Runspace' and 'PowerShell' so callers can dispose/stop it.
.PARAMETER Tooltip
Text shown when hovering over the tray icon.
.PARAMETER MenuItems
Array of [PSCustomObject]@{Label='...'; Command={...}} entries.
Each Command can be a ScriptBlock or a string (passed to Invoke-Expression).
.PARAMETER IconPath
Optional path to a .ico file. Defaults to the PowerShell icon.
.PARAMETER HotKey
Optional global keyboard shortcut to open the menu (e.g., 'Win+Alt+T', 'Ctrl+Shift+M').
Supported modifiers: Win, Ctrl, Alt, Shift. Example: 'Ctrl+Alt+T'
.EXAMPLE
$items = @(
    [PSCustomObject]@{ Label = 'Run Notebook'; Command = { jupyter nbconvert --to notebook --execute mynotebook.ipynb } }
    [PSCustomObject]@{ Label = 'Open Shell';   Command = 'pwsh' }
)
$tray = New-TrayIcon -Tooltip 'My App' -MenuItems $items -HotKey 'Ctrl+Alt+T'
# To remove it later:
# $tray.PowerShell.Stop(); $tray.Runspace.Dispose()
#>
    param(
        [string]$Tooltip = 'PowerShell Tray',
        [Parameter(Mandatory)]
        [object[]]$MenuItems,
        [string]$IconPath = '',
        [string]$HotKey = ''
    )
    $mutexName = 'Global\PSTrayIcon_Mutex'
    $mutex = [System.Threading.Mutex]::new($false, $mutexName)
    $acquired = $mutex.WaitOne(0)   # non-blocking try
    if (-not $acquired) {
        Write-Warning "Tray icon is already running in another process."
            $mutex.Dispose()
            return
    }
if ($global:tray) {
    $global:tray.Dispose()
        $global:tray = $null
}

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $runspace.ThreadOptions   = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()

    $runspace.SessionStateProxy.SetVariable('TrayTooltip',  $Tooltip)
    $runspace.SessionStateProxy.SetVariable('TrayMenuItems', $MenuItems)
    $runspace.SessionStateProxy.SetVariable('TrayIconPath',  $IconPath)
    $runspace.SessionStateProxy.SetVariable('TrayHotKey',    $HotKey)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $runspace

    $null = $ps.AddScript({
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Register global hotkey API
        if ($TrayHotKey) {
            $HotKeyCode = @"
using System;
using System.Runtime.InteropServices;
public static class HotKeyHelper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    public static extern uint MapVirtualKey(uint uCode, uint uMapType);

    public const uint MOD_CONTROL = 0x0002;
    public const uint MOD_ALT = 0x0001;
    public const uint MOD_SHIFT = 0x0004;
    public const uint MOD_WIN = 0x0008;
}
"@
            Add-Type -TypeDefinition $HotKeyCode -ErrorAction SilentlyContinue
        }

        if ($TrayIconPath -and (Test-Path $TrayIconPath)) {
            $icon = [System.Drawing.Icon]::new($TrayIconPath)
        } else {
            $psExe = (Get-Process -Id $PID).Path
            $icon  = [System.Drawing.Icon]::ExtractAssociatedIcon($psExe)
        }

        $notifyIcon         = [System.Windows.Forms.NotifyIcon]::new()
        $notifyIcon.Icon    = $icon
        $notifyIcon.Text    = $TrayTooltip
        $notifyIcon.Visible = $true

        $menu = [System.Windows.Forms.ContextMenuStrip]::new()

        foreach ($item in $TrayMenuItems) {
            $menuItem = [System.Windows.Forms.ToolStripMenuItem]::new($item.Label)
            $cmd = $item.Command
            $menuItem.add_Click({
                if ($cmd -is [scriptblock]) {
                    & $cmd
                } else {
                    Invoke-Expression ([string]$cmd)
                }
            }.GetNewClosure())
            $null = $menu.Items.Add($menuItem)
        }

        $null = $menu.Items.Add([System.Windows.Forms.ToolStripSeparator]::new())

        $exitItem = [System.Windows.Forms.ToolStripMenuItem]::new('Exit')
        $exitItem.add_Click({
            $notifyIcon.Visible = $false
            $notifyIcon.Dispose()
            [System.Windows.Forms.Application]::Exit()
        })
        $null = $menu.Items.Add($exitItem)

        $notifyIcon.ContextMenuStrip = $menu

        # Ensure tray icon is removed if the message loop exits for any reason
        # (graceful runspace shutdown when parent pwsh exits cleanly).
        $cleanup = {
            try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch {}
        }
        [System.Windows.Forms.Application]::add_ApplicationExit($cleanup)

        # Setup hotkey if provided
        if ($TrayHotKey) {
            $hiddenForm = New-Object System.Windows.Forms.Form
            $hiddenForm.ShowInTaskbar = $false
            $hiddenForm.WindowState = 'Minimized'
            $hiddenForm.FormBorderStyle = 'None'
            $hiddenForm.Size = @{ Width = 1; Height = 1 }

            # Parse hotkey string (e.g., "Ctrl+Alt+T")
            $hotKeyParts = $TrayHotKey -split '\+'
            $modifiers = 0
            $vk = 0

            foreach ($part in $hotKeyParts) {
                $part = $part.Trim()
                switch -Exact ($part) {
                    'Ctrl'  { $modifiers += [HotKeyHelper]::MOD_CONTROL }
                    'Alt'   { $modifiers += [HotKeyHelper]::MOD_ALT }
                    'Shift' { $modifiers += [HotKeyHelper]::MOD_SHIFT }
                    'Win'   { $modifiers += [HotKeyHelper]::MOD_WIN }
                    default {
                        # Parse key name to virtual key code
                        if ($part.Length -eq 1) {
                            $vk = [System.Windows.Forms.Keys]::($part.ToUpper())
                        } else {
                            $vk = [System.Windows.Forms.Keys]::$part
                        }
                    }
                }
            }

            $hotKeyId = 9999
            $registered = [HotKeyHelper]::RegisterHotKey($hiddenForm.Handle, $hotKeyId, $modifiers, $vk)

            if ($registered) {
                $hiddenForm.add_Load({
                    $form = $this
                    $form.add_FormClosing({ [HotKeyHelper]::UnregisterHotKey($form.Handle, $hotKeyId) })
                })

                # Override WndProc to capture WM_HOTKEY
                $form = $hiddenForm
                $originalWndProc = $form.WndProc
                $form.WndProc = {
                    param($m)
                    if ($m.Msg -eq 0x0312 -and $m.WParam.ToInt32() -eq $hotKeyId) {
                        $menu.Show([System.Windows.Forms.Cursor]::Position)
                    }
                    & $originalWndProc ([ref]$m)
                }
            }
        }

        try {
            [System.Windows.Forms.Application]::Run()
        } finally {
            try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch {}
        }
    })

    $null = $ps.BeginInvoke()

    $global:tray=$ps
    $global:trayRunspace = $runspace
    $global:trayMutex = $mutex

    # Dispose tray cleanly when this pwsh exits, so we don't leak a ghost icon.
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -SupportEvent -Action {
        try {
            if ($global:tray)         { $global:tray.Stop(); $global:tray.Dispose() }
            if ($global:trayRunspace) { $global:trayRunspace.Close(); $global:trayRunspace.Dispose() }
            if ($global:trayMutex)    { try { $global:trayMutex.ReleaseMutex() } catch {}; $global:trayMutex.Dispose() }
        } catch {}
    } | Out-Null

    Write-Host "Tray icon '$Tooltip' created. Right-click and choose 'Exit' to remove it."
    return @{ PowerShell = $ps; Runspace = $runspace }
}


# ---------- Auto-Claude Docker helpers ----------
function Find-AutoClaudeContainer {
    $container = docker ps --format "{{.Names}}" 2>$null | Where-Object { $_ -match 'claude' } | Select-Object -First 1
    if (-not $container) {
        Write-Error "No running auto-claude container found"
        return $null
    }
    return $container
}

function Copy-ToAutoClaude {
    param(
        [Parameter(Mandatory, Position = 0)][string]$HostPath,
        [Parameter(Position = 1)][string]$ContainerPath = "/workspace"
    )

    $container = Find-AutoClaudeContainer
    if (-not $container) { return }

    $resolved = Resolve-Path $HostPath -ErrorAction Stop
    $name = Split-Path $resolved -Leaf
    $dest = "$ContainerPath/$name"

    Write-Host "Copying '$resolved' -> '${container}:${dest}' ..."
    docker cp "$resolved" "${container}:${dest}"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Done. Files available at $dest inside container." -ForegroundColor Green
    } else {
        Write-Error "docker cp failed with exit code $LASTEXITCODE"
    }
}

function Test-Port([int]$Port, [int]$TimeoutMs = 200) {
<#
.SYNOPSIS
Fast check whether a TCP port is listening on localhost.
.PARAMETER Port
Port number to test.
.PARAMETER TimeoutMs
Connection timeout in milliseconds (default 200).
#>
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $task = $tcp.ConnectAsync('127.0.0.1', $Port)
        $connected = $task.Wait($TimeoutMs) -and $tcp.Connected
        $tcp.Dispose()
        return $connected
    } catch { return $false }
}
function RunClaudeDir($d)
{
    cd $d
    claude
}

# --- Generic parameter-forwarding helpers (from previous monolithic profile) ---

function AddWrapper([parameter(mandatory=$true, position=0)][string]$For,[parameter(mandatory=$true, position=1)][string]$To)
{
    $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    $paramset= $(Get-Command $For).Parameters.Values | %{[System.Management.Automation.RuntimeDefinedParameter]::new($_.Name,$_.ParameterType,$_.Attributes)}
    $paramsetlet= $(Get-Command empt).Parameters.Keys
    $paramsetlet+= $(Get-Command $To).ScriptBlock.Ast.Body.ParamBlock.Parameters.Name | %{ $_.VariablePath.UserPath }
    $paramset | %{ if ( -not ($paramsetlet -contains $_.Name) )
        {$paramDictionary.Add($_.Name,$_)
        }}
    return $paramDictionary
}

function GetRestOfParams()
{
    Param([parameter(mandatory=$true, position=1)][hashtable]$params,
        [parameter(mandatory=$true, position=0)][string]$dstsource,
        [parameter(mandatory=$false, position=2)][switch][bool]$dontincludecommon=$true)
    $dstorgparams=$(Get-Command $dstsource).Parameters.Keys
    $z= $params
    if ( -not $dontincludecommon)
    {
        $z.Keys | %{ if ( -not ($dstorgparams -contains $_) )
            {$z.Remove($_)
            } } | Out-Null
    } else
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
    DynamicParam
    {
        AddWrapper -For Get -To $MyInvocation.MyCommand.Name
    }
    Begin
    {
        $params = GetRestOfParams Let $PSBoundParameters -dontincludecommon
    }
    Process
    {
        Get @params -OptionB ( $OptionB + "1" )
    }
}

function Get
{
    [CmdLetBinding()]
    Param([parameter(mandatory=$false, position=0)][string]$OptionA,
        [parameter(mandatory=$false, position=1)][string]$OptionB)
    Write-Host "opta",$OptionA
    Write-Host "optb",$OptionB
}

function Grant-UserRW {
<#
.SYNOPSIS
Recursively grants the current user read/write (Modify) permission on a path.
Takes ownership first if the initial icacls grant fails (typical for files
owned by SYSTEM/TrustedInstaller or another user).
.PARAMETER Path
Root file or directory to fix. Defaults to current directory.
#>
    param(
        [Parameter(Position=0)][string]$Path = (Get-Location).Path
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Error "Path not found: $Path"
        return
    }
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $user = "$env:USERDOMAIN\$env:USERNAME"
    Write-Host "Granting Modify to $user on $resolved ..."
    icacls $resolved /grant "${user}:(OI)(CI)M" /T /C | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Initial grant had errors ($LASTEXITCODE). Taking ownership and retrying..."
        if (Test-Path -LiteralPath $resolved -PathType Container) {
            takeown /F $resolved /R /D Y | Out-Null
        } else {
            takeown /F $resolved | Out-Null
        }
        icacls $resolved /grant "${user}:(OI)(CI)M" /T /C | Out-Null
    }
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Done."
    } else {
        Write-Warning "icacls finished with exit code $LASTEXITCODE - some items may not have been updated."
    }
}

function Repair-Bluetooth {
<#
.SYNOPSIS
Restarts Bluetooth radio adapters stuck in CM_PROB_FAILED_START (Code 10).
.DESCRIPTION
When Windows reports Bluetooth as off but bthserv is running, the radio
adapter itself may have failed to start (typically NTStatus 0xC000025E
STATUS_DEVICE_POWER_FAILURE). Disable/Enable-PnpDevice often does NOT clear
this; pnputil /restart-device does. Requires elevation.
#>
    param(
        [string]$Match = '*Bluetooth*'
    )
    $radios = Get-PnpDevice -Class Bluetooth | Where-Object {
        $_.FriendlyName -like $Match -and $_.InstanceId -like 'USB\*'
    }
    if (-not $radios) { Write-Warning "No Bluetooth USB radio matched '$Match'."; return }
    foreach ($r in $radios) {
        Write-Host "$($r.FriendlyName)  Status=$($r.Status)  Problem=$($r.Problem)"
        if ($r.Status -ne 'OK') {
            pnputil /restart-device $r.InstanceId
            Start-Sleep -Seconds 2
            Get-PnpDevice -InstanceId $r.InstanceId | Format-List FriendlyName,Status,Problem
        } else {
            Write-Host "  (already OK, skipping)"
        }
    }
}

function ConvertTo-LineEnding {
<#
.SYNOPSIS
Convert line endings of a file in place (LF <-> CRLF). Preserves bytes otherwise.
#>
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][string]$Path,
        [Parameter(Mandatory=$true)][ValidateSet('LF','CRLF')][string]$Eol
    )
    process {
        $resolved = (Resolve-Path -LiteralPath $Path).Path
        $bytes = [System.IO.File]::ReadAllBytes($resolved)
        # Strip CR (0x0D); for CRLF, re-insert before each LF (0x0A).
        $out = New-Object System.Collections.Generic.List[byte]
        foreach ($b in $bytes) {
            if ($b -eq 0x0D) { continue }
            if ($Eol -eq 'CRLF' -and $b -eq 0x0A) { $out.Add([byte]0x0D) }
            $out.Add($b)
        }
        [System.IO.File]::WriteAllBytes($resolved, $out.ToArray())
        Write-Host "$Eol  $resolved"
    }
}

function dos2unix {
    param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][string[]]$Path)
    process { foreach ($p in $Path) { ConvertTo-LineEnding -Path $p -Eol LF } }
}

function unix2dos {
    param([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][string[]]$Path)
    process { foreach ($p in $Path) { ConvertTo-LineEnding -Path $p -Eol CRLF } }
}

Set-Alias windows2dos unix2dos

function tail {
<#
.SYNOPSIS
Unix-like tail. Prints the last N lines of a file; -f follows for new content.
#>
    param(
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][string]$Path,
        [Alias('n')][int]$Lines = 10,
        [Alias('f')][switch]$Follow
    )
    process {
        if ($Follow) { Get-Content -LiteralPath $Path -Tail $Lines -Wait }
        else         { Get-Content -LiteralPath $Path -Tail $Lines }
    }
}

function Invoke-NotebookCell {
<#
.SYNOPSIS
Run a single Jupyter notebook cell (by index or substring match) with all preceding code cells included so dependencies are satisfied.
.PARAMETER Path
Path to the .ipynb file.
.PARAMETER Index
Zero-based cell index to run.
.PARAMETER Match
Substring matched against each code cell's source; first hit wins.
.PARAMETER Python
Python executable used to run nbconvert. Defaults to .\.venv11\Scripts\python.exe if present, else 'python'.
.PARAMETER Kernel
Jupyter kernel name. Default 'python3'.
.PARAMETER Timeout
Per-cell execution timeout in seconds. Default 120.
#>
    [CmdletBinding(DefaultParameterSetName='ByIndex')]
    param(
        [Parameter(Mandatory=$true, Position=0)][string]$Path,
        [Parameter(Mandatory=$true, ParameterSetName='ByIndex')][int]$Index,
        [Parameter(Mandatory=$true, ParameterSetName='ByMatch')][string]$Match,
        [string]$Python,
        [string]$Kernel = 'python3',
        [int]$Timeout = 120
    )

    if (-not (Test-Path $Path)) { throw "Notebook not found: $Path" }
    $Path = (Resolve-Path $Path).Path

    if (-not $Python) {
        $cand = Join-Path (Get-Location) '.venv11\Scripts\python.exe'
        $Python = if (Test-Path $cand) { $cand } else { 'python' }
    }

    $tmpIn  = [IO.Path]::GetTempFileName() + '.ipynb'
    $tmpOut = [IO.Path]::GetTempFileName() + '.ipynb'

    $py = @"
import json, sys
src   = json.load(open(r'$Path','r',encoding='utf-8'))
mode  = '$($PSCmdlet.ParameterSetName)'
idx   = $Index
match = r'''$Match'''
cells = src['cells']
target = None
if mode == 'ByIndex':
    if idx < 0 or idx >= len(cells):
        sys.exit(f'index {idx} out of range (0..{len(cells)-1})')
    target = idx
else:
    for i,c in enumerate(cells):
        if c.get('cell_type') == 'code' and match in ''.join(c.get('source', [])):
            target = i; break
    if target is None:
        sys.exit(f'no code cell matched: {match!r}')
keep = []
for i,c in enumerate(cells[:target+1]):
    if c.get('cell_type') == 'code':
        nc = dict(c)
        nc['outputs'] = []
        nc['execution_count'] = None
        keep.append(nc)
src['cells'] = keep
src.setdefault('metadata', {}).setdefault('kernelspec', {'name':'$Kernel','display_name':'$Kernel'})
json.dump(src, open(r'$tmpIn','w',encoding='utf-8'), ensure_ascii=False)
print(f'target={target} kept={len(keep)}')
"@
    $py | & $Python -

    if ($LASTEXITCODE -ne 0) { Remove-Item -EA SilentlyContinue $tmpIn,$tmpOut; return }

    $env:PYTHONNOUSERSITE = '1'
    $exec = @"
import sys, nbformat
from nbclient import NotebookClient
nb = nbformat.read(r'$tmpIn', as_version=4)
client = NotebookClient(nb, timeout=$Timeout, kernel_name='$Kernel', allow_errors=True)
client.execute()
nbformat.write(nb, r'$tmpOut')
print('[exec] done')
"@
    $exec | & $Python -

    $report = @"
import json
nb = json.load(open(r'$tmpOut','r',encoding='utf-8'))
last = nb['cells'][-1]
print('--- target cell source ---')
print(''.join(last.get('source', []))[:400])
print('--- outputs ---')
err = False
for o in last.get('outputs', []):
    t = o.get('output_type')
    if t == 'stream':
        name = o.get('name','stdout')
        print('[' + name + ']')
        print(''.join(o.get('text', [])))
    elif t == 'error':
        err = True
        print('[ERROR] ' + str(o.get('ename')) + ': ' + str(o.get('evalue')))
        for line in o.get('traceback', []):
            print(line)
    elif t in ('execute_result','display_data'):
        d = o.get('data', {})
        txt = d.get('text/plain')
        if isinstance(txt, list): txt = ''.join(txt)
        head = (txt[:400] if txt else str(list(d.keys())))
        print('[' + t + '] ' + head)
import sys; sys.exit(2 if err else 0)
"@
    $report | & $Python -
    $rc = $LASTEXITCODE
    Remove-Item -EA SilentlyContinue $tmpIn,$tmpOut
    if ($rc -eq 2) { Write-Host 'cell raised an exception' -ForegroundColor Red }
}
