
# ---------- Track last command per PID via PSReadLine ----------
$script:_LastCmdFile = "$env:TEMP\ps_last_command.json"
$script:_LastCmdData = @{}

# Load existing data once at startup (not on every command)
if (Test-Path $script:_LastCmdFile) {
try {
$loaded = Get-Content $script:_LastCmdFile -Raw | ConvertFrom-Json
$loaded.PSObject.Properties | ForEach-Object {
$script:_LastCmdData[$_.Name] = $_.Value
}
} catch {}
}

Set-PSReadLineOption -AddToHistoryHandler {
param([string]$line)
    $script:_LastCmdData["$PID"] = [ordered]@{
        Command = $line
            Time    = (Get-Date -Format 'o')
            CWD = $PWD.Path
    }
$WarningPreference='SilentlyContinue'
    $script:_LastCmdData | ConvertTo-Json  -Depth 5 | Set-Content $script:_LastCmdFile -Encoding UTF8
    return $true
}
# Shared shell-state file consumed by window_switcher: map PID -> {title, cwd, time, processid, command}.
# Concurrent shells coordinate via a named mutex.
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

$ExecutionContext.InvokeCommand.PostCommandLookupAction = {
try{
    $cmdLine = $MyInvocation.Line
    if ($args[1].CommandOrigin -ne 'Runspace' -or $cmdLine -match 'PostCommandLookupAction|^prompt$')
    { return
    }

    Add-CmdLineRecord -Dir (Get-Location).Path -CommandLine $cmdLine
    # Keep the window_switcher export alive — this handler replaces the one common.psm1 installs.
    Update-WindowSwitcherState -CommandLine $cmdLine
    }catch {
Write-Debug "error in PostCommandLookupAction: $_"
    }
}
