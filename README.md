# PowerShell Profile

Personal PowerShell profile and modules for Windows, with productivity utilities, git helpers, WSL integration, and more.
The profile is loaded in each powershell session automatically. 


## Structure


## Aliases

| Alias  | Maps To              | Description                                  |
|--------|----------------------|----------------------------------------------|
| `ss`   | `Select-String`      | Grep-like string search                      |
| `grep` | `Select-String`      | Grep-like string search                      |
| `z`    | `Get-Help`           | Quick help lookup                            |
| `m`    | `Get-Member`         | Inspect object members                       |
| `cd`   | `MyCD`               | Custom cd that records history               |
| `q`    | `CdLast`             | Jump to a previously visited directory       |
| `gitp` | `GitPullKeepLocal`   | Git pull keeping local changes               |

## Install

Clone into the PowerShell profile directory (or pull if it already exists):

```powershell
$ErrorActionPreference = 'Stop'
$repoUrl  = 'https://github.com/eyalbold/.powershell.git'
$destDir  = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell'

if (-not (Test-Path $destDir)) {
    git clone $repoUrl $destDir
} else {
    Write-Warning "'$destDir' already exists — pulling latest instead"
    Push-Location $destDir
    if (-not $(Test-Path .git))
    { git init }
    git remote add upstream $repoUrl
    git fetch upstream
    if (-not $(Test-Path Microsoft.PowerShell_profile.ps1) )
    { git pull upstream/main} else { Write-Error "profile exists, please merge" ; Exit 1 } 
    Pop-Location
}

# Hook into Windows PowerShell 5 profile
$ps5Profile = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Microsoft.PowerShell_profile.ps1'
$dotSource  = ". $destDir\Microsoft.PowerShell_profile.ps1"
if (-not (Test-Path $ps5Profile)) {
    New-Item -ItemType File -Path $ps5Profile -Force | Out-Null
}
Add-Content -Path $ps5Profile -Value $dotSource
```


## Key Bindings (PSReadLine)

> **Opt-in only** — disabled by default.

To enable, run once in any PowerShell session:

```powershell
OptInWtKeys
```

This writes `$global:enable_wt_keys=$true` to `config.ps1`, which is loaded automatically on every subsequent session.

| Key       | Action                                                        |
|-----------|---------------------------------------------------------------|
| `Alt+q`   | Jump to a previously visited directory (fuzzy via fzf)       |
| `Alt+e`   | Insert a command previously run in the **current** directory  |
| `Alt+h`   | Fuzzy-search the full global command history                  |
| `Alt+c`   | Pick a previously visited directory via fzf and open Claude there |

## Functions

### Profile

| Function          | Description                                                        |
|-------------------|--------------------------------------------------------------------|
| `ProfileCommands` | List all profile functions with their synopsis and description     |

### Navigation & History

| Function       | Description                                                                 |
|----------------|-----------------------------------------------------------------------------|
| `MyCD`         | Wraps `Set-Location`, records every navigation in PSReadLine history        |
| `CdLast`       | Pick a previously visited directory via fzf and jump to it                  |
| `StupidHist`   | Returns list of previously visited directories from history                  |
| `SimpHist`     | Fuzzy-search full command history (returns selected entry)                   |
| `SimpHistEx`   | Like `SimpHist` but immediately executes the selected command                |
| `GrepOnCurDir` | Fuzzy-select a command previously run in the current directory               |

### File & Search

| Function        | Description                                                               |
|-----------------|---------------------------------------------------------------------------|
| `Find`          | Unix-like `find`: filter by name, type, path; supports `-exec` callbacks  |
| `FindFileRg`    | Recursively find files and search their contents with ripgrep              |
| `FindGitFile`   | Show full git log for a file, following renames (`git log --follow`)       |
| `Which`         | Locate the full path of an executable (like Unix `which`)                  |
| `Get-MD5`       | Compute MD5 hash of a string or file                                       |
| `lastls`        | List directory entries newest-first (by `LastWriteTime`)                   |

### Git Utilities

| Function                        | Description                                                        |
|---------------------------------|--------------------------------------------------------------------|
| `Checkout-FileWithDifferentName`| Checkout a file from a branch and save it under a new name        |
| `GetGitStash`                   | List current git stashes                                           |
| `CheckCommit`                   | Inspect a commit by number or line                                 |
| `RemoveCommit`                  | Remove a specific commit                                           |
| `ExtractFromLastStash`          | Pull a single file out of the last stash                           |
| `Checkout-FileFromStash`        | Checkout a file from a named stash                                 |
| `StashAll`                      | Stash all changes with a name                                      |
| `FilesInCommit`                 | List files changed in a specific commit                            |

### Process & Window Management

| Function            | Description                                                             |
|---------------------|-------------------------------------------------------------------------|
| `Show-Window`       | Bring a process window to the foreground (restore if minimized)        |
| `Get-LockingProcess`| Find processes locking a file (uses Sysinternals `handle.exe`)         |
| `KillByName`        | Kill all processes matching a name pattern                              |
| `Find-HungUI`       | List UI processes that are Not Responding (alias `hungui`)              |
| `Start-ChatGPTDevTools`| Relaunch the ChatGPT desktop app with remote debugging and open DevTools in Chrome |
| `Kill-Port`         | Kill the process(es) listening on a TCP port (alias `killport`, `-WhatIf` supported) |
| `Test-Port`         | Fast check whether a TCP port is listening on localhost                 |
| `Respawn`           | Kill the main process(es) matching a name and relaunch each from its original path |
| `Get-ProcessCwd`    | Read a process's current working directory out of its PEB (64-bit only) |

### Claude

| Function       | Description                                                                          |
|----------------|--------------------------------------------------------------------------------------|
| `ClaudeGo`     | Block until the Claude usage limit resets, then resume with a message                |
| `ClaudeGoTask` | Same, but via a `-WakeToRun` scheduled task that wakes the machine (`-Cancel` to drop)|
| `Enable-WakeTimers` | Turn on "Allow wake timers" — without it `ClaudeGoTask` never wakes the box     |
| `SearchConv`   | Search past Claude Code conversations and resume the best match here                 |

### WSL Integration

| Function        | Description                                       |
|-----------------|---------------------------------------------------|
| `TranslatePath` | Convert a WSL/Linux path to a Windows path        |
| `RunBash`       | Execute a command in WSL with bash profile loaded |
| `RestartWsl`    | Restart the WSL service (`LxssManager`)           |

### Hyper-V / VM

| Function                    | Description                                              |
|-----------------------------|----------------------------------------------------------|
| `ConVM`                     | Open a PSSession to the local `win10` Hyper-V VM        |
| `copy-foldertovirtualmachine`| Copy a folder's contents into a Hyper-V VM             |
| `NewVMDrive`                | Create a new VM drive                                    |

### Data & Utilities

| Function                   | Description                                                               |
|----------------------------|---------------------------------------------------------------------------|
| `IIF`                      | Inline ternary: `IIF $cond $true $false`                                  |
| `ConvertPSObjectToHashtable`| Recursively convert a PSObject/JSON result to a hashtable                |
| `Get-Histogram`            | Generate a histogram (with optional ASCII visualization) from pipeline data|
| `Select-Zip` / `Tuple-Zip` | Zip two collections together element-by-element                           |
| `Ext2`                     | Run a script block in a new PowerShell window                             |
| `DebugIt`                  | Debug helper                                                              |
| `Add-ToPath`               | Permanently add a directory to the system PATH                            |
| `Format-ErrorWithStackTrace`| Format an error record with enhanced stack trace and source lines        |

### Dynamic Parameter Helpers

| Function          | Description                                                              |
|-------------------|--------------------------------------------------------------------------|
| `AddWrapper`      | Build a `DynamicParam` dictionary forwarding params from another function|
| `GetRestOfParams` | Filter a bound-params hashtable to only those accepted by a target function|
| `Let` / `Get`     | Wrapper functions using dynamic parameter forwarding                     |

## Modules

### Microsoft.PowerToys.Configure

A PowerShell DSC module for configuring PowerToys settings declaratively. Covers:

- **AdvancedPaste** – AI paste, clipboard preview, shortcuts
- **Awake** – keep-awake modes (Passive, Indefinite, Timed, Expirable)
- **ColorPicker** – activation actions, click behavior
- **Hosts** – hosts file editor settings
- **PowerAccent** – accent character activation key
- And many more PowerToys modules

## Optional Dependencies

These are only needed for specific functions — the profile loads fine without them.

| Dependency                                                                                   | Used By                                                                      |
|----------------------------------------------------------------------------------------------|------------------------------------------------------------------------------|
| [fzf](https://github.com/junegunn/fzf)                                                      | `CdLast`, `SimpHist`, `SimpHistEx`, `GrepOnCurDir`, `Alt+q/e/h` keybindings |
| [ripgrep (`rg`)](https://github.com/BurntSushi/ripgrep)                                      | `FindFileRg`                                                                 |
| [Sysinternals handle.exe](https://learn.microsoft.com/en-us/sysinternals/downloads/handle)  | `Get-LockingProcess`                                                         |
| PowerShell 7                                                                                 | keybindings                                                                  |
| WSL (Ubuntu)                                                                                 | `TranslatePath`, `RunBash`, `RestartWsl`                                     |
