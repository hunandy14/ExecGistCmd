# 執行 gist 上的 cmd 代碼
function ExecGistCmd {
    [Alias("ecgist")]
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $GistId,
        [Parameter(ParameterSetName = "")]
        [UInt16] $Index = 0,
        [Parameter(ParameterSetName = "")]
        [string[]] $ArgumentList,
        [Parameter(ParameterSetName = "")]
        [Switch] $RunAsAdministrator,
        [Switch] $Pause,
        [string] $SaveOnWorkDir
    )

    # 獲取 gist 上的最新資源
    $gist = Invoke-RestMethod "api.github.com/gists/$GistId"
    $gistUrl = $gist.html_url
    $author = $gist.owner.login
    $description = $gist.description
    $file = @($gist.files.PSObject.Properties.Value)[$Index]
    $content = $file.content
    $filename = $file.filename
    if ($filename -notmatch '\.(cmd|bat)$') { Write-Error "The file does not with '*.cmd' or '*.bat' file" -ErrorAction Stop }

    # 儲存到暫存檔案
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    $enc = [Text.Encoding]::GetEncoding([int](PowerShell -Nop "& {return [Text.Encoding]::Default.CodePage}"))
    if ($SaveOnWorkDir) {
        $filepath = [IO.Path]::GetFullPath($filename)
    } else { $filepath = "$env:TEMP\$filename" }
    [IO.File]::WriteAllText($filepath, $content, $enc)

    # 執行前確認
    Write-Host "The script that will be executed." -ForegroundColor DarkGreen
    Write-Host "  gisturl     : $gistUrl"
    Write-Host "  author      : $author"
    Write-Host "  description : $description"
    Write-Host "  filename    : $filename"
    Write-Host ""
    Write-Host "Press Enter to continue or Esc to exit..."  -ForegroundColor Yellow
    $keyInfo = [System.Console]::ReadKey($true)
    if ($keyInfo.Key -eq "Enter") { } elseif ($keyInfo.Key -eq "Escape") {
        Write-Host "Exiting..."; return
    }

    # 組裝命令
    $cmdStr = @()
    if ($content -notmatch "^\s*title\b") { $cmdStr += "title $description" }
    if ($ArgumentList) { $argument = ' ' + ($ArgumentList -join " ") }
    $cmdStr += [IO.Path]::GetFullPath($filepath) + $argument
    if ($Pause) { $cmdStr += "Pause" }
    $cmdStr = $cmdStr -join "&"
    if ($RunAsAdministrator) { $verb = 'RunAs' } else { $verb = 'Open' }

    # 執行檔案
    Write-Host $cmdStr -ForegroundColor DarkGray
    Start-Process 'cmd.exe' -ArgumentList "/c $cmdStr" -Verb $verb
} # ExecGistCmd 6a290cc77609c13a899b9a6e0801d008 -Pause
