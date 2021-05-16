#_EDCBX_HIDE_
#Requires -Version 5
#Requires -PSEdition Desktop

#--------------------ユーザ設定--------------------
$Settings =
@{
    Log =
    @{
        # ログ出力先
        Path = 'C:\logs\PostRecEnd'
        # ログを残す件数
        CntMax = 500
    }
    Post =
    @{
        # Webhook Url
        WebhookUrl = "https://discord.com/api/webhooks/XXXXXXXXXX"
    }
    Profiles =
    @(
        @{
            # 実行条件例 番組名(ファイル名) $env:FileName の部分一致の例
            Conditional = {$env:FileName -match "ほげほげぷー|ふがふがぽー"}

            # 処理内容 tsをHDDに移動
            ScriptBlock =
            {
                Copy-Item -LiteralPath "$env:FilePath" "E:\ts" -ErrorAction SilentlyContinue
            }
        }
        @{
            # 実行条件例 ジャンルの部分一致
            Conditional = {Get-ProgramInfoGenre -match "帝国内アニメ"}
            ScriptBlock =
            {
                # 引数を変えてエンコする, 画像を出力する等
            }
        }
        # $env:BatFileTag = ts: HDDにtsを移動, HDDの容量圧迫警告
        @{
            # 実行条件 録画後実行batタグが"ts"(tsをHDDに永久保存)のとき
            Conditional =
            {
                # 通常は$env:BatFileTagで十分
                # Get-ImmediateBatFileTagforEpgAutoAddは予約一覧にタグが反映されていなくても、自動予約登録を参照する
                # EpgTimerの予約一覧/自動予約登録で"タグ"プロパティを追加すると分かりやすい
                $env:BatFileTag -eq "ts" -Or (Get-ImmediateBatFileTagforEpgAutoAdd -eq "ts")
            }
            ScriptBlock =
            {
                # HDDにtsを移動
                Move-Item -LiteralPath $env:FilePath -Destination "E:\ts" -ErrorAction Stop

                # HDDの容量監視 -Round超過で通知
                $Text = FolderRound -Mode 'Warning' -Ext '*' -Path "E:\ts" -Round 13TB
                if (![string]::IsNullOrEmpty($Text))
                {
                    Send-Webhook -Text $Text
                    Send-BalloonTip -Icon 'Warning' -Text $Text
                }
            }
        }
        # $env:BatFileTag = enc,encremove: エンコード
        @{
            Conditional = {Get-ImmediateBatFileTagforEpgAutoAdd -in "enc","encremove"}
            ScriptBlock =
            {
                # エンコード
                # 関数 Get-ArgumentsDualMono はステレオかデュアルモノかを判別し、引数を補完する NHKニュース7やその前後の番組で確認するとよい
                # 関数 Get-ArgumentsPID はtsから必要なPIDを取得してFFmpegに渡す インターミッションで確認するとよい
                $Process = Invoke-Process -FileName 'ffmpeg.exe' -Arguments "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"env:FilePath`" -c:a libfdk_aac -vbr 5 -max_muxing_queue_size 4000 $(Get-ArgumentsDualMono -Stereo '-ac 2' -DualMono '-ac 1 -filter_complex channelsplit') -vf dejudder,fps=30000/1001:round=zero,fieldmatch=mode=pc:combmatch=full:combpel=70,yadif=mode=send_frame:parity=auto:deint=interlaced -c:v hevc_nvenc -preset:v p7 -profile:v main10 -rc:v constqp -rc-lookahead 1 -spatial-aq 0 -temporal-aq 1 -weighted_pred 0 -init_qpI 21 -init_qpP 21 -init_qpB 23 -b_ref_mode 1 -dpb_size 4 -multipass 2 -g 60 -bf 3 -pix_fmt yuv420p10le $(Get-ArgumentsPID) -movflags +faststart `"$env:USERPROFILE\Videos\encoded\$env:FileName.mp4`""
                $Process = Invoke-Process
                $Process | Format-List -Property *

                # FFmpegの終了コードが0でなければ通知する
                $Text = "Invoke-Process: エンコードに失敗 終了コード: $($Process.ExitCode)"
                if ($Process.ExitCode -ne 0)
                {
                    Send-Webhook -Text $Text
                    Send-BalloonTip -Icon 'Error' -Text $Text
                }
            }
        }
        # $env:BatFileTag = enc: HDDにmp4を移動, HDDの容量圧迫警告
        @{
            Conditional = {Get-ImmediateBatFileTagforEpgAutoAdd -eq "enc"}
            ScriptBlock =
            {
                # HDDにmp4をコピー
                Move-Item -LiteralPath "$env:USERPROFILE\Videos\encoded\$env:FileName.mp4" -Destination "E:\ts"

                # HDDの容量監視 -Round超過で通知
                $Text = FolderRound -Mode 'Warning' -Ext '*' -Path "E:\ts" -Round 13TB
                if (![string]::IsNullOrEmpty($Text))
                {
                    Send-Webhook -Text $Text
                    Send-BalloonTip -Icon 'Warning' -Text $Text
                }
            }
        }
        # $env:BatFileTag = encremove: エンコード先のローテ
        @{
            Conditional = {Get-ImmediateBatFileTagforEpgAutoAdd -eq "encremove"}
            ScriptBlock =
            {
                # 一時エンコード先のmp4を閾値を超えたら削除
                FolderRound -Mode 'Delete' -Ext "mp4" -Path "$env:USERPROFILE\Videos\encoded" -Round 50GB
            }
        }
        # tsremove,enc,encremove ≒ always: 録画先のローテ
        @{
            # 実行条件 常に実行
            Conditional = {$True}
            # 処理内容 ts容量監視
            ScriptBlock =
            {
                # 録画保存フォルダのtsを閾値を超えたら削除
                FolderRound -Mode 'Delete' -Ext "ts" -Path "$env:FolderPath" -Round 200GB
            }
        }
    )
}

#--------------------関数--------------------
# NotifiIcon.BalloonTipを表示する
function Send-BalloonTip
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Icon = "Warning",
        [String]
        $Title = "$($MyInvocation.MyCommand.Name)",
        [String]
        $Text = "WARN Send-BalloonTip:`nUse -Text <String>"
    )

    #[System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
    $NotifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::$Icon
    $NotifyIcon.BalloonTipTitle = $Title
    $NotifyIcon.BalloonTipText = $Text

    # 5000msバルーンチップを表示
    $NotifyIcon.ShowBalloonTip(5000)
    Start-Sleep -Milliseconds 5000
}

#視聴予約なら終了
if ($env:RecMode -eq 4) {
    return "視聴予約の為終了"
}
if ("${env:FilePath}" -eq $null) {
    Post -Exc $True -Toggle $True -Content "Error:${env:Title}`n[EDCB] 録画失敗によりエンコード不可" -TipIcon 'Error' -TipTitle '録画失敗'
}

# ffmpeg, &ffmpeg, .\ffmpegでは複雑な引数に対応できない
# Start-ProcessのStandardOutput, Errorはファイル出力のみ
function Invoke-Process
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Priority = 'Normal',
        [Int32]
        $Affinity =  [Convert]::ToInt32(('1' * $env:NUMBER_OF_PROCESSORS), 2), 
        [String]
        $FileName = 'powershell.exe',
        [String]
        $Arguments = 'ls',
        [String[]]
        $ArgumentList
    )
    Write-Host "DEBUG Invoke-Process`nFile: $FileName`nArguments: $Arguments`nArgumentList: $ArgumentList`n"

    #cf. https://github.com/guitarrapc/PowerShellUtil/blob/master/Invoke-Process/Invoke-Process.ps1 

    # new Process
    $ps = New-Object System.Diagnostics.Process
    $ps.StartInfo.UseShellExecute = $False
    $ps.StartInfo.RedirectStandardInput = $False
    $ps.StartInfo.RedirectStandardOutput = $True
    $ps.StartInfo.RedirectStandardError = $True
    $ps.StartInfo.CreateNoWindow = $True
    $ps.StartInfo.Filename = $FileName
    if ($Arguments)
    {
        # Windows
        $ps.StartInfo.Arguments = $Arguments
    }
    elseif ($ArgumentList)
    {
        # Linux
        $ArgumentList | ForEach-Object {
            $ps.StartInfo.ArgumentList.Add("$_")
        }
    }

    # Event Handler for Output
    $stdSb = New-Object -TypeName System.Text.StringBuilder
    $errorSb = New-Object -TypeName System.Text.StringBuilder
    $scripBlock = 
    {
        $x = $Event.SourceEventArgs.Data
        if (-not [String]::IsNullOrEmpty($x))
        {
            [System.Console]::WriteLine($x)
            $Event.MessageData.AppendLine($x)
        }
    }
    $stdEvent = Register-ObjectEvent -InputObject $ps -EventName OutputDataReceived -Action $scripBlock -MessageData $stdSb
    $errorEvent = Register-ObjectEvent -InputObject $ps -EventName ErrorDataReceived -Action $scripBlock -MessageData $errorSb

    # execution
    $ps.Start() > $Null
    $ps.PriorityClass = $Priority
    $ps.ProcessorAffinity = $Affinity
    $ps.BeginOutputReadLine()
    $ps.BeginErrorReadLine()

    # wait for complete
    $ps.WaitForExit()
    $ps.CancelOutputRead()
    $ps.CancelErrorRead()

    # verbose Event Result
    $stdEvent, $errorEvent | Out-String -Stream | Write-Verbose

    # Unregister Event to recieve Asynchronous Event output (You should call before process.Dispose())
    Unregister-Event -SourceIdentifier $stdEvent.Name
    Unregister-Event -SourceIdentifier $errorEvent.Name

    # verbose Event Result
    $stdEvent, $errorEvent | Out-String -Stream | Write-Verbose

    # Get Process result
    return [PSCustomObject]@{
        StartTime = $ps.StartTime
        StandardOutput = $stdSb.ToString().Trim()
        ErrorOutput = $errorSb.ToString().Trim()
        ExitCode = $ps.ExitCode
    }

    if ($Null -ne $process)
    {
        $ps.Dispose()
    }
    if ($Null -ne $stdEvent)
    {
        $stdEvent.StopJob()
        $stdEvent.Dispose()
    }
    if ($Null -ne $errorEvent)
    {
        $errorEvent.StopJob()
        $errorEvent.Dispose()
    }
}

# DiscordやSlackにWebhookする
function Send-Webhook
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Text = 'ERROR Send-Webhook: Use -Text <String> or -Payload <Object>',
        [System.Object]
        $Payload,
        [string]
        $WebhookUrl = $Settings.Post.WebhookUrl
    )

    if ([string]::IsNullOrEmpty($WebhookUrl))
    {
        return "ERROR Send-Webhook: Webhook URL is empty"
    }

    # $Payload が指定されていなければ $Text を使う
    if (!$Payload)
    {
        $Payload = [PSCustomObject]@{
            content = $Text
        }
    }
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -Headers @{ "Content-Type" = "application/json" } -Body ([System.Text.Encoding]::UTF8.GetBytes(($Payload | ConvertTo-Json -Depth 5)))
}

# System.Windows.Forms.NotifyIconを使う
Add-Type -AssemblyName System.Windows.Forms
$NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$ContextMenu = New-Object System.Windows.Forms.ContextMenu

# Windows PowerShellのアイコンを使用
$NotifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')

# マウスオーバー時に表示されるヒントにファイル名(64字未満)を表示
$NotifyIcon.Text = [Regex]::Replace($MyInvocation.MyCommand.Name + ": $env:FileName.ts", "^(.{63}).*$", { $args.Groups[1].Value })

# タスクトレイアイコン表示
$NotifyIcon.Visible = $True

# タスクトレイアイコンクリック時に表示されるコンテキストメニュー内のアイテムをクリックした時の動作

$MenuItemViewLog = New-Object System.Windows.Forms.MenuItem
$MenuItemViewLog.Text = "View log"
$MenuItemViewLog.add_Click({
    # ログファイルを既定のテキストエディタで開く
    &"$($Settings.Log.Path)\$env:FileName.log"
})

$MenuItemOpenScriptLoc = New-Object System.Windows.Forms.MenuItem
$MenuItemOpenScriptLoc.Text = "Open script location"
$MenuItemOpenScriptLoc.add_Click({
    # エクスプローラでこのスクリプトの場所を開く
    &explorer.exe "$PSScriptRoot"
})

$MenuItemKillFFmpeg = New-Object System.Windows.Forms.MenuItem
$MenuItemKillFFmpeg.Text = "Kill ffmpeg.exe under this process"
$MenuItemKillFFmpeg.add_Click({
    # エクスプローラでこのスクリプトの場所を開く
    $ParentProcessIds = Get-CimInstance -Class Win32_Process -Filter "Name = 'ffmpeg.exe'"
    Stop-Process -Id ($ParentProcessIds | Where-Object ParentProcessId -eq $PID).ProcessId
})

# タスクトレイアイコンクリック時に表示されるコンテキストメニュー
$ContextMenu.MenuItems.Add($MenuItemViewLog)
$ContextMenu.MenuItems.Add($MenuItemOpenScriptLoc)
$ContextMenu.MenuItems.Add($MenuItemKillFFmpeg)
$NotifyIcon.ContextMenu = $ContextMenu

# ログ取り開始
Start-Transcript -LiteralPath "$($Settings.Log.Path)\$env:FileName.log"

"#--------------------ログ--------------------"
# 古いログの削除
Get-ChildItem -LiteralPath "$($Settings.Log.Path)\" -Include *.log | Sort-Object LastWriteTime -Descending | Select-Object -Skip $Settings.Log.CntMax | ForEach-Object {
    Remove-Item -LiteralPath "$_"
    "INFO Remove-Item: $_"
}

"#--------------------ユーザ設定--------------------"
# ユーザ設定をログに記述
foreach ($line in (Get-Content -LiteralPath $PSCommandPath) -split "`n")
{
    if ($line -match '#--------------------関数--------------------')
    {
        break
    }
    $line
}

# フォルダの合計サイズを設定値以下に丸め込む関数
function FolderRound
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Mode = "Warning",
        [String]
        $Ext = "ts",
        [String]
        $Path = $env:FolderPath,
        [String]
        $Round = 10GB
    )
    # ディレクトリ内のファイルを日付順ソートで取得
    $sortTsFolder = Get-ChildItem "$Path\*.$Ext" | Sort-Object LastWriteTime

    # ディレクトリ内のファイルサイズの合計が$Roundより大きい場合実行し続ける
    for ($i = 0; ($sortTsFolder | Select-Object -Skip $i | Measure-Object -Sum Length).Sum -gt $Round; $i++)
    {
        Write-Output "WARN FolderRound: $Path is over $Round."
        if ($Mode -eq "Delete")
        {
            # Deleteモードの場合
            # 削除対象
            $removeItem = ($sortTsFolder | Select-Object -Skip $i | Select-Object -Index 0).FullName
            Remove-Item -LiteralPath $removeItem
            Write-Host "INFO Remove-Item: $removeItem"
        } elseif ($Mode -eq "Warning")
        {
            # Warningモードの場合
            # forループから抜けてfunction内に戻る
            break
        }
    }
}

# EpgTimer.exeのアセンブリを読む
[void][Reflection.Assembly]::LoadFile("$PSScriptRoot\EpgTimer.exe")

# 予約一覧にタグ $env:BatFileTag が反映されていなくても、 $env:AddKey を含む自動予約登録 EpgTimer.CtrlCmdUtil.SendEnumEpgAutoAdd.searchInfo.andKey から searchInfo.BatFileTag の一致を行う
# 録画後実行batタグ https://github.com/xtne6f/EDCB/blob/work-plus-s/Document/Readme_Mod.txt#L160-L163
function Get-ImmediateBatFileTagforEpgAutoAdd
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $AddKey = $env:AddKey,
        # サービス名 $EpgTimer.EpgAutoAddData.searchInfo.serviceList: OriginalNetworkID($env:ONID16), TransportStreamID(TSID16), ServiceID(SID16)
        [String]
        $OnTsSID10 = [Convert]::ToString("0x0$env:ONID16$env:TSID16$env:SID16", 10)
    )

    # EPG自動予約キーワードが空でなければ
    if ($AddKey -And $OnTsSID10 -ne 0)
    {
        $EpgTimer =
        @{
            CtrlCmdUtil = New-Object EpgTimer.CtrlCmdUtil
            EpgAutoAddData = New-Object Collections.Generic.List[EpgTimer.EpgAutoAddData]
        }

        # 自動予約登録条件一覧を取得する
        Write-Host "DEBUG Get-ImmediateBatFileTagforEpgAutoAdd:" ($EpgTimer.CtrlCmdUtil.SendEnumEpgAutoAdd([ref]$EpgTimer.EpgAutoAddData) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

        # 自動予約登録条件一覧からサービス名とEPG自動予約キーワードが一致する最初の項目を選ぶ
        $BatFilePath = ($EpgTimer.EpgAutoAddData | Where-Object {$OnTsSID10 -in $_.searchInfo.serviceList -And $_.searchInfo.andKey -match $AddKey}).recSetting.BatFilePath | Select-Object -Index 0

        # BatFilePathからBatFileTagを抽出
        $BatFileTag = $BatFilePath -replace (".*\*","")
        Write-Host "DEBUG Get-ImmediateBatFileTagforEpgAutoAdd: $BatFileTag"
        return $BatFileTag
    }
    else
    {
        Write-Host "WARN Get-ImmediateBatFileTagforEpgAutoAdd: -AddKey '$AddKey' and -OnTsSID10 '$OnTsSID10' are required"
        return $False
    }
}

# 録画情報からジャンルを取得
function Get-ProgramInfoGenre
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [UInt32]
        $RecInfoID = $env:RecInfoID
    )

    # EpgTimer.CtrlCmdUtil API https://github.com/xtne6f/EDCB/blob/work-plus-s/Document/Readme_Mod.txt#L630-L634
    $EpgTimer =
    @{
        CtrlCmdUtil = New-Object EpgTimer.CtrlCmdUtil
        RecFileInfo = New-Object EpgTimer.RecFileInfo
    }
    # 録画済み情報取得 https://github.com/xtne6f/EDCB/blob/work-plus-s/EpgTimer/EpgTimer/Common/CtrlCmd.cs#L616-L617
    Write-Host "DEBUG Get-ProgramInfoGenre:" ($EpgTimer.CtrlCmdUtil.SendGetRecInfo([uint32]$RecInfoID, [ref]$EpgTimer.RecFileInfo) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

    # 番組情報からジャンルを抽出し、KeywordGenreに一致させる
    return ($EpgTimer.RecFileInfo.ProgramInfo -split '\r?\n' | Select-String -Pattern "ジャンル" -Context 0,3).Context.PostContext
}

# デュアルモノ判別
function Get-ArgumentsDualMono
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Stereo,
        [String]
        $DualMono,
        [UInt32]
        $RecInfoID = $env:RecInfoID
    )

    # EpgTimer.CtrlCmdUtil
    $EpgTimer =
    @{
        CtrlCmdUtil = New-Object EpgTimer.CtrlCmdUtil
        RecFileInfo = New-Object Collections.Generic.List[EpgTimer.RecFileInfo]
    }
    # 録画済み情報取得
    [void]($EpgTimer.CtrlCmdUtil.SendGetRecInfo([uint32]$RecInfoID, [ref]$EpgTimer.RecFileInfo) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

    # 番組情報からジャンルを抽出し、KeywordGenreに一致させる
    if ($EpgTimer.RecFileInfo.ProgramInfo | Select-String -Pattern "デュアルモノ")
    {
        Write-Host "DEBUG Get-ArgumentsDualMono: $ArgPID"
        return $DualMono
    }
    else
    {
        Write-Host "DEBUG Get-ArgumentsDualMono: $ArgPID"
        return $Stereo
    }
}

# PID引数の設定
function Get-ArgumentsPID
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $FilePath = $env:FilePath
    )

    # ffprobeでcodec_type,height,idをソート
    $stream = (&"ffprobe.exe" -v quiet -analyzeduration 30M -probesize 100M -i "$FilePath" -show_entries stream=codec_type,height,id,channels -print_format json 2>&1 | ConvertFrom-Json).programs.streams
    $stream | Format-Table -Property codec_type,height,id,channels

    # 解像度の大きいVideoストリームを選ぶ
    [string[]]$ArgPID = ($stream | Where-Object {$_.codec_type -eq "video"} | Sort-Object -Property height -Descending | Select-Object -Index 0).id

    # VideoのPIDの先頭(0x1..)と一致するAudioストリームを選ぶ
    $ArgPID += ($stream | Where-Object {$_.codec_type -eq "audio" -And $_.channels -ne "0" -And $_.id -match ($ArgPID).Substring(0,3)}).id

    # FFmpeg引数のフォーマットに直す
    [string]$ArgPID = "-map i:" + ($ArgPID -join " -map i:")
    Write-Host "DEBUG Get-ArgumentsPID: $ArgPID"
    return $ArgPID
}

"#--------------------プロファイル別処理(メインルーチン)--------------------"
# $Settings.Profilesの実行条件と処理内容を回す
$Settings.Profiles | Where-Object Conditional -And ScriptBlock | ForEach-Object {
    if (Invoke-Command -ScriptBlock $_.Conditional)
    {
        Invoke-Command -ScriptBlock $_.ScriptBlock
    }
}

#タスクトレイアイコン非表示(異常終了時は実行されずトレイに亡霊が残る仕様)
$NotifyIcon.Visible = $False
$NotifyIcon.Dispose()

# ログ取り停止
Stop-Transcript
