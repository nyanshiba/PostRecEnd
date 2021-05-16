#_EDCBX_HIDE_
#Requires -Version 5
#Requires -PSEdition Desktop



#ffmpeg.exe、ffprobe.exeがあるディレクトリ
$ffpath='C:\bin\ffmpeg'

#--------------------ログ--------------------
#ログ出力ディレクトリ
$log_path='C:\logs\PostRecEnd'
#ログを残す数
$logcnt_max=500

#--------------------tsの自動削除--------------------
#閾値を超過した場合、Warning=容量警告、Delete=tsを自動削除
$TsFolderRound="Delete"
#録画フォルダの上限
$ts_folder_max=150GB

#--------------------mp4の自動削除--------------------
#閾値を超過した場合、Warning=容量警告、Delete=tsを自動削除
$Mp4FolderRound="Delete"
#mp4用フォルダの上限
$mp4_folder_max=50GB

#--------------------tsファイルサイズ判別--------------------
#映像の品質引数をtsファイルサイズによって適応的に変える($ArgQual)
#適応品質機能 $False=無効(エンコード引数内に記述)、$True=通常・低品質を閾値で切り替え
$tssize_toggle=$True
#閾値
$tssize_max=20GB #くらいがおすすめ
#通常品質(LA-ICQ:27,x265:25)
$quality_normal='-init_qpI 21 -init_qpP 21 -init_qpB 23'
#低品質(LA-ICQ:30,x265:27)
$quality_low='-init_qpI 23 -init_qpP 23 -init_qpB 26'

#--------------------デュアルモノの判別--------------------
#音声引数をデュアルモノか否かで変える($ArgAudio)
#デュアルモノ
$audio_dualmono='-strict -2 -c:a aac -b:a 128k -aac_coder twoloop -ac 1 -max_muxing_queue_size 4000 -filter_complex channelsplit'
#通常
$audio_normal='-strict -2 -c:a aac -b:a 256k -aac_coder twoloop -ac 2 -max_muxing_queue_size 4000' #失敗しない、ただし再エンコ
#$audio_normal='-strict -2 -c:a flac -ac 2 -max_muxing_queue_size 4000'
#$audio_normal='-c:a copy' #失敗する上ExitCode=0
#$audio_normal='-c:a copy -bsf:a aac_adtstoasc' 失敗する上ExitCode=0

#--------------------PIDの判別--------------------
#必要なPIDを取得し-map引数に加える($ArgPid)

#--------------------エンコード--------------------
#mp4の一時出力ディレクトリ
$tmp_folder_path='C:\Rec\tmp'
#mp4保存(Backup and Sync、ローカル保存)用ディレクトリ
$mp4_folder_path='C:\Rec\mp4'
#例外ディレクトリ(ループしてもffmpegの処理に失敗、mp4が10GBより大きい場合 etc…にts、ts.program.txt、ts.err、mp4を退避)
$err_folder_path='C:\Rec\Err'
#mp4の10GBファイルサイズ上限 $True=有効 $False=無効
$googledrive=$True
#mp4用ffmpeg引数 
<#
-File: 実行ファイルのパス
-Arg: 引数
    $ArgAudio(エンコ失敗しない為に必須)
    $ArgQual(エンコード引数内に記述し品質やビットレートを固定する場合不要)
    $ArgPid(エンコ失敗しない為に必須)
-Priority: プロセス優先度 MSDNのProcess.PriorityClass参照 (Normal,Idle,High,RealTime,BelowNormal,AboveNormal) ※必須ではない
-Affinity: 使用する論理コアの指定 MSDNのProcess.ProcessorAffinity参照 コア5(10000)～12(100000000000)を使用=0000111111110000(2進)=4080(10進)=0xFF0(16進) ※必須ではない

NVEnc H.264 VBR MinQP
-Arg "-y -nostats -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1 -c:v h264_nvenc -preset:v slow -profile:v high -rc:v vbr_minqp -rc-lookahead 32 -spatial-aq 1 -aq-strength 1 -qmin:v 23 -qmax:v 25 -b:v 1500k -maxrate:v 3500k -pix_fmt yuv420p ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
QSV H.264 LA-ICQ
-Arg "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -ss 5 -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac,hqdn3d=2.0 -global_quality ${ArgQual} -c:v h264_qsv -preset:v veryslow -g 300 -bf 6 -refs 4 -b_strategy 1 -look_ahead 1 -look_ahead_depth 60 -pix_fmt nv12 -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x265 fast
-Arg "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -crf ${ArgQual} -preset:v fast -g 15 -bf 2 -refs 4 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x265 fast bel9r inspire
-Arg "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -preset:v fast -x265-params crf=${ArgQual}:rc-lookahead=40:psy-rd=0.3:keyint=15:no-open-gop:bframes=2:rect=1:amp=1:me=umh:subme=3:ref=3:rd=3 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x264 placebo by bel9r
-Arg "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx264 -preset:v placebo -x264-params crf=${ArgQual}:rc-lookahead=60:qpmin=5:qpmax=40:qpstep=16:qcomp=0.85:mbtree=0:vbv-bufsize=31250:vbv-maxrate=25000:aq-strength=0.35:psy-rd=0.35:keyint=300:bframes=6:partitions=p8x8,b8x8,i8x8,i4x4:merange=64:ref=4:no-dct-decimate=1 -pix_fmt yuv420p -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
#>
$VideoEncode =
({
    #hevc_nvenc constqp (qpI,P,Bはtsファイルサイズ判別を参照)
    Invoke-Process -File "${ffpath}\ffmpeg.exe" -Arg "-y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" $ArgAudio -vf pullup,dejudder,idet=intl_thres=1.38:prog_thres=1.5,yadif=mode=send_field:parity=auto:deint=interlaced,fps=fps=30000/1001:round=zero -c:v hevc_nvenc -preset:v p7 -profile:v main10 -rc:v constqp -rc-lookahead 1 -spatial-aq 0 -temporal-aq 1 -weighted_pred 0 $ArgQual -b_ref_mode 1 -dpb_size 4 -multipass 2 -g 60 -bf 3 -pix_fmt yuv420p10le $ArgPid -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`"" -Priority 'BelowNormal' -Affinity '0xFFF'
})

#--------------------Post--------------------
#$False=Error時のみ、$True=常時 Twitter、DiscordにPost
$InfoPostToggle=$False

#Discord機能 $False=無効、$True=有効
$discord_toggle=$True
#webhook url
$hookUrl='https://discordapp.com/api/webhooks/XXXXXXXXXX'

#BalloonTip機能 $False=無効、$True=有効
$balloontip_toggle=$True


#--------------------関数--------------------
function Post
{
    param
    (
        [bool]$exc,
        [bool]$toggle,
        [string]$content,
        [string]$tipicon,
        [string]$tiptitle
    )
    #例外。自動削除が有効の場合、ts、ts.program.txt、ts.err、mp4を退避
    if ($exc)
    {
        if ($TsFolderRound)
        {
            Move-Item -LiteralPath "${env:FilePath}" "${err_folder_path}" -ErrorAction SilentlyContinue
            Move-Item -LiteralPath "${env:FilePath}.program.txt" "${err_folder_path}" -ErrorAction SilentlyContinue
            Move-Item -LiteralPath "${env:FilePath}.err" "${err_folder_path}" -ErrorAction SilentlyContinue
        }
        if ($Mp4FolderRound)
        {
            Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${err_folder_path}" -ErrorAction SilentlyContinue
        }
    }
    #Error時だけでなく、Info時もPostできるようにするトグル
    if ($toggle)
    {
        #Discord警告
        if ($discord_toggle)
        {
            $payload = [PSCustomObject]@{
                content = $content
            }
            $payload = ($payload | ConvertTo-Json)
            $payload = [System.Text.Encoding]::UTF8.GetBytes($payload)
            Invoke-RestMethod -Uri $hookUrl -Method Post -Headers @{ "Content-Type" = "application/json" } -Body $payload
        }
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


"DEBUG ArgPid: $ArgPid"

"#--------------------エンコード--------------------"
#終了コードが1且つループカウントが50未満までの間、エンコードを試みる
for ($i = 0; $i -lt 50 -And $ExitCode -ne 0; $i++)
{
    #再試行時からクールタイムを追加
    if ($i -gt 0)
    {
        Start-Sleep -s 60
    }
    
    #エンコ mp4用ffmpeg引数を遅延展開
    Invoke-Command -Command $VideoEncode
    "DEBUG ExitCode:$ExitCode"

    #エンコ1回目と成功時(ExitCode:0)のログだけで十分
    if ($i -eq 0 -Or $ExitCode -eq 0)
    {
        #プロセスの標準エラー出力をシェルの標準出力に出力
        $StdErr
    }
}

#タスクトレイアイコン非表示(異常終了時は実行されずトレイに亡霊が残る仕様)
$NotifyIcon.Visible = $False
$NotifyIcon.Dispose()

# ログ取り停止
Stop-Transcript
