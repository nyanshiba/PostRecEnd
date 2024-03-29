#_EDCBX_HIDE_
#Requires -Version 5
#Requires -PSEdition Desktop

param
(
    # $Settings.Profiles.Conditional に条件式を書いておけば、そのScriptBlockだけ呼び出して使用できる
    [String]$Conditional
)

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
    MenuItem =
    @(
        @{
            Text = "View Log"
            Click =
            {
                # ログファイルを既定のテキストエディタで開く
                &"$($Settings.Log.Path)\$env:FileName.log"
            }
        }
        @{
            Text = "Open script location"
            Click =
            {
                # エクスプローラでこのスクリプトの場所を開く
                &explorer.exe $PSScriptRoot
            }
        }
        @{
            Text = "Kill ffmpeg.exe under this process"
            Click =
            {
                # エクスプローラでこのスクリプトの場所を開く
                # $ParentProcessIds = Get-CimInstance -Class Win32_Process -Filter "Name = 'ffmpeg.exe'"
                # Stop-Process -Id ($ParentProcessIds | Where-Object ParentProcessId -eq $PID).ProcessId
                
                # 信頼性に乏しいのでタスクマネージャを開いておく
                &taskmgr.exe
            }
        }
    )
    Post =
    @{
        # Webhook Url
        WebhookUrl = "https://discord.com/api/webhooks/XXXXXXXXXX"
    }
    Profiles =
    @(
        @{
            # powershell.exe PostRecEnd.ps1 -Conditional "ManualEncode"
            Conditional = {$Conditional -eq "ManualEncode"}
            ScriptBlock =
            {
            }
        }
        @{
            # EpgTimerから呼び出し(EpgTimerが追加する環境変数なら何でもよい)の場合、最初に実行
            Conditional = {$env:RecMode}
            ScriptBlock =
            {
                #視聴予約なら終了
                if ($env:RecMode -eq 4)
                {
                    Send-BalloonTip -Icon 'Info' -Text "視聴予約`n$env:Title"
                    exit
                }
                # ファイル名が空なら
                elseif ([string]::IsNullOrEmpty($env:FilePath))
                {
                    $Text = "❌録画失敗`n$env:Result`n$env:Title"
                    Send-Webhook -Text ('<@379045222451249154> ' + $Text)
                    Send-BalloonTip -Icon 'Error' -Text $Text
                }
            }
        }
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
            Conditional = {(Get-ProgramInfoGenre) -match "帝国内アニメ"}
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
                # 通常は $env:BatFileTag で十分
                # Get-ImmediateBatFileTagforEpgAutoAdd を使うことで、予約一覧にタグが反映されていなくても自動予約登録を参照する
                # EpgTimerの予約一覧/自動予約登録で"タグ"プロパティを追加すると分かりやすい
                (Get-ImmediateBatFileTagforEpgAutoAdd) -eq "ts" -Or $env:Scrambles -ne 0
            }
            ScriptBlock =
            {
                "INFO `$Settings.Profiles.Conditional: ts"
                # HDDにtsを移動
                Move-Item -LiteralPath $env:FilePath -Destination "E:\ts"

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
            Conditional = {(Get-ImmediateBatFileTagforEpgAutoAdd) -in "enc","encremove"}
            ScriptBlock =
            {
                "INFO `$Settings.Profiles.Conditional: enc, encremove"
                
                # obs64.exe が実行中なら4時間までは待機
                for ($min = 0; $min -le 240 -And (Get-Process -Name obs64 -ErrorAction SilentlyContinue).Count -ne 0; $min++)
                {
                    "DEBUG $(Get-Date) obs64.exeが実行中です"
                    Start-Sleep -Seconds 60
                }
                
                # エンコード
                # 関数 Get-ArgumentsDualMono はデュアルモノかステレオか再エンコード不要かを判別して引数に補完する。既定値はあるが、-Copy, -Stereo, -DualMonoそれぞれ好みの引数を指定してもよい。デュアルモノはNHKニュース7、で確認するとよい。
                # 関数 Get-ArgumentsPID はtsから必要なPIDを取得して引数に補完する。インターミッションで確認するとよい。
                # FFmpegのエンコード設定は https://github.com/nyanshiba/best-ffmpeg-arguments#hevc_nvenc に基づいている。hevc_nvencの例では、300M-1.2GB程度でVMAF97辺りのアニメエンコードが可能。
                $Process = Invoke-Process -FileName 'ffmpeg.exe' -Arguments "-loglevel -repeat+level+info -y -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"$env:FilePath`" -f mpegts -scan_all_pmts $(Get-ArgumentsDualMono -Copy '-bsf:a aac_adtstoasc -c:a copy' -Stereo '-c:a libfdk_aac -ac 2 -vbr 5 -max_muxing_queue_size 4000' -DualMono '-c:a libfdk_aac -ac 1 -vbr 5 -max_muxing_queue_size 4000 -filter_complex channelsplit') -vf dejudder,fps=30000/1001:round=zero,fieldmatch=mode=pc:combmatch=full:mchroma=0:cthresh=35:combpel=47,yadif=mode=send_frame:parity=auto:deint=interlaced -c:v hevc_nvenc -preset:v p7 -profile:v main10 -rc:v constqp -rc-lookahead 1 -spatial-aq 0 -temporal-aq 1 -weighted_pred 0 -init_qpI 21 -init_qpP 21 -init_qpB 23 -b_ref_mode 1 -dpb_size 4 -multipass 2 -g 60 -bf 3 -pix_fmt yuv420p10le $(Get-ArgumentsPID) -movflags +faststart `"$env:USERPROFILE\Videos\encoded\$env:FileName.mp4`""

                # 逆テレシネとフィールド補間に欠かせない fieldmatch の行はうるさいので削除してから出力する
                $Process.ErrorOutput = $Process.ErrorOutput -split '\r?\n' | Select-String -NotMatch "Parsed_fieldmatch_2|aac bitstream error|Last message repeated" | Out-String -Width 1024
                $Process | Format-List -Property *

                # エンコード失敗時
                if ($Process.ExitCode -ne 0)
                {
                    # $Settings.Post.WebhookUrl のチャンネルにユーザIDへメンション
                    Send-Webhook -Payload ([PSCustomObject]@{
                        content = "<@379045222451249154> ❌エンコード失敗`n$env:FileName"
                        username = "PostRecEnd.ps1"
                        avatar_url = "https://cdn.discordapp.com/emojis/912607140571668490.png"
                        embeds =
                        @(
                            @{
                                title = "PostRecEnd.ps1"
                                description = $env:Title
                                color = 0xc21e54
                                fields =
                                @(
                                    @{
                                        name = "Drops"
                                        value = $env:Drops
                                        inline = 'true'
                                    }
                                    @{
                                        name = "Scrambles"
                                        value = $env:Scrambles
                                        inline = 'true'
                                    }
                                    @{
                                        name = "Result"
                                        value = $env:Result
                                        inline = 'true'
                                    }
                                    @{
                                        name = "ErrorOutput"
                                        value = (
                                            $Process.ErrorOutput -split '\r?\n' | Select-String -Pattern "\[(error|fatal)\]|Conversion failed" | Select-Object -Last 5 | ForEach-Object {
                                                $_ -replace ' @ \w{16}', ''
                                            } | Out-String
                                        )
                                        inline = 'false'
                                    }
                                    @{
                                        name = "ExitCode"
                                        value = $Process.ExitCode
                                        inline = 'true'
                                    }
                                    @{
                                        name = "FileSize"
                                        value = (Get-Item -LiteralPath "$env:USERPROFILE\Videos\encoded\$env:FileName.mp4").Length
                                        inline = 'true'
                                    }
                                )
                            }
                        )
                    })

                    # HDDにts, ログを退避
                    Move-Item -LiteralPath $env:FilePath -Destination "E:\ts"
                    Copy-Item -LiteralPath "$($Settings.Log.Path)\$env:FileName.log" "E:\ts"

                    # シイタケ
                    Send-BalloonTip -Icon 'Error' -Text "Invoke-Process: エンコード失敗`n$env:FileName.mp4`n終了コード: $($Process.ExitCode)"
                }
                # エンコード成功時
                elseif ($Process.ExitCode -eq 0)
                {
                    # エンコード後のファイルサイズ
                    $Length = (Get-ChildItem -LiteralPath "$env:USERPROFILE\Videos\encoded\$env:FileName.mp4").Length
                    switch ([Math]::Truncate([Math]::Log($Length, 1024)))
                    {
                        '2' {$Length = "{0:n1} MB" -f ($Length / 1MB)}
                        '3' {$Length = "{0:n2} GB" -f ($Length / 1GB)}
                    }

                    # 正常にエンコード終了したことをバルーンチップで通知
                    Send-BalloonTip -Icon 'Info' -Text "Invoke-Process: エンコード終了`n$env:FileName.mp4`n$Length"
                }
            }
        }
        # $env:BatFileTag = enc: HDDにmp4を移動, HDDの容量圧迫警告
        @{
            Conditional = {(Get-ImmediateBatFileTagforEpgAutoAdd) -eq "enc"}
            ScriptBlock =
            {
                "INFO `$Settings.Profiles.Conditional: enc"
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
            Conditional = {(Get-ImmediateBatFileTagforEpgAutoAdd) -eq "encremove"}
            ScriptBlock =
            {
                "INFO `$Settings.Profiles.Conditional: encremove"
                # 一時エンコード先のmp4を閾値を超えたら削除
                FolderRound -Mode 'Delete' -Ext "mp4" -Path "$env:USERPROFILE\Videos\encoded" -Round 50GB
            }
        }
        # tsremove,enc,encremove ≒ always: 録画先のローテ
        @{
            # 実行条件 常に実行
            Conditional = {$env:TSID16}
            # 処理内容 ts容量監視
            ScriptBlock =
            {
                "INFO `$Settings.Profiles.Conditional: tsremove, enc, encremove (always)"
                # 録画保存フォルダのtsを閾値を超えたら削除
                FolderRound -Mode 'Delete' -Ext "ts" -Path "$env:FolderPath" -Round 200GB
            }
        }
    )
}

#--------------------関数--------------------
# 効率よく同時エンコードするため、実行中のプロセスが使用するスレッドを均等割りする
function Get-EquallyDividedProcessorAffinity
{
    param
    (
        [int]$Threads = $env:NUMBER_OF_PROCESSORS,
        [string]$ProcessName = 'ffmpeg'
    )

    # 同じ名前のプロセスを取得
    $ProcessList = Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension("$ProcessName")) -ErrorAction SilentlyContinue

    # プロセスがない
    if (!$ProcessList)
    {
        Write-Host "WARN Get-EquallyDividedProcessorAffinity: $('1' * $env:NUMBER_OF_PROCESSORS)"
        return $False
    }

    # スレッドをプロセス数で割る
    $ThreadsforProcess = $Threads / $ProcessList.Count

    # 同じ名前のプロセスすべてにスレッド数を均等に割り振る
    for ($i = 0; $i -lt $ProcessList.Count; $i++)
    {
        # 00000000 -repace [Regex](0{1*4})0{4}, 11110000 ？？？
        $Base2Affinity = ('0' * $Threads) -replace "(?<=^0{$($i * $ThreadsforProcess)})0{$ThreadsforProcess}", ('1' * $ThreadsforProcess)
        Write-Host "DEBUG Get-EquallyDividedProcessorAffinity: $Base2Affinity"
        $ProcessList[$i].ProcessorAffinity = [Convert]::ToInt32($Base2Affinity, 2)
    }
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
        [Bool]
        $Affinity =  $True, 
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
    if ($Affinity)
    {
        Get-EquallyDividedProcessorAffinity -ProcessName $FileName
    }
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
        [UInt64]
        $OnTsSID10 = [Convert]::ToString("0x0$env:ONID16$env:TSID16$env:SID16", 10)
    )

    # EPG自動予約キーワードが空でなければ
    $EpgTimer =
    @{
        CtrlCmdUtil = New-Object EpgTimer.CtrlCmdUtil
        EpgAutoAddData = New-Object Collections.Generic.List[EpgTimer.EpgAutoAddData]
    }

    # 予約キーワードによる録画ならば
    if ($AddKey)
    {
        Write-Host "DEBUG Get-ImmediateBatFileTagforEpgAutoAdd: `$env:AddKey: $AddKey"

        # 自動予約登録条件一覧を取得する
        [void]($EpgTimer.CtrlCmdUtil.SendEnumEpgAutoAdd([ref]$EpgTimer.EpgAutoAddData) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

        # 自動予約登録条件一覧からサービス名と予約キーワードが一致する最初の項目を選ぶ
        $BatFilePath = ($EpgTimer.EpgAutoAddData | Where-Object {$OnTsSID10 -in $_.searchInfo.serviceList -And $_.searchInfo.andKey -match $AddKey}).recSetting.BatFilePath | Select-Object -Index 0

        # BatFilePathからBatFileTagを抽出
        $BatFileTag = $BatFilePath -replace (".*\*","")

        # 自動予約登録の BatFileTag があれば優先する
        if (![string]::IsNullOrEmpty($BatFileTag))
        {
            Write-Host "DEBUG Get-ImmediateBatFileTagforEpgAutoAdd: $BatFileTag"
            return $BatFileTag
        }
    }

    # $env:BatFileTag があれば代わりに
    if (![string]::IsNullOrEmpty($env:BatFileTag))
    {
        Write-Host 'DEBUG Get-ImmediateBatFileTagforEpgAutoAdd: $env:BatFileTag'
        return $env:BatFileTag
    }
    # BatFileTagはいずれも設定されていない
    else
    {
        Write-Host "DEBUG Get-ImmediateBatFileTagforEpgAutoAdd: BatFileTag is not set"
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
    [void]($EpgTimer.CtrlCmdUtil.SendGetRecInfo([uint32]$RecInfoID, [ref]$EpgTimer.RecFileInfo) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

    # 番組情報からジャンル辺りを抽出する
    if (![string]::IsNullOrEmpty($EpgTimer.RecFileInfo.ProgramInfo))
    {
        return ($EpgTimer.RecFileInfo.ProgramInfo -split '\r?\n' | Select-String -Pattern "ジャンル" -Context 0,3).Context.PostContext
    }
    else
    {
        Write-Host "WARN Get-ProgramInfoGenre: EpgTimerSrv設定で「番組情報を出力する」が無効か、ts.program.txtがありません"
        return $False
    }
}

# デュアルモノと-bsf:a aac_adtstoasc -c:a copyが失敗する番組を判別して、適切な引数を返す
function Get-ArgumentsDualMono
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Copy = '-bsf:a aac_adtstoasc -c:a copy',
        [String]
        $Stereo = '-c:a aac -ac 2 -b:a 256k -max_muxing_queue_size 4000',
        [String]
        $DualMono = '-c:a aac -ac 1 -b:a 128k -max_muxing_queue_size 4000 -filter_complex channelsplit',
        [String]
        $FilePath = $env:FilePath,
        [UInt32]
        $RecInfoID = $env:RecInfoID,
        [UInt64]
        $pgID = "0x0$env:ONID16$env:TSID16$env:SID16$env:EID16"
    )

    # EpgTimer.CtrlCmdUtil
    $EpgTimer =
    @{
        CtrlCmdUtil = New-Object EpgTimer.CtrlCmdUtil
        RecFileInfo = New-Object EpgTimer.RecFileInfo
        EpgEventInfo = New-Object EpgTimer.EpgEventInfo
    }
    # 録画済み情報取得
    Write-Host "DEBUG Get-ArgumentsDualMono: use SendGetRecInfo (ts.program.txt)"
    [void]($EpgTimer.CtrlCmdUtil.SendGetRecInfo([uint32]$RecInfoID, [ref]$EpgTimer.RecFileInfo) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)

    # EpgTimerSrv設定 番組情報を出力する が無効の場合, ts.program.txt がない場合
    if ([string]::IsNullOrEmpty($EpgTimer.RecFileInfo.ProgramInfo))
    {
        # 指定イベントの番組情報を取得する
        # 録画直後は SendEnumPgArc に含まれない
        Write-Host "DEBUG Get-ArgumentsDualMono: use SendGetPgInfo"
        [void]($EpgTimer.CtrlCmdUtil.SendGetPgInfo([uint64]$pgID, [ref]$EpgTimer.EpgEventInfo) -eq [EpgTimer.ErrCode]::CMD_SUCCESS)
        Write-Host "DEBUG Get-ArgumentsDualMono: ES_multi_lingual_flag:" $EpgTimer.EpgEventInfo.AudioInfo.componentList.ES_multi_lingual_flag
    }

    # ES_multi_lingual_flagか番組情報からデュアルモノかを判断して、適切な引数を返す
    if ($EpgTimer.EpgEventInfo.AudioInfo.componentList.ES_multi_lingual_flag -eq 1 -Or ($EpgTimer.RecFileInfo.ProgramInfo | Select-String -Pattern "デュアルモノ"))
    {
        Write-Host "DEBUG Get-ArgumentsDualMono: $DualMono"
        return $DualMono
    }
    else
    {
        # 再エンコしない
        Write-Host "DEBUG Get-ArgumentsDualMono: $Copy"
        return $Copy
    }
}

# PID引数の設定
# https://www.tele.soumu.go.jp/horei/reiki_honbun/a72ab04601.html PMTの構成
function Get-ArgumentsPID
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $FilePath = $env:FilePath,
        [Bool]
        $Video = $True,
        [Bool]
        $Audio = $True
    )

    # ffprobeで ストリームPID, ストリームの種類, 映像の高さ, 音声のチャンネル数 をソート
    # -analyzeduration...がないと、videoのheightやaudioのchanelsが0になる場合がインターミッションで確認できる
    $streams = (&"ffprobe.exe" -v quiet -analyzeduration 30M -probesize 100M -i "$FilePath" -show_entries stream=id,codec_type,height,channels -print_format json 2>&1 | ConvertFrom-Json)
    $streams = $streams.programs.streams | Select-Object -Property id,codec_type,height,channels | Where-Object codec_type -in 'video','audio'
    Write-Host ($streams | Out-String -Width 1024)

    # 解像度の大きいvideoストリームのPID(エレメンタリーPID？)上位4bit分(0x1)
    $prefix = ($streams | Sort-Object -Property height -Descending | Select-Object -Index 0).id.Substring(0,3)
    # prefixが同じvideo, audioをソートし、0 channelsなaudioは除外
    $streams = $streams | Where-Object {$_.id -match $prefix -And $_.channels -ne "0"}

    # 映像のPIDのみを返す Get-ArgumentsPID -Video $True -Audio $False
    if ($Video -And !$Audio)
    {
        $streams = $streams | Where-Object codec_type -eq 'video'
    }
    # 音声のPIDのみ
    elseif (!$Video -And $Audio)
    {
        $streams = $streams | Where-Object codec_type -eq 'audio'
    }
    # 映像・音声のPID(デフォルト) 何もしない

    # FFmpeg引数のフォーマットに直す
    [string]$ArgPID = "-map i:" + ([string[]]$streams.id -join " -map i:")
    Write-Host "DEBUG Get-ArgumentsPID (-Video `$$Video -Audio `$$Audio): $ArgPID"
    return $ArgPID
}

# NotifiIcon.BalloonTipを表示する
function Send-BalloonTip
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Icon = "Warning",
        [String]
        $Title = $PSCommandPath.Replace("$PSScriptRoot\",''),
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

# MenuItem内容がバックグラウンド実行できるようにイベント登録
function Register-ContextMenuItem
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $Text,
        [ScriptBlock]
        $Click
    )

    $MenuItem = New-Object System.Windows.Forms.MenuItem
    $MenuItem.Text = $_.Text
    # $MenuItem.add_Click([ScriptBlock]$_.add_Click)
    # Get-EventSubscriber
    Write-Host "DEBUG Invoke-NotifyIcon MenuItem Event:" (Register-ObjectEvent -InputObject $MenuItem -EventName Click -Action $_.Click).Name
    return $MenuItem
}


# System.Windows.Forms.NotifyIcon, ContextMenu, MenuItemを使う
Add-Type -AssemblyName System.Windows.Forms
$NotifyIcon = New-Object System.Windows.Forms.NotifyIcon

# Windows PowerShellのアイコンを使用
$NotifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')

# マウスオーバー時に表示されるヒントにファイル名(64字未満)を表示
$NotifyIcon.Text = [Regex]::Replace($MyInvocation.MyCommand.Name + ": $env:FileName.ts", "^(.{63}).*$", { $args.Groups[1].Value })

# タスクトレイアイコン表示
$NotifyIcon.Visible = $True

# タスクトレイアイコン右クリック時に表示されるコンテキストメニュー
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$NotifyIcon.ContextMenu = $ContextMenu

# コンテキストメニュー内容を設定
$Settings.MenuItem | ForEach-Object {
    $ContextMenu.MenuItems.Add((
        Register-ContextMenuItem -Text $_.Text -Click ([ScriptBlock]$_.Click)
    ))
}

# ログ取り開始
Start-Transcript -LiteralPath "$($Settings.Log.Path)\$env:FileName.log"
ls env:

"#--------------------ログローテ--------------------"
# 古いログの削除
Get-ChildItem -LiteralPath "$($Settings.Log.Path)\" -Include *.log,*.txt | Sort-Object LastWriteTime -Descending | Select-Object -Skip $Settings.Log.CntMax | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName
    "INFO Remove-Item: $_"
}

# ユーザ設定をログに記述
foreach ($line in (Get-Content -LiteralPath $PSCommandPath) -split "`n")
{
    if ($line -match '#--------------------関数--------------------')
    {
        break
    }
    $line
}

"#--------------------プロファイル別処理(メインルーチン)--------------------"
# $Settings.Profilesの実行条件と処理内容を回す
$Settings.Profiles | Where-Object {$_.Conditional -And $_.ScriptBlock} | ForEach-Object {
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
