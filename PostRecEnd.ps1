#180802
#_EDCBX_HIDE_
#視聴予約なら終了
if ($env:RecMode -eq 4) { exit }
#Powershellプロセスの優先度を高に
#(Get-Process -Id $pid).PriorityClass='High'

<#
##EDCBの設定
・指定サービスのみ録画:全サービスはPID判別が非対応
・番組情報を出力する:デュアルモノの判別
・録画情報保存フォルダを指定しない:場所が$env:FilePath(録画フォルダ)になっているため、要望があれば設定項目に加える
・録画終了後のデフォルト動作:何もしない:Backup and Syncを動かすため
・録画後動作の抑制条件なし:自動エンコを録画中も実行しないと詰まる
・録画マージン:デフォルト(前5秒後ろ2秒)でおk
・xtne6f版recname_macro.dllで半角リネーム(ZtoH):全角英数記号でffmpegがエラーを吐く多分。$SubTitle2$を使う場合はHead文字数を使用し、意図しない長いサブタイトルがヒットし、長いファイル名になってうp出来なくなる等のエラーを避ける。例:$SDYY$$SDMM$$SDDD$_$ZtoH(Title)$$Head10~(ZtoH(SubTitle2))$.ts
・EpgTimerSrvをサービス登録しない(QSV使用時):WindowsサービスからffmpegでQSVを使用しようと試みると'Failed to create Direct3D device'エラーが出る、どうやらそういう仕様らしい。

##ユーザ設定
・各自の環境に合わせる。設定項目のそれぞれの意味は記事を参照し、少なくともこの部分だけは理解して使用すること。
・特にmp4用ffmpeg引数($arg_mp4)は、Haswell以降のQSV対応のそこそこスペックのあるマシン向けの設定になっているのでそのまま使えるとは限らない。
・私がffmpegの良さげなエンコ設定を見つけた時は最初ここに反映される。

##エラーメッセージ
・No such file or directory. そもそもtsファイル無いやんけ
・[googledrive] Can not upload because it exceeds 10GB. mp4が10GBより大きくなっちゃったよ
・[mpegts] Could not find codec parameters. PID判別に失敗して余分なストリームが紛れ込んだか、引数の-analyzedurationと-probesizeが足りず解析できないか
・[aac] Unsupported channel layout. デュアルモノ-filter_complex channelsplit失敗？ffmpeg4.0で起こることを確認
・[AVBSFContext] Error parsing ADTS frame header!. ADTSのフレームヘッダ解析エラーは恐らく問題ない ExitCode:0
・[h264_qsv] Error during encoding: device failed (-17). ループしてもQSVの機嫌は戻らなかった、CPUスペックに対して背伸びした引数でないことを確認
・Unknown. おま環
#>

#====================ユーザ設定====================
#--------------------バルーンチップ表示--------------------
#0=無効、1=有効
$balloontip_toggle=1
#--------------------ログ--------------------
#0=無効、1=有効
$log_toggle=1
#ログ出力ディレクトリ
$log_path='C:\DTV\EncLog'
#ログを残す数を指定
$logcnt_max=1000
#--------------------tsフォルダサイズを一定に保つファイルの削除--------------------
#0=無効、1=有効
$ts_del_toggle=1
#録画フォルダの最大サイズを指定(0～8EB、単位:任意)
$ts_folder_max=150GB
#--------------------mp4フォルダサイズを一定に保つファイルの削除--------------------
#0=無効、1=有効
$mp4_del_toggle=1
#backup and sync用フォルダの最大サイズを指定(0～8EB、単位:任意)
$mp4_folder_max=50GB
#--------------------jpg出力--------------------
#0=無効、1=有効
$jpg_toggle=1
#連番jpgを出力するフォルダ用のディレクトリ
$jpg_path='C:\Users\sbn\Desktop\TVTest'
#jpg出力したい自動予約キーワード(全てjpg出力したい場合は$jpg_addkey=''のようにして下さい)
$jpg_addkey='フランキス|BEATLESS|夏目友人帳|キズナアイのBEATスクランブル'
#jpg用ffmpeg引数(横pxが1440のとき、$scale=',scale=1920:1080'が使用可能)(ここの引数を弄ればpng等の別の出力を行うことも可能)
function arg_jpg {
    $script:arg="-y -hide_banner -nostats -an -skip_frame nokey -i `"${env:FilePath}`" -vf yadif=0:-1:1,hqdn3d=4.0${scale} -f image2 -q:v 0 -vsync 0 `"${jpg_path}\${env:FileName}\%05d.jpg`""
}
#--------------------tsファイルサイズ判別--------------------
#通常品質
$quality_normal=24
#0=無効(通常品質のみ使用)、1=有効(通常・低品質を閾値を元に切り替える)
$tssize_toggle=1
#閾値
$tssize_max=20GB
#低品質
$quality_low=26
#--------------------エンコード--------------------
#プロセス優先度(High,Abobe,Normal,Bellow,Idle)
$Priority='Below'
#ffmpeg.exe、ffprobe.exeがあるディレクトリ
$ffpath='C:\DTV\ffmpeg'
#一時的にmp4を吐き出すディレクトリ
$tmp_folder_path='C:\DTV\tmp'
#backup and sync用ディレクトリ
$bas_folder_path='C:\DTV\backupandsync'
#25回試行してもffmpegの処理に失敗、mp4が10GBより大きい場合、tsファイル、番組情報ファイル、エンコしたmp4ファイルを退避するディレクトリ
$err_folder_path='C:\Users\sbn\Desktop'
#mp4用ffmpeg引数 使用可能:$audio_option(デュアルモノの判別)、$quality(tsファイルサイズ判別)、$pid_need(PID判別)
function arg_mp4 {
    #$script:arg="-y -hide_banner -nostats -threads 1 -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf bwdif=0:-1:1,pp=ac -global_quality ${quality} -c:v h264_qsv -preset:v veryslow -g 15 -bf 2 -refs 4 -b_strategy 1 -look_ahead 1 -look_ahead_depth 60 -pix_fmt nv12 -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${pid_need} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
    $script:arg="-y -hide_banner -nostats -threads 1 -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -crf ${quality} -preset:v fast -g 15 -bf 2 -refs 4 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${pid_need} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
}
#--------------------Twitter--------------------
#0=無効、1=有効
$tweet_toggle=1
#ruby.exe
$ruby_path='C:\Ruby24-x64\bin\ruby.exe'
#tweet.rb
$tweet_rb_path='C:\DTV\EDCB\tweet.rb'
#SSL証明書(環境変数)
$env:ssl_cert_file='C:\DTV\EDCB\cacert.pem'
#--------------------Discord--------------------
#0=無効、1=有効
$discord_toggle=1
#webhook url
$hookUrl='https://discordapp.com/api/webhooks/XXXXXXXXXX


#====================NotifyIcon====================
#System.Windows.FormsクラスをPowerShellセッションに追加
Add-Type -AssemblyName System.Windows.Forms
#NotifyIconクラスをインスタンス化
$balloon=New-Object System.Windows.Forms.NotifyIcon
#powershellのアイコンを使用
$balloon.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')
#タスクトレイアイコンのヒントにファイル名を表示
#NotifyIcon.Textが64文字を超えると例外、String.Substringの開始値~終了値が文字数を超えると例外
switch (([string]($MyInvocation.MyCommand.Name) + ":${env:FileName}.ts").Length) {
    {$_ -ge 64} {$TextLength="63"}
    {$_ -lt 64} {$TextLength="$_"}
}
$balloon.Text=([string]($MyInvocation.MyCommand.Name) + ":${env:FileName}.ts").SubString(0,$TextLength)
#タスクトレイアイコン表示
$balloon.Visible=$True
#ファイル名をタイトルバーに表示
#(Get-Host).UI.RawUI.WindowTitle=$MyInvocation.MyCommand.Name + ":${env:FileName}.ts"
#====================BalloonTip関数====================
function BalloonTip {
    if ("${balloontip_toggle}" -eq "1") {
        #特定のTipIconのみを使用可
        #[System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
        $balloon.BalloonTipIcon=[System.Windows.Forms.ToolTipIcon]::${ToolTipIcon}
        #表示するタイトル
        $balloon.BalloonTipTitle="エンコード${enc_result}"
        #表示するメッセージ
        $balloon.BalloonTipText="${env:FileName}`nts:$([math]::round(${ts_size}/1GB,2))GB mp4:$([math]::round(${mp4_size}/1MB,0))MB${err_detail}"
        #balloontip_toggle=1なら5000ミリ秒バルーンチップ表示
        $balloon.ShowBalloonTip(5000)
        #5秒待って
        Start-Sleep -Seconds 5
    }
    #タスクトレイアイコン非表示(異常終了時は実行されずトレイに亡霊が残る仕様)
    $balloon.Visible=$False
}

#====================Process====================
#ffmpeg、&ffmpeg、.\ffmpeg:ffmpegが引数を正しく認識しない(ファイル名くらいなら-f mpegtsで行けるけどもういいです)
#Start-Process ffmpeg:-NoNewWindowはWrite-Host？-RedirectStandardOutput、Errorはファイルのみ、-PassThruはExitCodeは受け取れても.StandardOutput、Errorは受け取れない仕様
function ffprocess {
    #設定
    #ProcessStartInfoクラスをインスタンス化
    $psi=New-Object System.Diagnostics.ProcessStartInfo
    #アプリケーションファイル名
    $psi.FileName="$file"
    #引数
    $psi.Arguments="$arg"
    #標準エラー出力だけを同期出力(注意:$trueは1つだけにしないとデッドロックします)
    $psi.UseShellExecute=$false
    $psi.RedirectStandardInput=$false
    $psi.RedirectStandardOutput=$false
    $psi.RedirectStandardError=$true
    $psi.WindowStyle=[System.Diagnostics.ProcessWindowStyle]::Hidden

    #実行
    #Processクラスをインスタンス化
    $p=New-Object System.Diagnostics.Process
    #設定を読み込む
    $p.StartInfo=$psi
    #プロセス開始
    $p.Start() | Out-Null
    #プロセス優先度を高に
    (Get-Process -Id $p.Id).PriorityClass="$Priority"
    #Write-Output $p.Id
    #プロセスの標準エラー出力を変数に格納(注意:WaitForExitの前に書かないとデッドロックします)
    $script:StdErr=$p.StandardError.ReadToEnd()
    #プロセス終了まで待機
    $p.WaitForExit()
    #終了コードを変数に格納
    $script:ExitCode=$p.ExitCode
    #リソースを開放
    $p.Close()
}


#====================ログ====================
#log_toggle=1ならば実行
if ("${log_toggle}" -eq "1") {
    #ログ取り開始
    Start-Transcript -LiteralPath "${log_path}\${env:FileName}.txt"
    #録画用アプリの起動数を取得
    $RecCount=(Get-Process -ErrorAction 0 "EpgDataCap_bon","TVTest").Count
    Write-Output "同時録画数:$RecCount"
    #Get-ChildItemでログフォルダのtxtファイルを取得、更新日降順でソートし、logcnt_max個飛ばし、ForEach-ObjectでRemove-Itemループ
    Get-ChildItem "${log_path}\*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -Skip ${logcnt_max} | ForEach-Object {
        Remove-Item -LiteralPath "$_"
        Write-Output "ログ削除:$_"
    }
}

#====================tsフォルダサイズを一定に保つファイルの削除====================
#ts_del_toggle=1なら実行
if ("${ts_del_toggle}" -eq "1") {
    #録画フォルダの合計サイズを変数"ts_folder_size"に指定
    $ts_folder_size=$(Get-ChildItem "${env:FolderPath}" | Measure-Object -Sum Length).Sum
    Write-Output "録画フォルダ:$([math]::round(${ts_folder_size}/1GB,2))GB"
    #録画フォルダの合計サイズがts_folder_maxGBより大きいならファイルの削除
    while ($ts_folder_size -gt $ts_folder_max) {
        #録画フォルダ内の1番古いtsファイルのファイル名を取得
        #録画フォルダ内のtsファイルに対し、最終更新年月日でソートした1番最初にくるやつ、ファイル名(拡張子なし)を取得
        $ts_del_name=$(Get-ChildItem "${env:FolderPath}\*.ts" | Sort-Object LastWriteTime | Select-Object BaseName -First 1).BaseName
        #ts、同名のts.program.txt削除
        Remove-Item -LiteralPath "${env:FolderPath}\${ts_del_name}.ts"
        Remove-Item -LiteralPath "${env:FolderPath}\${ts_del_name}.ts.program.txt"
        Remove-Item -LiteralPath "${env:FolderPath}\${ts_del_name}.ts.err"
        Write-Output "削除:${ts_del_name}.ts、ts.program.txt"
        #録画フォルダの合計サイズを取得
        $ts_folder_size=$(Get-ChildItem "${env:FolderPath}" | Measure-Object -Sum Length).Sum
        Write-Output "録画フォルダ:$([math]::round(${ts_folder_size}/1GB,2))GB"
    }
}

#====================mp4フォルダサイズを一定に保つファイルの削除====================
#mp4_del_toggle=1なら実行
if ("${mp4_del_toggle}" -eq "1") {
    #backup and sync用フォルダの合計サイズを変数"mp4_folder_size"に指定
    $mp4_folder_size=$(Get-ChildItem "${bas_folder_path}" | Measure-Object -Sum Length).Sum
    Write-Output "backup and sync用フォルダ:$([math]::round(${mp4_folder_size}/1GB,2))GB"
    #backup and sync用フォルダの合計サイズがmp4_folder_maxGBより大きいならファイルの削除
    while ($mp4_folder_size -gt $mp4_folder_max) {
        #backup and sync用フォルダ内の1番古いmp4ファイルのファイル名を取得
        #backup and sync用フォルダ内のmp4ファイルに対し、最終更新年月日でソートした1番最初にくるやつ、ファイル名(拡張子なし)を取得
        $mp4_del_name=$(Get-ChildItem "${bas_folder_path}\*.mp4" | Sort-Object LastWriteTime | Select-Object BaseName -First 1).BaseName
        #mp4削除
        Remove-Item -LiteralPath "${bas_folder_path}\${mp4_del_name}.mp4"
        Write-Output "削除:${mp4_del_name}.mp4"
        #backup and sync用フォルダの合計サイズを取得
        $mp4_folder_size=$(Get-ChildItem "${bas_folder_path}" | Measure-Object -Sum Length).Sum
        Write-Output "backup and sync用フォルダ:$([math]::round(${mp4_folder_size}/1GB,2))GB"
    }
}

#====================jpg出力====================
#jpg出力機能が有効(jpg_toggle=1)且つenv:Addkey(自動予約時のキーワード)にjpg_addkey(指定の文字)が含まれている場合は連番jpgも出力
if (("$jpg_toggle" -eq "1") -And ("$env:Addkey" -match "$jpg_addkey")) {
    #出力フォルダ作成
    New-Item "${jpg_path}\${env:FileName}" -ItemType Directory
    Write-Output "jpg出力:${env:FileName}.ts"
    #生TSの横が1920か1440か調べる
    #xml形式で扱い、tsファイル特有のwidthがメタデータの2箇所にあり2つ出力されちゃう問題を解決
    $ts_width=[xml](&"${ffpath}\ffprobe.exe" -v quiet -i "${env:FilePath}" -show_entries stream=width -print_format xml 2>&1)
    $ts_width=$ts_width.ffprobe.streams.stream.width
    #SAR比(1440x1080しか想定してないけど)によるフィルタ設定、jpg出力
    if ("${ts_width}" -eq "1440") { $scale=',scale=1920:1080' }
    #jpg用ffmpeg引数を遅延展開
    $file="${ffpath}\ffmpeg.exe"
    arg_jpg
    #ffmpegprocessを起動
    ffprocess
}


#====================tsファイルサイズ判別====================
switch ("$tssize_toggle") {
    {$_ -eq "1"} {
        #閾値$tssize_max以下なら通常品質$quality_normal、より大きいなら低品質$quality_low
        switch ((Get-ChildItem -LiteralPath "${env:FilePath}").Length) {
            {$_ -le $tssize_max} {$quality="$quality_normal"}
            {$_ -gt $tssize_max} {$quality="$quality_low"}
            {$true} {$ts_size=$_}
        }
    }
    {$_ -eq "0"} {$quality="$quality_normal"}
}
Write-Output "quality:$quality"

#====================デュアルモノの判別====================
#番組情報ファイルがありデュアルモノという文字列があればTrue、文字列がない場合はFalse、番組情報ファイルが無ければNull
if (Get-Content -LiteralPath "${env:FilePath}.program.txt" | Select-String -SimpleMatch 'デュアルモノ' -quiet) {
    $audio_option='-c:a aac -b:a 128k -filter_complex channelsplit'
} else {
    #$audio_option='-c:a aac -b:a 256k'
    $audio_option='-c:a copy -bsf:a aac_adtstoasc'
}
Write-Output "audio_option:$audio_option"

#====================PIDの判別====================
#前の番組、裏番組等の音声や映像のPIDを引数に入れないため
#-analyzeduration 30M -probesize 100Mで適切にストリームを読み込む
$StdErr=[string](&"${ffpath}\ffmpeg.exe" -hide_banner -nostats -analyzeduration 30M -probesize 100M -i "${env:FilePath}" 2>&1)
#スペース、CRを消す
$StdErr=($StdErr -replace " ","")
$StdErr=($StdErr -replace "`r","")
#if x480だけの場合はx480、else x1080だけ・x480とx1080がある場合はx1080
if (($StdErr -match 'x480') -And ($StdErr -notmatch 'x1080')) {
    $res_need='x480'
} else {
    $res_need='x1080'
}
#LFで分割して配列として格納
$StdErr=($StdErr -split "`n")
#配列を展開(映像)
foreach ($a in $StdErr) {
    #"Video:"and"${res_need}"が含まれ、'none'が含まれない行の場合実行
    if (($a -match "^(?=.*Video:)(?=.*${res_need})") -And ($a -notmatch 'none')) {
        #引数に追記
        $pid_need+=' -map i:0x'
        #PIDの部分だけ切り取り
        $pid_need+=($a -split '0x|]')[1]
    }
}
#0x1なら音声も0x1**を選ぶ
#0x2なら音声も0x2**を選ぶ
if ("$pid_need" -match '0x1') {
    $audio_need='0x1'
} elseif ("$pid_need" -match '0x2') {
    $audio_need='0x2'
}
#配列を展開(音声)
foreach ($a in $StdErr) {
    #"Audio:"and"${audio_need}"が含まれ、'0channels'が含まれない行の場合実行
    if (($a -match "^(?=.*Audio:)(?=.*${audio_need})") -And ($a -notmatch '0channels')) {
        #引数に追記
        $pid_need+=' -map i:0x'
        #PIDの部分だけ切り取り
        $pid_need+=($a -split '0x|]')[1]
    }
}
Write-Output "PID:${pid_need}"


#====================エンコード====================
#プロセス開始用の変数
$file="${ffpath}\ffmpeg.exe"
#mp4用ffmpeg引数を遅延展開
arg_mp4
#エンコードに使用する引数を表示
Write-Output "Arguments:ffmpeg $arg"
#カウントを0にリセット
$cnt=0
#終了コードが1且つループカウントが10未満までの間、エンコードを試みる
do {
    #ループカウント
    $cnt++
    #再試行時のクールタイム
    if (($cnt -ge 2) -And ($cnt -lt 5)) {
        Start-Sleep -s 60
    }
    #エンコ
    ffprocess
    #エンコ回数が1以下か、終了コードが0の場合実行(何度もループした場合にログが肥大しないため)
    #whileの後に書くのもありだけど
    if (($cnt -le 1) -Or ($ExitCode -eq 0)) {
        #プロセスの標準エラー出力をシェルの標準出力に出力
        Write-Output $StdErr
        #エンコ後mp4のサイズを出力
        $mp4_size=$(Get-ChildItem -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4").Length
        Write-Output "mp4:$([math]::round(${mp4_size}/1MB,0))MB"
    }
} while (($ExitCode -eq 1) -And ($cnt -lt 10))
#最終的なエンコード回数、終了コードを追記
Write-Output "エンコード回数:$cnt"
Write-Output "ExitCode:$ExitCode"

#====================mp4ファイルサイズ判別====================
#ffmpegの終了コード、mp4のファイルサイズによる条件分岐
if (($ExitCode -eq 1) -Or ($mp4_size -gt 10GB)) {
    #25回試行してもffmpegの終了コードが1、mp4が10GBより大きい場合、ts、ts.program.txt、mp4を退避する
    Move-Item -LiteralPath "${env:FilePath}" "${err_folder_path}"
    Move-Item -LiteralPath "${env:FilePath}.program.txt" "${err_folder_path}"
    #Move-Item -LiteralPath "${env:FilePath}.err" "${err_folder_path}"
    Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${err_folder_path}"
    Write-Output "エンコード失敗:ts=$([math]::round(${ts_size}/1GB,2))GB mp4=$([math]::round(${mp4_size}/1MB,0))MB"
    #エラー詳細
    if (("${tweet_toggle}" -eq "1") -Or ("${balloontip_toggle}" -eq "1")) {
        #ffprocessから渡された$StdErrからスペースを消す
        $StdErr=($StdErr -replace " ","")
        #$StdErrをソートし分岐
        switch ($StdErr) {
            {$_ -match 'Nosuchfileordirectory'} {$err_detail+="`nNo such file or directory."}
            {$mp4_size -gt 10GB} {$err_detail+="`n[googledrive] Can not upload because it exceeds 10GB."}
            {$_ -match 'Errorduringencoding:devicefailed'} {$err_detail+="`n[h264_qsv] Error during encoding: device failed (-17)."}
            {$_ -match 'Couldnotfindcodecparameters'} {$err_detail+="`n[mpegts] Could not find codec parameters."}
            {$_ -match 'Unsupportedchannellayout'} {$err_detail+="`n[aac] Unsupported channel layout."}
            #{$_ -match 'ErrorparsingADTSframeheader'} {$err_detail+='[AVBSFContext] Error parsing ADTS frame header!.'} #ADTSのフレームヘッダ解析エラーは恐らく問題ない ExitCode:0
            default {$err_detail="`nUnknown."}
        }
    }
    #投稿内容
    $env:content="Error:${env:FileName}.ts`n${err_detail}`n同時録画数:$RecCount"
    #Twitter警告
    if ("${tweet_toggle}" -eq "1") {
        &"${ruby_path}" "${tweet_rb_path}"
        #Start-Process "${ruby_path}" "${tweet_rb_path}" -WindowStyle Hidden -Wait
    }
    #Discord警告
    if ("${discord_toggle}" -eq "1") {
        $payload=[PSCustomObject]@{
            content = $env:content
        }
        $payload=($payload | ConvertTo-Json)
        $payload=[System.Text.Encoding]::UTF8.GetBytes($payload)
        Invoke-RestMethod -Uri $hookUrl -Method Post -Body $payload
    }
    #NotifyIcon
    $ToolTipIcon='Error'
    $enc_result='失敗'
    BalloonTip
} elseif ($mp4_size -le 10GB) {
    #それ以外でmp4が10GB以下ならmp4をbas_folder_pathに投げる
    Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${bas_folder_path}"
    Write-Output "エンコード終了:ts=$([math]::round(${ts_size}/1GB,2))GB mp4=$([math]::round(${mp4_size}/1MB,0))MB"
    #NotifyIcon
    $ToolTipIcon='Info'
    $enc_result='終了'
    BalloonTip
}