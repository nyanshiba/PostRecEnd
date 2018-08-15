#180815
#_EDCBX_HIDE_
#Powershellプロセスの優先度を高に
#(Get-Process -Id $pid).PriorityClass='High'

#####################ユーザ設定####################################################################################################

<#
エラーメッセージ
・No such file or directory. そもそもtsファイル無いやんけ
・[googledrive] Can not upload because it exceeds 10GB. mp4が10GBより大きくなっちゃったよ
・[mpegts] Could not find codec parameters. PID判別に失敗して余分なストリームが紛れ込んだか、引数の-analyzedurationと-probesizeが足りず解析できないか
・[aac] Unsupported channel layout. デュアルモノ-filter_complex channelsplit失敗？ffmpeg4.0で起こることを確認
・[AVBSFContext] Error parsing ADTS frame header!. -c:a copy時にこのADTSのフレームヘッダ解析エラーが2度出ると問題がある ExitCode:0
・[h264_qsv] Error during encoding: device failed (-17). ループしてもQSVの機嫌は戻らなかった、CPUスペックに対して背伸びした引数でないことを確認
・Unknown. おま環
#>

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
#--------------------tsの自動削除--------------------
#0=無効(tsをローカルに残す)、1=有効(tsを自動削除)
$ts_del_toggle=1
#録画フォルダの最大サイズを指定(0～8EB、単位:任意)
$ts_folder_max=150GB
#--------------------mp4の自動削除--------------------
#0=無効(mp4をローカルに残す)、1=有効(mp4を自動削除)
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
    $script:arg="-y -hide_banner -nostats -an -skip_frame nokey -i `"${env:FilePath}`" -vf bwdif=0:-1:1,pp=ac,hqdn3d=2.0${scale} -f image2 -q:v 0 -vsync 0 `"${jpg_path}\${env:FileName}\%05d.jpg`""
}
#--------------------tsファイルサイズ判別--------------------
#通常品質(LA-ICQ:29,x265:25)
$quality_normal=27
#0=無効(通常品質のみ使用)、1=有効(通常・低品質を閾値を元に切り替える)
$tssize_toggle=1
#閾値
$tssize_max=20GB
#低品質(LA-ICQ:31,x265:27)
$quality_low=30
#--------------------エンコード--------------------
#プロセス優先度 (Normal,Idle,High,RealTime,BelowNormal,AboveNormal) Process.PriorityClass参照
$Priority='BelowNormal'
#使用する論理コアの指定 コア5(10000)～12(100000000000)を使用=0000111111110000(2進)=4080(10進)=0x0FF0(16進) Process.ProcessorAffinity参照
$Affinity='0x0FF0'
#ffmpeg.exe、ffprobe.exeがあるディレクトリ
$ffpath='C:\DTV\ffmpeg'
#一時的にmp4を吐き出すディレクトリ
$tmp_folder_path='C:\DTV\tmp'
#Backup and Sync、ローカル保存用ディレクトリ
$bas_folder_path='C:\DTV\backupandsync'
#ループしてもffmpegの処理に失敗、mp4が10GBより大きい場合、tsファイル、番組情報ファイル、エンコしたmp4ファイルを退避するディレクトリ
$err_folder_path='C:\Users\sbn\Desktop'
#mp4用ffmpeg引数 使用可能:$audio_option(デュアルモノの判別)、$quality(tsファイルサイズ判別)、$pid_need(PID判別)
function arg_mp4 {
    #QSV H.264 LA-ICQ
    $script:arg="-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf bwdif=0:-1:1,pp=ac,hqdn3d=2.0 -global_quality ${quality} -c:v h264_qsv -preset:v veryslow -g 300 -bf 6 -refs 4 -b_strategy 1 -look_ahead 1 -look_ahead_depth 60 -pix_fmt nv12 -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${pid_need} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
    #x265 fast
    #$script:arg="-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -crf ${quality} -preset:v fast -g 15 -bf 2 -refs 4 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${pid_need} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
    #x265 bel9r
    #$script:arg="-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -preset:v fast -x265-params crf=${quality}:rc-lookahead=40:psy-rd=0.3:keyint=15:no-open-gop:bframes=2:rect=1:amp=1:me=umh:subme=3:ref=3:rd=3 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${pid_need} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
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
$hookUrl='https://discordapp.com/api/webhooks/466346185456746506/6DvzeWQVXT7oqUVCu5T4fN_ShkgVkXpvy-r4_ztxQPgfaJJtCM_FM-HjQc5CAZDu49Ok'

#########################################################################################################################

#====================Move関数====================
function Post {
    #自動削除が有効の場合、ts、ts.program.txt、ts.err、mp4を退避
    if ("${ts_del_toggle}" -eq "1") {
        Move-Item -LiteralPath "${env:FilePath}" "${err_folder_path}" -ErrorAction SilentlyContinue
        Move-Item -LiteralPath "${env:FilePath}.program.txt" "${err_folder_path}" -ErrorAction SilentlyContinue
        Move-Item -LiteralPath "${env:FilePath}.err" "${err_folder_path}" -ErrorAction SilentlyContinue
    }
    if ("${mp4_del_toggle}" -eq "1") {
        Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${err_folder_path}" -ErrorAction SilentlyContinue
    }
}

#====================Post関数====================
function Post {
    #投稿内容
    $env:content="Error:${env:FileName}.ts${err_detail}"
    #Twitter警告
    if ("$tweet_toggle" -eq "1") {
        &"${ruby_path}" "${tweet_rb_path}"
        #Start-Process "${ruby_path}" "${tweet_rb_path}" -WindowStyle Hidden -Wait
    }
    #Discord警告
    if ("$discord_toggle" -eq "1") {
        $payload=[PSCustomObject]@{
            content = $env:content
        }
        $payload=($payload | ConvertTo-Json)
        $payload=[System.Text.Encoding]::UTF8.GetBytes($payload)
        Invoke-RestMethod -Uri $hookUrl -Method Post -Body $payload
    }
}

#視聴予約なら終了
if ($env:RecMode -eq 4) {
    exit
}
if ("${env:FilePath}" -eq $null) {
    $err_detail+="`n[EDCB] 録画失敗によりエンコード不可"
    Post
}

#====================ffprocess関数====================
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
    #使用コア お手上げorz
    #(Get-Process -Id $p.Id).ProcessorAffinity=[int]"$Affinity"
    $p.ProcessorAffinity=[int]"$Affinity"
    Write-Output "ProcessorAffinity:$($p.ProcessorAffinity)"
    #Get-Process "ffmpeg" | select Id,ProcessorAffinity
    #プロセス優先度
    $p.PriorityClass=$Priority
    #プロセスの標準エラー出力を変数に格納(注意:WaitForExitの前に書かないとデッドロックします)
    $script:StdErr=$p.StandardError.ReadToEnd()
    #プロセス終了まで待機
    $p.WaitForExit()
    #終了コードを変数に格納
    $script:ExitCode=$p.ExitCode
    #リソースを開放
    $p.Close()
}

#====================NotifyIcon====================
#System.Windows.FormsクラスをPowerShellセッションに追加
Add-Type -AssemblyName System.Windows.Forms
#NotifyIconクラスをインスタンス化
$balloon=New-Object System.Windows.Forms.NotifyIcon
#powershellのアイコンを使用
$balloon.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')
#タスクトレイアイコンのヒントにファイル名を表示
#NotifyIcon.Textが64文字を超えると例外、String.Substringの開始値~終了値が文字数を超えると例外
#[string]($MyInvocation.MyCommand.Name) + ":${env:FileName}.ts"
switch (("$($MyInvocation.MyCommand.Name):${env:FileName}.ts").Length) {
    {$_ -ge 64} {$TextLength="63"}
    {$_ -lt 64} {$TextLength="$_"}
}
$balloon.Text=([string]($MyInvocation.MyCommand.Name) + ":${env:FileName}.ts").SubString(0,$TextLength)
#タスクトレイアイコン表示
$balloon.Visible=$True
#ファイル名をタイトルバーに表示
#(Get-Host).UI.RawUI.WindowTitle=$MyInvocation.MyCommand.Name + ":${env:FileName}.ts"
#(Get-Host).UI.RawUI.WindowTitle="$($MyInvocation.MyCommand.Name):${env:FileName}.ts"

#====================ログ====================
#log_toggle=1ならば実行
if ("${log_toggle}" -eq "1") {
    #ログ取り開始
    Start-Transcript -LiteralPath "${log_path}\${env:FileName}.txt"
    #録画用アプリの起動数を取得
    #$RecCount=(Get-Process -ErrorAction 0 "EpgDataCap_bon","TVTest").Count
    #Write-Output "同時録画数:$RecCount"
    #Get-ChildItemでログフォルダのtxtファイルを取得、更新日降順でソートし、logcnt_max個飛ばし、ForEach-ObjectでRemove-Itemループ
    Get-ChildItem "${log_path}\*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -Skip $logcnt_max | ForEach-Object {
        Remove-Item -LiteralPath "$_"
        Write-Output "ログ削除:$_"
    }
}

#====================ts・mp4の自動削除====================
#フォルダの合計サイズを取得する関数
function FolderSize {
    $script:delsize=(Get-ChildItem "$delpath" | Measure-Object -Sum Length).Sum
}
#フォルダの合計サイズを設定値以下に丸め込む関数
function FolderRound {
    #フォルダの合計サイズを取得
    FolderSize
    #合計サイズがdelroundより大きい間ループ
    while ($delsize -gt $delround) {
        #録画フォルダ内の1番古いtsファイルのファイル名を取得
        #録画フォルダ内のtsファイルに対し、最終更新年月日でソートした1番最初にくるやつ、ファイル名(拡張子なし)を取得
        $delname=$(Get-ChildItem "$delpath\*.$delext" | Sort-Object LastWriteTime | Select-Object BaseName -First 1).BaseName
        #tsかmp4を削除
        Remove-Item -LiteralPath "$delpath\$delname.$delext" -ErrorAction SilentlyContinue
        $dellog="削除:$delname.$delext"
        #tsフォルダを削除中の場合、同名のts.program.txt、ts.errも削除
        if ("$delext" -eq "ts") {
            Remove-Item -LiteralPath "$delpath\$delname.$delext.program.txt" -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath "$delpath\$delname.$delext.err" -ErrorAction SilentlyContinue
            $dellog+="、.program.txt、.err"
        }
        Write-Output $dellog
        #フォルダの合計サイズを取得
        FolderSize
    }
    Write-Output "${delext}フォルダ:$([math]::round(${delsize}/1GB,2))GB"
}
#ts_del_toggle=1なら実行
if ("$ts_del_toggle" -eq "1") {
    $delext='ts'
    $delround=$ts_folder_max
    $delpath="$env:FolderPath"
    FolderRound
}
#mp4_del_toggle=1なら実行
if ("$mp4_del_toggle" -eq "1") {
    $delext='mp4'
    $delround=$mp4_folder_max
    $delpath="$bas_folder_path"
    FolderRound
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
    if ("$ts_width" -eq "1440") {
        $scale=',scale=1920:1080'
    }
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
#-analyzeduration 30M -probesize 100M -ss 20
$StdErr=[string](&"${ffpath}\ffmpeg.exe" -hide_banner -nostats -i "${env:FilePath}" 2>&1)
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
    #'NoProgram'が含まれる行以降の行は不要とみなし抜ける
    if ($a -match 'NoProgram') {
        break
    }
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
    #'NoProgram'が含まれる行以降の行は不要とみなし抜ける
    if ($a -match 'NoProgram') {
        break
    }
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

#====================Backup and Sync====================
#ffprocessから渡された$StdErrからスペースを消す
$StdErr=($StdErr -replace " ","")
#ffmpegの終了コード、mp4のファイルサイズによる条件分岐
if (($ExitCode -ne 0) -Or ($mp4_size -gt 10GB)) {
    #ts、ts.program.txt、ts.err、mp4を退避
    Move
    #$StdErrをソートし分岐
    switch ($StdErr) {
        {$mp4_size -gt 10GB} {$err_detail+="`n[GoogleDrive] 10GB以上の為アップロードできません"}
        {$_ -match 'Errorduringencoding:devicefailed'} {$err_detail+="`n[h264_qsv] device failed (-17)"}
        {$_ -match 'Couldnotfindcodecparameters'} {$err_detail+="`n[mpegts] PIDの判別に失敗"}
        {$_ -match 'Unsupportedchannellayout'} {$err_detail+="`n[aac] 非対応のチャンネルレイアウト"}
        {$_ -match 'Toomanypacketsbuffered'} {$err_detail+="`n[-c:a aac] PIDの判別に失敗"}
        {$_ -match 'Inputpackettoosmall'} {$err_detail+="`n[-c:a copy] PIDの判別に失敗"}
        default {$err_detail="`n不明なエラー"}
    }
    #BalloonTip
    $TipIcon='Error'
    $TipTitle='エンコード失敗'
} else {
    #mp4をbas_folder_pathに投げる
    Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${bas_folder_path}"
    #$StdErrをソートし分岐
    switch ($StdErr) {
        {$_ -match 'Toomanypacketsbuffered'} {
            $err_detail+="`n[-c:a aac] PIDの判別に失敗？"
            #Move
            Post
        }
        {$_ -match 'Inputpackettoosmall'} {
            $err_detail+="`n[-c:a copy] PIDの判別に失敗？"
            #Move
            Post
        }
    }
    #BalloonTip
    $TipIcon='Info'
    $TipTitle='エンコード終了'
}

#====================BalloonTip====================
$TipText="ts:$([math]::round(${ts_size}/1GB,2))GB mp4:$([math]::round(${mp4_size}/1MB,0))MB"
Write-Output $TipTitle
Write-Output $TipText
if ("$balloontip_toggle" -eq "1") {
    #特定のTipIconのみを使用可
    #[System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
    $balloon.BalloonTipIcon=[System.Windows.Forms.ToolTipIcon]::$TipIcon
    #表示するタイトル
    $balloon.BalloonTipTitle="$TipTitle"
    #表示するメッセージ
    $balloon.BalloonTipText="${env:FileName}`n${TipText}${err_detail}"
    #balloontip_toggle=1なら5000ミリ秒バルーンチップ表示
    $balloon.ShowBalloonTip(5000)
    #5秒待って
    Start-Sleep -Seconds 5
}
#タスクトレイアイコン非表示(異常終了時は実行されずトレイに亡霊が残る仕様)
$balloon.Visible=$False