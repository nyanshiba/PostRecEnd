#180518
#_EDCBX_HIDE_
#視聴予約なら終了
if ($env:RecMode -eq 4) { exit }

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
$ts_folder_max=100GB
#--------------------mp4フォルダサイズを一定に保つファイルの削除--------------------
#0=無効、1=有効
$mp4_del_toggle=1
#backup and sync用フォルダの最大サイズを指定(0～8EB、単位:任意)
$mp4_folder_max=20GB
#--------------------jpg出力--------------------
#0=無効、1=有効
$jpg_toggle=1
#連番jpgを出力するフォルダ用のディレクトリ
$jpg_path='C:\Users\Shibanyan\Desktop\TVTest'
#jpg出力したい自動予約キーワード
$jpg_addkey='フランキス|BEATLESS|エヴァーガーデン|夏目友人帳|キズナアイのBEATスクランブル|Lostorage|フルメタル|ソードアート'
#jpg用ffmpeg引数(横pxが1440のとき、$scale=',scale=1920:1080'が使用可能)
function arg_jpg {
    $global:arg="-y -hide_banner -nostats -an -skip_frame nokey -i `"${env:FilePath}`" -vf yadif=0:-1:1,hqdn3d=4.0${scale} -f image2 -q:v 0 -vsync 0 `"${jpg_path}\${env:FileName}\%05d.jpg`""
}
#--------------------tsファイルサイズ判別--------------------
#通常品質
$quality_normal=27
#0=無効(通常品質のみ使用)、1=有効(通常・低品質を閾値を元に切り替える)
$tssize_toggle=1
#閾値
$tssize_max=20GB
#低品質
$quality_low=29
#--------------------エンコード--------------------
#ffmpeg.exe、ffprobe.exeがあるディレクトリ
$ffpath='C:\DTV\ffmpeg'
#一時的にmp4を吐き出すディレクトリ
$tmp_folder_path='C:\DTV\tmp'
#backup and sync用ディレクトリ
$bas_folder_path='C:\DTV\backupandsync'
#エンコしたファイルが10GBより大きい、処理に25回失敗した場合にts、ts.program.txt、mp4を退避するディレクトリ
$err_folder_path='C:\Users\Shibanyan\Desktop'
#mp4用ffmpeg引数
function arg_mp4 {
    $global:arg="-y -hide_banner -nostats -fflags +discardcorrupt -i `"${env:FilePath}`" ${audio_option} -vf yadif=0:-1:1,hqdn3d=3.0,scale=1280:720:flags=lanczos+accurate_rnd,unsharp=3:3:1:3:3:1:0 -global_quality ${quality} -c:v h264_qsv -preset:v veryslow -g 300 -bf 16 -refs 9 -b_strategy 1 -look_ahead 1 -look_ahead_downsampling off -pix_fmt nv12 -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 -map 0:p:${env:SID10} -sn -dn -ignore_unknown -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
}
#--------------------ツイート警告--------------------
#0=無効、1=有効
$tweet_toggle=1
#ruby.exe
$ruby_path='C:\Ruby24-x64\bin\ruby.exe'
#tweet.rb
$tweet_rb_path='C:\DTV\EDCB\tweet.rb'
#SSL証明書(環境変数)
$env:ssl_cert_file='C:\DTV\EDCB\cacert.pem'



#====================NotifyIcon====================
#System.Windows.FormsクラスをPowerShellセッションに追加
Add-Type -AssemblyName System.Windows.Forms
#NotifyIconクラスをインスタンス化
$balloon=New-Object System.Windows.Forms.NotifyIcon
#powershellのアイコンを使用
$balloon.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')
#タスクトレイアイコンのヒントにファイル名を表示
$balloon.Text=$MyInvocation.MyCommand.Name + ":${env:FileName}.ts"
#タスクトレイアイコン表示
$balloon.Visible=$True
#ファイル名をタイトルバーに表示
#(Get-Host).UI.RawUI.WindowTitle=$MyInvocation.MyCommand.Name + ":${env:FileName}.ts"
#====================BalloonTip関数====================
function BalloonTip {
    #特定のTipIconのみを使用可
    #[System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
    $balloon.BalloonTipIcon=[System.Windows.Forms.ToolTipIcon]::${ToolTipIcon}
    #表示するタイトル
    $balloon.BalloonTipTitle="エンコード${enc_result}"
    #表示するメッセージ
    $balloon.BalloonTipText="${env:FileName}`nts:$([math]::round(${ts_size}/1GB,2))GB mp4:$([math]::round(${mp4_size}/1MB,0))MB"
    #balloontip_toggle=1なら5000ミリ秒バルーンチップ表示
    if ("${balloontip_toggle}" -eq "1") { $balloon.ShowBalloonTip(5000) }
    #5秒待ってタスクトレイアイコン非表示
    Start-Sleep -Seconds 5
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
    #プロセスの標準エラー出力をシェルの標準出力に出力(注意:WaitForExitの前に書かないとデッドロックします)
    Write-Output $p.StandardError.ReadToEnd()
    #プロセス終了まで待機
    $p.WaitForExit()
    #終了コードを変数に格納
    $global:ExitCode=$p.ExitCode
    #リソースを開放
    $p.Close()
}

#====================ログ====================
#log_toggle=1ならば実行
if ("${log_toggle}" -eq "1") {
    #ログ取り開始
    Start-Transcript -LiteralPath "${log_path}\${env:FileName}.txt"
    #Get-ChildItemでログフォルダのtxtファイルを取得、更新日降順でソートし、logcnt_max個飛ばし、Foreach-ObjectでRemove-Itemループ
    Get-ChildItem "${log_path}\*.txt" | Sort-Object LastWriteTime -Descending | Select-Object -Skip ${logcnt_max} | Foreach-Object {
        Remove-Item -LiteralPath "$_"
        Write-Output "ログ削除:$_"
    }
}

#====================jpg出力====================
#jpg出力機能が有効(jpg_toggle=1)且つenv:Addkey(自動予約時のキーワード)にjpg_addkey(指定の文字)が含まれている場合は連番jpgも出力
if (("${jpg_toggle}" -eq "1") -And ($("${env:Addkey}" -match "${jpg_addkey}") -eq $True)) {
    #出力フォルダ作成
    New-Item "${jpg_path}\${env:FileName}" -ItemType Directory
    Write-output "jpg出力:${env:FileName}.ts"
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

#====================tsファイルサイズ判別====================
if ("${tssize_toggle}" -eq "1") {
    #tsファイルのサイズを変数ts_sizeに格納
    $ts_size=$(Get-ChildItem -LiteralPath "${env:FilePath}").Length
    #20GB以下ならquality 26、より大きいなら28
    if ($ts_size -le $tssize_max) {
        $quality="$quality_normal"
    } elseif ($ts_size -gt $tssize_max) {
        $quality="$quality_low"
    }
} elseif ("${tssize_toggle}" -eq "0") {
    $quality="$quality_normal"
}
Write-Output "quality:$quality"

#====================音声形式の判別====================
#番組情報ファイルがありデュアルモノという文字列があればTrue、文字列がない場合はFalse、番組情報ファイルが無ければNull
if ($(Get-Content -LiteralPath "${env:FilePath}.program.txt" | Select-String -SimpleMatch 'デュアルモノ' -quiet) -eq $True) {
    $audio_option='-c:a aac -b:a 128k -filter_complex channelsplit'
} else {
    $audio_option='-c:a aac -b:a 256k'
}
Write-Output "audio_option:$audio_option"

#====================エンコード====================
#mp4用ffmpeg引数を遅延展開
$file="${ffpath}\ffmpeg.exe"
arg_mp4
Write-Output "Arguments:$arg"

#終了コードが1且つループカウントが25以下までの間、エンコードを試みる
do {
    #録画の開始終了のビジー時を避ける、再試行の効果を出すためにちょっと待つ
    Start-Sleep -s 5
    #ffmpegprocessを起動
    ffprocess
    Write-Output "ExitCode:$ExitCode"
    #エンコ後mp4のサイズを取得
    $mp4_size=$(Get-ChildItem -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4").Length
    Write-Output "mp4:$([math]::round(${mp4_size}/1MB,0))MB"
    #ループカウント
    $cnt++
    Write-Output "エンコード回数:$cnt"
} while (($ExitCode -eq 1) -And ($cnt -le 25))

#====================mp4ファイルサイズ判別====================
#ffmpegの終了コード、mp4のファイルサイズによる条件分岐
if (($ExitCode -eq 1) -Or ($mp4_size -gt 10GB)) {
    #25回試行してもffmpegの終了コードが1、mp4が10GBより大きい場合、ts、ts.program.txt、mp4を退避する
    Move-Item -LiteralPath "${env:FilePath}" "${err_folder_path}"
    Move-Item -LiteralPath "${env:FilePath}.program.txt" "${err_folder_path}"
    Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${err_folder_path}"
    Write-Output "エンコード失敗:ts=$([math]::round(${ts_size}/1GB,2))GB mp4=$([math]::round(${mp4_size}/1MB,0))MB"
    #ツイート警告
    if ("${tweet_toggle}" -eq "1") {
        $env:tweet_content="ERROR:${env:FileName}.tsと関連ファイルを退避しました。ログを確認して下さい。"
        #Start-Process "${ruby_path}" "${tweet_rb_path}" -WindowStyle Hidden -Wait
        &"${ruby_path}" "${tweet_rb_path}"
    }
    #バルーンチップ
    $ToolTipIcon='Error'
    $enc_result='失敗'
    BalloonTip
} elseif ($mp4_size -le 10GB) {
    #それ以外でmp4が10GB以下ならmp4をbas_folder_pathに投げる
    Move-Item -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4" "${bas_folder_path}"
    Write-Output "エンコード終了:ts=$([math]::round(${ts_size}/1GB,2))GB mp4=$([math]::round(${mp4_size}/1MB,0))MB"
    #バルーンチップ
    $ToolTipIcon='Info'
    $enc_result='終了'
    BalloonTip
}