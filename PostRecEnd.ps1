#2020-10-08
#_EDCBX_HIDE_
#ファイル名をタイトルバーに表示
#(Get-Host).UI.RawUI.WindowTitle="$($MyInvocation.MyCommand.Name):${env:FileName}.ts"

#####################ユーザ設定####################################################################################################

<#
設定の書き方ルール(PowerShellの仕様)

X だめ
$Toggle= イコール残したままにしない！テストでXにされるよ！
O おk
$Toggle=$False 無効
$Toggle=$false 大文字小文字の区別はない
$Toggle=0 無効
$Toggle=$True 有効
$Toggle=1 有効
$Toggle = $true =や+=の前後にスペースがあっても良い(コーディング的にはこっちが推奨っぽい)
$Toggle 空欄
$Toggle=$Null null
#$Toggle コメント

X だめ
$Path=C:\DTV\EncLog エラー出る
$Path="C:\DTV\エンコード　ログ" 処理できるかもしれないけど基本的にパスに半角や全角のスペースは非推奨
O おk
$Path='C:\DTV\EncLog' 変数が無ければリテラルでおk
$Path = "C:\DTV\EncLog" もちのろん

X だめ
$Arg='-quality $ArgQual' シングルクォートでは'$ArgQual'という文字列になってしまうので変数の中身が展開されないよ！
$Arg="-vf bwdif=0:-1:1$ArgScale" 変数名とコードが紛らわしいよ！
$Arg="-i "${FilePath}"" ダブルクオートの範囲が滅茶滅茶だよ><
$Arg='-i "${FilePath}"' 変数が展開されないよ！
O おk
$Arg="-quality $ArgQual" ダブルクォートでは変数の中身が展開される
$Arg="-vf bwdif=0:-1:1${ArgScale}" ${}を使おう
$Arg="-i `"${FilePath}`"" バッククオート`でエスケープしよう

X だめ
$logcnt_max="1000" こういうのはString(文字列)じゃないよ！int(数値)だよ！
O おk
$logcnt_max=1000
$logcnt_max=[int]"1000"
$Size=200GB
$Size="200GB" OKらしい(^^;;
$Size=0.2TB
#>

#ffmpeg.exe、ffprobe.exeがあるディレクトリ
$ffpath='C:\bin\ffmpeg'

#--------------------ログ--------------------
#$False=無効、$True=有効
$log_toggle=$True
#ログ出力ディレクトリ
$log_path='C:\logs\PostRecEnd'
#ログを残す数
$logcnt_max=500

#--------------------tsの自動削除--------------------
#閾値を超過した場合、$False=容量警告、$True=tsを自動削除
$TsFolderRound=$True
#録画フォルダの上限
$ts_folder_max=150GB

#--------------------mp4の自動削除--------------------
#閾値を超過した場合、$False=容量警告、$True=mp4を自動削除
$Mp4FolderRound=$True
#mp4用フォルダの上限
$mp4_folder_max=50GB

#--------------------jpg出力--------------------
#$True=有効 $False=無効
$jpg_toggle=$True
#連番jpgを出力するフォルダ用のディレクトリ
$jpg_path='C:\Users\sbn\Desktop\TVTest'
#jpg出力したい自動予約キーワード(常にjpg出力する場合: $jpg_addkey='')
$jpg_addkey='宇宙戦艦ヤマト|アリシゼーション|ダンジョンに出会いを求める|冴えない彼女の育てかた|キルラキル|サイコパス|PSYCHO|鬼滅の刃|とある科学の一方通行|戦姫絶唱シンフォギア|プリンシパル|魔法少女まどか|打ち上げ花火|オーディナル|ソードアート|Angel Beats|革命機ヴァルヴレイヴ|パンツァー|新世紀エヴァンゲリオン|ジョジョの奇妙な冒険|涼宮ハルヒ|戦姫絶唱シンフォギア|デート・ア・ライブ|終わりのセラフ|灼眼のシャナ|アズールレーン|新世紀エヴァンゲリオン|翠星のガルガンティア|グリザイアの果実|シャーロット|ゼロから始める|魔法科高校の劣等生|マギアレコード|とある科学の超電磁砲|デンドログラム|恋する小惑星|へやキャン|フランキス|ガンダム|グリザイア|新サクラ大戦|エヴァーガーデン|銀河英雄伝説|政見放送|インターミッション|ハイビジョンウインドー|さわやかウインドー|シドニアの騎士|ゼロから始める異世界生活|Lapis Re:LiGHTs|魔女の旅々|エルメロイ|Charlotte|ゆるキャン|フリート|はいふり|庶民サンプル|ストライクウィッチーズ|戦翼のシグルドリーヴァ|ダンジョンに出会いを求めるのは間違っているだろうか|キミと僕の最後の戦場|アサルトリリィ|五等分の花嫁'
#自動予約キーワードに引っ掛かった場合に実行するコード 使用可能:$ArgScale(横が1440pxの場合のみ",scale=1920:1080"が格納される。画像にはSARとか無いので)
function ImageEncode {
    #連番jpg出力する例
    <#
    #出力ディレクトリ
    New-Item "${jpg_path}\${env:FileName}" -ItemType Directory
    #プロセス実行
    Invoke-Process -File "${ffpath}\ffmpeg.exe" -Arg "-y -hide_banner -nostats -an -skip_frame nokey -i `"${env:FilePath}`" -vf bwdif=0:-1:1,pp=ac,hqdn3d=2.0${ArgScale} -f image2 -q:v 0 -vsync 0 `"$jpg_path\$env:FileName\%05d.jpg`""
    #>

    #waifu2xをここで使用することは一応可能ですが、処理時間が非現実的です

    #tsを保持用ディレクトリにコピーする例
    Copy-Item -LiteralPath "${env:FilePath}" "E:\ts" -ErrorAction SilentlyContinue
}

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
-Arg "-y -hide_banner -nostats -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1 -c:v h264_nvenc -preset:v slow -profile:v high -rc:v vbr_minqp -rc-lookahead 32 -spatial-aq 1 -aq-strength 1 -qmin:v 23 -qmax:v 25 -b:v 1500k -maxrate:v 3500k -pix_fmt yuv420p ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
QSV H.264 LA-ICQ
-Arg "-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -ss 5 -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac,hqdn3d=2.0 -global_quality ${ArgQual} -c:v h264_qsv -preset:v veryslow -g 300 -bf 6 -refs 4 -b_strategy 1 -look_ahead 1 -look_ahead_depth 60 -pix_fmt nv12 -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x265 fast
-Arg "-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -crf ${ArgQual} -preset:v fast -g 15 -bf 2 -refs 4 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x265 fast bel9r inspire
-Arg "-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx265 -preset:v fast -x265-params crf=${ArgQual}:rc-lookahead=40:psy-rd=0.3:keyint=15:no-open-gop:bframes=2:rect=1:amp=1:me=umh:subme=3:ref=3:rd=3 -pix_fmt yuv420p -bsf:v hevc_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
x264 placebo by bel9r
-Arg "-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" ${ArgAudio} -vf bwdif=0:-1:1,pp=ac -c:v libx264 -preset:v placebo -x264-params crf=${ArgQual}:rc-lookahead=60:qpmin=5:qpmax=40:qpstep=16:qcomp=0.85:mbtree=0:vbv-bufsize=31250:vbv-maxrate=25000:aq-strength=0.35:psy-rd=0.35:keyint=300:bframes=6:partitions=p8x8,b8x8,i8x8,i4x4:merange=64:ref=4:no-dct-decimate=1 -pix_fmt yuv420p -bsf:v h264_metadata=colour_primaries=1:transfer_characteristics=1:matrix_coefficients=1 ${ArgPid} -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`""
#>
function VideoEncode {
    #hevc_nvenc constqp (qpI,P,Bはtsファイルサイズ判別を参照)
    Invoke-Process -File "${ffpath}\ffmpeg.exe" -Arg "-y -hide_banner -nostats -analyzeduration 30M -probesize 100M -fflags +discardcorrupt -i `"${env:FilePath}`" $ArgAudio -vf pullup,dejudder,idet=intl_thres=1.38:prog_thres=1.5,yadif=mode=send_field:parity=auto:deint=interlaced,fps=fps=30000/1001:round=zero -c:v hevc_nvenc -preset:v p7 -profile:v main10 -rc:v constqp -rc-lookahead 1 -spatial-aq 0 -temporal-aq 1 -weighted_pred 0 $ArgQual -b_ref_mode 1 -dpb_size 4 -multipass 2 -g 60 -bf 3 -pix_fmt yuv420p10le $ArgPid -movflags +faststart `"${tmp_folder_path}\${env:FileName}.mp4`"" -Priority 'BelowNormal' -Affinity '0xFFF'
}

#--------------------Post--------------------
#$False=Error時のみ、$True=常時 Twitter、DiscordにPost
$InfoPostToggle=$False

#Twitter機能 $False=無効、$True=有効
$tweet_toggle=$False
#ruby.exe
$ruby_path='wsl ruby'
#tweet.rb
$tweet_rb_path='C:\DTV\EDCB\tweet.rb'
#SSL証明書(環境変数)
$env:ssl_cert_file='C:\DTV\EDCB\cacert.pem'

#Discord機能 $False=無効、$True=有効
$discord_toggle=$True
#webhook url
$hookUrl='https://discordapp.com/api/webhooks/XXXXXXXXXX'

#BalloonTip機能 $False=無効、$True=有効
$balloontip_toggle=$True

<#
エラーメッセージ一覧

・[EDCB] 録画失敗によりエンコード不可: tsファイルが無い(パスが渡されない)場合。録画失敗？
・[EDCB] PIDの判別不可: ストリームの解析が失敗以前に不可能。Drop過多orスクランブル解除失敗？
・[GoogleDrive] 10GB以上の為アップロードできません: GoogleDriveの仕様に合わせる。
・[h264_qsv] device failed (-17): QSVのエラー。ループして復帰を試みるも失敗した場合。
・[mpegts] コーデックパラメータが見つかりません: PID判別から渡されたPIDが適切でないorffmpegが非対応のストリーム。
・[aac] 非対応のチャンネルレイアウト: ffmpeg4.0～デュアルモノを少なくとも従来の引数では扱えなくなった。
・[-c:a aac] PIDの判別に失敗: -c:a aac時。指定サービスのみ(全サービスでない)録画になっていなければ必ず発生。また一通りのストリームに対応させた筈だけど起こるかもしれない。
・[-c:a copy] PIDの判別に失敗: -c:a copy時。上に同じ。
・[-c:a aac] PIDの判別に失敗？ ExitCode:0: -c:a aac時。ffmpegの終了コードは0だが異常がある？場合。
・[-c:a copy] -c:a aacか-ss 1で治るやつ ExitCode:0: -c:a copy時。上に同じ。
・[FFmpeg] 無効な引数
・不明なエラー
#>

#########################################################################################################################

"#--------------------Post関数--------------------"
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
        #Twitter警告
        if ($tweet_toggle)
        {
            $env:content = $content
            &"$ruby_path" "$tweet_rb_path"
            #Start-Process "${ruby_path}" "${tweet_rb_path}" -WindowStyle Hidden -Wait
        }
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
    #BalloonTip
    if ($balloontip_toggle)
    {
        #特定のTipIconのみを使用可
        #[System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::$tipicon
        #表示するタイトル
        $balloon.BalloonTipTitle = $tiptitle
        #表示するメッセージ
        $balloon.BalloonTipText = $content
        #balloontip_toggle=1なら5000ミリ秒バルーンチップ表示
        $balloon.ShowBalloonTip(5000)
        #5秒待って
        Start-Sleep -Seconds 5
    }
    #タスクトレイアイコン非表示(異常終了時は実行されずトレイに亡霊が残る仕様)
    $balloon.Visible = $False
}

#視聴予約なら終了
if ($env:RecMode -eq 4) {
    return "視聴予約の為終了"
}
if ("${env:FilePath}" -eq $null) {
    Post -Exc $True -Toggle $True -Content "Error:${env:Title}`n[EDCB] 録画失敗によりエンコード不可" -TipIcon 'Error' -TipTitle '録画失敗'
}

"#--------------------Invoke-Process関数--------------------"
#ffmpeg、&ffmpeg、.\ffmpeg:ffmpegが引数を正しく認識しない(ファイル名くらいなら-f mpegtsで行けるけどもういいです)
#Start-Process ffmpeg:-NoNewWindowはWrite-Host？-RedirectStandardOutput、Errorはファイルのみ、-PassThruはExitCodeは受け取れても.StandardOutput、Errorは受け取れない仕様
function Invoke-Process {
    param
    (
        [string]$priority,
        [int]$affinity,
        [string]$file,
        [string]$arg
    )
    Write-Output "Invoke-Process:$file $arg"

    #設定
    #ProcessStartInfoクラスをインスタンス化
    $psi=New-Object System.Diagnostics.ProcessStartInfo
    #アプリケーションファイル名
    $psi.FileName = $file
    #引数
    $psi.Arguments = $arg
    #標準エラー出力だけを同期出力(注意:$trueは1つだけにしないとデッドロックします)
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $false
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

    #実行
    #Processクラスをインスタンス化
    $p=New-Object System.Diagnostics.Process
    #設定を読み込む
    $p.StartInfo = $psi
    #プロセス開始
    $p.Start() > $Null
    #プロセッサ親和性
    if ($affinity)
    {
        #(Get-Process -Id $p.Id).ProcessorAffinity=[int]"$Affinity"
        $p.ProcessorAffinity = [int]"$affinity"
    }
    #プロセス優先度
    if ($priority)
    {
        $p.PriorityClass = $priority
    }
    #標準エラー出力をプロセス終了まで読む
    $script:StdErr = $Null
    while (!$p.HasExited)
    {
        $script:StdErr += "$($p.StandardError.ReadLine())`n"
    }
    #プロセスの標準エラー出力を変数に格納(注意:WaitForExitの前に書かないとデッドロックします)
    #$script:StdErr=$p.StandardError.ReadToEnd()
    #プロセス終了まで待機
    #$p.WaitForExit()
    #終了コードを変数に格納
    $script:ExitCode = $p.ExitCode
    #リソースを開放
    $p.Close()
}

"#--------------------NotifyIcon--------------------"
#System.Windows.FormsクラスをPowerShellセッションに追加
Add-Type -AssemblyName System.Windows.Forms
#NotifyIconクラスをインスタンス化
$balloon=New-Object System.Windows.Forms.NotifyIcon
#powershellのアイコンを使用
$balloon.Icon=[System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe')
#NotifyIcon.Textが64文字を超えると例外、String.Substringの開始値~終了値が文字数を超えると例外
switch (("$($MyInvocation.MyCommand.Name):${env:FileName}.ts").Length) {
    {$_ -ge 64} {$TextLength="63"}
    {$_ -lt 64} {$TextLength="$_"}
}
#タスクトレイアイコンのヒントにファイル名を表示
$balloon.Text=([string]($MyInvocation.MyCommand.Name) + ":${env:FileName}.ts").SubString(0,$TextLength)
#タスクトレイアイコン表示
$balloon.Visible=$True

#ログ有効時、NotifyIconクリックでログを既定のテキストエディタで開く
if ($log_toggle)
{
    $balloon.add_Click({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left)
        {
            &"${log_path}\${env:FileName}.txt"
        }
    })
}

"#--------------------ログ--------------------"
#ログのソート例: (sls -path "$log_path\*.txt" 'faild' -SimpleMatch).Path
#log_toggle=$Trueならば実行
if ($log_toggle) {
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


"#--------------------ts・mp4の自動削除--------------------"
#フォルダの合計サイズを設定値以下に丸め込む関数
function FolderRound {
    param
    (
        [bool]$toggle,
        [string]$ext,
        [string]$path,
        [string]$round
    )
    if ($toggle)
    {
        #初期値
        $delcnt=-1
        #必ず1回は実行、フォルダ内の新しいファイルをSkipする数$iを増やしていって$maintsizeを$round以下に丸め込むループ
        do {
            $delcnt++
            $maintsize=(Get-ChildItem "$path\*.$ext" | Sort-Object LastWriteTime -Descending | Select-Object -Skip $delcnt | Measure-Object -Sum Length).Sum
        } while ($maintsize -gt $round)
        #先程Skipしたファイルを実際に削除
        Get-ChildItem "$path\*.$ext" | Sort-Object LastWriteTime | Select-Object -First $delcnt | ForEach-Object {
            #tsかmp4を削除
            Remove-Item -LiteralPath "$path\$($_.BaseName).$ext" -ErrorAction SilentlyContinue
            $dellog="削除:$($_.BaseName).$ext"
            #tsを削除中の場合、同名のts.program.txt、ts.errも削除
            if ("$ext" -eq "ts") {
                Remove-Item -LiteralPath "$path\$($_.BaseName).$ext.program.txt" -ErrorAction SilentlyContinue
                Remove-Item -LiteralPath "$path\$($_.BaseName).$ext.err" -ErrorAction SilentlyContinue
                $dellog+="、.program.txt、.err"
            }
            Write-Output $dellog
        }
        Write-Output "${ext}フォルダ:$([math]::round(${maintsize}/1GB,2))GB"
    } elseif (!($toggle))
    {
        #超過時の警告
        if ($((Get-ChildItem "$path" | Measure-Object -Sum Length).Sum) -gt $round) {
            $err_detail="`n[FolderRound] ${ext}ディレクトリが${round}を超過"
        }
    }

}
#ts
FolderRound -Toggle $TsFolderRound -Ext "ts" -Path "$env:FolderPath" -Round $ts_folder_max
#mp4
FolderRound -Toggle $Mp4FolderRound -Ext "mkv" -Path "$mp4_folder_path" -Round $mp4_folder_max

"#--------------------jpg出力--------------------"
#jpg出力機能が有効(jpg_toggle=1)且つenv:Addkey(自動予約時のキーワード)にjpg_addkey(指定の文字)が含まれている場合は連番jpgも出力
if (($jpg_toggle) -And ("$env:Addkey" -match "$jpg_addkey")) {
    Write-Output "jpg出力"
    #生TSの横が1920か1440か調べる
    $ts_width=[xml](&"${ffpath}\ffprobe.exe" -v quiet -i "${env:FilePath}" -show_entries stream=width -print_format xml 2>&1)
    $ts_width=$ts_width.ffprobe.streams.stream.width
    #SAR比(1440x1080しか想定してないけど)によるフィルタ設定、jpg出力
    if ("$ts_width" -eq "1440") {
        $ArgScale=',scale=1920:1080'
    }
    #jpg用ffmpeg引数を遅延展開
    ImageEncode
}

"#--------------------tsファイルサイズ判別--------------------"
#tsファイルサイズを取得
$ts_size=(Get-ChildItem -LiteralPath "${env:FilePath}").Length
if ($tssize_toggle) {
    #閾値$tssize_max以下なら通常品質$quality_normal、より大きいなら低品質$quality_low
    switch ($ts_size) {
        {$_ -le $tssize_max} {$ArgQual="$quality_normal"}
        {$_ -gt $tssize_max} {$ArgQual="$quality_low"}
    }
    Write-Output "ArgQual:$ArgQual"
}

"#--------------------デュアルモノの判別--------------------"
#番組情報ファイルがありデュアルモノという文字列があればTrue、文字列がない場合はFalse、番組情報ファイルが無ければNull
if (Get-Content -LiteralPath "${env:FilePath}.program.txt" | Select-String -SimpleMatch 'デュアルモノ' -quiet) {
    $ArgAudio=$audio_dualmono
} else {
    $ArgAudio=$audio_normal
}
Write-Output "ArgAudio:$ArgAudio"

"#--------------------PIDの判別--------------------"
#PID引数の設定
#ffprobeでcodec_type,height,idをソート
$programs = [xml](&"$ffpath\ffprobe.exe" -v quiet -analyzeduration 30M -probesize 100M -i "${env:FilePath}" -show_entries stream=codec_type,height,id,channels -print_format xml 2>&1)
$programs.ffprobe.streams.stream
$programs.ffprobe.streams.stream | foreach {
    #解像度の大きいVideoストリームを選ぶ
    #xmlの要素はStringなのでintにする必要あり
    if (($_.codec_type -eq "video") -And ($_.height -ne "0") -And ([int]($_.height) -gt [int]($otherheight)))
    {
        $otherheight = $_.height
        $ArgPid = "-map i:$($_.id)"
    }
}
#VideoのPIDの先頭(0x1..)と一致するAudioストリームを引数に追加
$programs.ffprobe.streams.stream | foreach {
    if (($_.codec_type -eq "audio") -And ($_.channels -ne "0") -And ($($_.id).Substring(0,3) -eq $ArgPid.Substring(7,3)))
    {
        $ArgPid += " -map i:$($_.id)"
    }
}
Write-Output "ArgPid:$ArgPid"

"#--------------------エンコード--------------------"
#カウントを0にリセット
$cnt=0
#終了コードが1且つループカウントが50未満までの間、エンコードを試みる
do {
    $cnt++
    #再試行時からクールタイムを追加
    if ($cnt -ge 2) {
        Start-Sleep -s 60
    }
    #エンコ mp4用ffmpeg引数を遅延展開
    VideoEncode
    #エンコ1回目と成功時(ExitCode:0)のログだけで十分
    if (($cnt -le 1) -Or ($ExitCode -eq 0)) {
        #プロセスの標準エラー出力をシェルの標準出力に出力
        Write-Output $StdErr
        #エンコ後のmp4のファイルサイズ
        $mp4_size=$(Get-ChildItem -LiteralPath "${tmp_folder_path}\${env:FileName}.mp4").Length
    }
} while (($ExitCode -eq 1) -And ($cnt -lt 50))
#最終的なエンコード回数、終了コード、ファイルサイズ
Write-Output "エンコード回数:$cnt"
Write-Output "ExitCode:$ExitCode"
$PostFileSize="`nts:$([math]::round(${ts_size}/1GB,2))GB mp4:$([math]::round(${mp4_size}/1MB,0))MB"
$PostFileSize

"#--------------------Backup and Sync--------------------"
#Invoke-Processから渡された$StdErrからスペースを消す
$StdErr=($StdErr -replace " ","")
#ffmpegの終了コード、mp4のファイルサイズによる条件分岐
if ($ExitCode -gt 0)
{
    #$StdErrをソートしPost内容を決める
    switch ($StdErr)
    {
        {$_ -match 'Errorduringencoding:devicefailed'} {$err_detail+="`n[h264_qsv] device failed (-17)"}
        {$_ -match 'Couldnotfindcodecparameters'} {$err_detail+="`n[mpegts] PIDの判別に失敗"}
        {$_ -match 'Unsupportedchannellayout'} {$err_detail+="`n[aac] 非対応のチャンネルレイアウト"}
        {$_ -match 'Toomanypacketsbuffered'} {$err_detail+="`n[-c:a aac] max_muxing_queue_size"}
        {$_ -match 'Inputpackettoosmall'} {$err_detail+="`n[-c:a copy] PIDの判別に失敗"}
        {$_ -match 'Invalidargument'} {$err_detail+="`n[FFmpeg] 無効な引数"}
        default {$err_detail+="`n不明なエラー"}
    }
    #Twitter、Discord、BalloonTip
    Post -Exc $True -Toggle $True -Content "Error:${env:FileName}.ts${err_detail}${PostFileSize}" -TipIcon 'Error' -TipTitle 'エンコード失敗'
} elseif (($googledrive) -And ($mp4_size -gt 10GB))
{
    #Post内容
    $err_detail+="`n[GoogleDrive] 10GB以上の為アップロードできません"
    #Twitter、Discord、BalloonTip
    Post -Exc $True -Toggle $True -Content "Error:${env:FileName}.ts${err_detail}${PostFileSize}" -TipIcon 'Error' -TipTitle 'アップロード失敗'
} else
{
    #mp4をmp4_folder_pathに投げる
    Move-Item -LiteralPath "$tmp_folder_path\$env:FileName.mp4" -Destination "$mp4_folder_path\$env:FileName.mkv"
    #エラーメッセージが格納されていればTipIconをWarningに変える
    if ($err_detail)
    {
        $TipIcon='Warning'
    } else
    {
        $TipIcon='Info'
    }
    #Twitter、Discord、BalloonTip
    Post -Exc $False -Toggle $InfoPostToggle -Content "${env:FileName}.ts${err_detail}${PostFileSize}" -TipIcon "$TipIcon" -TipTitle 'エンコード終了'
}
