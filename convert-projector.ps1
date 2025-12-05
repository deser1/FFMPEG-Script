param(
    [Parameter(Position=0)]
    [string]$Path,
    [ValidateSet('hevc','h264','vp9','av1','mpeg2','mpeg1')]
    [string]$VideoCodec = 'hevc',
    [ValidateSet('aac','mp3','wma')]
    [string]$AudioCodec = 'aac',
    [switch]$Recursive,
    [switch]$Allow4K,
    [string]$OutputDir,
    [switch]$Overwrite,
    [switch]$Force,
    [string]$VideoBitrate,
    [string]$AudioBitrate,
    [switch]$DryRun,
    [switch]$Help,
    [switch]$AutoProfile,
    [switch]$AiAdvisor,
    [switch]$QualityCheck,
    [string]$LogPath,
    [int]$SampleSeconds = 30,
    [int]$SampleStart = 60
)

function Show-Help {
    Write-Host "Użycie: powershell -File convert-projector.ps1 [-Path <plik|katalog>] [-Recursive] [-Allow4K] [-VideoCodec <hevc|h264|vp9|av1|mpeg2|mpeg1>] [-AudioCodec <aac|mp3|wma>] [-OutputDir <katalog>] [-Overwrite] [-Force] [-VideoBitrate <np. 4M>] [-AudioBitrate <np. 192k>] [-DryRun] [-AutoProfile] [-AiAdvisor] [-QualityCheck] [-LogPath <plik>] [-SampleSeconds <int>] [-SampleStart <int>]" -ForegroundColor Yellow
}

if ($Help) { Show-Help; exit 0 }

if (-not $PSBoundParameters.ContainsKey('AutoProfile')) { $AutoProfile = $true }

function Ensure-Tool($name) {
    $p = (Get-Command $name -ErrorAction SilentlyContinue)
    if (-not $p) { throw "Brak narzędzia '$name' w PATH" }
}

Ensure-Tool ffmpeg
Ensure-Tool ffprobe

function Has-Encoder($enc) {
    try { $o = & ffmpeg -hide_banner -encoders; $s = $o | Out-String; return ($s -match [regex]::Escape($enc)) } catch { return $false }
}

function Select-Path {
    param([string]$Mode)
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    if ($Mode -eq 'file') {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Wideo|*.mp4;*.mkv;*.mov;*.avi;*.wmv;*.mpeg;*.mpg;*.webm;*.m4v;*.ts;*.m2ts|Wszystkie|*.*"
        $dlg.Multiselect = $false
        $null = $dlg.ShowDialog()
        return $dlg.FileName
    }
    else {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $null = $dlg.ShowDialog()
        return $dlg.SelectedPath
    }
}

function Get-InputItems {
    param([string]$InPath,[switch]$Recursive,[string]$ExcludePath)
    $ext = @('*.mp4','*.mkv','*.mov','*.avi','*.wmv','*.mpeg','*.mpg','*.webm','*.m4v','*.ts','*.m2ts')
    if (Test-Path -LiteralPath $InPath) {
        if ((Get-Item -LiteralPath $InPath).PSIsContainer) {
            $items = foreach($e in $ext){ if ($Recursive) { Get-ChildItem -LiteralPath $InPath -Filter $e -File -Recurse } else { Get-ChildItem -LiteralPath $InPath -Filter $e -File } }
            if ($ExcludePath) { $items = $items | Where-Object { -not $_.FullName.StartsWith($ExcludePath,[System.StringComparison]::OrdinalIgnoreCase) } }
            $items
        } else { Get-Item -LiteralPath $InPath }
    } else { @() }
}

function Probe-Video($file) {
    $vj = ffprobe -v error -select_streams v:0 -show_entries stream=codec_name,width,height,r_frame_rate -of json -- "$file" | ConvertFrom-Json
    $aj = ffprobe -v error -select_streams a:0 -show_entries stream=codec_name -of json -- "$file" | ConvertFrom-Json
    $durRaw = ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 -- "$file"
    $fpsVal = 0
    try {
        $r = $vj.streams[0].r_frame_rate
        if ($r -match '/') { $a,$b = $r -split '/'; if ([int]$b -ne 0) { $fpsVal = [double]$a/[double]$b } else { $fpsVal = 0 } }
        else { $fpsVal = [double]$r }
    } catch { $fpsVal = 0 }
    $durVal = 0
    try { $durVal = [double]$durRaw } catch { $durVal = 0 }
    [pscustomobject]@{
        vcodec = $vj.streams[0].codec_name
        width = [int]$vj.streams[0].width
        height = [int]$vj.streams[0].height
        fps = $fpsVal
        acodec = $aj.streams[0].codec_name
        duration = $durVal
    }
}

function Is-Compatible($meta,[switch]$Allow4K) {
    $okV = @('hevc','h265','h264','avc','vp9','av1','mpeg1video','mpeg2video','mpeg1','mpeg2','avs2')
    $okA = @('aac','mp3','wmav2','wma')
    $vok = $okV -contains $meta.vcodec
    $aok = $okA -contains $meta.acodec
    if (-not $vok -or -not $aok) { return $false }
    if ($meta.width -ge 3840 -or $meta.height -ge 2160) {
        if ($Allow4K) { return ($meta.fps -le 30) } else { return $false }
    }
    if ($meta.width -le 1920 -and $meta.height -le 1080) { return $true }
    return $false
}

function Get-Container($VideoCodec) {
    switch ($VideoCodec) {
        'hevc' { 'mp4' }
        'h264' { 'mp4' }
        'vp9' { 'webm' }
        'av1' { 'mkv' }
        'mpeg2' { 'mpg' }
        'mpeg1' { 'mpg' }
        default { 'mp4' }
    }
}

function Build-Args($meta,$VideoCodec,$AudioCodec,[switch]$Allow4K,$VideoBitrate,$AudioBitrate,[switch]$Overwrite) {
    $vargs = @()
    switch ($VideoCodec) {
        'hevc' { $vargs += '-c:v','libx265'; $vargs += '-tag:v','hvc1' }
        'h264' { $vargs += '-c:v','libx264'; $vargs += '-profile:v','high'; $vargs += '-level','4.1' }
        'vp9'  { $vargs += '-c:v','libvpx-vp9' }
        'av1'  { $vargs += '-c:v','libaom-av1' }
        'mpeg2' { $vargs += '-c:v','mpeg2video' }
        'mpeg1' { $vargs += '-c:v','mpeg1video' }
        default { $vargs += '-c:v','libx265'; $vargs += '-tag:v','hvc1' }
    }
    if ($VideoBitrate) { $vargs += '-b:v',$VideoBitrate }
    $tW = 1920; $tH = 1080
    if ($Allow4K) { $tW = 3840; $tH = 2160 }
    $needScale = ($meta.width -gt $tW -or $meta.height -gt $tH)
    if ($needScale) { $vargs += '-vf',"scale=${tW}:${tH}:force_original_aspect_ratio=decrease" }
    if (($meta.width -ge 3840 -or $meta.height -ge 2160) -and $Allow4K -and ($meta.fps -gt 30)) { $vargs += '-r','30' }
    $aargs = @()
    switch ($AudioCodec) {
        'aac' { $aargs += '-c:a','aac' }
        'mp3' { $aargs += '-c:a','libmp3lame' }
        'wma' { $aargs += '-c:a','wmav2' }
        default { $aargs += '-c:a','aac' }
    }
    if ($AudioBitrate) { $aargs += '-b:a',$AudioBitrate } else { $aargs += '-b:a','192k' }
    $args = @('-y')
    if (-not $Overwrite) { $args = @() }
    $args + $vargs + $aargs
}

function Build-AutoArgs($meta,[switch]$Allow4K,[string]$TargetContainer,[switch]$Overwrite) {
    $okV = @('hevc','h265','h264','avc','vp9','av1','mpeg1video','mpeg2video','avs2')
    $okA = @('aac','mp3','wmav2','wma')
    $canCopyV1080 = ($okV -contains $meta.vcodec) -and ($meta.width -le 1920 -and $meta.height -le 1080)
    $canCopyV4K = ($okV -contains $meta.vcodec) -and ($meta.width -ge 3840 -or $meta.height -ge 2160) -and ($meta.fps -le 30) -and $Allow4K
    $copyV = $canCopyV1080 -or $canCopyV4K
    $copyA = ($okA -contains $meta.acodec)
    $args = @()
    if ($Overwrite) { $args += '-y' }
    if ($copyV -and $copyA) {
        $args += '-c:v','copy','-c:a','copy'
        if ($AiAdvisor) { Write-Host ("Ai: kopiowanie v={0}, a={1}, kontener={2}" -f $meta.vcodec,$meta.acodec,$TargetContainer) -ForegroundColor DarkCyan }
        return $args
    }
    $tW = 1920; $tH = 1080
    if ($Allow4K) { $tW = 3840; $tH = 2160 }
    $needScale = ($meta.width -gt $tW -or $meta.height -gt $tH)
    $vcodec = 'libx265'; $vtag = 'hvc1'
    if (-not (Has-Encoder 'libx265')) { if (Has-Encoder 'libx264') { $vcodec = 'libx264'; $vtag = $null } elseif (Has-Encoder 'libvpx-vp9') { $vcodec = 'libvpx-vp9'; $vtag = $null } else { $vcodec = 'mpeg2video'; $vtag = $null } }
    $vargs = @('-c:v',$vcodec,'-pix_fmt','yuv420p')
    if ($vtag) { $vargs += '-tag:v',$vtag }
    if ($needScale) { $vargs += '-vf',"scale=${tW}:${tH}:force_original_aspect_ratio=decrease" }
    if (($meta.width -ge 3840 -or $meta.height -ge 2160) -and $Allow4K -and ($meta.fps -gt 30)) { $vargs += '-r','30' }
    $crf = 24
    if ($tW -ge 3840 -or $tH -ge 2160) { $crf = 26 }
    elseif ($tW -le 1280 -and $tH -le 720) { $crf = 22 }
    $vargs += '-crf',[string]$crf,'-preset','medium'
    $aargs = @()
    if ($copyA) { $aargs += '-c:a','copy' } else { if (Has-Encoder 'aac') { $aargs += '-c:a','aac','-b:a','192k' } elseif (Has-Encoder 'libmp3lame') { $aargs += '-c:a','libmp3lame','-b:a','192k' } elseif (Has-Encoder 'wmav2') { $aargs += '-c:a','wmav2','-b:a','192k' } else { $aargs += '-c:a','aac','-b:a','192k' } }
    if ($AiAdvisor) {
        $sc = if ($needScale) { "${tW}x${tH}" } else { "${meta.width}x${meta.height}" }
        $fpsInfo = if (($meta.width -ge 3840 -or $meta.height -ge 2160) -and $Allow4K -and ($meta.fps -gt 30)) { "${meta.fps}->30" } else { "${meta.fps}" }
        $aSel = $(if ($copyA) { 'copy' } else { $aargs[1] })
        Write-Host ("Ai: transkod v={0} crf={1} preset=medium pix=yuv420p scale={2} fps={3} a={4} kontener={5}" -f $vcodec,$crf,$sc,$fpsInfo,$aSel,$TargetContainer) -ForegroundColor DarkCyan
    }
    $args + $vargs + $aargs
}

function Compute-VMAF($ref,$dist,$tw,$th,$startSec,$durSec) {
    try {
        $guid = [System.Guid]::NewGuid().ToString()
        $log = "vmaf_${guid}.json"
        $fcs = "vmaf_${guid}.fcs"
        $modelOpt = ":model=version=vmaf_v0.6.1"
        $fg = "[0:v]scale=${tw}:${th}:force_original_aspect_ratio=decrease:flags=bicubic[rf];[1:v]scale=${tw}:${th}:force_original_aspect_ratio=decrease:flags=bicubic[ds];[ds][rf]libvmaf=log_path='${log}':log_fmt=json:n_threads=4:shortest=1${modelOpt}"
        $fg | Out-File -LiteralPath $fcs -Encoding ascii
        & ffmpeg -hide_banner -nostdin -ss $startSec -t $durSec -i $ref -ss $startSec -t $durSec -i $dist -filter_complex_script $fcs -f null - | Out-Null
        $j = Get-Content -LiteralPath $log -Raw | ConvertFrom-Json
        Remove-Item -LiteralPath $log -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $fcs -ErrorAction SilentlyContinue
        $score = $null
        if ($j.metrics -and $j.metrics.VMAF_mean) { $score = [double]$j.metrics.VMAF_mean }
        elseif ($j.pooled_metrics -and $j.pooled_metrics.vmaf -and $j.pooled_metrics.vmaf.mean) { $score = [double]$j.pooled_metrics.vmaf.mean }
        return $score
    } catch { return $null }
}

function Compute-SSIM($ref,$dist,$tw,$th,$startSec,$durSec) {
    try {
        $fltc = "[0:v]scale=${tw}:${th}:force_original_aspect_ratio=decrease:flags=bicubic[rf];[1:v]scale=${tw}:${th}:force_original_aspect_ratio=decrease:flags=bicubic[ds];[rf][ds]ssim"
        $o = & ffmpeg -hide_banner -nostdin -ss $startSec -t $durSec -i $ref -ss $startSec -t $durSec -i $dist -filter_complex $fltc -f null - 2>&1
        $text = ($o | Out-String)
        $m = [regex]::Match($text,'SSIM.*All:\s*([0-9\.]+)')
        if ($m.Success) { return [double]$m.Groups[1].Value } else { return $null }
    } catch { return $null }
}

function Append-Metrics($obj,$LogPath) {
    try {
        $line = ($obj | ConvertTo-Json -Compress)
        if ($LogPath) {
            $dir = Split-Path -Parent $LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            if (-not (Test-Path -LiteralPath $LogPath)) { New-Item -ItemType File -Path $LogPath | Out-Null }
            $line | Out-File -LiteralPath $LogPath -Append -Encoding utf8
        } else {
            $lp = Join-Path (Get-Location) 'conversion-log.jsonl'
            if (-not (Test-Path -LiteralPath $lp)) { New-Item -ItemType File -Path $lp | Out-Null }
            $line | Out-File -LiteralPath $lp -Append -Encoding utf8
        }
    } catch { Write-Host "Nie udało się zapisać logu" -ForegroundColor Yellow }
}

function Start-FFmpegWithProgress($ffmpegPath,$args,$duration,$activityName,$parentId) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ffmpegPath
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $false
    $psi.UseShellExecute = $false
    $psi.Arguments = ($args -join ' ')
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $p.Start() | Out-Null
    $spinner = '|','/','-','\'
    $si = 0
    while(-not $p.HasExited){
        if (-not $p.StandardOutput.EndOfStream) {
            $line = $p.StandardOutput.ReadLine()
            if ($null -ne $line){
                if ($line -like 'out_time_ms=*'){
                    $ms = [double]($line.Split('=')[1])
                    $sec = [double]($ms/1000000.0)
                    if ($duration -gt 0){
                        $pct = [math]::Min(100,[math]::Round(($sec/$duration)*100,2))
                        Write-Progress -Id 1 -ParentId $parentId -Activity $activityName -Status ("{0}% ({1}s)" -f $pct,[int]$sec) -PercentComplete $pct
                    } else {
                        $si = ($si + 1) % $spinner.Count
                        Write-Progress -Id 1 -ParentId $parentId -Activity $activityName -Status $spinner[$si] -PercentComplete 0
                    }
                }
            }
        } else {
            Start-Sleep -Milliseconds 100
        }
    }
    Write-Progress -Id 1 -ParentId $parentId -Activity $activityName -PercentComplete 100
    [pscustomobject]@{ ExitCode = $p.ExitCode }
}

function Convert-File($file,$VideoCodec,$AudioCodec,[switch]$Allow4K,$OutputDir,[switch]$Overwrite,[switch]$Force,$VideoBitrate,$AudioBitrate,[switch]$DryRun) {
    $meta = Probe-Video $file.FullName
    $compatible = Is-Compatible $meta -Allow4K:$Allow4K
    if ($compatible -and -not $Force) { Write-Host "Pomijanie (kompatybilny): $($file.FullName)" -ForegroundColor Green; return }
    $ext = Get-Container $VideoCodec
    if ($AutoProfile) {
        $okV = @('hevc','h265','h264','avc','vp9','av1','mpeg1video','mpeg2video','avs2')
        $canCopyV1080 = ($okV -contains $meta.vcodec) -and ($meta.width -le 1920 -and $meta.height -le 1080)
        $canCopyV4K = ($okV -contains $meta.vcodec) -and ($meta.width -ge 3840 -or $meta.height -ge 2160) -and ($meta.fps -le 30) -and $Allow4K
        $copyV = $canCopyV1080 -or $canCopyV4K
        if ($copyV) { $ext = Get-Container $meta.vcodec } else { $ext = 'mp4' }
    }
    $odir = if ($OutputDir) { $OutputDir } else { Join-Path $file.Directory.FullName 'converted' }
    if (-not (Test-Path $odir)) { New-Item -ItemType Directory -Path $odir | Out-Null }
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + '_proj.' + $ext
    $outPath = Join-Path $odir $outName
    $finalOut = $outPath
    if ((Test-Path -LiteralPath $finalOut) -and -not $Overwrite) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($outName)
        $extOut = [System.IO.Path]::GetExtension($outName)
        $n = 1
        do {
            $candidate = Join-Path $odir ("{0} ({1}){2}" -f $base,$n,$extOut)
            $n++
        } while (Test-Path -LiteralPath $candidate)
        $finalOut = $candidate
    }
    $args = if ($AutoProfile) { Build-AutoArgs $meta -Allow4K:$Allow4K $ext -Overwrite:$Overwrite } else { Build-Args $meta $VideoCodec $AudioCodec -Allow4K:$Allow4K $VideoBitrate $AudioBitrate -Overwrite:$Overwrite }
    $ffArgs = @('-hide_banner','-nostdin','-v','warning','-i',('"' + $file.FullName + '"')) + $args + @(('"' + $finalOut + '"'))
    Write-Host ("Konwersja: {0} -> {1}" -f $file.FullName,$finalOut) -ForegroundColor Cyan
    $cmdStr = "ffmpeg " + ($ffArgs -join ' ')
    Write-Host ("Polecenie: " + $cmdStr) -ForegroundColor DarkGray
    if ($DryRun) { Write-Host ("[DRY] pomijam uruchomienie") -ForegroundColor Yellow; return }
    & ffmpeg $ffArgs
    $ok = $false
    if (Test-Path -LiteralPath $finalOut) {
        $okProbe = $false
        try {
            $probe = ffprobe -hide_banner -v error -select_streams v:0 -show_entries stream=codec_name -of default=nw=1:nk=1 -- "$finalOut"
            if ($LASTEXITCODE -eq 0 -and $probe) { $okProbe = $true }
        } catch { $okProbe = $false }
        if ($okProbe) {
            $ok = $true
            Write-Host ("Zapisano: " + $finalOut) -ForegroundColor Green
        } else {
            Write-Host ("Nieprawidłowy plik wyjściowy: " + $finalOut) -ForegroundColor Yellow
            $plainArgs = @('-hide_banner','-nostdin','-v','warning','-i',$file.FullName) + $args + @($finalOut)
            & ffmpeg $plainArgs
            try {
                $probe2 = ffprobe -hide_banner -v error -select_streams v:0 -show_entries stream=codec_name -of default=nw=1:nk=1 -- "$finalOut"
                if ($LASTEXITCODE -eq 0 -and $probe2) { $ok = $true; Write-Host ("Naprawiono: " + $finalOut) -ForegroundColor Green }
            } catch { }
        }
    } else {
        Write-Host ("Błąd konwersji (ExitCode=" + $res.ExitCode + "; LASTEXITCODE=" + $LASTEXITCODE + ")") -ForegroundColor Red
        Write-Host ("Polecenie: " + $cmdStr) -ForegroundColor DarkYellow
    }
    $tw = 1920; $th = 1080
    if ($Allow4K) { $tw = 3840; $th = 2160 }
    $start = 0
    if ($meta.duration -gt ($SampleStart + $SampleSeconds)) { $start = $SampleStart } elseif ($meta.duration -gt $SampleSeconds) { $start = [int]([double]$meta.duration - $SampleSeconds) }
    $vmaf = $null
    $ssim = $null
    if ($QualityCheck -and $ok) {
        $vmaf = Compute-VMAF $file.FullName $finalOut $tw $th $start $SampleSeconds
        if ($null -eq $vmaf) { $ssim = Compute-SSIM $file.FullName $finalOut $tw $th $start $SampleSeconds }
    }
    $rec = [pscustomobject]@{
        source = $file.FullName
        output = $finalOut
        width = $meta.width
        height = $meta.height
        fps = $meta.fps
        vcodec_in = $meta.vcodec
        acodec_in = $meta.acodec
        allow4k = [bool]$Allow4K
        success = [bool]$ok
        vmaf = $vmaf
        ssim = $ssim
        time = (Get-Date).ToString('s')
    }
    Append-Metrics $rec $LogPath
    if ($LogPath) {
        if (-not (Test-Path -LiteralPath $LogPath)) {
            $rec | ConvertTo-Json -Compress | Out-File -LiteralPath $LogPath -Append -Encoding utf8
        }
    }
}

if (-not $Path) {
    $choice = Read-Host "Wybierz tryb [file/folder]"
    if ($choice -match '^f') { $Path = Select-Path -Mode 'file' } else { $Path = Select-Path -Mode 'folder' }
}

if (-not (Test-Path -LiteralPath $Path)) { Write-Error "Ścieżka nie istnieje: $Path"; exit 1 }

$items = Get-InputItems -InPath $Path -Recursive:$Recursive -ExcludePath $OutputDir
if (-not $items -or $items.Count -eq 0) { Write-Error "Brak plików wideo do przetworzenia"; exit 1 }

$i = 0
$total = $items.Count
foreach($f in $items) {
    $pctAll = if ($total -gt 0) { [math]::Round(($i/$total)*100,2) } else { 0 }
    Write-Progress -Id 0 -Activity 'Postep zadania' -PercentComplete $pctAll
    Convert-File $f $VideoCodec $AudioCodec -Allow4K:$Allow4K $OutputDir -Overwrite:$Overwrite -Force:$Force $VideoBitrate $AudioBitrate -DryRun:$DryRun
    $i++
}
Write-Progress -Id 0 -Activity 'Postep zadania' -Completed

Write-Host 'Zakonczono' -ForegroundColor Green
