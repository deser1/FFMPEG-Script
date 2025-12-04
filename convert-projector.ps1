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
    [switch]$AutoProfile
)

function Show-Help {
    Write-Host "Użycie: powershell -File convert-projector.ps1 [-Path <plik|katalog>] [-Recursive] [-Allow4K] [-VideoCodec <hevc|h264|vp9|av1|mpeg2|mpeg1>] [-AudioCodec <aac|mp3|wma>] [-OutputDir <katalog>] [-Overwrite] [-Force] [-VideoBitrate <np. 4M>] [-AudioBitrate <np. 192k>] [-DryRun]" -ForegroundColor Yellow
}

if ($Help) { Show-Help; exit 0 }

if (-not $PSBoundParameters.ContainsKey('AutoProfile')) { $AutoProfile = $true }

function Ensure-Tool($name) {
    $p = (Get-Command $name -ErrorAction SilentlyContinue)
    if (-not $p) { throw "Brak narzędzia '$name' w PATH" }
}

Ensure-Tool ffmpeg
Ensure-Tool ffprobe

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
    param([string]$InPath,[switch]$Recursive)
    $ext = @('*.mp4','*.mkv','*.mov','*.avi','*.wmv','*.mpeg','*.mpg','*.webm','*.m4v','*.ts','*.m2ts')
    if (Test-Path -LiteralPath $InPath) {
        if ((Get-Item -LiteralPath $InPath).PSIsContainer) {
            $items = foreach($e in $ext){ if ($Recursive) { Get-ChildItem -LiteralPath $InPath -Filter $e -File -Recurse } else { Get-ChildItem -LiteralPath $InPath -Filter $e -File } }
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
        return $args
    }
    $tW = 1920; $tH = 1080
    if ($Allow4K) { $tW = 3840; $tH = 2160 }
    $needScale = ($meta.width -gt $tW -or $meta.height -gt $tH)
    $vargs = @('-c:v','libx265','-tag:v','hvc1','-pix_fmt','yuv420p')
    if ($needScale) { $vargs += '-vf',"scale=${tW}:${tH}:force_original_aspect_ratio=decrease" }
    if (($meta.width -ge 3840 -or $meta.height -ge 2160) -and $Allow4K -and ($meta.fps -gt 30)) { $vargs += '-r','30' }
    $crf = 24
    if ($tW -ge 3840 -or $tH -ge 2160) { $crf = 26 }
    elseif ($tW -le 1280 -and $tH -le 720) { $crf = 22 }
    $vargs += '-crf',[string]$crf,'-preset','medium'
    $aargs = @()
    if ($copyA) { $aargs += '-c:a','copy' } else { $aargs += '-c:a','aac','-b:a','192k' }
    $args + $vargs + $aargs
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
    $args = if ($AutoProfile) { Build-AutoArgs $meta -Allow4K:$Allow4K $ext -Overwrite:$Overwrite } else { Build-Args $meta $VideoCodec $AudioCodec -Allow4K:$Allow4K $VideoBitrate $AudioBitrate -Overwrite:$Overwrite }
    $ffArgs = @('-hide_banner','-nostdin','-v','warning','-i',('"' + $file.FullName + '"')) + $args + @('-progress','pipe:1','-nostats',('"' + $finalOut + '"'))
    Write-Host ("Konwersja: {0} -> {1}" -f $file.FullName,$outPath) -ForegroundColor Cyan
    $cmdStr = "ffmpeg " + ($ffArgs -join ' ')
    Write-Host ("Polecenie: " + $cmdStr) -ForegroundColor DarkGray
    if ($DryRun) { Write-Host ("[DRY] pomijam uruchomienie") -ForegroundColor Yellow; return }
    $res = Start-FFmpegWithProgress 'ffmpeg' $ffArgs $meta.duration ("Konwersja: " + $file.Name) 0
    if (-not (Test-Path -LiteralPath $outPath)) {
        Write-Host "Ponowna próba bez progresu (diagnostyka)" -ForegroundColor Yellow
        $plainArgs = @('-hide_banner','-nostdin','-v','warning','-i',$file.FullName) + $args + @($outPath)
        & ffmpeg $plainArgs
    }
    if (Test-Path -LiteralPath $outPath) {
        Write-Host ("Zapisano: " + $outPath) -ForegroundColor Green
    } else {
        Write-Host ("Błąd konwersji (ExitCode=" + $res.ExitCode + "; LASTEXITCODE=" + $LASTEXITCODE + ")") -ForegroundColor Red
        Write-Host ("Polecenie: " + $cmdStr) -ForegroundColor DarkYellow
    }
}

if (-not $Path) {
    $choice = Read-Host "Wybierz tryb [file/folder]"
    if ($choice -match '^f') { $Path = Select-Path -Mode 'file' } else { $Path = Select-Path -Mode 'folder' }
}

if (-not (Test-Path -LiteralPath $Path)) { Write-Error "Ścieżka nie istnieje: $Path"; exit 1 }

$items = Get-InputItems -InPath $Path -Recursive:$Recursive
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
