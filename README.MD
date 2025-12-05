Przewodnik użytkownika – convert-projector.ps1

Opis

- Konwersja plików wideo do formatu zgodnego z projektorem: kontener `mp4`, wideo `HEVC (H.265)` lub automatyczny wybór, audio `AAC`, skalowanie do `1080p` lub opcjonalnie `4K` (≤30 fps).
- Obsługa pojedynczego pliku lub katalogu, tryb rekurencyjny, pomijanie już kompatybilnych plików, wymuszenie rekodowania i nadpisywania.
- Kontrola jakości: VMAF (model `version=vmaf_v0.6.1`) z automatycznym fallbackiem do SSIM, zapis wyników do `JSONL`.
- Progres w czasie rzeczywistym: wyświetlanie „X GB / Y GB” (ilość danych wyjściowych względem rozmiaru oryginału).

Wymagania

- `ffmpeg` i `ffprobe` w `PATH`.
- Wspierane wejścia: `*.mp4, *.mkv, *.mov, *.avi, *.wmv, *.mpeg, *.mpg, *.webm, *.m4v, *.ts, *.m2ts`.

Parametry i przełączniki

- `-Path <ścieżka>`: plik lub katalog do przetworzenia. Gdy brak, pojawia się wybór interaktywny.
- `-Recursive`: przetwarza podkatalogi.
- `-Allow4K`: pozwala zachować 4K (przy `fps ≤ 30`), w razie wyższego `fps` ustawia `30 fps`.
- `-VideoCodec <hevc|h264|vp9|av1|mpeg2|mpeg1>`: docelowy kodek wideo (gdy `-AutoProfile` nieaktywny).
- `-AudioCodec <aac|mp3|wma>`: docelowy kodek audio.
- `-OutputDir <katalog>`: docelowy katalog (domyślnie `converted` obok źródła).
- `-Overwrite`: nadpisuje istniejące pliki wynikowe.
- `-Force`: wymusza konwersję nawet gdy plik jest kompatybilny.
- `-VideoBitrate <np. 4M>`: ustawia bitrate wideo (tryb ręczny).
- `-AudioBitrate <np. 192k>`: ustawia bitrate audio.
- `-DryRun`: nie uruchamia konwersji, pokazuje polecenia.
- `-AutoProfile`: automatycznie decyduje o kopiowaniu vs transkodowaniu; domyślnie aktywne.
- `-AiAdvisor`: pokazuje rekomendacje dla wybranych parametrów.
- `-QualityCheck`: oblicza VMAF (lub SSIM) dla 60-sekundowej próbki.
- `-LogPath <plik.jsonl>`: ścieżka dla logu metryk.
- `-SampleSeconds <int>`: długość próbki do oceny jakości (domyślnie 30).
- `-SampleStart <int>`: start próbki w sekundach (domyślnie 60).

Domyślne zachowanie

- Wideo: `libx265`, `pix_fmt yuv420p`, tag `hvc1`, `-crf` dostosowany do celu (`1080p/4K`), `preset medium`.
- Audio: `AAC 192k` lub kopiowanie gdy kompatybilne.
- Skalowanie: `force_original_aspect_ratio=decrease` do `1920x1080` (lub do `3840x2160` przy `-Allow4K`).
- `-movflags +faststart`: optymalizacja MP4 (szybsze otwieranie).
- Pomiń kompatybilne pliki bez `-Force`.
- Dobór kontenera: `hevc/h264 -> mp4`, `vp9 -> webm`, `av1 -> mkv`, `mpeg1/mpeg2 -> mpg`.

Jakość (VMAF/SSIM)

- VMAF: `libvmaf` z `model=version=vmaf_v0.6.1`, kolejność wejść `[zrekodowany][referencja]`.
- Fallback: gdy VMAF nie powiedzie się, obliczany jest `SSIM` z komunikatu konsoli FFmpeg.
- Wyniki zapisują się do `JSONL` (każda linia to jeden rekord, z polami `vmaf`, `ssim`).

Logi i wyjście

- Struktura rekordów: `source`, `output`, wymiary/fps wejścia, kodeki wejściowe, `allow4k`, `success`, `vmaf`, `ssim`, `time`.
- Pliki wynikowe mają sufiks `_proj` i trafiają do podanego `OutputDir`.

Przykłady

- Interaktywnie z wyborem:
  - `powershell -File .\convert-projector.ps1`
- Katalog rekurencyjnie (1080p HEVC/AAC):
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Recursive`
- Zachowaj 4K (≤30 fps):
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy4K" -Allow4K`
- Wymuś konwersję:
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Force`
- Nadpisz wyniki:
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Overwrite`
- Ręcznie ustaw kodeki/bitrate:
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -VideoCodec h264 -AudioCodec mp3 -VideoBitrate 4M -AudioBitrate 192k`
- Własny katalog wyjściowy:
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -OutputDir "D:\Wyjscie"`
- Sprawdzenie jakości (próbka 60 s):
  - `powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -QualityCheck -SampleSeconds 60`

Uwagi techniczne

- Kodowanie `wmav2` może nie być dostępne we wszystkich buildach FFmpeg; alternatywnie użyj `-AudioCodec aac`.
- Skrypt nie podnosi rozdzielczości (tylko skalowanie w dół, zachowanie proporcji).
