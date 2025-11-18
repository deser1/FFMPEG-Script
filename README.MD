Co zrobiłem

- Dodałem skrypt convert-projector.ps1 w bieżącym folderze do konwersji filmów do ustawień zgodnych z Twoim projektorem.
- Domyślnie koduje do HEVC (H.265) + AAC w kontenerze MP4 , ograniczając rozdzielczość do 1080p (z zachowaniem proporcji). Opcjonalnie pozwala zachować 4K , o ile nie przekracza 30 fps .
- Obsługuje wybór pojedynczego pliku lub katalogu (interaktywnie lub przez parametr), przetwarzanie rekurencyjne, pomijanie już kompatybilnych plików, wymuszenie rekodowania, nadpisywanie i tryb podglądu poleceń.
Jak używać

- Interaktywnie (wybór pliku lub folderu w okienku):
  - powershell -File .\convert-projector.ps1
- Katalog rekurencyjnie do 1080p HEVC/AAC:
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Recursive
- Zachowanie 4K jeśli ≤30 fps (inaczej ogranicza do 30 fps):
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy4K" -Allow4K
- Wymuszenie konwersji nawet gdy plik już kompatybilny:
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Force
- Nadpisywanie istniejących wyników:
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -Overwrite
- Własne kodeki i bitrate:
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -VideoCodec h264 -AudioCodec mp3 -VideoBitrate 4M -AudioBitrate 192k
- Własny katalog wyjściowy:
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -OutputDir "D:\Wyjscie"
- Podgląd, bez faktycznej konwersji (pokazuje polecenia ffmpeg):
  - powershell -File .\convert-projector.ps1 -Path "D:\Filmy" -DryRun
- Pomoc:
  - powershell -File .\convert-projector.ps1 -Help
Domyślne decyzje i zgodność

- Wideo: HEVC (libx265) z tagiem hvc1 w MP4 ( convert-projector.ps1:98–108 , convert-projector.ps1:86–95 ).
- Audio: AAC 192 kb/s ( convert-projector.ps1:116–123 ).
- Rozdzielczość: skaluje do 1920x1080 (bez rozciągania, force_original_aspect_ratio=decrease ) lub do 3840x2160 przy -Allow4K ( convert-projector.ps1:109–114 ).
- 4K powyżej 30 fps: obniża fps do 30 jeśli -Allow4K ( convert-projector.ps1:114 ).
- Pomija pliki już kompatybilne, chyba że -Force ( convert-projector.ps1:128–133 ).
- Kontenery wg kodeka wideo:
  - hevc/h264 -> mp4 , vp9 -> webm , av1 -> mkv , mpeg1/mpeg2 -> mpg ( convert-projector.ps1:86–95 ).
- Dozwolone wejściowe kodeki wideo: HEVC/VP9/AVS2/AVC/H.264/AV1/MPEG1/MPEG2 ; audio: MP3/WMA/AAC . Sprawdzenie kompatybilności: convert-projector.ps1:73–84 .
Wymagania

- ffmpeg i ffprobe muszą być w PATH . Skrypt sprawdza to na starcie ( convert-projector.ps1:25–31 ).
- Katalog wyników domyślnie tworzy się jako converted obok źródeł, a pliki mają sufiks _proj ( convert-projector.ps1:133–136 ).
Ważne funkcje w skrypcie

- Show-Help : convert-projector.ps1:19
- Wybór ścieżki (plik/folder): convert-projector.ps1:33–48
- Zbieranie plików wejściowych (z rekurencją): convert-projector.ps1:50–59
- Odczyt metadanych ffprobe: convert-projector.ps1:61–71
- Sprawdzanie zgodności: convert-projector.ps1:73–84
- Budowa parametrów ffmpeg: convert-projector.ps1:98–126
- Konwersja pojedynczego pliku: convert-projector.ps1:128–142
- Pętla przetwarzania i walidacja ścieżki: convert-projector.ps1:144–156
Uwagi techniczne

- WMA kodowanie ( wmav2 ) może nie być dostępne w każdej kompilacji ffmpeg; w razie problemów użyj -AudioCodec aac .
- AVS2 traktuję jako akceptowalny format wejściowy; docelowe kodowanie do AVS2 nie jest konfigurowane.
- Skrypt nie podnosi rozdzielczości; skaluje tylko w dół, zachowując proporcje.