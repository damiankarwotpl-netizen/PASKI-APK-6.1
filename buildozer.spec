[app]

# (str) Tytuł Twojej aplikacji
title = Future 9.4 ULTRA PRO

# (str) Nazwa pakietu (bez spacji i znaków specjalnych)
package.name = future_ultra_pro

# (str) Domena pakietu (używana do identyfikacji w Google Play)
package.domain = org.future.polska

# (str) Katalog źródłowy (tam gdzie masz main.py)
source.dir = .

# (list) Rozszerzenia plików, które mają zostać dołączone do APK
source.include_exts = py,png,jpg,kv,atlas,json,xlsx,db

# (str) Wersja aplikacji
version = 9.4

# (list) Biblioteki wymagane do działania (Kluczowe dla openpyxl i sqlite)
# sqlite3 jest wbudowane w python3, ale dodajemy go dla pewności.
# jdcal i et_xmlfile to zależności biblioteki openpyxl.
requirements = python3,kivy==2.3.0,openpyxl,jdcal,et_xmlfile,jnius,android

# (str) Orientacja ekranu
orientation = portrait

# -----------------------------------------------------------------------------
# Ustawienia Androida
# -----------------------------------------------------------------------------

# (list) Uprawnienia Androida - Kluczowe dla Internetu (SMTP) i plików
android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# (int) Android API (33 to Android 13)
android.api = 33

# (int) Minimalne API (21 to Android 5.0 - zapewnia dużą kompatybilność)
android.minapi = 21

# (int) Wersja NDK (zalecana dla aktualnego Buildozera)
android.ndk = 25b

# (bool) Czy aplikacja ma być pełnoekranowa
android.fullscreen = 0

# (list) Architektury procesorów (arm64 to standard dla nowych telefonów)
android.archs = arm64-v8a, armeabi-v7a

# (bool) Włączenie obsługi SQLite
android.enable_androidx = True

# (str) Ikona aplikacji (jeśli masz plik icon.png)
# icon.filename = %(source.dir)s/icon.png

# (str) Ekran powitalny (jeśli masz plik presplash.png)
# presplash.filename = %(source.dir)s/presplash.png

# -----------------------------------------------------------------------------
# Ustawienia Buildozera
# -----------------------------------------------------------------------------

[buildozer]

# (int) Poziom logowania (2 = debugowanie, pomocne przy błędach)
log_level = 2

# (int) Czy ostrzegać przed uruchomieniem jako root
warn_on_root = 1
