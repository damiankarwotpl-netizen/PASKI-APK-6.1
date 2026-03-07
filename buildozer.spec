[app]

# (str) Title of your application
title = Paski Future

# (str) Package name
package.name = paskifuture10

# (str) Package domain
package.domain = org.test

# (str) Source code where the main.py lives
source.dir = .

# (list) Source files to include
source.include_exts = py,png,jpg,kv,json,txt

# (str) Application version
version = 1.0.0

# (list) Application requirements
requirements = python3,kivy,plyer,openpyxl,et_xmlfile,jdcal,xlrd==1.2.0,pyjnius,reportlab

# (str) Orientation
orientation = portrait

# (bool) Fullscreen
fullscreen = 0

# (list) Permissions
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# (int) Android API
android.api = 33

# (int) Minimum API
android.minapi = 24

# (int) NDK API
android.ndk_api = 24

# (str) Android NDK version
android.ndk = 25b

# (list) Architectures
android.archs = arm64-v8a, armeabi-v7a

# (bool) Enable AndroidX
android.enable_androidx = True

# (bool) Allow backup
android.allow_backup = False

# (str) Entry point
android.entrypoint = org.kivy.android.PythonActivity

# (str) Logcat filters
android.logcat_filters = *:S python:D

# (bool) Copy libs
android.copy_libs = 1

# (int) Window soft input
android.window_softinput_mode = adjustResize

# (bool) Use logcat
android.logcat = True

# (bool) Enable multiwindow
android.multiwindow = False


# ------------------------------------------------------------------

[buildozer]

# (int) Log level
log_level = 2

# (int) Warn on root
warn_on_root = 1
