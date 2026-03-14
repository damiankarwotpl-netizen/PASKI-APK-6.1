[app]
# (str) Title of your application
title = Future 10.1 ULTIMATE

# (str) Package name
package.name = futureultipro

# (str) Package domain (needed for android packaging)
package.domain = org.future.hr

# (str) Source code where the main.py live
source.dir = .

# (list) Source files to include (let empty to include all the files)
source.include_exts = py,png,jpg,kv,atlas,json

# (str) Application versioning (method 1)
version = 10.1

# (list) Application requirements
# UWAGA: Dodano jdcal i et_xmlfile - są niezbędne dla stabilności openpyxl!
requirements = python3, kivy==2.3.0, openpyxl, jdcal, et_xmlfile, jnius, android, sqlite3, requests, urllib3, xlrd, openssl, reportlab

# (str) Supported orientations
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

# (list) Permissions
android.permissions = READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, INTERNET

# (int) Target Android API, should be as high as possible.
android.api = 33

# (int) Minimum API your APK will support.
android.minapi = 21

# (str) Android NDK version to use
android.ndk = 25b

# (list) The Android architectures to build for
android.archs = armeabi-v7a, arm64-v8a

# (bool) enables Android auto backup feature (Android API >= 23)
android.allow_backup = True

# (str) Path to a custom manifest template
# android.manifest_template = manifest.tmpl

# (list) Android additionnal libraries to copy into libs/armeabi
# android.add_libs_armeabi = lib/armeabi/libgnustl_shared.so

[buildozer]
# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root (0 = NO, 1 = YES)
warn_on_root = 1
