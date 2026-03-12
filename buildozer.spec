[app]

# (str) Title of your application
title = Future Ultimate

# (str) Package name
package.name = futureultimate

# (str) Package domain (needed for android/ios packaging)
package.domain = com.yourcompany

# (str) Application versioning (method 1)
version = 20.0

# (list) Source files to include (let empty to include all the files
# in the current directory)
source.include_exts = py,png,jpg,kv,atlas

# (list) List of modules to be included (blacklist - usually empty)
# modules =

# (list) Application requirements
# comma separated e.g. requirements = sqlite3,kivy
requirements = python3,kivy==2.2.1,openpyxl,xlrd,requests,pandas,jnius,android

# (str) Kivy version if target is Android
# This is automatically picked up from `requirements` now.
# kivy.version = 2.2.1

# (str) Presplash file.
# presplash.filename = %(source.dir)s/data/presplash.png

# (str) Icon file.
# icon.filename = %(source.dir)s/data/icon.png

# (list) Permissions
# (https://developer.android.com/reference/android/Manifest.permission.html)
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# (int) Android API level to use
android.api = 33

# (int) Minimum API level for Android
android.minapi = 21

# (str) Android NDK version
android.ndk = 25b

# (str) Android SDK version (if not specified, the latest available is used)
# android.sdk = 26

# (str) Python version to use for Android build
android.python_version = 3.9

# (bool) Enable or disable Android debugging
android.debug = 1

# (bool) If set to 1, will compile your application with the
# Python 3 bootstrap.
android.enable_python3 = 1

# (list) Android device architectures to build for
# Default is armeabi-v7a.
# android.arch = arm64-v8a

# (list) Java libraries to include in the build
# android.add_libs_armeabi-v7a =

# (list) Java classes to include in the build
# android.add_classes_armeabi-v7a =

# (list) Libraries to include in the build
# android.add_libraries =

# (bool) If set to 1, will enable the use of the Android support library
# android.enable_multidex = 1

# (str) The default value for the orientation of the screen. Can be one of
# 'landscape', 'portrait', 'sensor'.
orientation = portrait

# (bool) If set to 1, will force the application to be fullscreen
fullscreen = 1

# (str) The name of the main application file (usually main.py)
main.py = main.py

#
# Python for android (p4a) configuration
#
[buildozer]

# (int) Log level (0 = error, 1 = warning, 2 = info, 3 = debug)
log_level = 2

# (str) The directory where buildozer stores all the build stuff
build_dir = .buildozer

# (str) The directory where buildozer stores all the distributions
dist_dir = .dist

# (list) List of targets to build
# target = android debug
