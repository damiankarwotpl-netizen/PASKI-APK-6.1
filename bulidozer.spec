[app]

title = Paski Future
package.name = paskifuture10
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json

version = 1.0.0

requirements = python3,kivy,plyer,openpyxl,et_xmlfile,jdcal,xlrd==1.2.0,pyjnius

orientation = portrait

fullscreen = 0

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,INTERNET

android.api = 33
android.minapi = 24
android.ndk_api = 24

android.archs = arm64-v8a, armeabi-v7a

android.enable_androidx = True

android.allow_backup = False

android.entrypoint = org.kivy.android.PythonActivity

android.gradle_dependencies =

android.logcat_filters = *:S python:D

# buildozer
log_level = 2
warn_on_root = 1


[buildozer]

log_level = 2
