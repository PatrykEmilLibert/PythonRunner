QT += core gui widgets

CONFIG += c++17 cmdline

TEMPLATE = app
TARGET = PythonRunner

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += main.cpp

CONFIG += recheck_dependencies

RC_ICONS = icon.ico

win32 {
    QMAKE_LFLAGS += -Wl,-subsystem,windows
}
