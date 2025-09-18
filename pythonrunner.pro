QT += core gui widgets

CONFIG += c++17

TEMPLATE = app
TARGET = PythonRunner

SOURCES += main.cpp

macx {
    ICON = icon.icns
    QMAKE_INFO_PLIST = Info.plist
}
