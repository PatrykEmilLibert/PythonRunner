QT += core gui widgets

CONFIG += c++17 cmdline

TEMPLATE = app
TARGET = PythonRunner

SOURCES += main.cpp

CONFIG += recheck_dependencies

# Use .icns icon for macOS
ICON = icon.icns

# macOS-specific Info.plist file
macx {
    QMAKE_INFO_PLIST = Info.plist
}
