/******************************************************************************
 *
 * Nazwa pliku: main.cpp
 * Autor: Gemini
 * Data: 18.09.2025
 *
 * Opis:
 * Finalna, wieloplatformowa wersja aplikacji "Centrum Dowodzenia".
 * Kod został dostosowany do poprawnego działania zarówno na systemie Windows,
 * jak i macOS, poprzez warunkowe kompilowanie ścieżek do zasobów.
 *
 ******************************************************************************/

#include <QApplication>
#include <QMainWindow>
#include <QWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QPushButton>
#include <QScrollArea>
#include <QFileDialog>
#include <QProcess>
#include <QMessageBox>
#include <QFileInfo>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QFile>
#include <QStandardPaths>
#include <QSplitter>
#include <QTextEdit>
#include <QLabel>
#include <QDir>
#include <QStyle>
#include <QMouseEvent>
#include <QResizeEvent>
#include <QTimer>
#include <QMenu>
#include <QInputDialog>

// Definicja struktury przechowującej informacje o skrypcie
struct ScriptEntry {
    QString name;
    QString path;
};

// Niestandardowa klasa przycisku, aby obsłużyć podwójne kliknięcie
class DoubleClickButton : public QPushButton {
    Q_OBJECT
public:
    using QPushButton::QPushButton; // Użyj konstruktorów z klasy bazowej

signals:
    void doubleClicked();

protected:
    void mouseDoubleClickEvent(QMouseEvent *event) override {
        if (event->button() == Qt::LeftButton) {
            emit doubleClicked();
        }
        QPushButton::mouseDoubleClickEvent(event);
    }
};


class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr) : QMainWindow(parent) {
        setWindowTitle("Centrum Dowodzenia Skryptami Python");
        setMinimumSize(800, 600);

        setupUI();
        applyModernStyle();
        connectSignalsAndSlots();
        loadConfiguration();

        m_resizeTimer = new QTimer(this);
        m_resizeTimer->setSingleShot(true);
        m_resizeTimer->setInterval(150); // Opóźnienie 150ms
        connect(m_resizeTimer, &QTimer::timeout, this, &MainWindow::updateGridLayout);
    }

protected:
    void closeEvent(QCloseEvent *event) override {
        saveConfiguration();
        QMainWindow::closeEvent(event);
    }

    void resizeEvent(QResizeEvent *event) override {
        QMainWindow::resizeEvent(event);
        m_resizeTimer->start(); // Uruchom timer przy każdej zmianie rozmiaru
    }

    void showEvent(QShowEvent* event) override {
        QMainWindow::showEvent(event);
        // Upewnij się, że siatka jest poprawnie ułożona przy pierwszym wyświetleniu
        QTimer::singleShot(0, this, &MainWindow::updateGridLayout);
    }


private slots:
    void addScript() {
        QString filePath = QFileDialog::getOpenFileName(this, "Wybierz skrypt Pythona", "", "Skrypty Python (*.py);;Wszystkie pliki (*.*)");
        if (!filePath.isEmpty()) {
            QFileInfo fileInfo(filePath);
            ScriptEntry newEntry = {fileInfo.fileName(), filePath};
            for(const auto& entry : m_scripts) {
                if (entry.path == newEntry.path) {
                    QMessageBox::warning(this, "Duplikat", "Ten skrypt już znajduje się na liście.");
                    return;
                }
            }
            m_scripts.append(newEntry);
            repopulateScriptGrid();
        }
    }

    void removeScript() {
        if (m_currentSelectedIndex >= 0 && m_currentSelectedIndex < m_scripts.size()) {
            m_scripts.removeAt(m_currentSelectedIndex);
            repopulateScriptGrid();
        }
    }

    void runScript() {
        if (m_pythonInterpreterPath.isEmpty() || !QFile::exists(m_pythonInterpreterPath)) {
            QMessageBox::warning(this, "Brak interpretera", "Nie znaleziono interpretera Pythona. Sprawdź, czy jest on dołączony do programu lub wskaż go ręcznie.");
            return;
        }

        int currentIndex = m_currentSelectedIndex;
        if (currentIndex >= 0 && currentIndex < m_scripts.size()) {
            QString scriptPath = m_scripts[currentIndex].path;
            m_outputView->clear();
            m_outputView->append(QString("--- Uruchamianie: %1 ---\n").arg(scriptPath));
            m_outputView->append(QString("--- Interpreter: %1 ---\n").arg(m_pythonInterpreterPath));

            m_process->start(m_pythonInterpreterPath, {scriptPath});

            if (!m_process->waitForStarted()) {
                m_outputView->append(QString("\n--- BŁĄD: Nie można uruchomić procesu: %1 ---").arg(m_pythonInterpreterPath));
                return;
            }
            m_runButton->setEnabled(false);
            m_stopButton->setEnabled(true);
        }
    }

    void stopScript() {
        if (m_process->state() == QProcess::Running) {
            m_process->kill();
            m_outputView->append("\n--- Proces został zatrzymany przez użytkownika. ---");
        }
    }

    void readProcessOutput() { m_outputView->append(m_process->readAllStandardOutput()); }
    void readProcessError() { m_outputView->append(m_process->readAllStandardError()); }

    void onProcessFinished(int exitCode, QProcess::ExitStatus exitStatus) {
        QString status = (exitStatus == QProcess::NormalExit) ? "zakończony normalnie" : "zakończony awaryjnie";
        m_outputView->append(QString("\n--- Proces %1 (kod wyjścia: %2) ---\n").arg(status).arg(exitCode));
        updateButtonStates();
        m_stopButton->setEnabled(false);
    }

    void updateButtonStates() {
        bool hasSelection = m_currentSelectedIndex != -1;
        m_runButton->setEnabled(hasSelection && m_process->state() == QProcess::NotRunning);
        m_removeButton->setEnabled(hasSelection);
    }

    void selectPythonInterpreter() {
#ifdef Q_OS_WIN
        QString filter = "Plik wykonywalny (*.exe);;Wszystkie pliki (*)";
#else
        QString filter = "Wszystkie pliki (*)";
#endif
        QString newPath = QFileDialog::getOpenFileName(this, "Wybierz plik wykonywalny Pythona", "", filter);
        if (!newPath.isEmpty()) {
            m_pythonInterpreterPath = newPath;
            updatePythonPathLabel();
        }
    }

    void updateGridLayout() {
        if (!m_scriptGridPanel) return;
        const int buttonMinWidth = 180;
        const int gridSpacing = m_scriptGridLayout->spacing();
        int availableWidth = m_scriptGridPanel->width();
        if (availableWidth <= 0) return;
        int newColumnCount = availableWidth / (buttonMinWidth + gridSpacing);
        if (newColumnCount < 1) newColumnCount = 1;
        if (newColumnCount != m_columnCount) {
            m_columnCount = newColumnCount;
            repopulateScriptGrid();
        }
    }

    void showScriptContextMenu(const QPoint &pos) {
        auto* button = qobject_cast<QPushButton*>(sender());
        if (!button) return;
        int index = m_scriptButtons.indexOf(button);
        if (index == -1) return;
        m_contextMenuScriptIndex = index;
        QMenu contextMenu(this);
        QAction *renameAction = contextMenu.addAction("Zmień nazwę");
        connect(renameAction, &QAction::triggered, this, &MainWindow::renameScript);
        contextMenu.exec(button->mapToGlobal(pos));
    }

    void renameScript() {
        if (m_contextMenuScriptIndex < 0 || m_contextMenuScriptIndex >= m_scripts.size()) return;
        const QString currentName = m_scripts[m_contextMenuScriptIndex].name;
        bool ok;
        QString newName = QInputDialog::getText(this, "Zmień nazwę skryptu", "Nowa nazwa:", QLineEdit::Normal, currentName, &ok);
        if (ok && !newName.isEmpty()) {
            m_scripts[m_contextMenuScriptIndex].name = newName;
            m_scriptButtons[m_contextMenuScriptIndex]->setText(newName);
        }
    }

private:
    void setupUI() {
        // Ta funkcja pozostaje bez zmian
        QWidget *centralWidget = new QWidget;
        setCentralWidget(centralWidget);
        QVBoxLayout *mainLayout = new QVBoxLayout(centralWidget);
        QWidget* topBar = new QWidget;
        QHBoxLayout* topBarLayout = new QHBoxLayout(topBar);
        topBarLayout->setContentsMargins(0, 0, 0, 0);
        m_selectPythonButton = new QPushButton("Zmień...");
        m_pythonPathLabel = new QLabel;
        m_pythonPathLabel->setWordWrap(true);
        QWidget* pythonSelectorWidget = new QWidget;
        QVBoxLayout* pythonSelectorLayout = new QVBoxLayout(pythonSelectorWidget);
        pythonSelectorLayout->addWidget(new QLabel("Interpreter Pythona:"));
        QHBoxLayout* pythonPathLayout = new QHBoxLayout;
        pythonPathLayout->addWidget(m_pythonPathLabel, 1);
        pythonPathLayout->addWidget(m_selectPythonButton);
        pythonSelectorLayout->addLayout(pythonPathLayout);
        m_addButton = new QPushButton("Dodaj skrypt");
        m_removeButton = new QPushButton("Usuń skrypt");
        m_runButton = new QPushButton("Uruchom");
        m_stopButton = new QPushButton("Zatrzymaj");
        m_addButton->setObjectName("addButton");
        m_runButton->setObjectName("runButton");
        m_removeButton->setObjectName("removeButton");
        m_stopButton->setObjectName("stopButton");
        topBarLayout->addWidget(pythonSelectorWidget, 1);
        topBarLayout->addStretch();
        topBarLayout->addWidget(m_addButton);
        topBarLayout->addWidget(m_removeButton);
        topBarLayout->addWidget(m_runButton);
        topBarLayout->addWidget(m_stopButton);
        mainLayout->addWidget(topBar);
        QSplitter *splitter = new QSplitter(Qt::Vertical, this);
        mainLayout->addWidget(splitter);
        QScrollArea* scrollArea = new QScrollArea;
        scrollArea->setWidgetResizable(true);
        scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
        scrollArea->setObjectName("scriptScrollArea");
        m_scriptGridPanel = new QWidget;
        m_scriptGridLayout = new QGridLayout(m_scriptGridPanel);
        m_scriptGridLayout->setSpacing(15);
        m_scriptGridLayout->setAlignment(Qt::AlignTop | Qt::AlignLeft);
        scrollArea->setWidget(m_scriptGridPanel);
        QWidget *consolePanel = new QWidget;
        QVBoxLayout *consoleLayout = new QVBoxLayout(consolePanel);
        consoleLayout->setContentsMargins(0,5,0,0);
        QLabel* outputLabel = new QLabel("Konsola wyjściowa:");
        m_outputView = new QTextEdit;
        m_outputView->setReadOnly(true);
        m_outputView->setFontFamily("Courier");
        consoleLayout->addWidget(outputLabel);
        consoleLayout->addWidget(m_outputView);
        splitter->addWidget(scrollArea);
        splitter->addWidget(consolePanel);
        splitter->setSizes({400, 300});
        m_process = new QProcess(this);
        updateButtonStates();
        m_stopButton->setEnabled(false);
    }

    void connectSignalsAndSlots() {
        connect(m_selectPythonButton, &QPushButton::clicked, this, &MainWindow::selectPythonInterpreter);
        connect(m_addButton, &QPushButton::clicked, this, &MainWindow::addScript);
        connect(m_removeButton, &QPushButton::clicked, this, &MainWindow::removeScript);
        connect(m_runButton, &QPushButton::clicked, this, &MainWindow::runScript);
        connect(m_stopButton, &QPushButton::clicked, this, &MainWindow::stopScript);
        connect(m_process, &QProcess::readyReadStandardOutput, this, &MainWindow::readProcessOutput);
        connect(m_process, &QProcess::readyReadStandardError, this, &MainWindow::readProcessError);
        connect(m_process, &QProcess::finished, this, &MainWindow::onProcessFinished);
    }

    void clearGridLayout(QLayout* layout) {
        if (!layout) return;
        QLayoutItem* item;
        while ((item = layout->takeAt(0)) != nullptr) {
            if (item->widget()) delete item->widget();
            delete item;
        }
    }

    void repopulateScriptGrid() {
        clearGridLayout(m_scriptGridLayout);
        m_scriptButtons.clear();
        m_currentSelectedIndex = -1;

        for (int i = 0; i < m_scripts.size(); ++i) {
            const auto& script = m_scripts[i];
            auto* button = new DoubleClickButton(script.name);
            button->setMinimumHeight(60);
            button->setToolTip(script.path);
            button->setProperty("isScriptButton", true);
            button->setProperty("selected", false);
            button->setContextMenuPolicy(Qt::CustomContextMenu);
            connect(button, &QPushButton::customContextMenuRequested, this, &MainWindow::showScriptContextMenu);
            connect(button, &QPushButton::clicked, this, [this, button, i]() {
                if (m_currentSelectedIndex != -1 && m_currentSelectedIndex < m_scriptButtons.size()) {
                    m_scriptButtons[m_currentSelectedIndex]->setProperty("selected", false);
                    style()->unpolish(m_scriptButtons[m_currentSelectedIndex]);
                    style()->polish(m_scriptButtons[m_currentSelectedIndex]);
                }
                m_currentSelectedIndex = i;
                button->setProperty("selected", true);
                style()->unpolish(button); style()->polish(button);
                updateButtonStates();
            });
            connect(button, &DoubleClickButton::doubleClicked, this, [this, i]() {
                if (m_currentSelectedIndex != i) {
                    if (m_currentSelectedIndex != -1 && m_currentSelectedIndex < m_scriptButtons.size()) {
                        m_scriptButtons[m_currentSelectedIndex]->setProperty("selected", false);
                        style()->unpolish(m_scriptButtons[m_currentSelectedIndex]); style()->polish(m_scriptButtons[m_currentSelectedIndex]);
                    }
                    m_currentSelectedIndex = i;
                    m_scriptButtons[i]->setProperty("selected", true);
                    style()->unpolish(m_scriptButtons[i]); style()->polish(m_scriptButtons[i]);
                }
                if (m_process->state() == QProcess::NotRunning) runScript();
            });
            m_scriptGridLayout->addWidget(button, i / m_columnCount, i % m_columnCount);
            m_scriptButtons.append(button);
        }
        updateButtonStates();
    }

    void applyModernStyle() {
        // Styl pozostaje bez zmian
        this->setStyleSheet(R"(
            QMainWindow, QWidget{background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2c2c3e, stop:1 #3a3a52); color: #f2f2f2; font-family: 'Segoe UI', Arial, sans-serif; font-size: 11pt;}
            #scriptScrollArea { border: none; background-color: transparent; }
            QLabel{font-weight: bold; padding-bottom: 5px; background-color: transparent;}
            QLabel#m_pythonPathLabel{font-weight: normal; font-style: italic; color: #a2a2c2;}
            QTextEdit{background-color: #1e1e2f; color: #ff79c6; border: 1px solid #4a4a68; border-radius: 5px;}
            QPushButton{background-color: #4a4a68; color: white; border: none; padding: 10px 15px; border-radius: 5px; font-weight: bold;}
            QPushButton:hover{background-color: #5f5f7a;} QPushButton:pressed{background-color: #3a3a52;}
            QPushButton:disabled{background-color: #3a3a52; color: #888;}
            #addButton, #runButton{background-color: #e91e63;} #addButton:hover, #runButton:hover{background-color: #c2185b;}
            #removeButton, #stopButton{background-color: #5d6d7e;} #removeButton:hover, #stopButton:hover{background-color: #808b96;}
            QPushButton[isScriptButton="true"]{background-color: #3a3a52; border: 1px solid #5f5f7a; text-align: center;}
            QPushButton[isScriptButton="true"]:hover{background-color: #4a4a68;}
            QPushButton[isScriptButton="true"][selected="true"]{background-color: #e91e63; border: 2px solid #ff79c6;}
            QSplitter::handle{background: #4a4a68;} QSplitter::handle:hover{background: #5f5f7a;} QSplitter::handle:pressed{background: #e91e63;}
            QMenu { background-color: #3a3a52; border: 1px solid #4a4a68; }
            QMenu::item { color: #f2f2f2; padding: 5px 25px 5px 25px; }
            QMenu::item:selected { background-color: #e91e63; }
        )");
        m_pythonPathLabel->setObjectName("m_pythonPathLabel");
    }

    void updatePythonPathLabel() {
        if (m_pythonInterpreterPath.isEmpty()) {
            m_pythonPathLabel->setText("Nie znaleziono");
            m_pythonPathLabel->setToolTip("Nie znaleziono osadzonego interpretera Pythona. Wskaż go ręcznie.");
        } else {
            QFileInfo fileInfo(m_pythonInterpreterPath);
            m_pythonPathLabel->setText(fileInfo.fileName());
            m_pythonPathLabel->setToolTip(m_pythonInterpreterPath);
        }
    }

    QString getConfigFilePath() {
        QString dataPath = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
        QDir dir(dataPath);
        if (!dir.exists()) dir.mkpath(".");
        return dataPath + "/python_runner_config.json";
    }

    void saveConfiguration() {
        QJsonObject configObject;
        configObject["python_path"] = m_pythonInterpreterPath;
        QJsonArray scriptsArray;
        for (const auto& entry : m_scripts) {
            QJsonObject scriptObject;
            scriptObject["name"] = entry.name;
            scriptObject["path"] = entry.path;
            scriptsArray.append(scriptObject);
        }
        configObject["scripts"] = scriptsArray;
        QFile file(getConfigFilePath());
        if (file.open(QIODevice::WriteOnly)) {
            file.write(QJsonDocument(configObject).toJson());
        }
    }

    void scanForInitialScripts() {
        QString scriptsPath;
#ifdef Q_OS_MAC
        // Na macOS, zasoby są w innym miejscu względem pliku .exe
        scriptsPath = QApplication::applicationDirPath() + "/../Resources/scripts";
#else
        // Ścieżka dla Windows
        scriptsPath = QApplication::applicationDirPath() + "/scripts";
#endif

        QDir scriptsDir(scriptsPath);
        if (!scriptsDir.exists()) return;

        QStringList nameFilters;
        nameFilters << "*.py";
        QFileInfoList scriptFiles = scriptsDir.entryInfoList(nameFilters, QDir::Files);

        for (const QFileInfo& fileInfo : scriptFiles) {
            m_scripts.append({fileInfo.fileName(), fileInfo.absoluteFilePath()});
        }
    }

    void loadConfiguration() {
        QString defaultPythonPath = "";
        QString embeddedPythonPath = "";

// Sprawdź, na jakim systemie operacyjnym jesteśmy
#ifdef Q_OS_MAC
        QString resourcePath = QApplication::applicationDirPath() + "/../Resources";
        embeddedPythonPath = resourcePath + "/python/bin/python3";
#else
        embeddedPythonPath = QApplication::applicationDirPath() + "/python/python.exe";
#endif

        if (QFile::exists(embeddedPythonPath)) {
            defaultPythonPath = embeddedPythonPath;
        }

        QFile file(getConfigFilePath());
        if (!file.open(QIODevice::ReadOnly)) {
            m_pythonInterpreterPath = defaultPythonPath;
        } else {
            QJsonDocument doc = QJsonDocument::fromJson(file.readAll());
            QJsonObject configObject = doc.object();
            m_pythonInterpreterPath = configObject.value("python_path").toString(defaultPythonPath);
            if (configObject.contains("scripts") && configObject["scripts"].isArray()) {
                m_scripts.clear();
                for (const auto& val : configObject["scripts"].toArray()) {
                    QJsonObject obj = val.toObject();
                    if (obj.contains("path") && QFile::exists(obj["path"].toString())) {
                        m_scripts.append({obj["name"].toString(), obj["path"].toString()});
                    }
                }
            }
        }

        if (m_scripts.isEmpty()) {
            scanForInitialScripts();
        }

        updatePythonPathLabel();
        repopulateScriptGrid();
    }

    // Elementy interfejsu i logiki
    QPushButton *m_addButton, *m_removeButton, *m_runButton, *m_stopButton, *m_selectPythonButton;
    QTextEdit *m_outputView;
    QLabel *m_pythonPathLabel;
    QWidget* m_scriptGridPanel;
    QGridLayout* m_scriptGridLayout;
    QString m_pythonInterpreterPath;
    QList<ScriptEntry> m_scripts;
    QList<QPushButton*> m_scriptButtons;
    int m_currentSelectedIndex = -1;
    int m_contextMenuScriptIndex = -1;
    QProcess *m_process;
    int m_columnCount = 4;
    QTimer *m_resizeTimer;
};

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);
    MainWindow window;
    window.show();
    return app.exec();
}

#include "main.moc"

