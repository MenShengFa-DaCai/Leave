#ifndef QINGJIA2_MAINWIN_H
#define QINGJIA2_MAINWIN_H

#include <QWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QLabel>
#include <QFileDialog>
#include <QMessageBox>
#include <QProgressBar>
#include <QTableWidget>
#include <QHeaderView>
#include <QFileInfo>
#include <QDateTime>
#include <memory>

// QXlsx 头文件
#include "xlsxdocument.h"
#include "xlsxformat.h"

class mainwin : public QWidget {
    Q_OBJECT

public:
    explicit mainwin(QWidget* parent = nullptr);
    ~mainwin() override;

private slots:
    void onSelectPatientFile();
    void onGenerateLeaveSlips();

private:
    void setupUI();
    bool readPatientData(const QString& filePath, QList<QStringList>& patientList);
    bool generateLeaveSlips(const QList<QStringList>& patientList, const QString& outputPath);
    void updatePatientTable(const QList<QStringList>& patientList);

    QPushButton* selectFileButton;
    QPushButton* generateButton;
    QLabel* filePathLabel;
    QLabel* statusLabel;
    QProgressBar* progressBar;
    QTableWidget* patientTable;
    QString currentPatientFilePath;
};

#endif // QINGJIA2_MAINWIN_H