#include "mainwin.h"
#include <QDir>
#include <QCoreApplication>

mainwin::mainwin(QWidget* parent) : QWidget(parent) {
    setupUI();
}

mainwin::~mainwin() {
}

void mainwin::setupUI() {
    // 设置窗口属性
    setWindowTitle("患者请假条生成系统");
    setFixedSize(800, 600);

    // 创建主布局
    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    // 标题
    QLabel* titleLabel = new QLabel("患者请假条生成系统", this);
    titleLabel->setAlignment(Qt::AlignCenter);
    QFont titleFont = titleLabel->font();
    titleFont.setPointSize(16);
    titleFont.setBold(true);
    titleLabel->setFont(titleFont);
    mainLayout->addWidget(titleLabel);

    // 文件选择区域
    QHBoxLayout* fileLayout = new QHBoxLayout();
    selectFileButton = new QPushButton("选择在院病人明细文件", this);
    filePathLabel = new QLabel("未选择文件", this);
    filePathLabel->setFrameStyle(QFrame::Panel | QFrame::Sunken);
    filePathLabel->setMinimumWidth(400);

    fileLayout->addWidget(selectFileButton);
    fileLayout->addWidget(filePathLabel);
    mainLayout->addLayout(fileLayout);

    // 患者表格
    QLabel* tableLabel = new QLabel("患者列表:", this);
    mainLayout->addWidget(tableLabel);

    patientTable = new QTableWidget(this);
    patientTable->setColumnCount(5);
    patientTable->setHorizontalHeaderLabels(QStringList() << "床位" << "姓名" << "性别" << "年龄" << "联系电话");
    patientTable->horizontalHeader()->setStretchLastSection(true);
    patientTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    patientTable->setSelectionBehavior(QAbstractItemView::SelectRows);
    patientTable->setMinimumHeight(200);
    mainLayout->addWidget(patientTable);

    // 按钮区域
    QHBoxLayout* buttonLayout = new QHBoxLayout();
    generateButton = new QPushButton("生成请假条", this);
    generateButton->setEnabled(false);
    buttonLayout->addStretch();
    buttonLayout->addWidget(generateButton);
    mainLayout->addLayout(buttonLayout);

    // 进度条
    progressBar = new QProgressBar(this);
    progressBar->setVisible(false);
    mainLayout->addWidget(progressBar);

    // 状态标签
    statusLabel = new QLabel("就绪", this);
    statusLabel->setFrameStyle(QFrame::Panel | QFrame::Sunken);
    mainLayout->addWidget(statusLabel);

    // 连接信号槽
    connect(selectFileButton, &QPushButton::clicked, this, &mainwin::onSelectPatientFile);
    connect(generateButton, &QPushButton::clicked, this, &mainwin::onGenerateLeaveSlips);
}

void mainwin::onSelectPatientFile() {
    QString filePath = QFileDialog::getOpenFileName(this,
        "选择在院病人明细文件",
        "",
        "Excel Files (*.xlsx *.xls)");

    if (!filePath.isEmpty()) {
        currentPatientFilePath = filePath;
        filePathLabel->setText(QFileInfo(filePath).fileName());

        QList<QStringList> patientList;
        if (readPatientData(filePath, patientList)) {
            updatePatientTable(patientList);
            generateButton->setEnabled(true);
            statusLabel->setText(QString("成功加载 %1 名患者信息").arg(patientList.size()));
        } else {
            generateButton->setEnabled(false);
            statusLabel->setText("读取患者数据失败");
        }
    }
}

bool mainwin::readPatientData(const QString& filePath, QList<QStringList>& patientList) {
    try {
        QXlsx::Document xlsx(filePath);
        if (!xlsx.load()) {
            return false;
        }

        // 读取第一个工作表
        int row = 3; // 从第3行开始读取数据
        while (true) {
            // 使用 read() 方法而不是 cellAt() 来避免智能指针问题
            QVariant bedValue = xlsx.read(row, 2); // B列
            if (bedValue.isNull() || bedValue.toString().trimmed().isEmpty()) {
                break;
            }

            QString bed = bedValue.toString();
            QString name = xlsx.read(row, 3).toString(); // C列
            QString gender = xlsx.read(row, 4).toString(); // D列
            QString age = xlsx.read(row, 5).toString(); // E列
            QString phone = xlsx.read(row, 6).toString();

            if (!name.isEmpty()) {
                patientList.append(QStringList() << bed << name << gender << age << phone);
            }

            row++;
        }

        return !patientList.isEmpty();
    } catch (...) {
        return false;
    }
}

void mainwin::updatePatientTable(const QList<QStringList>& patientList) {
    patientTable->setRowCount(patientList.size());

    for (int i = 0; i < patientList.size(); ++i) {
        const QStringList& patient = patientList[i];
        for (int j = 0; j < patient.size(); ++j) {
            QTableWidgetItem* item = new QTableWidgetItem(patient[j]);
            patientTable->setItem(i, j, item);
        }
    }

    patientTable->resizeColumnsToContents();
}

void mainwin::onGenerateLeaveSlips() {
    if (currentPatientFilePath.isEmpty()) {
        QMessageBox::warning(this, "警告", "请先选择在院病人明细文件");
        return;
    }

    QString outputPath = QFileDialog::getSaveFileName(this,
        "保存请假条文件",
        "患者请假条_" + QDateTime::currentDateTime().toString("yyyyMMdd_hhmmss") + ".xlsx",
        "Excel Files (*.xlsx)");

    if (outputPath.isEmpty()) {
        return;
    }

    // 读取患者数据
    QList<QStringList> patientList;
    if (!readPatientData(currentPatientFilePath, patientList)) {
        QMessageBox::critical(this, "错误", "读取患者数据失败");
        return;
    }

    progressBar->setVisible(true);
    progressBar->setRange(0, patientList.size());
    statusLabel->setText("正在生成请假条...");

    // 生成请假条
    if (generateLeaveSlips(patientList, outputPath)) {
        progressBar->setValue(patientList.size());
        statusLabel->setText(QString("成功生成 %1 名患者的请假条").arg(patientList.size()));
        QMessageBox::information(this, "成功", "请假条生成完成！\n文件已保存至: " + outputPath);
    } else {
        statusLabel->setText("生成请假条失败");
        QMessageBox::critical(this, "错误", "生成请假条失败");
    }

    progressBar->setVisible(false);
}

bool mainwin::generateLeaveSlips(const QList<QStringList>& patientList, const QString& outputPath) {
    try {
        QXlsx::Document outputDoc;

        // 创建格式
        QXlsx::Format wrapFormat;
        wrapFormat.setTextWrap(true);
        wrapFormat.setFontName("宋体");
        wrapFormat.setFontSize(12);
        // 不设置边框，请假条内部无边框

        QXlsx::Format titleFormat;
        titleFormat.setTextWrap(true);
        titleFormat.setFontBold(true);
        titleFormat.setFontName("宋体");
        titleFormat.setFontSize(12);
        titleFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        // 不设置边框，请假条内部无边框

        // 添加请假条之间的分隔线格式
        QXlsx::Format borderFormat;
        borderFormat.setBottomBorderStyle(QXlsx::Format::BorderThick); // 设置粗底边框

        // 每2个患者一个工作表
        int sheetCount = (patientList.size() + 1) / 2; // 向上取整

        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            QString sheetName = QString("Sheet%1").arg(sheetIndex + 1);

            // 如果不是第一个工作表，需要添加新工作表
            if (sheetIndex > 0) {
                outputDoc.addSheet(sheetName);
            }

            // 使用 selectSheet 而不是 setCurrentWorksheet
            outputDoc.selectSheet(sheetName);

            // 设置列宽为47字符
            outputDoc.setColumnWidth(1, 48); // A列
            outputDoc.setColumnWidth(2, 48); // B列

            // 计算当前工作表的两个患者索引
            int patient1Index = sheetIndex * 2;
            int patient2Index = patient1Index + 1;

            // 填充第一个患者的三份请假条（左侧）
            if (patient1Index < patientList.size()) {
                const QStringList& patient1 = patientList[patient1Index];

                // 第一份请假条
                outputDoc.write("A1", "新湖镇卫生院住院病人请假条", titleFormat);

                // 填写患者信息
                outputDoc.write("A2", "姓名：" + patient1[1], wrapFormat); // 姓名在索引1
                outputDoc.write("A3", "性别：" + patient1[2], wrapFormat); // 性别在索引2
                outputDoc.write("A4", "年龄：" + patient1[3], wrapFormat); // 年龄在索引3
                outputDoc.write("A5", "床号：" + patient1[0], wrapFormat); // 床号在索引0

                // 请假条内容
                outputDoc.write("A6", "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("A7", "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A8", "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A9",  "4. 离院期间联系方式：" + patient1[4], wrapFormat);
                outputDoc.write("A10", "患者(或家属)签字：_____________", wrapFormat);
                outputDoc.write("A11", "日期：____年____月____日", wrapFormat);
                outputDoc.write("A12", "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("A13", "电话： 2440120    2443720", borderFormat);

                // 第二份请假条（向下偏移13行）
                int secondStart = 14;
                outputDoc.write("A" + QString::number(secondStart), "新湖镇卫生院住院病人请假条", titleFormat);
                outputDoc.write("A" + QString::number(secondStart + 1), "姓名：" + patient1[1], wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 2), "性别：" + patient1[2], wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 3), "年龄：" + patient1[3], wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 4), "床号：" + patient1[0], wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 5), "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 6), "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 7), "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 8), "4. 离院期间联系方式：" + patient1[4], wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 9),  "患者(或家属)签字：_____________", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 10), "日期：____年____月____日", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 11), "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("A" + QString::number(secondStart + 12), "电话： 2440120    2443720", borderFormat);

                // 第三份请假条（再向下偏移13行）
                int thirdStart = 27;
                outputDoc.write("A" + QString::number(thirdStart), "新湖镇卫生院住院病人请假条", titleFormat);
                outputDoc.write("A" + QString::number(thirdStart + 1), "姓名：" + patient1[1], wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 2), "性别：" + patient1[2], wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 3), "年龄：" + patient1[3], wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 4), "床号：" + patient1[0], wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 5), "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 6), "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 7), "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 8),  "4. 离院期间联系方式：" + patient1[4], wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 9), "患者(或家属)签字：_____________", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 10), "日期：____年____月____日", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 11), "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("A" + QString::number(thirdStart + 12), "电话： 2440120    2443720", borderFormat);
            }

            // 填充第二个患者的三份请假条（右侧）
            if (patient2Index < patientList.size()) {
                const QStringList& patient2 = patientList[patient2Index];

                // 第一份请假条
                outputDoc.write("B1", "新湖镇卫生院住院病人请假条", titleFormat);

                // 填写患者信息
                outputDoc.write("B2", "姓名：" + patient2[1], wrapFormat);
                outputDoc.write("B3", "性别：" + patient2[2], wrapFormat);
                outputDoc.write("B4", "年龄：" + patient2[3], wrapFormat);
                outputDoc.write("B5", "床号：" + patient2[0], wrapFormat);

                // 请假条内容
                outputDoc.write("B6", "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("B7", "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B8", "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B9",  "4. 离院期间联系方式：" + patient2[4], wrapFormat);
                outputDoc.write("B10", "患者（或家属）签字：_____________", wrapFormat);
                outputDoc.write("B11", "日期：____年____月____日", wrapFormat);
                outputDoc.write("B12", "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("B13", "电话： 2440120    2443720", borderFormat);

                // 第二份请假条
                int secondStart = 14;
                outputDoc.write("B" + QString::number(secondStart), "新湖镇卫生院住院病人请假条", titleFormat);
                outputDoc.write("B" + QString::number(secondStart + 1), "姓名：" + patient2[1], wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 2), "性别：" + patient2[2], wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 3), "年龄：" + patient2[3], wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 4), "床号：" + patient2[0], wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 5), "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 6), "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 7), "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 8),  "4. 离院期间联系方式：" + patient2[4], wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 9), "患者（或家属）签字：_____________", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 10), "日期：____年____月____日", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 11), "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("B" + QString::number(secondStart + 12), "电话： 2440120    2443720", borderFormat);

                // 第三份请假条
                int thirdStart = 27;
                outputDoc.write("B" + QString::number(thirdStart), "新湖镇卫生院住院病人请假条", titleFormat);
                outputDoc.write("B" + QString::number(thirdStart + 1), "姓名：" + patient2[1], wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 2), "性别：" + patient2[2], wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 3), "年龄：" + patient2[3], wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 4), "床号：" + patient2[0], wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 5), "1. 请假原因：__________________________（如：家庭事务、检查复查等）", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 6), "2. 离院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 7), "3. 预计返院时间：____年____月____日____时", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 8),  "4. 离院期间联系方式：" + patient2[4], wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 9), "患者（或家属）签字：_____________", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 10), "日期：____年____月____日", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 11), "1. 需经主治医生评估病情稳定性后批准；2.返院后需及时向报到", wrapFormat);
                outputDoc.write("B" + QString::number(thirdStart + 12), "电话： 2440120    2443720", borderFormat);
            }

            progressBar->setValue(sheetIndex + 1);
            QCoreApplication::processEvents(); // 使用 QCoreApplication 而不是 QApplication
        }

        return outputDoc.saveAs(outputPath);
    } catch (...) {
        return false;
    }
}