#include <QApplication>
#include "mainwin.h"

int main(int argc, char* argv[]) {
    QApplication a(argc, argv);

    // 设置应用程序信息
    QApplication::setApplicationName("患者请假条生成系统");
    QApplication::setApplicationVersion("1.0");
    QApplication::setOrganizationName("新湖镇卫生院");

    mainwin w;
    w.show();

    return QApplication::exec();
}