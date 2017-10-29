#include "qword.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    QApplication::setStyle(QStyleFactory::create("Fusion"));

    QWord w;

    Qt::WindowFlags flags = Qt::Dialog;
    flags |= Qt::WindowMinimizeButtonHint;
    flags |= Qt::WindowCloseButtonHint;
    w.setWindowFlags(flags);

    w.show();

    return a.exec();
}
