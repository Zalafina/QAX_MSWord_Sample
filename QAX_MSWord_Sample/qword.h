#ifndef QWORD_H
#define QWORD_H

#include <QDialog>
#include <QStyleFactory>
#include <QAxWidget>
#include <QAxObject>
#include <QTextDocument>
#include <QFileDialog>
#include <QDebug>

namespace Ui {
class QWord;
}

class QWord : public QDialog
{
    Q_OBJECT

public:
    explicit QWord(QWidget *parent = 0);
    ~QWord();

    static bool ExportHtml(const QString &filePath, const QTextDocument &document);

private slots:
    void on_createButton_clicked();

    void on_readButton_clicked();

    void on_writeButton_clicked();

private:
    Ui::QWord *ui;
};

#endif // QWORD_H
