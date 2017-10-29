#include "qword.h"
#include "ui_qword.h"

QWord::QWord(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::QWord)
{
    ui->setupUi(this);
}

QWord::~QWord()
{
    delete ui;
}

bool QWord::ExportHtml(const QString &filePath, const QTextDocument &document)
{
     if(filePath.endsWith(".html"))
     {
          QFile file(filePath);
          if(file.open(QIODevice::WriteOnly|QIODevice::Text))
          {
                QTextStream out(&file);

                out.setCodec("utf-8");
                out<<document.toHtml("utf-8");
                file.close();
                return true;
          }
          return false;
     }
     return false;
}

void QWord::on_createButton_clicked()
{
    QString createtext = ui->createTextEdit->toPlainText();
    if(true == createtext.isEmpty()){
        return;
    }

    QString fileName = QFileDialog::getSaveFileName(this, tr("Create Word File"),
                                 "D:/qword_text.docx",
                                 tr("Word Document (*.docx *.doc)"));

    //qDebug() << "Create file name:" << fileName;

    if (true == fileName.isEmpty()){
        return;
    }

    QAxWidget *word = new QAxWidget("Word.Application", this, Qt::MSWindowsOwnDC);
    word->setProperty("DisplayAlerts", false);
    word->setProperty("Visible", false);

    QAxObject *documents = new QAxObject;
    documents = word->querySubObject("Documents");

    documents->dynamicCall("Add(void)");

    QAxObject *document = word->querySubObject("ActiveDocument");

    QAxObject *selection;

    selection = word->querySubObject("Selection");
    selection->dynamicCall("TypeText(const QString&)", createtext);

    document->dynamicCall("SaveAs (const QString&)", fileName);

    document->dynamicCall("Close (boolean)", false);
    word->dynamicCall("Quit(void)");
}

void QWord::on_readButton_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Open Word File"),
                                                      "D:/qword_text.docx",
                                                      tr("Word Document (*.docx *.doc)"));

    if (true == fileName.isEmpty()){
        return;
    }

    QAxWidget *word = new QAxWidget("Word.Application", this, Qt::MSWindowsOwnDC);
    word->setProperty("DisplayAlerts", false);
    word->setProperty("Visible", false);

    QAxObject *documents = new QAxObject;
    documents = word->querySubObject("Documents");

    documents->dynamicCall("Open(const QString&)", fileName);

    QAxObject *document = word->querySubObject("ActiveDocument");

    QAxObject *pRange = document->querySubObject("Range()");
    pRange->dynamicCall("WholeStory()");

    QString text = pRange->property("Text").toString();

    if (false == text.isEmpty()){
        ui->readTextEdit->setPlainText(text);
    }

#if 0   // Generate Interface Documents
    QAxObject *selection = word->querySubObject("Selection");
    QAxObject* characters = document->querySubObject("Characters");
    int count = characters->property("Count").toInt();

    QTextDocument tempDocument;
    QString document_html = document->generateDocumentation();
    tempDocument.setHtml(document_html);
    ExportHtml(QString("D:/ActiveDocument.html"), tempDocument);

    QString selection_html = selection->generateDocumentation();
    tempDocument.setHtml(selection_html);
    ExportHtml(QString("D:/Selection.html"), tempDocument);

    QString characters_html = characters->generateDocumentation();
    tempDocument.setHtml(characters_html);
    ExportHtml(QString("D:/Characters.html"), tempDocument);

    QString range_html = pRange->generateDocumentation();
    tempDocument.setHtml(range_html);
    ExportHtml(QString("D:/Range.html"), tempDocument);
#endif

    document->dynamicCall("Close (boolean)", false);
    word->dynamicCall("Quit(void)");

}

void QWord::on_writeButton_clicked()
{
    QString inserttext = ui->writeTextEdit->toPlainText();
    if(true == inserttext.isEmpty()){
        return;
    }

    QString fileName = QFileDialog::getOpenFileName(this, tr("Open Word File"),
                                                      "D:/qword_text.docx",
                                                      tr("Word Document (*.docx *.doc)"));

    if (true == fileName.isEmpty()){
        return;
    }

    QAxWidget *word = new QAxWidget("Word.Application", this, Qt::MSWindowsOwnDC);
    word->setProperty("DisplayAlerts", false);
    word->setProperty("Visible", false);

    QAxObject *documents = new QAxObject;
    documents = word->querySubObject("Documents");

    documents->dynamicCall("Open(const QString&)", fileName);

    QAxObject *document = word->querySubObject("ActiveDocument");

    QAxObject *pRange = document->querySubObject("Range()");
    pRange->dynamicCall("WholeStory()");

    pRange->dynamicCall("InsertAfter(const QString&)", QString("\r"));
    pRange->dynamicCall("InsertAfter(const QString&)", inserttext);

    document->dynamicCall("Save()");
    document->dynamicCall("Close (boolean)", false);
    word->dynamicCall("Quit(void)");
}
