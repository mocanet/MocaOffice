Imports System.IO
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Moca.Office
Imports Moca.Office.Word

<TestClass()>
Public Class WordTest

    <TestMethod()>
    Public Sub AppTest()
        Using app As WordWrapper = New WordWrapper()
            Dim documents As DocumentsWrapper
            Dim document As DocumentWrapper

            Assert.AreEqual(WdAlertLevel.wdAlertsNone, app.DisplayAlerts)
            Assert.AreEqual(True, app.ScreenUpdating)

            documents = app.Documents
            Assert.IsNotNull(documents)
            Assert.AreEqual(app, documents.App)

            document = documents.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Word.docx"))
            Assert.IsNotNull(document)
            Assert.AreEqual(app, document.App)

            Assert.IsFalse(app.Visible)
        End Using
    End Sub

    <TestMethod()>
    Public Sub ExportAsFixedFormatTest()
        Using app As WordWrapper = New WordWrapper()
            Dim document As DocumentWrapper
            document = app.Documents.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Word.docx"))


            document.ExportAsFixedFormat("C:\Temp\Test.xps", WdExportFormat.wdExportFormatXPS)
            document.ExportAsFixedFormat("C:\Temp\Test.pdf", WdExportFormat.wdExportFormatPDF)

            'document.Saved = True
        End Using
    End Sub

End Class
