Imports System.Text
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Moca.Office.Excel

<TestClass()> Public Class ExcelTest

    <TestMethod()>
    Public Sub ExportAsFixedFormatTest()
        Using xlsx As ExcelWrapper = New ExcelWrapper()
            Dim book As BookWrapper
            book = xlsx.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "Book1.xlsx"))

            book.ActiveSheet.Range("9:9").Copy()
            book.ActiveSheet.Range("10:10").Insert()

            book.ActiveSheet.Cells(10, 2).Value = 5
            book.ActiveSheet.Cells(10, 3).Value = 14
            book.ActiveSheet.Cells(10, 4).Value = 24
            book.ActiveSheet.Cells(10, 5).Value = 34
            book.ActiveSheet.Cells(10, 6).Value = 44

            book.ExportAsFixedFormat(FixedFormatType.XPS, "C:\Temp\Test.xps")
            book.ExportAsFixedFormat(FixedFormatType.PDF, "C:\Temp\Test.pdf", , , , , , True)

            book.Saved = True
            xlsx.DisplayAlerts = False
        End Using
    End Sub

End Class
