Imports System.Text
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Moca.Office.Excel

<TestClass()> Public Class ExcelTest

    <TestMethod()>
    Public Sub ExportAsFixedFormatTest()
        Using xlsx As ExcelWrapper = New ExcelWrapper()
            Dim book As BookWrapper
            book = xlsx.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Book1.xlsx"))

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

    <TestMethod()>
    Public Sub NamesTest()
        Using xlsx As ExcelWrapper = New ExcelWrapper()
            Dim book As BookWrapper
            Dim names As NamesWrapper
            xlsx.DisplayAlerts = False
            book = xlsx.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Book1.xlsx"))
            names = book.Names

            Assert.AreEqual(xlsx, names.App)
            Assert.AreEqual(book, names.Book)
            Assert.AreEqual("Book1.xlsx", names.Parent.Name)

            Assert.AreEqual(5, names.Count)

            names.Add("TestBook", "$A$1")
            Assert.AreEqual(6, names.Count)
        End Using
    End Sub

    <TestMethod()>
    Public Sub NameTest()
        Using xlsx As ExcelWrapper = New ExcelWrapper()
            Dim book As BookWrapper
            Dim names As NamesWrapper
            Dim name As NameWrapper
            xlsx.DisplayAlerts = False
            book = xlsx.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Book1.xlsx"))
            names = book.Names
            name = names.Item(2)

            Assert.AreEqual(xlsx, name.App)
            Assert.AreEqual(names, name.Names)
            Assert.AreEqual("Book1.xlsx", name.Parent.Name)

            Assert.AreEqual("ColB", name.Name)
            Assert.AreEqual(2, name.Index)
            'Debug.Print(name.Category)
            'Debug.Print(name.CategoryLocal)
            Assert.AreEqual(CType(-4142, XlXLMMacroType), name.MacroType)
            Assert.AreEqual("ColB", name.NameLocal)
            Assert.AreEqual("=Sheet1!$C$5:$C$9", name.RefersTo)
            Assert.AreEqual("=Sheet1!$C$5:$C$9", name.RefersToLocal)
            Assert.AreEqual("=Sheet1!R5C3:R9C3", name.RefersToR1C1)
            Assert.AreEqual("B", CType(name.RefersToRange.Value, Object(,)).GetValue(1, 1))
            Assert.AreEqual(10.0, CType(name.RefersToRange.Value, Object(,)).GetValue(2, 1))
            Assert.AreEqual(11.0, CType(name.RefersToRange.Value, Object(,)).GetValue(3, 1))
            Assert.AreEqual(12.0, CType(name.RefersToRange.Value, Object(,)).GetValue(4, 1))
            Assert.AreEqual(13.0, CType(name.RefersToRange.Value, Object(,)).GetValue(5, 1))
            'Debug.Print(name.ShortcutKey)
            Assert.AreEqual("=Sheet1!$C$5:$C$9", name.Value)
            Assert.IsTrue(name.Visible)
        End Using
    End Sub

    <TestMethod()>
    Public Sub RangeTest()
        Using xlsx As ExcelWrapper = New ExcelWrapper()
            Dim book As BookWrapper
            Dim sheet As SheetWrapper
            Dim range As RangeWrapper
            xlsx.DisplayAlerts = False
            book = xlsx.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "Doc\Book1.xlsx"))
            sheet = book.Sheets(1)
            range = sheet.Cells(11, 6)

            Assert.AreEqual(xlsx, range.App)
            Assert.AreEqual(sheet, range.Sheet)
            Assert.AreEqual(-2146826273, range.Value)
            Assert.AreEqual(-2146826273, range.Value2)
            Assert.AreEqual("#VALUE!", range.Text)
        End Using
    End Sub

End Class
