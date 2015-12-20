# Moca.NET Office

[![Build status](https://ci.appveyor.com/api/projects/status/4903crpjjo64th3s?svg=true)](https://ci.appveyor.com/project/miyabis/mocaoffice-dsr0a)
[![NuGet](https://img.shields.io/nuget/v/Moca.NETOffice.svg)](https://www.nuget.org/packages/Moca.NETOffice/)

Wrapper library of Excel. without the memory release of the Excel object.  
References of excel library is not required.  
we do not have wrapping all of the functions.

How to get
==========

URL:https://www.nuget.org/packages/Moca.NETOffice/
```
PM> Install-Package Moca.NETOffice
```

Programming
=======

```vb
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
```



License
=======

Microsoft Public License (MS-PL)

http://opensource.org/licenses/MS-PL
