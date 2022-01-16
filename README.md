# Moca.NET Office

[![NuGet](https://img.shields.io/nuget/v/Moca.NETOffice.svg)](https://www.nuget.org/packages/Moca.NETOffice/)
[![NuGet Pre Release](https://img.shields.io/nuget/vpre/Moca.NETOffice.svg)](https://www.nuget.org/packages/Moca.NETOffice/)
[![NuGet](https://img.shields.io/nuget/dt/Moca.NETOffice.svg)](https://www.nuget.org/packages/Moca.NETOffice/)
[![license](https://img.shields.io/badge/License-MS--PL-blue.svg)](https://opensource.org/licenses/MS-PL)

## Overview
Wrapper library of Excel. without the memory release of the Excel object.  
References of excel library is not required.  
we do not have wrapping all of the functions.

## Support for multiple frameworks
### Microsoft .NET Framework
* 2.0
* 3.5
* 4.0
* 4.5
* 4.5.2
* 4.6
* 4.6.2
* 4.7
* 4.7.2
* 4.8

## How to get

URL:https://www.nuget.org/packages/Moca.NETOffice/
```
PM> Install-Package Moca.NETOffice
```

## Programming

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


## Other Libraries

[Moca.NET Organization](https://github.com/mocanet)

## Visual Studio Extensions

* [Moca.NET Template Extension](https://marketplace.visualstudio.com/items?itemName=MiYABiS.MocaNETTemplate30)
* [Moca.NET Snippets Extension](https://marketplace.visualstudio.com/items?itemName=MiYABiS.MocaNETCodeSnippet)

## Sample

* Web Form Application  
  * http://miyabis.github.io/Moca.NET-WebAppDemo/  
  * https://code.msdn.microsoft.com/vstudio/MocaNET-Framework-Web-0e8d6dd7

* Windows Form Application  
  * http://miyabis.github.io/Moca.NET-WinAppDemo/  
  * https://code.msdn.microsoft.com/vstudio/MocaNET-Framework-Windows-7174d250

## For Development

* Visual Studio 2019

## License

Microsoft Public License (MS-PL)

http://opensource.org/licenses/MS-PL
