
Imports System.IO
Imports System.Reflection

Namespace Excel

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContentsFactory

		''' <summary>log4net logger</summary>
		Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

		''' <summary>
		''' 
		''' </summary>
		''' <param name="contents"></param>
		''' <param name="sheet"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function Create(ByVal contents As ISheetContents, ByVal sheet As SheetWrapper) As SheetContents
			Dim obj As Object
			Dim name As String

			obj = contents
			name = obj.GetType().GetInterfaces(0).Namespace & Type.Delimiter & obj.GetType().GetInterfaces(0).Name.Substring(1)

			_mylog.DebugFormat("SheetContentsFactory CreateInstance {0}", name)
			Try
				Return DirectCast(Activator.CreateInstance(Type.GetType(name), contents, sheet), SheetContents)
			Catch ex As Exception
				Throw New ExcelException(sheet.App, ex, String.Format("シートコンテンツクラスをインスタンス化できませんでした。[{0}]", name))
			End Try
		End Function

	End Class

End Namespace
