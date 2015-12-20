
Namespace Excel

	''' <summary>
	''' シート内容を構成
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContents

		''' <summary></summary>
		Protected contents As ISheetContents

		''' <summary></summary>
		Protected sheet As SheetWrapper

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			Me.contents = sheetContents
			Me.sheet = sheet
		End Sub

#End Region

		''' <summary>
		''' 出力内容をセルへ設定する
		''' </summary>
		''' <remarks>
		''' </remarks>
		Public Overridable Sub WriteContents()
			' シート名設定
			If contents.SaveSheetName.Length > 0 Then
				Me.sheet.Name = contents.SaveSheetName
			End If

			contents.WriteContents(sheet)
		End Sub

	End Class

End Namespace
