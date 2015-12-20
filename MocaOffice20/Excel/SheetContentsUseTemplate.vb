
Namespace Excel

	''' <summary>
	''' 一覧形式のテンプレートシートを使用したときに、
	''' データを多次元配列に変換し値を設定する手法。
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContentsUseTemplate
		Inherits SheetContents

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			MyBase.New(sheetContents, sheet)
		End Sub

#End Region

		''' <summary>
		''' シートコンテンツを当クラスで使用するクラスへキャストする
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _cType() As ISheetContentsUseTemplate
			Return DirectCast(MyBase.contents, ISheetContentsUseTemplate)
		End Function

		''' <summary>
		''' 出力内容をセルへ設定する
		''' </summary>
		''' <remarks>
		''' </remarks>
		Public Overrides Sub WriteContents()
			MyBase.WriteContents()

			writeContentsTemplate()
		End Sub

		''' <summary>
		''' 一覧部分の設定
		''' </summary>
		''' <remarks></remarks>
		Protected Overridable Sub writeContentsTemplate()
			If _cType.DataCount <= 0 Then
				Exit Sub
			End If

			' 先頭行をデータ数分コピー
			rowCopy(_cType.DataCount _
			 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow + 1, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount - 1, _cType.StartCol))

			' 
			_writeContents()
		End Sub

		''' <summary>
		''' テンプレートとなる行を指定された内容でコピーする
		''' </summary>
		''' <remarks>
		''' </remarks>
		Protected Sub rowCopy(ByVal dataCount As Integer, ByVal rangeF1 As RangeWrapper, ByVal rangeF2 As RangeWrapper, ByVal rangeT1 As RangeWrapper, ByVal rangeT2 As RangeWrapper)
			' データが１件の場合はコピー不要
			If dataCount <= 1 Then
				Exit Sub
			End If

			' コピー元の行をコピー
			MyBase.sheet.Range(rangeF1, rangeF2).EntireRow.Select()
			MyBase.sheet.App.Selection.Copy()

			' コピー先の行を選択
			MyBase.sheet.Range(rangeT1, rangeT2).EntireRow.Select()

			' 出力予定の行数をインサート
			MyBase.sheet.App.Selection.Insert()
		End Sub

		''' <summary>
		''' リスト部出力
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents()
			Dim range1 As RangeWrapper
			Dim range2 As RangeWrapper

			range1 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			range2 = MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount, _cType.StartCol + _cType.ColumnLength - 1)
			MyBase.sheet.Range(range1, range2).Value = _cType.MakeArrayData()
		End Sub

	End Class

End Namespace
