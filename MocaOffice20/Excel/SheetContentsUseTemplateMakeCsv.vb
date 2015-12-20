
Imports System.IO
Imports System.Text

Namespace Excel

	''' <summary>
	''' 一覧形式のテンプレートシートを使用したときに、
	''' データを一度CSVファイルへ出力し、CSVファイルを読込んでExcelへ貼り付ける手法
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContentsUseTemplateMakeCsv
		Inherits SheetContentsUseTemplate

		''' <summary>CSVファイル名</summary>
		Private _csvFilename As String

		''' <summary>CSV読み込み時の列フォーマット</summary>
		Private _openFormat As Array

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			MyBase.New(sheetContents, sheet)

			_openFormat = Nothing
			_csvFilename = String.Empty
		End Sub

#End Region

#Region " プロパティ "

		''' <summary>CSV出力する際のテンポラリファイル名</summary>
		Private ReadOnly Property CsvTempFilename() As String
			Get
				If _csvFilename.Length = 0 Then
					_csvFilename = Path.Combine(Path.GetTempPath, "~" & MyBase.contents.SaveSheetName & Format(Now(), "_yyyyMMdd_hhmmss") & ".txt")
				End If
				Return _csvFilename
			End Get
		End Property

		''' <summary>CSVを読込むときの列フォーマット</summary>
		Private ReadOnly Property CsvOpenFormat() As Array
			Get
				If _openFormat Is Nothing Then
					_openFormat = Array.CreateInstance(GetType(Object), _cType.ColumnLength)
					For ii As Integer = 0 To _openFormat.GetUpperBound(0)
						_openFormat.SetValue(New Integer() {ii + 1, XlColumnDataType.xlTextFormat}, ii)
					Next ii
				End If
				Return _openFormat
			End Get
		End Property

#End Region

		''' <summary>
		''' シートコンテンツを当クラスで使用するクラスへキャストする
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _cType() As ISheetContentsUseTemplateMakeCsv
			Return DirectCast(MyBase.contents, ISheetContentsUseTemplateMakeCsv)
		End Function

		''' <summary>
		''' CSVファイルをExcelにて開く時のフォーマットを指定する
		''' </summary>
		''' <remarks>
		''' CSVファイルをExcelにて開く時のフォーマットを指定する場合は指定してください。<br/>
		''' デフォルトでは「一般」のフォーマットにて読み込みます。
		''' </remarks>
		Private Sub _setCsvOpenFormat()
			Dim value As XlColumnDataType
			For ii As Integer = 0 To UBound(Me.CsvOpenFormat)
				value = _cType.SetCsvOpenFormat(ii)
				If value = XlColumnDataType.xlNone Then
					Continue For
				End If
				Me.CsvOpenFormat.SetValue(New Integer() {ii + 1, value}, ii)
			Next ii
		End Sub

		''' <summary>
		''' 出力内容をセルへ設定する
		''' </summary>
		''' <remarks>
		''' </remarks>
		Protected Overrides Sub writeContentsTemplate()
			Dim xlBookTmp As BookWrapper

			xlBookTmp = Nothing

			Try
				If _cType.DataCount <= 0 Then
					Exit Sub
				End If

				' 作業用ファイルのお掃除
				_clearCsvTempFile()
				' データをテキストファイルでテンポラリ出力
				_csvTempWrite()

				' 先頭行をデータ数分コピー
				rowCopy(_cType.DataCount _
				 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow + 1, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount - 1, _cType.StartCol))

				' Tempファイルの読込
				_setCsvOpenFormat()
				MyBase.sheet.App.Workbooks.OpenText(Filename:=CsvTempFilename, DataType:=XlTextParsingType.xlDelimited, TextQualifier:=XlTextQualifier.xlTextQualifierDoubleQuote, Comma:=True, FieldInfo:=CsvOpenFormat)
				xlBookTmp = MyBase.sheet.App.ActiveWorkbook

				' Tempファイル内容を帳票へコピー＆ペースト
				_writeContents(xlBookTmp)

			Finally
				' Tempファイルを閉じる
				If xlBookTmp IsNot Nothing Then
					xlBookTmp.Close(False)
				End If
				' 作業用ファイルのお掃除
				_clearCsvTempFile()
			End Try
		End Sub

		''' <summary>
		''' 作業用ファイルのお掃除
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _clearCsvTempFile()
			If Not File.Exists(CsvTempFilename) Then
				Exit Sub
			End If

			'既に存在している場合は削除する。
			File.Delete(CsvTempFilename)
		End Sub

		''' <summary>
		''' CSV形式でテンポラリファイル作成
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _csvTempWrite()
			If _cType.DataCount = 0 Then
				Exit Sub
			End If

			Using file As StreamWriter = New StreamWriter(CsvTempFilename, False, Encoding.GetEncoding("Shift_JIS"))
				Try
					'テンプファイルへ出力
					_cType.CsvWrite(file)
				Catch ex As Exception
					Throw New ExcelException(MyBase.sheet.App, ex)
				End Try
			End Using
		End Sub

		''' <summary>
		''' リスト部出力
		''' </summary>
		''' <param name="xlBookTmp">Excelブックインスタンス（CSV用）</param>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents(ByVal xlBookTmp As BookWrapper)
			Dim tmpSheet As SheetWrapper
			Dim selection As SelectionWrapper
			Dim rowIndex As Integer

			Dim range1 As RangeWrapper
			Dim range2 As RangeWrapper

			'元ネタコピー
			xlBookTmp.Activate()
			tmpSheet = xlBookTmp.Worksheets(1)
			tmpSheet.Select()
			rowIndex = tmpSheet.Range("A65536").End(XlDirection.xlUp).Row
			range1 = tmpSheet.Cells(1, 1)
			range2 = tmpSheet.Cells(rowIndex, _cType.ColumnLength)
			tmpSheet.Range(range1, range2).Select()
			selection = MyBase.sheet.App.Selection
			selection.Copy()

			'貼り付け
			MyBase.sheet.Book.Activate()
			MyBase.sheet.Select()
			range1 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			range2 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			MyBase.sheet.Range(range1, range2).Select()
			selection = MyBase.sheet.App.Selection
			selection.PasteSpecial(Paste:=XlPasteType.xlPasteValues, Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=True, Transpose:=False)
		End Sub

	End Class

End Namespace
