
Namespace Excel

	''' <summary>
	''' Excel.Worksheets のラッパークラス
	''' </summary>
	''' <remarks></remarks>
	Public Class WorksheetsWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Workbook のラッパー</summary>
		Private _book As BookWrapper

		''' <summary>Excel.Worksheets</summary>
		Private _sheets As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="book">親ブック</param>
		''' <remarks></remarks>
		Public Sub New(ByVal book As BookWrapper)
			MyBase.New(book.ApplicationWrapper)
			_init(book)
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_sheets)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _sheets
			End Get
		End Property

#End Region

#Region " プロパティ "

		''' <summary>
		''' Excel.Application のラッパー
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property App() As ExcelWrapper
			Get
				Return DirectCast(MyBase.xlsWrapper, ExcelWrapper)
			End Get
		End Property

		''' <summary>
		''' 親のブック
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Book() As BookWrapper
			Get
				Return _book
			End Get
		End Property

		''' <summary>
		''' シート数
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_sheets, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' 作業中のブックのすべてのシートを表す Sheets コレクションから指定されたシートを取得します。 
		''' </summary>
		''' <param name="name">シート名</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(ByVal name As String) As SheetWrapper
			Get
				Dim sheet As Object
				Dim wrapper As SheetWrapper
				sheet = InvokeGetProperty(_sheets, "Item", New Object() {name})
				wrapper = New SheetWrapper(_book, sheet)
				addXlsObject(wrapper)
				Return wrapper
			End Get
		End Property

		''' <summary>
		''' 作業中のブックのすべてのシートを表す Sheets コレクションから指定されたシートを取得します。 
		''' </summary>
		''' <param name="index">シート番号</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(ByVal index As Integer) As SheetWrapper
			Get
				Dim sheet As Object
				Dim wrapper As SheetWrapper
				sheet = InvokeGetProperty(_sheets, "Item", New Object() {index})
				wrapper = New SheetWrapper(_book, sheet)
				addXlsObject(wrapper)
				Return wrapper
			End Get
		End Property

#End Region

		''' <summary>
		''' 初期化
		''' </summary>
		''' <param name="book">親のブック</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal book As BookWrapper)
			_book = book
			' Worksheetsオブジェクトの作成
			_sheets = InvokeGetProperty(_book.OrigianlInstance, "Worksheets", Nothing)
		End Sub

		''' <summary>
		''' 新しいワークシートを作成します。作成したワークシートはアクティブになります。 
		''' </summary>
		''' <param name="sheetName">シート名</param>
		''' <param name="Before">省略可能です。オブジェクト型 (Object) の値を指定します。新しいシートを特定のシートの直前の位置に追加するときに、そのシートを指定します。</param>
		''' <param name="After">省略可能です。オブジェクト型 (Object) の値を指定します。新しいシートを特定のシートの直後の位置に追加するときに、そのシートを指定します。</param>
		''' <param name="Count">省略可能です。オブジェクト型 (Object) の値を指定します。追加するシートの数を指定します。既定値は 1 です。</param>
		''' <param name="Type">省略可能です。オブジェクト型 (Object) の値を指定します。シートの種類を指定します。使用できる定数は、XlSheetType 列挙型の xlWorksheet、xlChart、xlExcel4MacroSheet、xlExcel4IntlMacroSheet のいずれかです。既存のテンプレートを基にしたシートを挿入する場合は、そのテンプレートへのパスを指定します。既定値は xlWorksheet です。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add(ByVal sheetName As String _
		, Optional ByVal Before As SheetWrapper = Nothing _
		, Optional ByVal After As SheetWrapper = Nothing _
		, Optional ByVal Count As Integer = 1 _
		, Optional ByVal Type As Object = Nothing) As SheetWrapper
			Dim xls As Object
			Dim sheet As SheetWrapper

			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If After IsNot Nothing Then
				argsV.Add(After)
				argsN.Add("After")
			End If
			If Before IsNot Nothing Then
				argsV.Add(Before)
				argsN.Add("Before")
			End If
			If argsV.Count = 0 Then
				argsV.Add(Me.App.ActiveSheet.OrigianlInstance)
				argsN.Add("After")
			End If

			argsV.Add(Count)
			argsN.Add("Count")
			If Type IsNot Nothing Then
				argsV.Add(Type)
				argsN.Add("Type")
			End If

			xls = InvokeMethod(_sheets, "Add", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))

			sheet = New SheetWrapper(_book, xls)
			sheet.Name = sheetName
			addXlsObject(sheet)
			Return sheet
		End Function

		''' <summary>
		''' Excel.Worksheets.GetEnumeratorを返す
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator(Of SheetWrapper)
			Dim sheetEnum As IEnumerator
			Dim result As IList(Of SheetWrapper)

			result = New List(Of SheetWrapper)

			sheetEnum = DirectCast(InvokeMethod(_sheets, "GetEnumerator", Nothing), IEnumerator)
			While sheetEnum.MoveNext()
				Dim wrapper As SheetWrapper
				wrapper = New SheetWrapper(_book, sheetEnum.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

		''' <summary>
		''' デフォルトシート削除
		''' </summary>
		''' <remarks>
		''' Excelを新規で作成するときに出来るデフォルトのシート（Sheet1...）を削除します。
		''' </remarks>
		Public Sub ClearDefaultSheet()
			Dim sheetEnum As IEnumerator(Of SheetWrapper)
			Dim xlSheet As SheetWrapper = Nothing

			sheetEnum = GetEnumerator()
			While sheetEnum.MoveNext()
				xlSheet = sheetEnum.Current()

				If xlSheet.Name.StartsWith("Sheet") Then
					xlSheet.Delete()
				End If
			End While
		End Sub

		''' <summary>
		''' 同一のシート名が存在する場合の件数を算出する
		''' </summary>
		''' <param name="sheetName">シート名</param>
		''' <returns></returns>
		''' <remarks>
		''' 「指定された名称＋”＿”＋番号など」の名称シートが存在した数を返す
		''' </remarks>
		Public Function MultiSheetCount(ByVal sheetName As String) As Integer
			Return MultiSheetCount(sheetName, "_")
		End Function

		''' <summary>
		''' 同一のシート名が存在する場合の件数を算出する
		''' </summary>
		''' <param name="sheetName">シート名</param>
		''' <param name="delim">区切り文字</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function MultiSheetCount(ByVal sheetName As String, ByVal delim As String) As Integer
			Dim sheetEnum As IEnumerator(Of SheetWrapper)
			Dim xlSheet As SheetWrapper = Nothing
			Dim sheetCount As Integer

			sheetCount = 0
			sheetEnum = GetEnumerator()
			While sheetEnum.MoveNext()
				xlSheet = sheetEnum.Current()
				Try
					Dim aryName() As String

					aryName = xlSheet.Name.Split(CChar(delim))
					If aryName(0) = sheetName Then
						sheetCount += 1
					End If
				Finally
					xlSheet.Dispose()
				End Try
			End While

			Return sheetCount
		End Function

	End Class

End Namespace
