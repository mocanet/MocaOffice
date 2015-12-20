
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Sheets のラッパークラス
	''' </summary>
	''' <remarks></remarks>
	Public Class SheetsWrapper
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

		''' <summary>
		''' オブジェクトを表示するか、非表示にするかを決定します。値の取得および設定が可能です。オブジェクト型 (Object) の値を使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Visible() As Object
			Get
				Return InvokeGetProperty(_sheets, "Visible", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_sheets, "Visible", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' シートの垂直方向の改ページを表す VPageBreaks コレクションを取得します。値の取得のみ可能です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property VPageBreaks() As VPageBreaksWrapper
			Get
				Return DirectCast(InvokeGetProperty(_sheets, "VPageBreaks", Nothing), VPageBreaksWrapper)
			End Get
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' 初期化
		''' </summary>
		''' <param name="book">親のブック</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal book As BookWrapper)
			_book = book
			' Worksheetsオブジェクトの作成
			_sheets = InvokeGetProperty(_book.OrigianlInstance, "Sheets", Nothing)
		End Sub

		''' <summary>
		''' 新しいワークシート、グラフ シート、マクロ シートのいずれかを作成します。作成したワークシートはアクティブになります。
		''' </summary>
		''' <param name="Before">省略可能です。オブジェクト型 (Object) の値を指定します。新しいシートを特定のシートの直前の位置に追加するときに、そのシートを指定します。</param>
		''' <param name="After">省略可能です。オブジェクト型 (Object) の値を指定します。新しいシートを特定のシートの直後の位置に追加するときに、そのシートを指定します。</param>
		''' <param name="Count">省略可能です。オブジェクト型 (Object) の値を指定します。追加するシートの数を指定します。既定値は 1 です。</param>
		''' <param name="Type">省略可能です。オブジェクト型 (Object) の値を指定します。シートの種類を指定します。使用できる定数は、XlSheetType 列挙型の xlWorksheet、xlChart、xlExcel4MacroSheet、xlExcel4IntlMacroSheet のいずれかです。既存のテンプレートを基にしたシートを挿入する場合は、そのテンプレートへのパスを指定します。既定値は xlWorksheet です。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add( _
		  <InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal After As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal Count As Object = 1, _
		  <InAttribute()> Optional ByVal Type As XlSheetType = XlSheetType.xlWorksheet _
		 ) As SheetWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Before IsNot Nothing Then
				argsV.Add(Before.OrigianlInstance)
				argsN.Add("Before")
			End If
			If After IsNot Nothing Then
				argsV.Add(After.OrigianlInstance)
				argsN.Add("After")
			End If
			If Count IsNot Nothing Then
				argsV.Add(Count)
				argsN.Add("Count")
			End If
			argsV.Add(Type)
			argsN.Add("Type")

			Dim obj As Object
			Dim wrapper As SheetWrapper
			obj = InvokeMethod(_sheets, "Add", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			wrapper = New SheetWrapper(_book, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' シートをブック内の他の場所にコピーします。
		''' </summary>
		''' <param name="Before">省略可能です。オブジェクト型 (Object) の値を指定します。コピーするシートを特定のシートの直前の位置に挿入するときに、そのシートを指定します。After が指定されている場合、Before は指定できません。</param>
		''' <param name="After">省略可能です。オブジェクト型 (Object) の値を指定します。コピーするシートを特定のシートの直後の位置に挿入するときに、そのシートを指定します。Before が指定されている場合、After は指定できません。</param>
		''' <remarks>
		''' 引数 Before と引数 After を共に省略した場合は、新規ブックが自動的に作成され、シートはそのブック内に挿入されます。
		''' </remarks>
		Public Sub Copy( _
		  <InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal After As SheetWrapper = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Before IsNot Nothing Then
				argsV.Add(Before.OrigianlInstance)
				argsN.Add("Before")
			End If
			If After IsNot Nothing Then
				argsV.Add(After.OrigianlInstance)
				argsN.Add("After")
			End If

			InvokeMethod(_sheets, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' オブジェクトを削除します。
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_sheets, "Delete", Nothing)
		End Sub

		''' <summary>
		''' 指定されたセル範囲を、コレクション内の他のすべてのワークシートの同じ領域にコピーします。
		''' </summary>
		''' <param name="Range">必ず指定します。Range オブジェクトを指定します。コレクションに属するすべてのワークシートのフィルに使用するセル範囲を指定します。このセル範囲には、コレクション内のワークシートを指定する必要があります。</param>
		''' <param name="Type">省略可能です。XlFillWith の値を指定します。指定したセル範囲をコピーする方法を指定します。</param>
		''' <remarks></remarks>
		Public Sub FillAcrossSheets( _
		  <InAttribute()> ByVal Range As RangeWrapper, _
		  <InAttribute()> Optional ByVal Type As XlFillWith = XlFillWith.xlFillWithAll _
		 )
			InvokeMethod(_sheets, "FillAcrossSheets", New Object() {Range.OrigianlInstance, Type})
		End Sub

		''' <summary>
		''' Excel.Sheets.GetEnumeratorを返す
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
		''' シートをブック内の他の場所に移動します。
		''' </summary>
		''' <param name="Before">省略可能です。オブジェクト型 (Object) の値を指定します。移動するシートを特定のシートの直前の位置に挿入するときに、そのシートを指定します。After が指定されている場合、Before は指定できません。</param>
		''' <param name="After">省略可能です。オブジェクト型 (Object) の値を指定します。移動するシートを特定のシートの直後の位置に挿入するときに、そのシートを指定します。Before が指定されている場合、After は指定できません。</param>
		''' <remarks></remarks>
		Public Sub Move( _
		<InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		<InAttribute()> Optional ByVal After As SheetWrapper = Nothing _
		  )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Before IsNot Nothing Then
				argsV.Add(Before.OrigianlInstance)
				argsN.Add("Before")
			End If
			If After IsNot Nothing Then
				argsV.Add(After.OrigianlInstance)
				argsN.Add("After")
			End If

			InvokeMethod(_sheets, "Move", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' オブジェクトを印刷します。
		''' </summary>
		''' <param name="From">省略可能です。オブジェクト型 (Object) の値を指定します。印刷を開始するページ番号を指定します。この引数を省略すると、最初のページから印刷されます。</param>
		''' <param name="To">省略可能です。オブジェクト型 (Object) の値を指定します。印刷を終了するページ番号を指定します。この引数を省略すると、印刷は最後のページで終了します。</param>
		''' <param name="Copies">省略可能です。オブジェクト型 (Object) の値を指定します。印刷する部数を指定します。この引数を省略すると、1 部が印刷されます。</param>
		''' <param name="Preview">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、オブジェクトを印刷する前に印刷プレビューが実行されます。False を設定するか、または引数を省略すると、オブジェクトは直ちに印刷されます。</param>
		''' <param name="ActivePrinter">省略可能です。オブジェクト型 (Object) の値を指定します。現在使用しているプリンタの名前を設定します。</param>
		''' <param name="PrintToFile">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、ファイルを印刷します。引数 PrToFileName を指定しないと、出力ファイル名の入力を促すダイアログ ボックスが表示されます。</param>
		''' <param name="Collate">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、複数部単位で印刷されます。</param>
		''' <param name="PrToFileName">省略可能です。オブジェクト型 (Object) の値を指定します。引数 PrintToFile に True を設定した場合、印刷するファイルの名前をこの引数に指定します。</param>
		''' <remarks></remarks>
		Public Sub PrintOut( _
		  <InAttribute()> Optional ByVal From As Object = Nothing, _
		  <InAttribute()> Optional ByVal [To] As Object = Nothing, _
		  <InAttribute()> Optional ByVal Copies As Object = Nothing, _
		  <InAttribute()> Optional ByVal Preview As Object = Nothing, _
		  <InAttribute()> Optional ByVal ActivePrinter As Object = Nothing, _
		  <InAttribute()> Optional ByVal PrintToFile As Object = Nothing, _
		  <InAttribute()> Optional ByVal Collate As Object = Nothing, _
		  <InAttribute()> Optional ByVal PrToFileName As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If From IsNot Nothing Then
				argsV.Add(From)
				argsN.Add("From")
			End If
			If [To] IsNot Nothing Then
				argsV.Add([To])
				argsN.Add("To")
			End If
			If Copies IsNot Nothing Then
				argsV.Add(Copies)
				argsN.Add("Copies")
			End If
			If Preview IsNot Nothing Then
				argsV.Add(Preview)
				argsN.Add("Preview")
			End If
			If ActivePrinter IsNot Nothing Then
				argsV.Add(ActivePrinter)
				argsN.Add("ActivePrinter")
			End If
			If PrintToFile IsNot Nothing Then
				argsV.Add(PrintToFile)
				argsN.Add("PrintToFile")
			End If
			If Collate IsNot Nothing Then
				argsV.Add(Collate)
				argsN.Add("Collate")
			End If
			If PrToFileName IsNot Nothing Then
				argsV.Add(PrToFileName)
				argsN.Add("PrToFileName")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' オブジェクトの印刷プレビュー (印刷時のイメージ) を表示します。
		''' </summary>
		''' <param name="EnableChanges">オブジェクトの変更を可能にします。</param>
		''' <remarks></remarks>
		Public Sub PrintPreview( _
		  <InAttribute()> Optional ByVal EnableChanges As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If EnableChanges IsNot Nothing Then
				argsV.Add(EnableChanges)
				argsN.Add("EnableChanges")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' オブジェクトを選択します。
		''' </summary>
		''' <param name="Replace">省略可能です。オブジェクト型 (Object) の値を指定します。指定のオブジェクトを選択する際に、既に選択しているオブジェクトの選択を解除するかどうかを指定します。</param>
		''' <remarks></remarks>
		Public Sub [Select]( _
		  <InAttribute()> Optional ByVal Replace As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Replace IsNot Nothing Then
				argsV.Add(Replace)
				argsN.Add("Replace")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

#End Region

	End Class

End Namespace
