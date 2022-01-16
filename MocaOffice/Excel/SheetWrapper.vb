
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Worksheet のラッパークラス
	''' </summary>
	''' <remarks></remarks>
	Public Class SheetWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Workbook のラッパー</summary>
		Private _book As BookWrapper

		''' <summary>Excel.Worksheet</summary>
		Private _sheet As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="book">親のブック</param>
		''' <remarks>
		''' シートを新規で追加するときに使う
		''' </remarks>
		Private Sub New(ByVal book As BookWrapper)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = InvokeGetProperty(_book.Worksheets.OrigianlInstance, "Add", Nothing)
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="book">親のブック</param>
		''' <param name="sheetName">開くシート名</param>
		''' <remarks>
		''' シート名を指定して開くときに使う。
		''' </remarks>
		Public Sub New(ByVal book As BookWrapper, ByVal sheetName As String)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = InvokeGetProperty(_book.Worksheets.OrigianlInstance, "Item", New Object() {sheetName})
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="book">親のブック</param>
		''' <param name="xlSheet">Excel.Worksheet</param>
		''' <remarks>
		''' 既に開いたシートを管理するときに使う。
		''' </remarks>
		Public Sub New(ByVal book As BookWrapper, ByVal xlSheet As Object)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = xlSheet
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_sheet)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _sheet
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
		''' シート名
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Name() As String
			Get
				Return DirectCast(InvokeGetProperty(_sheet, "Name", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_sheet, "Name", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 単一のセルまたはセル範囲を表す Range オブジェクトを取得します。
		''' </summary>
		''' <param name="Cell1"></param>
		''' <param name="Cell2"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Range(ByVal Cell1 As RangeWrapper, Optional ByVal Cell2 As RangeWrapper = Nothing) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				If Cell2 Is Nothing Then
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1.OrigianlInstance})
				Else
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1.OrigianlInstance, Cell2.OrigianlInstance})
				End If
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' 単一のセルまたはセル範囲を表す Range オブジェクトを取得します。
		''' </summary>
		''' <param name="Cell1"></param>
		''' <param name="Cell2"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Range(ByVal Cell1 As String, Optional ByVal Cell2 As String = Nothing) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				If Cell2 Is Nothing Then
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1})
				Else
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1, Cell2})
				End If
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' アクティブなワークシートにあるすべてのセルを表す Range オブジェクトを取得します。アクティブな文書がワークシートでない場合、このプロパティは失敗します。
		''' </summary>
		''' <param name="row"></param>
		''' <param name="col"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Cells(ByVal row As Integer, ByVal col As Integer) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				rangeBuf = InvokeGetProperty(_sheet, "Cells", New Object() {row, col})
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' シートが削除されているか
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property IsDeleted() As Boolean
			Get
				Return (_sheet Is Nothing)
			End Get
		End Property

		''' <summary>
		''' ワークシート上のすべての列を表す Range オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Columns() As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				rangeBuf = InvokeGetProperty(_sheet, "Columns", New Object() {})
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' ワークシートのすべてのページ設定を含む <see cref="PageSetupWrapper"/>  を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overridable ReadOnly Property PageSetup() As PageSetupWrapper
			Get
				Dim obj As Object
				Dim wrap As PageSetupWrapper
				obj = InvokeGetProperty(_sheet, "PageSetup", New Object() {})
				wrap = New PageSetupWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ワークシート上のすべての図形を表す <see cref="ShapesWrapper"/> オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overridable ReadOnly Property Shapes() As ShapesWrapper
			Get
				Dim obj As Object
				Dim wrap As ShapesWrapper
				obj = InvokeGetProperty(_sheet, "Shapes", New Object() {})
				wrap = New ShapesWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ワークシート上の水平方向の改ページを表す <see cref="HPageBreaksWrapper"/> コレクションを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property HPageBreaks() As HPageBreaksWrapper
			Get
				Dim obj As Object
				Dim wrap As HPageBreaksWrapper
				obj = InvokeGetProperty(_sheet, "HPageBreaks", New Object() {})
				wrap = New HPageBreaksWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ワークシート上の垂直方向の改ページを表す <see cref="VPageBreaksWrapper"/> コレクションを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property VPageBreaks() As VPageBreaksWrapper
			Get
				Dim obj As Object
				Dim wrap As VPageBreaksWrapper
				obj = InvokeGetProperty(_sheet, "VPageBreaks", New Object() {})
				wrap = New VPageBreaksWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
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
		End Sub

		''' <summary>
		''' シートをアクティブにする
		''' </summary>
		''' <remarks></remarks>
		Public Sub Activate()
			InvokeMethod(_sheet, "Activate", Nothing)
		End Sub

		''' <summary>
		''' シート全体を選択にする
		''' </summary>
		''' <remarks></remarks>
		Public Sub [Select]()
			InvokeMethod(_sheet, "Select", Nothing)
		End Sub

		''' <summary>
		''' シートを削除する
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_sheet, "Delete", Nothing)
			MyDispose()
		End Sub

		''' <summary>
		''' クリップボードから貼付け
		''' </summary>
		''' <remarks></remarks>
		Public Sub Paste()
			InvokeMethod(_sheet, "Paste", Nothing)
		End Sub

		''' <summary>
		''' シートをブック内の他の場所にコピーします。
		''' </summary>
		''' <param name="Before">省略可能です。オブジェクト型 (Object) の値を指定します。コピーするシートを特定のシートの直前の位置に挿入するときに、そのシートを指定します。After が指定されている場合、Before は指定できません。</param>
		''' <param name="After">省略可能です。オブジェクト型 (Object) の値を指定します。コピーするシートを特定のシートの直後の位置に挿入するときに、そのシートを指定します。Before が指定されている場合、After は指定できません。</param>
		''' <remarks>引数 Before と引数 After を共に省略した場合は、新規ブックが自動的に作成され、シートはそのブック内に挿入されます。</remarks>
		Public Sub Copy(<InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, <InAttribute()> Optional ByVal After As SheetWrapper = Nothing)
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

			InvokeMethod(_sheet, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' 指定した形式のファイルにエクスポート
		''' </summary>
		''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff198122.aspx</remarks>
		Public Sub ExportAsFixedFormat(ByVal xlType As FixedFormatType,
									   Optional ByVal filename As String = Nothing,
									   Optional ByVal quality As FixedFormatQuality = FixedFormatQuality.QualityStandard,
									   Optional ByVal includeDocProperties As Boolean = False,
									   Optional ByVal ignorePrintAreas As Boolean = False,
									   Optional ByVal [from] As Integer = 0,
									   Optional ByVal [to] As Integer = 0,
									   Optional ByVal openAfterPublish As Boolean = False,
									   Optional ByVal fixedFormatExtClassPtr As Object = Nothing)
			Dim argsV As New List(Of Object)
			Dim argsN As New List(Of String)

			argsV.Add(xlType)
			argsN.Add("Type")
			If filename IsNot Nothing Then
				argsV.Add(filename)
				argsN.Add("Filename")
			End If
			argsV.Add(quality)
			argsN.Add("Quality")
			argsV.Add(includeDocProperties)
			argsN.Add("IncludeDocProperties")
			argsV.Add(ignorePrintAreas)
			argsN.Add("IgnorePrintAreas")
			If [from] > 0 Then
				argsV.Add([from])
				argsN.Add("From")
			End If
			If [to] > 0 Then
				argsV.Add([to])
				argsN.Add("To")
			End If
			argsV.Add(openAfterPublish)
			argsN.Add("OpenAfterPublish")
			If fixedFormatExtClassPtr IsNot Nothing Then
				argsV.Add(fixedFormatExtClassPtr)
				argsN.Add("FixedFormatExtClassPtr")
			End If

			'args = New Object() {xlType, filename, quality, includeDocProperties, ignorePrintAreas, [from], [to], openAfterPublish, fixedFormatExtClassPtr}

			InvokeMethod(_book, "ExportAsFixedFormat", argsV.ToArray, argsN.ToArray)
		End Sub

	End Class

End Namespace
