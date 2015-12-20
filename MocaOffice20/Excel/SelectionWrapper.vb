
Namespace Excel

	''' <summary>
	''' アクティブなウィンドウで選択されているオブジェクト
	''' </summary>
	''' <remarks></remarks>
	Public Class SelectionWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Application のラッパー</summary>
		Private _xls As ExcelWrapper
		''' <summary>Excel.Application Selection オブジェクト</summary>
		Private _selection As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="xls">Excel.Applicationラッパー</param>
		''' <param name="selection">Excel.Range など選択されているオブジェクト</param>
		''' <remarks></remarks>
		Public Sub New(ByVal xls As ExcelWrapper, ByVal selection As Object)
			MyBase.New(xls.ApplicationWrapper)
			_xls = xls
			_selection = selection
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_selection)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _selection
			End Get
		End Property

#End Region

		''' <summary>
		''' オブジェクトをコピーします
		''' </summary>
		''' <remarks></remarks>
		Public Sub Copy()
			If _selection Is Nothing Then
				Exit Sub
			End If
			InvokeMethod(_selection, "Copy", Nothing)
		End Sub

		''' <summary>
		''' 選択されているオブジェクトをインサートします
		''' </summary>
		''' <remarks></remarks>
		Public Sub Insert()
			If _selection Is Nothing Then
				Exit Sub
			End If
			InvokeMethod(_selection, "Insert", Nothing)
		End Sub

		''' <summary>
		''' クリップボードにある Range オブジェクトを、指定したセル範囲に貼り付けます。
		''' </summary>
		''' <param name="Paste">省略可能です。<see cref="XlPasteType" /> 列挙型の定数を指定します。セル範囲の中で貼り付ける部分を指定します。</param>
		''' <param name="Operation">省略可能です。<see cref="XlPasteSpecialOperation" /> 列挙型の値を指定します。貼り付けの操作を指定します。</param>
		''' <param name="SkipBlanks">省略可能です。オブジェクト型 (Object) の値を指定します。True を指定すると、クリップボードに含まれる空白のセルを対象セル範囲に貼り付けません。既定値は False です。</param>
		''' <param name="Transpose">省略可能です。オブジェクト型 (Object) の値を指定します。True を指定すると、貼り付けるときにセル範囲の行と列を入れ替えます。既定値は False です。</param>
		''' <remarks></remarks>
		Public Sub PasteSpecial( _
		 Optional ByVal Paste As XlPasteType = XlPasteType.xlPasteAll, _
		 Optional ByVal Operation As XlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
		 Optional ByVal SkipBlanks As Object = Nothing, _
		 Optional ByVal Transpose As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Paste)
			argsN.Add("Paste")

			argsV.Add(Operation)
			argsN.Add("Operation")

			If SkipBlanks IsNot Nothing Then
				argsV.Add(SkipBlanks)
				argsN.Add("SkipBlanks")
			End If
			If Transpose IsNot Nothing Then
				argsV.Add(Transpose)
				argsN.Add("Transpose")
			End If

			InvokeMethod(_selection, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		'''' <summary>
		'''' 指定した形式で、クリップボードの内容をシートに貼り付けます。他のアプリケーションからデータを貼り付けたり、特定の形式でデータを貼り付ける場合に使用します。
		'''' </summary>
		'''' <param name="Format">省略可能です。オブジェクト型 (Object) の値を指定します。クリップボードのデータの形式を文字列で指定します。</param>
		'''' <param name="Link">省略可能です。オブジェクト型 (Object) の値を指定します。元のデータと貼り付けたデータの間にリンクを設定するには、True を指定します。元のデータがリンクに適さないデータである場合や、元のデータを作成したアプリケーションがリンクをサポートしない場合、この引数は無視されます。既定値は False です。</param>
		'''' <param name="DisplayAsIcon">省略可能です。オブジェクト型 (Object) の値を指定します。貼り付けたデータをアイコンとして表示するには、True を指定します。既定値は False です。</param>
		'''' <param name="IconFileName">省略可能です。オブジェクト型 (Object) を指定します。DisplayAsIcon が True の場合に使用するアイコンを含むファイルの名前を指定します。</param>
		'''' <param name="IconIndex">省略可能です。オブジェクト型 (Object) の値を指定します。アイコンのファイル内のアイコンのインデックス番号を指定します。</param>
		'''' <param name="IconLabel">省略可能です。オブジェクト型 (Object) の値を指定します。アイコンのラベルの文字列を指定します。</param>
		'''' <param name="NoHTMLFormatting">省略可能です。オブジェクト型 (Object) の値を指定します。HTML から書式設定、ハイパーリンク、およびイメージをすべて削除するには、True を指定します。HTML をそのまま貼り付けるには、False を指定します。既定値は False です。</param>
		'''' <remarks>
		'''' このメソッドを使用する前に貼り付け先のセル範囲を選択する必要があります。<br/>
		'''' このメソッドを使用すると、クリップボードの内容によっては選択範囲が変更される場合があります。
		'''' </remarks>
		'Public Sub PasteSpecial( _
		' Optional ByVal Format As Object = Nothing, _
		' Optional ByVal Link As Object = Nothing, _
		' Optional ByVal DisplayAsIcon As Object = Nothing, _
		' Optional ByVal IconFileName As Object = Nothing, _
		' Optional ByVal IconIndex As Object = Nothing, _
		' Optional ByVal IconLabel As Object = Nothing, _
		' Optional ByVal NoHTMLFormatting As Object = Nothing)
		'	Dim argsV As ArrayList
		'	Dim argsN As ArrayList

		'	argsV = New ArrayList()
		'	argsN = New ArrayList()

		'	If Format Is Nothing Then
		'		argsV.Add(Format)
		'		argsN.Add("Format")
		'	End If
		'	If Link Is Nothing Then
		'		argsV.Add(Link)
		'		argsN.Add("Link")
		'	End If
		'	If DisplayAsIcon Is Nothing Then
		'		argsV.Add(DisplayAsIcon)
		'		argsN.Add("DisplayAsIcon")
		'	End If
		'	If IconFileName Is Nothing Then
		'		argsV.Add(IconFileName)
		'		argsN.Add("IconFileName")
		'	End If
		'	If IconIndex Is Nothing Then
		'		argsV.Add(IconIndex)
		'		argsN.Add("IconIndex")
		'	End If
		'	If IconLabel Is Nothing Then
		'		argsV.Add(IconLabel)
		'		argsN.Add("IconLabel")
		'	End If
		'	If NoHTMLFormatting Is Nothing Then
		'		argsV.Add(NoHTMLFormatting)
		'		argsN.Add("NoHTMLFormatting")
		'	End If

		'	InvokeMethod(_selection, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		'End Sub

	End Class

End Namespace
