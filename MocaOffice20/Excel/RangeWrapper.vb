
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' セル、行、列、1 つ以上のセル範囲を含む選択範囲、または 3-D 範囲を表します
	''' </summary>
	''' <remarks></remarks>
	Public Class RangeWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のシート</summary>
		Private _sheet As SheetWrapper

		''' <summary>Excel.Range</summary>
		Private _range As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="sheet">親のシート</param>
		''' <param name="range">Excel.Range</param>
		''' <remarks></remarks>
		Public Sub New(ByVal sheet As SheetWrapper, ByVal range As Object)
			MyBase.New(sheet.ApplicationWrapper)
			_sheet = sheet
			_range = range
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_range)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _range
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
		''' 親のシート
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Sheet() As SheetWrapper
			Get
				Return _sheet
			End Get
		End Property

		''' <summary>
		''' セル数
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' Excel.Range.Count
		''' </remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_range, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' 指定したセル範囲の最初の領域の先頭行の番号を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' Excel.Range.Count
		''' </remarks>
		Public ReadOnly Property Row() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_range, "Row", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' 指定したセル範囲の値
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Value() As Object
			' 2003 以降の時は下記も使用可能
			'''' <summary>
			'''' 指定したセル範囲の値
			'''' </summary>
			'''' <param name="RangeValueDataType">指定した Range オブジェクトのデータ型</param>
			'''' <value></value>
			'''' <returns></returns>
			'''' <remarks></remarks>
			'Public Property Value(Optional ByVal RangeValueDataType As XlRangeValueDataType = XlRangeValueDataType.xlRangeValueDefault) As Object
			'	Get
			'		Return InvokeGetProperty(_range, "Value", New Object() {RangeValueDataType})
			'	End Get
			'	Set(ByVal value As Object)
			'		InvokeSetProperty(_range, "Value", New Object() {RangeValueDataType, value})
			'	End Set
			'End Property
			Get
				Return InvokeGetProperty(_range, "Value", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "Value", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したセル範囲に含まれる 1 行または複数の行全体を表す Range オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property EntireRow() As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "EntireRow", Nothing)

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' セルが含まれる領域の終端のセルを表す Range オブジェクトを取得します。
		''' </summary>
		''' <param name="Direction"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' End + 方向キー (↑、↓、←、→ のいずれか) に相当します。
		''' </remarks>
		Public ReadOnly Property [End](ByVal Direction As XlDirection) As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "End", New Object() {Direction})

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) の全体を表します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Font() As FontWrapper
			Get
				Dim val As Object
				Dim wrap As FontWrapper

				val = InvokeGetProperty(_range, "Font", New Object() {})

				wrap = New FontWrapper(Me, val)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' ポイント単位のセル範囲の幅です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Width() As Object
			Get
				Return InvokeGetProperty(_range, "Width", Nothing)
			End Get
		End Property

		''' <summary>
		''' 指定したセル範囲内のすべての列の幅を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColumnWidth() As Object
			Get
				Return InvokeGetProperty(_range, "ColumnWidth", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "ColumnWidth", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したセル範囲の列を表す Range オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Columns() As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "Columns", Nothing)

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' スタイルまたはセル範囲 (条件付き書式の一部として定義された範囲を含む) の罫線を表す Borders コレクションを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property Borders() As BordersWrapper
			Get
				Dim range As Object
				Dim wrap As BordersWrapper

				range = InvokeGetProperty(_range, "Borders", Nothing)

				wrap = New BordersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' スタイルまたはセル範囲 (条件付き書式の一部として定義された範囲を含む) の罫線を表す Borders コレクションを取得します。
		''' </summary>
		''' <param name="index">罫線を識別する値</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property Borders(ByVal index As XlBordersIndex) As BordersWrapper
			Get
				Dim range As Object
				Dim wrap As BordersWrapper

				range = InvokeGetProperty(_range, "Borders", New Object() {index})

				wrap = New BordersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' オブジェクトの文字列内の文字の範囲を表す Characters オブジェクトを取得します。
		''' </summary>
		''' <param name="Start">略可能です。オブジェクト型 (Object) の値を指定します。取得する文字列範囲の最初の文字を指定します。この引数に 1 を指定するか、省略すると、先頭文字から始まる文字列範囲を取得します。</param>
		''' <param name="Length"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Characters( _
		  <InAttribute()> Optional ByVal Start As Object = Nothing, _
		  <InAttribute()> Optional ByVal Length As Object = Nothing _
		 ) As CharactersWrapper
			Get
				Dim range As Object
				Dim wrap As CharactersWrapper

				range = InvokeGetProperty(_range, "Characters", New Object() {Start, Length})

				wrap = New CharactersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' 指定したオブジェクトの水平方向の配置を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property HorizontalAlignment() As Object
			Get
				Return InvokeGetProperty(_range, "HorizontalAlignment", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "HorizontalAlignment", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したオブジェクトの内部を表す Interior オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Interior() As InteriorWrapper
			Get
				Dim range As Object
				Dim wrap As InteriorWrapper

				range = InvokeGetProperty(_range, "Interior", Nothing)

				wrap = New InteriorWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

#End Region

		''' <summary>
		''' レンジ指定したセルを選択状態にする
		''' </summary>
		''' <remarks></remarks>
		Public Sub [Select]()
			InvokeMethod(_range, "Select", Nothing)
		End Sub

		''' <summary>
		''' オブジェクトをコピーします
		''' </summary>
		''' <param name="destination">省略可能です。オブジェクト型 (Object) の値を指定します。コピー先のセル範囲を指定します。この引数を省略すると、クリップボードにコピーします。</param>
		''' <remarks></remarks>
		Public Sub Copy(Optional ByVal destination As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If destination IsNot Nothing Then
				argsV.Add(destination)
				argsN.Add("Destination")
			End If

			InvokeMethod(_range, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

        ''' <summary>
        ''' 指定された範囲に、空白のセルまたはセル範囲を挿入します。指定された範囲にあったセルはシフトされます。
        ''' </summary>
        ''' <param name="shift">セルをシフトする方向</param>
        ''' <param name="copyOrigin">省略可能です。オブジェクト型 (Object) の値を指定します。コピー元を指定します。</param>
        ''' <remarks></remarks>
        Public Sub Insert(Optional ByVal shift As XlInsertShiftDirection = XlInsertShiftDirection.none, Optional ByVal copyOrigin As Object = Nothing)
            Dim argsV As ArrayList
            Dim argsN As ArrayList

            argsV = New ArrayList()
            argsN = New ArrayList()

            If shift <> XlInsertShiftDirection.none Then
                argsV.Add(shift)
                argsN.Add("Shift")
            End If

            If copyOrigin IsNot Nothing Then
                argsV.Add(copyOrigin)
                argsN.Add("CopyOrigin")
            End If

            InvokeMethod(_range, "Insert", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
        End Sub

        ''' <summary>
        ''' レンジ指定したセルを削除する
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Delete()
			InvokeMethod(_range, "Delete", Nothing)
		End Sub

		''' <summary>
		''' 内容をクリアします
		''' </summary>
		''' <remarks></remarks>
		Public Sub ClearContents()
			InvokeMethod(_range, "ClearContents", Nothing)
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
			If SkipBlanks IsNot Nothing Then
				argsV.Add(SkipBlanks)
				argsN.Add("SkipBlanks")
			End If
			If Transpose IsNot Nothing Then
				argsV.Add(Transpose)
				argsN.Add("Transpose")
			End If

			InvokeMethod(_range, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

	End Class

End Namespace
