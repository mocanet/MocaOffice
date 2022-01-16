
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Public Class VPageBreaksWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Sheet のラッパー</summary>
		Private _sheet As SheetWrapper

		''' <summary>Excel.VPageBreaks</summary>
		Private _vPageBreaks As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="sheet">親のSheet</param>
		''' <param name="vPageBreaks">Excel.VPageBreaks</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal sheet As SheetWrapper, ByVal vPageBreaks As Object)
			MyBase.New(sheet.ApplicationWrapper)
			_sheet = sheet
			_vPageBreaks = vPageBreaks
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_vPageBreaks)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _vPageBreaks
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 親シート
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
		''' コレクション内のオブジェクトの数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_vPageBreaks, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' 指定したオブジェクトの作成元のアプリケーションを示す 32 ビットの整数値を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_vPageBreaks, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' 指定した調整値を設定します。
		''' </summary>
		''' <param name="Index">必ず指定します。整数型 (Integer) の値を指定します。調整のインデックス番号を指定します。</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public Property Item( _
		  <InAttribute()> ByVal Index As Integer _
		 ) As VPageBreakWrapper
			Get
				Dim obj As Object
				Dim wrapper As VPageBreakWrapper
				obj = InvokeGetProperty(_vPageBreaks, "Item", New Object() {Index})
				wrapper = New VPageBreakWrapper(Me, obj)
				addXlsObject(wrapper)
				Return wrapper
			End Get
			Set(ByVal value As VPageBreakWrapper)
				InvokeSetProperty(_vPageBreaks, "Item", New Object() {value})
			End Set
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' 垂直な改ページを追加します。VPageBreak オブジェクトを取得します。
		''' </summary>
		''' <param name="Before">必ず指定します。オブジェクト型 (Object) の値を指定します。Range オブジェクトを指定します。新しい改ページを追加する右側の範囲を指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add( _
		  <InAttribute()> ByVal Before As RangeWrapper _
		 ) As VPageBreakWrapper
			Dim obj As Object
			Dim wrapper As VPageBreakWrapper
			obj = InvokeMethod(_vPageBreaks, "Add", New Object() {Before.OrigianlInstance})
			wrapper = New VPageBreakWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' Excel.Sheets.GetEnumeratorを返す
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator(Of VPageBreakWrapper)
			Dim enume As IEnumerator
			Dim result As IList(Of VPageBreakWrapper)

			result = New List(Of VPageBreakWrapper)

			enume = DirectCast(InvokeMethod(_vPageBreaks, "GetEnumerator", Nothing), IEnumerator)
			While enume.MoveNext()
				Dim wrapper As VPageBreakWrapper
				wrapper = New VPageBreakWrapper(Me, enume.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

#End Region

	End Class

End Namespace
