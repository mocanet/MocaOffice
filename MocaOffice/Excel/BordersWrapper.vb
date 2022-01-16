
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Range オブジェクトまたは Style オブジェクトの四辺の罫線を表す 4 つの Border オブジェクトのコレクションです。
	''' </summary>
	''' <remarks></remarks>
	Public Class BordersWrapper
		Inherits AbstractExcelWrapper

		''' <summary>RangeWrapper</summary>
		Private _range As RangeWrapper

		''' <summary>Excel.Borders</summary>
		Private _borders As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="range">親のrange</param>
		''' <param name="borders">Excel.Borders</param>
		''' <remarks></remarks>
		Public Sub New(ByVal range As RangeWrapper, ByVal borders As Object)
			MyBase.New(range.ApplicationWrapper)
			_range = range
			_borders = borders
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_borders)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _borders
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 指定したセル範囲の四辺の罫線の基本色を設定します。線のすべての色が同じでない場合は、0 (ゼロ) を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Color() As Object
			Get
				Return InvokeGetProperty(_borders, "Color", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_borders, "Color", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 四辺の罫線の色を設定します。四辺のすべての罫線が同じ色でない場合は、Null 値を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColorIndex() As Object
			Get
				Return InvokeGetProperty(_borders, "ColorIndex", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_borders, "ColorIndex", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' コレクション内のオブジェクトの数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_borders, "Count", Nothing), Integer)
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
				Return DirectCast(InvokeGetProperty(_borders, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' セル範囲またはスタイルの罫線のいずれかを表す Border オブジェクトを取得します。
		''' </summary>
		''' <param name="Index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(<InAttribute()> ByVal Index As XlBordersIndex) As BorderWrapper
			Get
				Dim obj As Object
				Dim wrapper As BorderWrapper
				obj = InvokeGetProperty(_borders, "Item", New Object() {Index})
				wrapper = New BorderWrapper(Me, obj)
				addXlsObject(wrapper)
				Return wrapper
			End Get
		End Property

		''' <summary>
		''' 罫線の線のスタイルを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property LineStyle() As XlLineStyle
			Get
				Return CType(InvokeGetProperty(_borders, "LineStyle", Nothing), XlLineStyle)
			End Get
			Set(ByVal value As XlLineStyle)
				InvokeSetProperty(_borders, "LineStyle", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 罫線の線のスタイルを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Value() As Object
			Get
				Return CType(InvokeGetProperty(_borders, "Value", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_borders, "Value", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 罫線の太さを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Weight() As XlBorderWeight
			Get
				Return CType(InvokeGetProperty(_borders, "Weight", Nothing), XlBorderWeight)
			End Get
			Set(ByVal value As XlBorderWeight)
				InvokeSetProperty(_borders, "LineStyle", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
