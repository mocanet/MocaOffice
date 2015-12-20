
Namespace Excel

	''' <summary>
	''' オブジェクトの罫線を表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class BorderWrapper
		Inherits AbstractExcelWrapper

		''' <summary>BordersWrapper</summary>
		Private _borders As BordersWrapper

		''' <summary>Excel.Border</summary>
		Private _border As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="borders">親のBorders</param>
		''' <param name="border">Excel.Border</param>
		''' <remarks></remarks>
		Public Sub New(ByVal borders As BordersWrapper, ByVal border As Object)
			MyBase.New(borders.ApplicationWrapper)
			_borders = borders
			_border = border
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_border)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _border
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
				Return InvokeGetProperty(_border, "Color", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_border, "Color", New Object() {value})
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
				Return InvokeGetProperty(_border, "ColorIndex", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_border, "ColorIndex", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したオブジェクトの作成元のアプリケーションを示す 32 ビットの整数値を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_border, "Creator", Nothing), XlCreator)
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
				Return CType(InvokeGetProperty(_border, "LineStyle", Nothing), XlLineStyle)
			End Get
			Set(ByVal value As XlLineStyle)
				InvokeSetProperty(_border, "LineStyle", New Object() {value})
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
				Return CType(InvokeGetProperty(_border, "Weight", Nothing), XlBorderWeight)
			End Get
			Set(ByVal value As XlBorderWeight)
				InvokeSetProperty(_border, "LineStyle", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
