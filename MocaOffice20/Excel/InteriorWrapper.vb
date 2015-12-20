
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Public Class InteriorWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のシート</summary>
		Private _sheet As RangeWrapper

		''' <summary>Excel.Interior</summary>
		Private _interior As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="range">親のRange</param>
		''' <param name="interior">Excel.Interior</param>
		''' <remarks></remarks>
		Public Sub New(ByVal range As RangeWrapper, ByVal interior As Object)
			MyBase.New(range.ApplicationWrapper)
			_sheet = range
			_interior = interior
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_interior)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _interior
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' セルの網かけや描画オブジェクトの塗りつぶしの基本色を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Color() As Object
			Get
				Return InvokeGetProperty(_interior, "Color", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_interior, "Color", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 内部の色を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColorIndex() As Object
			Get
				Return InvokeGetProperty(_interior, "ColorIndex", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_interior, "ColorIndex", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、項目が負の数の場合にその項目のパターンが反転されます。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property InvertIfNegative() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_interior, "InvertIfNegative", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_interior, "InvertIfNegative", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 内部のパターンを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Pattern() As XlPattern
			Get
				Return DirectCast(InvokeGetProperty(_interior, "Pattern", Nothing), XlPattern)
			End Get
			Set(ByVal value As XlPattern)
				InvokeSetProperty(_interior, "Pattern", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 内部パターンの色を RGB 値で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PatternColor() As Object
			Get
				Return DirectCast(InvokeGetProperty(_interior, "PatternColor", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_interior, "PatternColor", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 内部パターンの色を、現在のカラー パレットのインデックスとして設定するか、XlColorIndex 列挙型の定数 xlColorIndexAutomatic または xlColorIndexNone として設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' このプロパティに xlColorIndexAutomatic を設定すると、セルの網かけのパターンや描画オブジェクトの塗りつぶしパターンが自動的に決定されます。xlColorIndexNone を設定すると、パターンは決定されません。これは、Interior オブジェクトの Pattern プロパティに xlPatternNone を設定するのと同じです。
		''' </remarks>
		Public Property PatternColorIndex() As Object
			Get
				Return DirectCast(InvokeGetProperty(_interior, "PatternColorIndex", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_interior, "PatternColorIndex", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
