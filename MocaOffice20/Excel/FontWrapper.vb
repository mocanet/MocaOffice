
Namespace Excel

	''' <summary>
	''' オブジェクトのフォント属性 (フォント名、フォント サイズ、色など) の全体を表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class FontWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Range</summary>
		Private _range As RangeWrapper

		''' <summary>Excel.Font</summary>
		Private _font As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="range">親のrange</param>
		''' <param name="font">Excel.font</param>
		''' <remarks></remarks>
		Public Sub New(ByVal range As RangeWrapper, ByVal font As Object)
			MyBase.New(range.ApplicationWrapper)
			_range = range
			_font = font
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_font)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _font
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' グラフで使用する文字列の背景の種類を設定します。XlBackground 列挙型の定数のいずれかを使用できます。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Background() As XlBackground
			Get
				Return DirectCast(InvokeGetProperty(_font, "Background", Nothing), XlBackground)
			End Get
			Set(ByVal value As XlBackground)
				InvokeSetProperty(_font, "Background", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、フォントが太字になります。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Bold() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Bold", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Bold", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' フォントの基本色を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Color() As Object
			Get
				Return InvokeGetProperty(_font, "Color", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Color", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' フォントの色を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColorIndex() As Object
			Get
				Return InvokeGetProperty(_font, "ColorIndex", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "ColorIndex", New Object() {value})
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
				Return DirectCast(InvokeGetProperty(_font, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' フォント スタイルを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FontStyle() As Object
			Get
				Return InvokeGetProperty(_font, "FontStyle", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "FontStyle", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、フォントが斜体になります。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Italic() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Italic", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Italic", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' オブジェクトの名前を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Name() As Object
			Get
				Return DirectCast(InvokeGetProperty(_font, "Name", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Name", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、フォントがアウトライン フォントになります。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property OutlineFont() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "OutlineFont", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "OutlineFont", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、フォントが影付きフォントに、または指定したオブジェクトが影付きに設定されます。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Shadow() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Shadow", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Shadow", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' フォント サイズを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Size() As Object
			Get
				Return DirectCast(InvokeGetProperty(_font, "Size", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Size", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、フォントに取り消し線が付けられます。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Strikethrough() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Strikethrough", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Strikethrough", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、指定したフォントが下付き文字になります。既定値は False です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Subscript() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Subscript", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Subscript", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、指定したフォントが上付き文字になります。既定値は False です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Superscript() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Superscript", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Superscript", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' フォントに適用する下線の種類を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Underline() As XlUnderlineStyle
			Get
				Return DirectCast(InvokeGetProperty(_font, "Underline", Nothing), XlUnderlineStyle)
			End Get
			Set(ByVal value As XlUnderlineStyle)
				InvokeSetProperty(_font, "Underline", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
