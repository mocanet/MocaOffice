
Namespace Excel

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Public Class GraphicWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のPageSetup</summary>
		Private _pageSetup As PageSetupWrapper

		''' <summary>Excel.Graphic</summary>
		Private _graphic As Object


		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="pageSetup">親のPageSetup</param>
		''' <param name="graphic">Excel.Graphic</param>
		''' <remarks></remarks>
		Public Sub New(ByVal pageSetup As PageSetupWrapper, ByVal graphic As Object)
			MyBase.New(pageSetup.ApplicationWrapper)
			_pageSetup = pageSetup
			_graphic = graphic
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_graphic)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _pageSetup
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 指定した図または OLE オブジェクトの明るさを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Brightness() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "Brightness", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "Brightness", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトに適用するイメージ コントロールの種類を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColorType() As MsoPictureColorType
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "ColorType", Nothing), MsoPictureColorType)
			End Get
			Set(ByVal value As MsoPictureColorType)
				InvokeSetProperty(_graphic, "ColorType", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトのコントラストを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Contrast() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "Contrast", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "Contrast", New Object() {value})
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
				Return DirectCast(InvokeGetProperty(_pageSetup, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトの下端からトリミングされるポイント数を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CropBottom() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "CropBottom", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "CropBottom", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトの左端からトリミングされるポイント数を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CropLeft() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "CropLeft", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "CropLeft", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトの右端からトリミングされるポイント数を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CropRight() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "CropRight", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "CropRight", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定した図または OLE オブジェクトの上端からトリミングされるポイント数を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CropTop() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "CropTop", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "CropTop", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したソース オブジェクトを保存する場所の URL (イントラネットまたは Web) あるいはパス (ローカルまたはネットワーク) を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Filename() As String
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "Filename", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_graphic, "Filename", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' オブジェクトのポイント単位の高さ。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Height() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "Height", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "Height", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True を設定すると、指定した図形のサイズを変更しても元の比率が保持されます。False を設定すると、サイズを変更するときに図形の高さと幅を個別に変更できます。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property LockAspectRatio() As MsoTriState
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "LockAspectRatio", Nothing), MsoTriState)
			End Get
			Set(ByVal value As MsoTriState)
				InvokeSetProperty(_graphic, "LockAspectRatio", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' オブジェクトの幅をポイント単位で指定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Width() As Single
			Get
				Return DirectCast(InvokeGetProperty(_graphic, "Width", Nothing), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_graphic, "Width", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
