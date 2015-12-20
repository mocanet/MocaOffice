
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' ページ設定の説明を表します。PageSetup オブジェクトには、すべてのページ設定の属性 (左余白、下余白、用紙サイズなど) が、プロパティとして含まれています。
	''' </summary>
	''' <remarks></remarks>
	Public Class PageSetupWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のシート</summary>
		Private _sheet As SheetWrapper

		''' <summary>Excel.PageSetup</summary>
		Private _pageSetup As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="sheet">親のシート</param>
		''' <param name="pageSetup">Excel.PageSetup</param>
		''' <remarks></remarks>
		Public Sub New(ByVal sheet As SheetWrapper, ByVal pageSetup As Object)
			MyBase.New(sheet.ApplicationWrapper)
			_sheet = sheet
			_pageSetup = pageSetup
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_pageSetup)
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
		''' True の場合、対象となるシートのセルや描画オブジェクトを白黒印刷します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property BlackAndWhite() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "BlackAndWhite", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "BlackAndWhite", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 下余白の大きさをポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property BottomMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "BottomMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "BottomMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 中央に配置するフッターを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CenterFooter() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "CenterFooter", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "CenterFooter", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 中央に配置するヘッダーを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CenterHeader() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "CenterHeader", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "CenterHeader", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 中央に配置するヘッダーの図を表す Graphic オブジェクトを取得します。 図の属性を設定するために使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property CenterHeaderPicture() As GraphicWrapper
			Get
				Dim graphic As Object
				Dim wrap As GraphicWrapper

				graphic = InvokeGetProperty(_pageSetup, "CenterHeaderPicture", Nothing)

				wrap = New GraphicWrapper(Me, graphic)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' True の場合、印刷時のシートのページ レイアウトの設定を、水平方向の中央寄せ (余白を除く) にします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CenterHorizontally() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "CenterHorizontally", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "CenterHorizontally", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True の場合、印刷時のシートのページ レイアウトの設定を、垂直方向の中央寄せ (余白を除く) にします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property CenterVertically() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "CenterVertically", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "CenterVertically", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' グラフの印刷サイズを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ChartSize() As XlObjectSize
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "ChartSize", Nothing), XlObjectSize)
			End Get
			Set(ByVal value As XlObjectSize)
				InvokeSetProperty(_pageSetup, "ChartSize", New Object() {value})
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
		''' True の場合、シートの印刷時にグラフィックスを印刷しない設定 (簡易印刷) になります。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Draft() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "Draft", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "Draft", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 印刷するシートの先頭ページで使用される番号を設定します。xlAutomatic の場合、自動的に先頭ページの番号が選択されます。既定値は xlAutomatic です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FirstPageNumber() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "FirstPageNumber", Nothing), Integer)
			End Get
			Set(ByVal value As Integer)
				InvokeSetProperty(_pageSetup, "FirstPageNumber", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ワークシートを印刷するときに、縦何ページ分で収めるかを示す値を指定します。このプロパティは、ワークシートだけを対象とします
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FitToPagesTall() As Object
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "FitToPagesTall", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_pageSetup, "FitToPagesTall", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ワークシートを印刷するときに、横何ページ分で収めるかを示す値を指定します。このプロパティは、ワークシートだけを対象とします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FitToPagesWide() As Object
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "FitToPagesWide", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_pageSetup, "FitToPagesWide", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' フッターの余白 (用紙の下端からフッターまでの距離) の値をポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FooterMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "FooterMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "FooterMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ヘッダーの余白 (用紙の上端からヘッダーまでの距離) の値をポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property HeaderMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "HeaderMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "HeaderMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 左側に配置するフッターを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property LeftFooter() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "LeftFooter", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "LeftFooter", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 左側に配置するフッターの図を表す Graphic オブジェクトを取得します。 図の属性を設定するために使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property LeftFooterPicture() As GraphicWrapper
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "LeftFooterPicture", Nothing), GraphicWrapper)
			End Get
		End Property

		''' <summary>
		''' 左側に配置するヘッダーを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property LeftHeader() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "LeftHeader", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "LeftHeader", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 左側に配置するヘッダーの図を表す Graphic オブジェクトを取得します。 図の属性を設定するために使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property LeftHeaderPicture() As GraphicWrapper
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "LeftHeaderPicture", Nothing), GraphicWrapper)
			End Get
		End Property

		''' <summary>
		''' 左余白の大きさをポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property LeftMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "LeftMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "LeftMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 大きいワークシートを複数ページに分けて印刷するときに、ページ番号を付ける順番を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Order() As XlOrder
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "Order", Nothing), XlOrder)
			End Get
			Set(ByVal value As XlOrder)
				InvokeSetProperty(_pageSetup, "Order", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 印刷の向き (縦と横) を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Orientation() As XlPageOrientation
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "Orientation", Nothing), XlPageOrientation)
			End Get
			Set(ByVal value As XlPageOrientation)
				InvokeSetProperty(_pageSetup, "Orientation", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 用紙サイズを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PaperSize() As XlPaperSize
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PaperSize", Nothing), XlPaperSize)
			End Get
			Set(ByVal value As XlPaperSize)
				InvokeSetProperty(_pageSetup, "PaperSize", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 印刷するセル範囲を、コード記述時の言語の A1 形式の文字列で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintArea() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintArea", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "PrintArea", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' シートへのコメントの印刷方法を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintComments() As XlPrintLocation
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintComments", Nothing), XlPrintLocation)
			End Get
			Set(ByVal value As XlPrintLocation)
				InvokeSetProperty(_pageSetup, "PrintComments", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 表示される印刷エラーの種類を指定する XlPrintErrors 定数を設定します。この機能を使用すると、ワークシートの印刷時にエラー値が表示されません。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintErrors() As XlPrintErrors
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintErrors", Nothing), XlPrintErrors)
			End Get
			Set(ByVal value As XlPrintErrors)
				InvokeSetProperty(_pageSetup, "PrintErrors", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True の場合、セルの枠線がページに印刷されます。このプロパティは、ワークシートだけを対象とします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintGridlines() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintGridlines", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "PrintGridlines", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True の場合、行と列の番号がページに印刷されます。このプロパティは、ワークシートだけを対象とします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintHeadings() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintHeadings", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "PrintHeadings", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True の場合、文末脚注のようにシート印刷時にセル内のコメントも印刷されます。このプロパティは、ワークシートだけを対象とします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintNotes() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintNotes", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_pageSetup, "PrintNotes", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 印刷品質を設定します。
		''' </summary>
		''' <param name="Index">省略可能です。オブジェクト型 (Object) の値を指定します。水平方向の印刷品質 (1) または垂直方向の印刷品質 (2) を指定します。プリンタによっては、垂直方向の印刷品質を制御していない場合があります。この引数を省略すると、PrintQuality プロパティは水平および垂直の両方向の印刷品質を含む 2 つの要素から成る配列を返します。</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintQuality(<InAttribute()> Optional ByVal Index As Object = Nothing) As Object
			Get
				Dim argsV As ArrayList
				Dim argsN As ArrayList

				argsV = New ArrayList()
				argsN = New ArrayList()

				If Index IsNot Nothing Then
					argsV.Add(Index)
					argsN.Add("Index")
				End If

				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintQuality", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String())), Boolean)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_pageSetup, "PrintQuality", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 各ページの左端に常に表示するセルを含む列を、コード記述時の言語の A1 形式の文字列で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintTitleColumns() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintTitleColumns", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "PrintTitleColumns", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 各ページの上端に常に表示するセルを含む行を、コード記述時の言語の A1 形式の文字列で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PrintTitleRows() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "PrintTitleRows", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "PrintTitleRows", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 右側に配置するフッターを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property RightFooter() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "RightFooter", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "RightFooter", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 右側に配置するフッターの図を表す Graphic オブジェクトを取得します。 図の属性を設定するために使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property RightFooterPicture() As GraphicWrapper
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "RightFooterPicture", Nothing), GraphicWrapper)
			End Get
		End Property

		''' <summary>
		''' 右側に配置するヘッダーを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property RightHeader() As String
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "RightHeader", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_pageSetup, "RightHeader", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 右側に配置するヘッダーの図を表す Graphic オブジェクトを取得します。 図の属性を設定するために使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property RightHeaderPicture() As GraphicWrapper
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "RightHeaderPicture", Nothing), GraphicWrapper)
			End Get
		End Property

		''' <summary>
		''' 右余白の大きさをポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property RightMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "RightMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "RightMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 上余白の大きさをポイント単位で設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property TopMargin() As Double
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "TopMargin", Nothing), Double)
			End Get
			Set(ByVal value As Double)
				InvokeSetProperty(_pageSetup, "TopMargin", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ワークシートを印刷する拡大率または縮小率 (%) の範囲を、10 ～ 400 の値で設定します。このプロパティは、ワークシートだけを対象とします。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Zoom() As Object
			Get
				Return DirectCast(InvokeGetProperty(_pageSetup, "Zoom", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_pageSetup, "Zoom", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
