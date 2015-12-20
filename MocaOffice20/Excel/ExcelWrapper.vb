
Imports System.Reflection

Namespace Excel

	''' <summary>
	''' Excel.Application のラッパークラス
	''' </summary>
	''' <remarks>
	''' Excel をレイトバインディングにて（参照設定することなく）操作出来ます。<br/>
	''' Excel を操作する上でインスタンス化されたオブジェクトは当クラスを開放することで全て開放するようになっています。<br/>
	''' Excel を終了するかどうかは、<see cref="ExcelWrapper.Visible"/> によって自動で判断します。<br/>
	''' 使用するときは、<c>Using</c>句を利用してください。<br/>
	''' <br/>
	''' <example>
	''' <code lang="vb">
	''' Using xls As ExcelWrapper = New ExcelWrapper()
	''' 	Try
	''' 		Dim book As BookWrapper
	''' 
	''' 		xls.Visible = False
	''' 		xls.DisplayAlerts = False
	''' 		xls.Interactive = False
	''' 		xls.ScreenUpdating = False
	''' 
	''' 		book = xls.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "test.xls"))
	''' 		book.Save()
	''' 		book.Close(False)
	''' 	Catch ex As Exception
	''' 		xls.Dispose()
	''' 	End Try
	''' End Using
	''' </code>
	''' </example>
	''' </remarks>
	Public Class ExcelWrapper
		Inherits AbstractExcelWrapper

		''' <summary>Excel.Workbooks インスタンス</summary>
		Private _workbooks As BooksWrapper


		Private _quit As Boolean

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' デフォルトコンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
			MyBase.New()

			Try

				' Excelクラス ProgID に関連付けられている型を取得
				typApplication = Type.GetTypeFromProgID("Excel.Application")

				' Excelの型がそれに関連付けられていない。(Excelが存在しない)
				If typApplication Is Nothing Then
					_mylog.Error("Excelが存在しません。インストールされているか確認してください。")
					Throw New NotSupportedException("Excelが存在しません。インストールされているか確認してください。")
				End If

				' 各種初期化
				Me.ApplicationWrapper = Me

				' Excelのインスタンスを作成します。
				xlsApp = Activator.CreateInstance(typApplication)

				_mylog.DebugFormat("{0} Version:{1} ProductCode:{2}", Me.Name, Me.Version, Me.ProductCode)

				' Booksオブジェクトの作成
				_workbooks = New BooksWrapper(Me)
				addXlsObject(_workbooks)
			Catch ex As ExcelException
				Me.MyDispose()
				Throw ex
			Catch ex As Exception
				Me.MyDispose()
				Throw New ExcelException(Me, ex, "ExcelWrapper のインスタンス生成時にエラーが発生しました。")
			End Try
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			If _quit Then
				Exit Sub
			End If
			Quit(Not Me.Visible)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return xlsApp
			End Get
		End Property

#End Region

#Region " プロパティ "

		''' <summary>
		''' アプリケーション名
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Name() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Name", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' Excelバージョン
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Version() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Version", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' プロダクトコード
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ProductCode() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "ProductCode", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' アクティブなウィンドウで選択されているオブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Selection() As SelectionWrapper
			Get
				Dim xl As Object
				Dim sel As SelectionWrapper

				sel = Nothing
				xl = InvokeGetProperty(xlsApp, "Selection", Nothing)
				If xl IsNot Nothing Then
					sel = New SelectionWrapper(Me, xl)
					addXlsObject(sel)
				End If

				Return sel
			End Get
		End Property

		''' <summary>
		''' 画面表示有無
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Visible() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Visible", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "Visible", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 確認ダイアログ表示有無
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property DisplayAlerts() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "DisplayAlerts", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "DisplayAlerts", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 画面更新を有無
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ScreenUpdating() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "ScreenUpdating", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "ScreenUpdating", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ユーザーからの干渉有無
		''' </summary>
		''' <value></value>
		''' <returns>
		''' True を設定すると、Microsoft Excel が対話モードになります。このプロパティは通常 True です。このプロパティに False を設定すると、キーボードおよびマウスからの入力を受け付けなくなります。ただし、コードによって表示されたダイアログ ボックスへの入力は可能です。入力できない状態にしておくと、コードを使用して Microsoft Excel のオブジェクトを移動したりアクティブにしたりしているときに、ユーザーからの干渉を防ぐことができます。<br/>
		''' このプロパティに False を設定した場合は、True に戻すのを忘れないようにしてください。コードの実行が終了しても、このプロパティは自動的に True に戻りません。
		''' </returns>
		''' <remarks></remarks>
		Public Property Interactive() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Interactive", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "Interactive", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' イベント発生の有無
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' True の場合、指定されたオブジェクトに対してイベントが発生します。値の取得および設定が可能です。ブール型 (Boolean) の値を使用します。
		''' </remarks>
		Public Property EnableEvents() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "EnableEvents", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "EnableEvents", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' Excel.Workbooks
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Workbooks() As BooksWrapper
			Get
				Return _workbooks
			End Get
		End Property

		''' <summary>
		''' 現在アクティブなブック
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ActiveWorkbook() As BookWrapper
			Get
				Dim book As Object
				Dim bookWrap As BookWrapper
				bookWrap = Nothing
				book = InvokeGetProperty(xlsApp, "ActiveWorkbook", Nothing)
				If book IsNot Nothing Then
					Dim nm As String
					nm = DirectCast(InvokeGetProperty(book, "Name", Nothing), String)
					bookWrap = _workbooks.GetMyBook(nm)
					If bookWrap Is Nothing Then
						bookWrap = New BookWrapper(Me, book)
					End If
				End If
				Return bookWrap
			End Get
		End Property

		''' <summary>
		''' 現在アクティブなシート
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ActiveSheet() As SheetWrapper
			Get
				Dim xl As Object
				Dim sheet As SheetWrapper
				sheet = Nothing
				xl = InvokeGetProperty(xlsApp, "ActiveSheet", Nothing)
				If xl IsNot Nothing Then
					sheet = New SheetWrapper(ActiveWorkbook, xl)
					addXlsObject(sheet)
				End If

				Return sheet
			End Get
		End Property

		''' <summary>
		''' Microsoft Excel で新規ブックに自動的に挿入されるシートの数を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property SheetsInNewWorkbook() As Integer
			Get
				Return CInt(InvokeGetProperty(xlsApp, "SheetsInNewWorkbook", Nothing))
			End Get
			Set(ByVal value As Integer)
				InvokeSetProperty(xlsApp, "SheetsInNewWorkbook", New Object() {value})
			End Set
		End Property

#End Region

		''' <summary>
		''' Excel終了
		''' </summary>
		''' <param name="windowClose">画面を閉じるかどうか</param>
		''' <remarks></remarks>
		Public Sub Quit(Optional ByVal windowClose As Boolean = False)
			ReleaseExcelObject(xlsApp, windowClose)
			_quit = True
		End Sub

	End Class

End Namespace
