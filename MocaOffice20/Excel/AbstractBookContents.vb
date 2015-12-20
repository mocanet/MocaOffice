
Imports System.IO

Namespace Excel

	''' <summary>
	''' ブック操作用のイベント引数
	''' </summary>
	''' <remarks></remarks>
	Public Class BookContentsEventArgs
		Inherits EventArgs

		''' <summary>操作中のシート名</summary>
		Public SheetName As String
	End Class

	''' <summary>
	''' 進捗
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks></remarks>
	Public Delegate Sub PerformStep(ByVal sender As AbstractBookContents, ByVal e As BookContentsEventArgs)

	''' <summary>
	''' Excelブックを操作する為の共通処理を保持する抽象クラス
	''' </summary>
	''' <remarks>
	''' 当抽象クラスでは、Excel出力する際のお決まりのブック操作を既に実装してあります。<br/>
	''' 当クラスを継承しサブクラス化して、ブックファイル名（<see cref="AbstractBookContents.SaveFilename"/>、<see cref="AbstractBookContents.TemplateFilename"/>）を設定し、
	''' <see cref="AbstractBookContents.Add"/> にて各シート出力クラスを追加することで、比較的簡単に Excel 出力機能を実装出来るようになってます。<br/>
	''' シート出力クラスは <seealso cref="ISheetContents"/>, <seealso cref="ISheetContentsUseTemplate"/>, <seealso cref="ISheetContentsUseTemplateMakeCsv"/>
	''' を実装してください。<br/>
	''' </remarks>
	Public MustInherit Class AbstractBookContents

		'Public Event PerformStep As PerformStep

		''' <summary>Excelアプリケーション</summary>
		Private _app As ExcelWrapper

		''' <summary>ファイル名</summary>
		Private _saveFilename As String

		''' <summary>テンプレートとなるExcelファイル名</summary>
		Protected myTemplateFilename As String

		''' <summary>Excelを画面表示するかどうかを判定する変数</summary>
		Private _display As Boolean
		''' <summary>Excelを保存するかどうかを判定する変数</summary>
		Private _save As Boolean
		''' <summary>Excelを印刷するかどうかを判定する変数</summary>
		Private _print As Boolean

		Private _performStep As PerformStep

		''' <summary>シート内容</summary>
		Protected sheetContents As IList(Of ISheetContents)

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' デフォルトコンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
			sheetContents = New List(Of ISheetContents)
		End Sub

#End Region

#Region " プロパティ "

		''' <summary>
		''' Excelアプリケーション
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected ReadOnly Property App() As ExcelWrapper
			Get
				Return _app
			End Get
		End Property

		''' <summary>ファイル名</summary>
		Public Property SaveFilename() As String
			Get
				Return _saveFilename
			End Get
			Set(ByVal value As String)
				_saveFilename = value
			End Set
		End Property

		''' <summary>テンプレートとなるExcelファイル名</summary>
		Public Property TemplateFilename() As String
			Get
				Return myTemplateFilename
			End Get
			Set(ByVal value As String)
				myTemplateFilename = value
			End Set
		End Property

		''' <summary>Excelを画面表示するかどうかを判定する変数</summary>
		Public Property Display() As Boolean
			Get
				Return _display
			End Get
			Set(ByVal value As Boolean)
				_display = value
			End Set
		End Property

		''' <summary>Excelを保存するかどうかを判定する変数</summary>
		Public Property Save() As Boolean
			Get
				Return _save
			End Get
			Set(ByVal value As Boolean)
				_save = value
			End Set
		End Property

		''' <summary>Excelを印刷するかどうかを判定する変数</summary>
		Public Property Print() As Boolean
			Get
				Return _print
			End Get
			Set(ByVal value As Boolean)
				_print = value
			End Set
		End Property

#End Region

		''' <summary>
		''' シートを追加する
		''' </summary>
		''' <param name="sheet">シートコンテンツ</param>
		''' <remarks></remarks>
		Public Sub Add(ByVal sheet As ISheetContents)
			sheetContents.Add(sheet)
		End Sub

		''' <summary>
		''' コンテンツ出力
		''' </summary>
		''' <param name="performStep">進捗（プログレスバーなど）を必要とするときは、<see cref="PerformStep"/> デリゲートを指定してください。</param>
		''' <remarks>
		''' </remarks>
		Public Sub Write(Optional ByVal performStep As PerformStep = Nothing)
			_performStep = performStep
			Using xls As ExcelWrapper = New ExcelWrapper
				Try
					_app = xls
					_writeContents()
				Catch ex As Exception
					xls.Dispose()
					Throw ex
				End Try
			End Using
			''DoSomethingメソッドを別のスレッドで実行する
			''Threadオブジェクトを作成する
			'Dim t As New System.Threading.Thread( _
			' New System.Threading.ThreadStart( _
			' AddressOf DoSomething))
			''スレッドを開始する
			't.Start()
			''t.Join()
		End Sub

		Private Sub DoSomething()
			Using xls As ExcelWrapper = New ExcelWrapper
				_app = xls
				_writeContents()
			End Using
		End Sub

		''' <summary>
		''' 出力ロジック
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents()
			Dim cbValue As String

			' シート指定が無いときは処理終了
			If sheetContents.Count = 0 Then
				Throw New ExcelException(Me.App, "シートが一つも指定されていません。")
			End If

			' クリップボードの退避
			cbValue = My.Computer.Clipboard.GetText()

			Try
				' Excelテンプレートファイルの存在チェック
				_xlsTemplateFileExists()

				' Excel出力
				_writeExcel()
			Finally
				Try
					' クリップボードの復元
					My.Computer.Clipboard.SetText(cbValue)
				Catch ex As Exception
				End Try
			End Try
		End Sub

		''' <summary>
		''' Excelテンプレートファイルの存在チェック
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _xlsTemplateFileExists()
			If TemplateFilename.Length = 0 Then
				Exit Sub
			End If

			' Excelテンプレートファイルの存在チェック
			If Not File.Exists(TemplateFilename) Then
				Throw New ExcelException(Me.App, String.Format("Excelテンプレートファイルが存在しません。({0})", TemplateFilename))
			End If
		End Sub

		''' <summary>
		''' 出力
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeExcel()
			Dim xlBook As BookWrapper
			Dim xlSheet As SheetWrapper
			Dim e As BookContentsEventArgs

			e = New BookContentsEventArgs()
			e.SheetName = String.Empty

			Try
				Me.App.Visible = False
				Me.App.DisplayAlerts = False
				Me.App.Interactive = False
				Me.App.ScreenUpdating = False

				' ブックを取得
				If TemplateFilename.Length = 0 Then
					' 新規ブック作成
					xlBook = Me.App.Workbooks.Add(Me.SaveFilename)
				Else
					' テンプレートファイルの読込
					xlBook = Me.App.Workbooks.Open(TemplateFilename)
				End If

				' 各シート処理
				For ii As Integer = 0 To sheetContents.Count - 1
					Dim contents As ISheetContents
					Dim contentsWriter As SheetContents

					contents = sheetContents(ii)

					If contents.BaseSheetName = String.Empty Then
						If contents.SaveSheetName = String.Empty Then
							xlSheet = xlBook.Worksheets(ii + 1)
							contents.BaseSheetName = xlSheet.Name
						Else
							xlSheet = xlBook.Worksheets.Add(contents.SaveSheetName)
						End If
					Else
						xlSheet = xlBook.Worksheets(contents.BaseSheetName)
					End If
					xlSheet.Activate()
					e.SheetName = xlSheet.Name

					' 出力
					contentsWriter = SheetContentsFactory.Create(contents, xlSheet)
					Dim tim As Stopwatch = New Stopwatch()
					tim.Start()
					contentsWriter.WriteContents()
					tim.Stop()
					_mylog.DebugFormat("[{0}] Write Time {1}", IIf(contents.SaveSheetName.Length = 0, contents.BaseSheetName, contents.SaveSheetName), tim.ElapsedMilliseconds)

					' ホームポジションに移動
					If Not xlSheet.IsDeleted Then
						xlSheet.Range("A1").Select()
					End If

					_runPerformStep(e)
				Next

				' 終了処理
				_endWrite(xlBook)
			Catch chex As ExcelException
				Throw chex
			Catch ex As Exception
				Throw New ExcelException(Me.App, ex)
			End Try
		End Sub

		''' <summary>
		''' Excel終了処理
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' </remarks>
		Private Function _endWrite(ByVal xlBook As BookWrapper) As Boolean
			Dim sheet As SheetWrapper

			' ホームポジションに移動
			sheet = Me.App.ActiveWorkbook.Worksheets(1)
			sheet.Activate()
			sheet.Range("A1").Select()

			' 自動保存
			If Save Then
				If Not _autoSave(xlBook) Then
					Exit Function
				End If
			End If
			' 自動印刷
			If Print Then
				If Not _autoPrint() Then
					Exit Function
				End If
			End If
			' 画面表示
			If Display Then
				If Not _autoDisplay() Then
					Exit Function
				End If
			End If
		End Function

		''' <summary>
		''' 自動保存
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' <see cref="SaveFilename" /> に指定されている名前で保存します。
		''' </remarks>
		Private Function _autoSave(ByVal xlBook As BookWrapper) As Boolean
			' 保存ファイル名が指定していないときは無視
			If _saveFilename.Length = 0 Then
				Exit Function
			End If

			Try
				_app.DisplayAlerts = False
				xlBook.SaveAs(_saveFilename)
				_app.DisplayAlerts = True
				Return True
			Catch ex As ExcelException
				Throw ex
			Catch ex As Exception
				Throw New ExcelException(_app, ex, "Excel 自動保存時にエラーが発生しました。")
			End Try
		End Function

		''' <summary>
		''' 自動印刷
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' </remarks>
		Private Function _autoPrint() As Boolean
			_app.ActiveWorkbook.PrintOut()	' 印刷
			Return True
		End Function

		''' <summary>
		''' 画面へ表示する
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' 下記の項目を <c>Ture</c> に設定します。<br/>
		''' <list>
		''' <item><description><see cref="ExcelWrapper.ScreenUpdating"/></description></item>
		''' <item><description><see cref="ExcelWrapper.Interactive"/></description></item>
		''' <item><description><see cref="ExcelWrapper.DisplayAlerts"/></description></item>
		''' <item><description><see cref="ExcelWrapper.Visible"/></description></item>
		''' </list>
		''' </remarks>
		Private Function _autoDisplay() As Boolean
			_app.ScreenUpdating = True
			_app.Interactive = True
			_app.DisplayAlerts = True
			_app.Visible = True
			Return True
		End Function

		''' <summary>
		''' 進捗報告
		''' </summary>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub _runPerformStep(ByVal e As BookContentsEventArgs)
			Try
				'RaiseEvent PerformStep(Me, e)
				If _performStep Is Nothing Then
					Exit Sub
				End If

				_performStep(Me, e)
				'_performStep.Target.GetType.InvokeMember(_performStep.GetType().Name, Reflection.BindingFlags.InvokeMethod, Nothing, _performStep.Target, New Object() {Me, e})
			Catch ex As Exception
				_mylog.ErrorFormat(ex.Message)
			End Try
		End Sub

	End Class

End Namespace
