
Imports System.Reflection

Namespace Excel

    ''' <summary>
    ''' Excel.Workbook のラッパークラス
    ''' </summary>
    ''' <remarks>
    ''' Microsoft Excel のブックを表します。
    ''' </remarks>
    Public Class BookWrapper
        Inherits AbstractExcelWrapper

        ''' <summary>Excel.Workbook</summary>
        Private _book As Object

        ''' <summary>Excel.Sheets</summary>
        Private _sheets As SheetsWrapper
        ''' <summary>Excel.Worksheets</summary>
        Private _worksheets As WorksheetsWrapper

        ''' <summary>ファイル名</summary>
        Private _filename As String

        ''' <summary>新規ファイル</summary>
        Private _new As Boolean
        ''' <summary>閉じたかフラグ</summary>
        Private _close As Boolean

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="xls">Excel.Applicationラッパー</param>
        ''' <param name="workbook">Excel.Workbook</param>
        ''' <param name="newBook">新規ブックかどうか</param>
        ''' <remarks>
        ''' 既に開いているブックを操作するときに使用します。
        ''' </remarks>
        Friend Sub New(ByVal xls As ExcelWrapper, ByVal workbook As Object, Optional ByVal newBook As Boolean = False)
            MyBase.New(xls)
            _init(xls)

            _book = workbook

            _new = newBook

            If _new Then
                Me.Saved = False
            Else
                _filename = FullName
            End If
        End Sub

#End Region
#Region " Overrides "

        ''' <summary>
        ''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub MyDispose()
            ReleaseExcelObject(_book)
        End Sub

        ''' <summary>
        ''' 取得した Excel インスタンス
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Overrides ReadOnly Property OrigianlInstance() As Object
            Get
                Return _book
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
        ''' ファイル名
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Filename() As String
            Get
                Return _filename
            End Get
            Set(ByVal value As String)
                _filename = value
            End Set
        End Property

        ''' <summary>
        ''' 保存済み
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Saved() As Boolean
            Get
                Return DirectCast(InvokeGetProperty(_book, "Saved", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(_book, "Saved", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' ブック内の全シート
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Sheets() As SheetsWrapper
            Get
                If _sheets Is Nothing Then
                    _sheets = New SheetsWrapper(Me)
                    addXlsObject(_sheets)
                End If
                Return _sheets
            End Get
        End Property

        ''' <summary>
        ''' ブック内の全ワークシート
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Worksheets() As WorksheetsWrapper
            Get
                If _worksheets Is Nothing Then
                    _worksheets = New WorksheetsWrapper(Me)
                    addXlsObject(_worksheets)
                End If
                Return _worksheets
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
                xl = InvokeGetProperty(_book, "ActiveSheet", Nothing)
                If xl IsNot Nothing Then
                    sheet = New SheetWrapper(Me, xl)
                    addXlsObject(sheet)
                End If

                Return sheet
            End Get
        End Property

        ''' <summary>
        ''' オブジェクトの名前を示す文字列を取得します。名前にはディスク上のパスが含まれます。値の取得のみ可能です。文字列型 (String) の値を使用します。 
        ''' </summary>
        ''' <value></value>
        ''' <returns>
        ''' このプロパティを使用すると、<see cref="Path"/> プロパティ、現在のファイル システムの区切り文字、<see cref="Name"/> プロパティを続けて記述するのと同じ結果が得られます。
        ''' </returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FullName() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "FullName", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' アプリケーションまでの絶対パスを取得します。このパスでは、最後の区切り文字とアプリケーション名が省かれます。値の取得のみ可能です。文字列型 (String) の値を使用します。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Path() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "Path", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' オブジェクトの名前を取得します。値の取得のみ可能です。文字列型 (String) の値を使用します。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "Name", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' オブジェクトの名前を取得します。値の取得のみ可能です。文字列型 (String) の値を使用します。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Names() As NamesWrapper
            Get
                Dim xl As Object
                Dim obj As NamesWrapper
                obj = Nothing
                xl = InvokeGetProperty(_book, "Names", Nothing)
                If xl IsNot Nothing Then
                    obj = New NamesWrapper(Me, xl)
                    addXlsObject(obj)
                End If
                Return obj
            End Get
        End Property

#End Region
#Region " メソッド "

        ''' <summary>
        ''' 初期化処理
        ''' </summary>
        ''' <param name="xls">Excel.Applicationラッパー</param>
        ''' <remarks></remarks>
        Private Sub _init(ByVal xls As ExcelWrapper)
            _new = False
            _close = False
        End Sub

        ''' <summary>
        ''' シートをアクティブにする
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Activate()
            InvokeMethod(_book, "Activate", Nothing)
        End Sub

        ''' <summary>
        ''' ブックを閉じる
        ''' </summary>
        ''' <remarks>
        ''' 閉じると同時に保存します。
        ''' </remarks>
        Public Sub Close()
            Close(True)
        End Sub

        ''' <summary>
        ''' ブックを閉じる
        ''' </summary>
        ''' <param name="save">保存するかどうか</param>
        ''' <remarks></remarks>
        Public Sub Close(ByVal save As Boolean)
            If _book Is Nothing Then
                Exit Sub
            End If
            If _close Then
                Exit Sub
            End If

            If save Then
                Me.Save()
            End If

            InvokeMethod(_book, "Close", New Object() {save})
            _close = True
        End Sub

        '''' <summary>
        '''' ブックを保存
        '''' </summary>
        '''' <remarks></remarks>
        'Friend Sub SaveTitle()
        '	If Not _new Then
        '		Exit Sub
        '	End If
        '	Dim flg As Boolean
        '	flg = Saved
        '	Saved = False
        '	SaveAs(_filename)
        '	Saved = flg
        '	System.IO.File.Delete(_filename)
        'End Sub

        ''' <summary>
        ''' ブックを保存
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Save()
            If Saved Then
                Exit Sub
            End If

            If _new Then
                SaveAs(_filename)
                Exit Sub
            End If

            InvokeMethod(_book, "Save", Nothing)
        End Sub

        ''' <summary>
        ''' ブックにファイル名を付けて保存
        ''' </summary>
        ''' <param name="filename">ファイル名</param>
        ''' <remarks></remarks>
        Public Sub SaveAs(ByVal filename As String)
            If Saved Then
                Exit Sub
            End If

            InvokeMethod(_book, "SaveAs", New Object() {filename})
        End Sub

        ''' <summary>
        ''' ブックを印刷
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub PrintOut()
            InvokeMethod(_book, "PrintOut", Nothing)
        End Sub

        ''' <summary>
        ''' 自動実行マクロの実行
        ''' </summary>
        ''' <param name="Which">実行する自動実行マクロを指定<br/>
        ''' XlRunAutoMacroで定義された定数を使用します。</param>
        ''' <remarks></remarks>
        Public Sub RunAutoMacros(ByVal which As XlRunAutoMacro)
            InvokeMethod(_book, "RunAutoMacros", New Object() {which})
        End Sub

        ''' <summary>
        ''' 指定した形式のファイルにエクスポート
        ''' </summary>
        ''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff198122.aspx</remarks>
        Public Sub ExportAsFixedFormat(ByVal xlType As FixedFormatType,
                                       Optional ByVal filename As String = Nothing,
                                       Optional ByVal quality As FixedFormatQuality = FixedFormatQuality.QualityStandard,
                                       Optional ByVal includeDocProperties As Boolean = False,
                                       Optional ByVal ignorePrintAreas As Boolean = False,
                                       Optional ByVal [from] As Integer = 0,
                                       Optional ByVal [to] As Integer = 0,
                                       Optional ByVal openAfterPublish As Boolean = False,
                                       Optional ByVal fixedFormatExtClassPtr As Object = Nothing)
            Dim argsV As New List(Of Object)
            Dim argsN As New List(Of String)

            argsV.Add(xlType)
            argsN.Add("Type")
            If filename IsNot Nothing Then
                argsV.Add(filename)
                argsN.Add("Filename")
            End If
            argsV.Add(quality)
            argsN.Add("Quality")
            argsV.Add(includeDocProperties)
            argsN.Add("IncludeDocProperties")
            argsV.Add(ignorePrintAreas)
            argsN.Add("IgnorePrintAreas")
            If [from] > 0 Then
                argsV.Add([from])
                argsN.Add("From")
            End If
            If [to] > 0 Then
                argsV.Add([to])
                argsN.Add("To")
            End If
            argsV.Add(openAfterPublish)
            argsN.Add("OpenAfterPublish")
            If fixedFormatExtClassPtr IsNot Nothing Then
                argsV.Add(fixedFormatExtClassPtr)
                argsN.Add("FixedFormatExtClassPtr")
            End If

            'args = New Object() {xlType, filename, quality, includeDocProperties, ignorePrintAreas, [from], [to], openAfterPublish, fixedFormatExtClassPtr}

            InvokeMethod(_book, "ExportAsFixedFormat", argsV.ToArray, argsN.ToArray)
        End Sub

#End Region

    End Class

End Namespace
