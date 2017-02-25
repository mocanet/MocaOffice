
Namespace Word

    ''' <summary>
	''' Word.Application のラッパークラス
    ''' </summary>
	''' Word をレイトバインディングにて（参照設定することなく）操作出来ます。<br/>
	''' Word を操作する上でインスタンス化されたオブジェクトは当クラスを開放することで全て開放するようになっています。<br/>
	''' Word を終了するかどうかは、<see cref="WordWrapper.Visible"/> によって自動で判断します。<br/>
	''' 使用するときは、<c>Using</c>句を利用してください。<br/>
    Public Class WordWrapper
        Inherits AbstractAppWrapper

#Region " Declare "

        ''' <summary>Excel.Documents インスタンス</summary>
        Private _documents As DocumentsWrapper

        ''' <summary>終了フラグ</summary>
        Private _quit As Boolean

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' デフォルトコンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New()

            Try

                ' Word クラス ProgID に関連付けられている型を取得
                typApplication = Type.GetTypeFromProgID("Word.Application")

                ' Wordの型がそれに関連付けられていない。(Wordが存在しない)
                If typApplication Is Nothing Then
                    _mylog.Error("Wordが存在しません。インストールされているか確認してください。")
                    Throw New NotSupportedException("Wordが存在しません。インストールされているか確認してください。")
                End If

                ' 各種初期化
                Me.ApplicationWrapper = Me

                ' Wordのインスタンスを作成します。
                officeApp = Activator.CreateInstance(typApplication)

                _mylog.DebugFormat("{0} Version:{1} ProductCode:{2}", Me.Name, Me.Version, Me.ProductCode)

                ' Documents オブジェクトの作成
                _documents = New DocumentsWrapper(Me)
                addObject(_documents)
            Catch ex As OfficeException
                Me.MyDispose()
                Throw ex
            Catch ex As Exception
                Me.MyDispose()
                Throw New OfficeException(Me, ex, "OfficeException のインスタンス生成時にエラーが発生しました。")
            End Try
        End Sub

#End Region

#Region " Overrides "

        ''' <summary>
        ''' 取得した Word インスタンス
        ''' </summary>
        ''' <returns></returns>
        Friend Overrides ReadOnly Property OrigianlInstance As Object
            Get
                Return officeApp
            End Get
        End Property

        ''' <summary>
		''' 自分自身で管理しているWord関係のオブジェクトのメモリ開放
        ''' </summary>
        Public Overrides Sub MyDispose()
            If _quit Then
                Exit Sub
            End If
            Quit(Not Me.Visible)
        End Sub

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
                Return DirectCast(InvokeGetProperty(officeApp, "Name", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' Wordバージョン
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Version() As String
            Get
                Return DirectCast(InvokeGetProperty(officeApp, "Version", Nothing), String)
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
                Return DirectCast(InvokeGetProperty(officeApp, "Visible", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(officeApp, "Visible", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' 確認ダイアログ表示有無
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DisplayAlerts() As WdAlertLevel
            Get
                Return DirectCast(InvokeGetProperty(officeApp, "DisplayAlerts", Nothing), WdAlertLevel)
            End Get
            Set(ByVal value As WdAlertLevel)
                InvokeSetProperty(officeApp, "DisplayAlerts", New Object() {value})
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
                Return DirectCast(InvokeGetProperty(officeApp, "ScreenUpdating", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(officeApp, "ScreenUpdating", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' Word.Documents
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Documents() As DocumentsWrapper
            Get
                Return _documents
            End Get
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' プロダクトコード
        ''' </summary>
        ''' <returns></returns>
        Public Function ProductCode() As String
            Return DirectCast(InvokeMethod(officeApp, "ProductCode", Nothing), String)
        End Function

        ''' <summary>
        ''' 終了
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Quit(Optional ByVal SaveChanges As WdSaveOptions = WdSaveOptions.wdSaveChanges,
                        Optional ByVal Format As WdOriginalFormat = WdOriginalFormat.wdWordDocument,
                        Optional ByVal RouteDocument As Boolean = True
                        )
            ReleaseOfficeObject(officeApp, True)
            _quit = True
        End Sub

#End Region

    End Class

End Namespace
