
Namespace Word

    ''' <summary>
    ''' Word で現在開かれている すべての Document オブジェクトのコレクション
    ''' </summary>
    ''' <remarks>
    ''' https://msdn.microsoft.com/JA-JP/library/office/ff840891.aspx
    ''' </remarks>
    Public Class DocumentsWrapper
        Inherits AbstractAppWrapper

#Region " Declare "

        ''' <summary>Word.Documents インスタンス</summary>
        Private _documents As Object

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="appWrapper"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal appWrapper As WordWrapper)
            MyBase.New(appWrapper)
            _init(appWrapper)
        End Sub

#End Region

#Region " Overrides "

        ''' <summary>
        ''' 取得した Excel インスタンス
        ''' </summary>
        ''' <returns></returns>
        Friend Overrides ReadOnly Property OrigianlInstance As Object
            Get
                Return _documents
            End Get
        End Property

        ''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
        ''' </summary>
        Public Overrides Sub MyDispose()
            ReleaseOfficeObject(_documents)
        End Sub

#End Region
#Region " Property "

        ''' <summary>
        ''' Word.Application のラッパー
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property App() As WordWrapper
            Get
                Return DirectCast(MyBase.appWrapper, WordWrapper)
            End Get
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' 指定された文書を開き、その文書を Documents コレクションに追加します。Document オブジェクトを返します。
        ''' </summary>
        ''' <returns>
        ''' https://msdn.microsoft.com/ja-jp/library/office/ff835182.aspx
        ''' </returns>
        Public Function Open(ByVal Filename As String,
                Optional ByVal ConfirmConversions As Object = Nothing,
                Optional ByVal [ReadOnly] As Object = Nothing,
                Optional ByVal AddToRecentFiles As Object = Nothing,
                Optional ByVal PasswordDocument As Object = Nothing,
                Optional ByVal PasswordTemplate As Object = Nothing,
                Optional ByVal Revert As Object = Nothing,
                Optional ByVal WritePasswordDocument As Object = Nothing,
                Optional ByVal WritePasswordTemplate As Object = Nothing,
                Optional ByVal Format As Object = Nothing,
                Optional ByVal Encoding As Object = Nothing,
                Optional ByVal Visible As Object = Nothing,
                Optional ByVal OpenConflictDocument As Object = Nothing,
                Optional ByVal OpenAndRepair As Object = Nothing,
                Optional ByVal DocumentDirection As Object = Nothing,
                Optional ByVal NoEncodingDialog As Object = Nothing
                ) As DocumentWrapper
            Dim obj As Object
            Dim document As DocumentWrapper

            argsClear()
            argsAdd("Filename", Filename, True)
            argsAdd("ConfirmConversions", ConfirmConversions)
            argsAdd("ReadOnly", [ReadOnly])
            argsAdd("AddToRecentFiles", AddToRecentFiles)
            argsAdd("PasswordDocument", PasswordDocument)
            argsAdd("PasswordTemplate", PasswordTemplate)
            argsAdd("Revert", Revert)
            argsAdd("WritePasswordDocument", WritePasswordDocument)
            argsAdd("WritePasswordTemplate", WritePasswordTemplate)
            argsAdd("Format", Format)
            argsAdd("Encoding", Encoding)
            argsAdd("Visible", Visible)
            argsAdd("OpenConflictDocument", OpenConflictDocument)
            argsAdd("OpenAndRepair", OpenAndRepair)
            argsAdd("DocumentDirection", DocumentDirection)
            argsAdd("NoEncodingDialog", NoEncodingDialog)

            obj = InvokeMethod(_documents, "Open",
                               argsV.ToArray(),
                               DirectCast(argsN.ToArray(GetType(String)), String())
                               )
            document = New DocumentWrapper(Me.App, obj)
            addObject(document)
            Return document
        End Function

        ''' <summary>
        ''' 初期化処理
        ''' </summary>
        ''' <param name="wrapper">Word.Applicationラッパー</param>
        ''' <remarks></remarks>
        Private Sub _init(ByVal wrapper As WordWrapper)
            ' documents オブジェクトの作成
            _documents = InvokeGetProperty(officeApp, "Documents", Nothing)
        End Sub

#End Region

    End Class

End Namespace
