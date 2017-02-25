

Namespace Word

    ''' <summary>
    ''' 文書を表します
    ''' </summary>
    ''' <remarks>
    ''' https://msdn.microsoft.com/JA-JP/library/office/ff822963.aspx
    ''' </remarks>
    Public Class DocumentWrapper
        Inherits AbstractAppWrapper

#Region " Declare "

        ''' <summary>Word.Document インスタンス</summary>
        Private _document As Object

        ''' <summary>ファイル名</summary>
        Private _filename As String

        ''' <summary>新規ファイル</summary>
        Private _new As Boolean
        ''' <summary>閉じたかフラグ</summary>
        Private _close As Boolean

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="word">Word.Applicationラッパー</param>
        ''' <param name="document">Word.Document</param>
        ''' <param name="newDocument">新規ブックかどうか</param>
        ''' <remarks>
        ''' 既に開いているブックを操作するときに使用します。
        ''' </remarks>
        Friend Sub New(ByVal word As WordWrapper, ByVal document As Object, Optional ByVal newDocument As Boolean = False)
            MyBase.New(word)
            _init(word)

            _document = document

            _new = newDocument

            If _new Then
                Me.Saved = False
            Else
                _filename = FullName
            End If
        End Sub

#End Region

#Region " Overrides "

        Friend Overrides ReadOnly Property OrigianlInstance As Object
            Get
                Return _document
            End Get
        End Property

        Public Overrides Sub MyDispose()
            ReleaseOfficeObject(_document)
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
        ''' オブジェクトの名前を示す文字列を取得します。名前にはディスク上のパスが含まれます。値の取得のみ可能です。文字列型 (String) の値を使用します。 
        ''' </summary>
        ''' <value></value>
        ''' <returns>
        ''' このプロパティを使用すると、<see cref="Path"/> プロパティ、現在のファイル システムの区切り文字、<see cref="Name"/> プロパティを続けて記述するのと同じ結果が得られます。
        ''' </returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FullName() As String
            Get
                Return DirectCast(InvokeGetProperty(_document, "FullName", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' 保存済み
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Saved() As Boolean
            Get
                Return DirectCast(InvokeGetProperty(_document, "Saved", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(_document, "Saved", New Object() {value})
            End Set
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' 文書を PDF 形式または XPS 形式で保存
        ''' </summary>
        ''' <remarks>
        ''' https://msdn.microsoft.com/ja-jp/library/office/ff840962.aspx
        ''' </remarks>
        Public Sub ExportAsFixedFormat(
                                      ByVal OutputFileName As String,
                                      ByVal ExportFormat As WdExportFormat,
                                      Optional ByVal OpenAfterExport As Boolean = False,
                                      Optional ByVal OptimizeFor As WdExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                                      Optional ByVal Range As WdExportRange = WdExportRange.wdExportAllDocument,
                                      Optional ByVal From As Integer = 0,
                                      Optional ByVal [To] As Integer = 0,
                                      Optional ByVal Item As WdExportItem = WdExportItem.wdExportDocumentContent,
                                      Optional ByVal IncludeDocProps As Boolean = False,
                                      Optional ByVal KeepIRM As Boolean = True,
                                      Optional ByVal CreateBookmarks As WdExportCreateBookmarks = WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                                      Optional ByVal DocStructureTags As Boolean = True,
                                      Optional ByVal BitmapMissingFonts As Boolean = True,
                                      Optional ByVal UseISO19005_1 As Boolean = False,
                                      Optional ByVal FixedFormatExtClassPtr As Object = Nothing
                                      )

            argsClear()
            argsAdd("OutputFileName", OutputFileName, True)
            argsAdd("ExportFormat", ExportFormat, True)
            argsAdd("OpenAfterExport", OpenAfterExport)
            argsAdd("OptimizeFor", OptimizeFor)
            argsAdd("Range", Range)
            argsAdd("From", From)
            argsAdd("To", [To])
            argsAdd("Item", Item)
            argsAdd("IncludeDocProps", IncludeDocProps)
            argsAdd("KeepIRM", KeepIRM)
            argsAdd("CreateBookmarks", CreateBookmarks)
            argsAdd("DocStructureTags", DocStructureTags)
            argsAdd("BitmapMissingFonts", BitmapMissingFonts)
            argsAdd("UseISO19005_1", UseISO19005_1)
            argsAdd("FixedFormatExtClassPtr", FixedFormatExtClassPtr)


            InvokeMethod(_document, "ExportAsFixedFormat",
                         argsV.ToArray,
                         DirectCast(argsN.ToArray(GetType(String)), String())
                         )
        End Sub

        ''' <summary>
        ''' 初期化処理
        ''' </summary>
        ''' <param name="wrapper">Word.Applicationラッパー</param>
        ''' <remarks></remarks>
        Private Sub _init(ByVal wrapper As WordWrapper)
            _new = False
            _close = False
        End Sub

#End Region
    End Class

End Namespace
