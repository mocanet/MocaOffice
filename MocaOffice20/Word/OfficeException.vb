
Namespace Word

    ''' <summary>
    ''' Office 操作時の例外
    ''' </summary>
    ''' <remarks></remarks>
    Public Class OfficeException
        Inherits ApplicationException

        ''' <summary>Word.Application のラッパー</summary>
        Private _appWrapper As AbstractAppWrapper

#Region " Constructor/DeConstructor "

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="appWrapper">Application のラッパー</param>
        ''' <param name="Message">エラーメッセージ</param>
        ''' <remarks>
        ''' </remarks>
        Public Sub New(ByVal appWrapper As AbstractAppWrapper, ByVal Message As String)
            MyBase.New(Message)
            _appWrapper = appWrapper
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="appWrapper">Application のラッパー</param>
        ''' <param name="ex">例外インスタンス</param>
        ''' <remarks>
        ''' </remarks>
        Public Sub New(ByVal appWrapper As AbstractAppWrapper, ByVal ex As Exception)
            MyBase.New("", ex)
            _appWrapper = appWrapper
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="appWrapper">Application のラッパー</param>
        ''' <param name="ex">例外インスタンス</param>
        ''' <param name="Message">エラーメッセージ</param>
        ''' <remarks>
        ''' </remarks>
        Public Sub New(ByVal appWrapper As AbstractAppWrapper, ByVal ex As Exception, ByVal Message As String)
            MyBase.New(Message, ex)
            _appWrapper = appWrapper
        End Sub

#End Region

    End Class

End Namespace
