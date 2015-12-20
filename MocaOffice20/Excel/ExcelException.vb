
Namespace Excel

	''' <summary>
	''' Excel 操作時の例外
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelException
        Inherits ApplicationException

        ''' <summary>Excel.Application のラッパー</summary>
        Private _xls As AbstractExcelWrapper

#Region " Constructor/DeConstructor "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="Message">エラーメッセージ</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal Message As String)
			MyBase.New(Message)
			_xls = xls
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="ex">例外インスタンス</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal ex As Exception)
            MyBase.New("", ex)
            _xls = xls
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="ex">例外インスタンス</param>
		''' <param name="Message">エラーメッセージ</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal ex As Exception, ByVal Message As String)
            MyBase.New(Message, ex)
            _xls = xls
		End Sub

#End Region

	End Class

End Namespace
