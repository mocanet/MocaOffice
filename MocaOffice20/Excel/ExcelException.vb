
Namespace Excel

	''' <summary>
	''' Excel ���쎞�̗�O
	''' </summary>
	''' <remarks></remarks>
	Public Class ExcelException
        Inherits ApplicationException

        ''' <summary>Excel.Application �̃��b�p�[</summary>
        Private _xls As AbstractExcelWrapper

#Region " Constructor/DeConstructor "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="Message">�G���[���b�Z�[�W</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal Message As String)
			MyBase.New(Message)
			_xls = xls
		End Sub

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="ex">��O�C���X�^���X</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal ex As Exception)
            MyBase.New("", ex)
            _xls = xls
		End Sub

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="ex">��O�C���X�^���X</param>
		''' <param name="Message">�G���[���b�Z�[�W</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal xls As AbstractExcelWrapper, ByVal ex As Exception, ByVal Message As String)
            MyBase.New(Message, ex)
            _xls = xls
		End Sub

#End Region

	End Class

End Namespace
