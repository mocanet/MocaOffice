
Imports System.IO

Namespace Excel

	''' <summary>
	''' �ꗗ�`���̃e���v���[�g�V�[�g���g�p�����Ƃ��ɁA
	''' �f�[�^����xCSV�t�@�C���֏o�͂��ACSV�t�@�C����Ǎ����Excel�֓\��t�����@�̃C���^�t�F�[�X
	''' </summary>
	''' <remarks></remarks>
	Public Interface ISheetContentsUseTemplateMakeCsv
		Inherits ISheetContents

		''' <summary>
		''' ���ו��̏o�͊J�n�s
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' �e���v���[�g�V�[�g��̈ꗗ�o�͂����ŏ��̍s��Ԃ��悤�ɂ���B
		''' </remarks>
		ReadOnly Property StartRow() As Integer

		''' <summary>
		''' ���ו��̏o�͊J�n��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' �e���v���[�g�V�[�g��̈ꗗ�o�͂����ŏ��̗��Ԃ��悤�ɂ���B
		''' </remarks>
		ReadOnly Property StartCol() As Integer

		''' <summary>
		''' ���ו��̏o�͗�
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property ColumnLength() As Integer

		''' <summary>
		''' �o�͂���f�[�^����
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property DataCount() As Integer

		''' <summary>
		''' �f�[�^��CSV���ԃt�@�C���Ƃ��ďo�͂���
		''' </summary>
		''' <param name="csv"></param>
		''' <remarks>
		''' �u�C�v�i�J���}�j��؂�̕�������o�͂��Ă��������B<br/>
		''' </remarks>
		Sub CsvWrite(ByRef csv As StreamWriter)

		''' <summary>
		''' CSV�t�@�C����Excel�ɂĊJ�����̃t�H�[�}�b�g���w�肷��
		''' </summary>
		''' <param name="columnIndex"></param>
		''' <remarks>
		''' CSV�t�@�C����Excel�ɂĊJ�����̃t�H�[�}�b�g���w�肷��ꍇ�͎w�肵�Ă��������B<br/>
		''' �f�t�H���g�ł́u��ʁv�̃t�H�[�}�b�g�ɂēǂݍ��݂܂��B
		''' </remarks>
		Function SetCsvOpenFormat(ByVal columnIndex As Integer) As XlColumnDataType

	End Interface

End Namespace
