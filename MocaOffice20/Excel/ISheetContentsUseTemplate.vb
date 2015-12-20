
Namespace Excel

	''' <summary>
	''' �ꗗ�`���̃e���v���[�g�V�[�g���g�p�����Ƃ��ɁA
	''' �f�[�^�𑽎����z��ɕϊ����l��ݒ肷���@�̃C���^�t�F�[�X
	''' </summary>
	''' <remarks>
	''' </remarks>
	Public Interface ISheetContentsUseTemplate
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
		''' �o�͂���f�[�^�z����쐬����
		''' </summary>
		''' <remarks>
		''' �o�͂���f�[�^��K�v�Ȑ��̍s��ɊY�����鑽�����z����쐬���߂�l�Ƃ��ĕԂ��܂��B<br/>
		''' �쐬����z��́A<see cref="DataCount"/>�s�C<see cref="ColumnLength"/>��Ƃ��Ă��������B<br/>
		''' �쐬���ꂽ�z����A<see cref="StartRow"/>�s�F<see cref="StartCol"/>������ɐݒ肳��܂��B
		''' </remarks>
		Function MakeArrayData() As Array

	End Interface

End Namespace
