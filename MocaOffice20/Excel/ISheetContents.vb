Namespace Excel

	''' <summary>
	''' �V�[�g���e���\������ׂ̃C���^�t�F�[�X
	''' </summary>
	''' <remarks>
	''' ���G�Ȓ��[�݌v�ȂǁA�Z���ɑ΂��ďڍׂɑ��삪�K�v�ȂƂ��Ɏg�p���܂��B<br/>
	''' </remarks>
	Public Interface ISheetContents

		''' <summary>
		''' �u�b�N�ɑ��݂���x�[�X�ƂȂ�V�[�g��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' �o�͑ΏۂƂȂ�u�b�N��ɑ��݂���V�[�g����Ԃ��悤�ɂ���B
		''' �V�K�ɒǉ�����V�[�g�̂Ƃ��͋󕶎���Ԃ��悤�ɂ���B
		''' </remarks>
		Property BaseSheetName() As String

		''' <summary>
		''' �ۑ����Ɏg�p����V�[�g��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' �ۑ����Ɏg�p����V�[�g����Ԃ��悤�ɂ���B
		''' <see cref="BaseSheetName"/> �Ŏw�肵���V�[�g���Ɠ���̏ꍇ�͋󕶎���Ԃ��B
		''' </remarks>
		ReadOnly Property SaveSheetName() As String

		''' <summary>
		''' �o�͓��e���Z���֐ݒ肷��
		''' </summary>
		''' <param name="sheet">�Y������V�[�g����C���X�^���X</param>
		''' <remarks>
		''' </remarks>
		Sub WriteContents(ByVal sheet As SheetWrapper)

	End Interface

End Namespace
