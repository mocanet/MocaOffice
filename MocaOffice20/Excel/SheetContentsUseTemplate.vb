
Namespace Excel

	''' <summary>
	''' �ꗗ�`���̃e���v���[�g�V�[�g���g�p�����Ƃ��ɁA
	''' �f�[�^�𑽎����z��ɕϊ����l��ݒ肷���@�B
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContentsUseTemplate
		Inherits SheetContents

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			MyBase.New(sheetContents, sheet)
		End Sub

#End Region

		''' <summary>
		''' �V�[�g�R���e���c�𓖃N���X�Ŏg�p����N���X�փL���X�g����
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _cType() As ISheetContentsUseTemplate
			Return DirectCast(MyBase.contents, ISheetContentsUseTemplate)
		End Function

		''' <summary>
		''' �o�͓��e���Z���֐ݒ肷��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Public Overrides Sub WriteContents()
			MyBase.WriteContents()

			writeContentsTemplate()
		End Sub

		''' <summary>
		''' �ꗗ�����̐ݒ�
		''' </summary>
		''' <remarks></remarks>
		Protected Overridable Sub writeContentsTemplate()
			If _cType.DataCount <= 0 Then
				Exit Sub
			End If

			' �擪�s���f�[�^�����R�s�[
			rowCopy(_cType.DataCount _
			 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow + 1, _cType.StartCol) _
			 , MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount - 1, _cType.StartCol))

			' 
			_writeContents()
		End Sub

		''' <summary>
		''' �e���v���[�g�ƂȂ�s���w�肳�ꂽ���e�ŃR�s�[����
		''' </summary>
		''' <remarks>
		''' </remarks>
		Protected Sub rowCopy(ByVal dataCount As Integer, ByVal rangeF1 As RangeWrapper, ByVal rangeF2 As RangeWrapper, ByVal rangeT1 As RangeWrapper, ByVal rangeT2 As RangeWrapper)
			' �f�[�^���P���̏ꍇ�̓R�s�[�s�v
			If dataCount <= 1 Then
				Exit Sub
			End If

			' �R�s�[���̍s���R�s�[
			MyBase.sheet.Range(rangeF1, rangeF2).EntireRow.Select()
			MyBase.sheet.App.Selection.Copy()

			' �R�s�[��̍s��I��
			MyBase.sheet.Range(rangeT1, rangeT2).EntireRow.Select()

			' �o�͗\��̍s�����C���T�[�g
			MyBase.sheet.App.Selection.Insert()
		End Sub

		''' <summary>
		''' ���X�g���o��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents()
			Dim range1 As RangeWrapper
			Dim range2 As RangeWrapper

			range1 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			range2 = MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount, _cType.StartCol + _cType.ColumnLength - 1)
			MyBase.sheet.Range(range1, range2).Value = _cType.MakeArrayData()
		End Sub

	End Class

End Namespace
