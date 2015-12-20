
Imports System.IO
Imports System.Text

Namespace Excel

	''' <summary>
	''' �ꗗ�`���̃e���v���[�g�V�[�g���g�p�����Ƃ��ɁA
	''' �f�[�^����xCSV�t�@�C���֏o�͂��ACSV�t�@�C����Ǎ����Excel�֓\��t�����@
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContentsUseTemplateMakeCsv
		Inherits SheetContentsUseTemplate

		''' <summary>CSV�t�@�C����</summary>
		Private _csvFilename As String

		''' <summary>CSV�ǂݍ��ݎ��̗�t�H�[�}�b�g</summary>
		Private _openFormat As Array

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			MyBase.New(sheetContents, sheet)

			_openFormat = Nothing
			_csvFilename = String.Empty
		End Sub

#End Region

#Region " �v���p�e�B "

		''' <summary>CSV�o�͂���ۂ̃e���|�����t�@�C����</summary>
		Private ReadOnly Property CsvTempFilename() As String
			Get
				If _csvFilename.Length = 0 Then
					_csvFilename = Path.Combine(Path.GetTempPath, "~" & MyBase.contents.SaveSheetName & Format(Now(), "_yyyyMMdd_hhmmss") & ".txt")
				End If
				Return _csvFilename
			End Get
		End Property

		''' <summary>CSV��Ǎ��ނƂ��̗�t�H�[�}�b�g</summary>
		Private ReadOnly Property CsvOpenFormat() As Array
			Get
				If _openFormat Is Nothing Then
					_openFormat = Array.CreateInstance(GetType(Object), _cType.ColumnLength)
					For ii As Integer = 0 To _openFormat.GetUpperBound(0)
						_openFormat.SetValue(New Integer() {ii + 1, XlColumnDataType.xlTextFormat}, ii)
					Next ii
				End If
				Return _openFormat
			End Get
		End Property

#End Region

		''' <summary>
		''' �V�[�g�R���e���c�𓖃N���X�Ŏg�p����N���X�փL���X�g����
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _cType() As ISheetContentsUseTemplateMakeCsv
			Return DirectCast(MyBase.contents, ISheetContentsUseTemplateMakeCsv)
		End Function

		''' <summary>
		''' CSV�t�@�C����Excel�ɂĊJ�����̃t�H�[�}�b�g���w�肷��
		''' </summary>
		''' <remarks>
		''' CSV�t�@�C����Excel�ɂĊJ�����̃t�H�[�}�b�g���w�肷��ꍇ�͎w�肵�Ă��������B<br/>
		''' �f�t�H���g�ł́u��ʁv�̃t�H�[�}�b�g�ɂēǂݍ��݂܂��B
		''' </remarks>
		Private Sub _setCsvOpenFormat()
			Dim value As XlColumnDataType
			For ii As Integer = 0 To UBound(Me.CsvOpenFormat)
				value = _cType.SetCsvOpenFormat(ii)
				If value = XlColumnDataType.xlNone Then
					Continue For
				End If
				Me.CsvOpenFormat.SetValue(New Integer() {ii + 1, value}, ii)
			Next ii
		End Sub

		''' <summary>
		''' �o�͓��e���Z���֐ݒ肷��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Protected Overrides Sub writeContentsTemplate()
			Dim xlBookTmp As BookWrapper

			xlBookTmp = Nothing

			Try
				If _cType.DataCount <= 0 Then
					Exit Sub
				End If

				' ��Ɨp�t�@�C���̂��|��
				_clearCsvTempFile()
				' �f�[�^���e�L�X�g�t�@�C���Ńe���|�����o��
				_csvTempWrite()

				' �擪�s���f�[�^�����R�s�[
				rowCopy(_cType.DataCount _
				 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow + 1, _cType.StartCol) _
				 , MyBase.sheet.Cells(_cType.StartRow + _cType.DataCount - 1, _cType.StartCol))

				' Temp�t�@�C���̓Ǎ�
				_setCsvOpenFormat()
				MyBase.sheet.App.Workbooks.OpenText(Filename:=CsvTempFilename, DataType:=XlTextParsingType.xlDelimited, TextQualifier:=XlTextQualifier.xlTextQualifierDoubleQuote, Comma:=True, FieldInfo:=CsvOpenFormat)
				xlBookTmp = MyBase.sheet.App.ActiveWorkbook

				' Temp�t�@�C�����e�𒠕[�փR�s�[���y�[�X�g
				_writeContents(xlBookTmp)

			Finally
				' Temp�t�@�C�������
				If xlBookTmp IsNot Nothing Then
					xlBookTmp.Close(False)
				End If
				' ��Ɨp�t�@�C���̂��|��
				_clearCsvTempFile()
			End Try
		End Sub

		''' <summary>
		''' ��Ɨp�t�@�C���̂��|��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _clearCsvTempFile()
			If Not File.Exists(CsvTempFilename) Then
				Exit Sub
			End If

			'���ɑ��݂��Ă���ꍇ�͍폜����B
			File.Delete(CsvTempFilename)
		End Sub

		''' <summary>
		''' CSV�`���Ńe���|�����t�@�C���쐬
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _csvTempWrite()
			If _cType.DataCount = 0 Then
				Exit Sub
			End If

			Using file As StreamWriter = New StreamWriter(CsvTempFilename, False, Encoding.GetEncoding("Shift_JIS"))
				Try
					'�e���v�t�@�C���֏o��
					_cType.CsvWrite(file)
				Catch ex As Exception
					Throw New ExcelException(MyBase.sheet.App, ex)
				End Try
			End Using
		End Sub

		''' <summary>
		''' ���X�g���o��
		''' </summary>
		''' <param name="xlBookTmp">Excel�u�b�N�C���X�^���X�iCSV�p�j</param>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents(ByVal xlBookTmp As BookWrapper)
			Dim tmpSheet As SheetWrapper
			Dim selection As SelectionWrapper
			Dim rowIndex As Integer

			Dim range1 As RangeWrapper
			Dim range2 As RangeWrapper

			'���l�^�R�s�[
			xlBookTmp.Activate()
			tmpSheet = xlBookTmp.Worksheets(1)
			tmpSheet.Select()
			rowIndex = tmpSheet.Range("A65536").End(XlDirection.xlUp).Row
			range1 = tmpSheet.Cells(1, 1)
			range2 = tmpSheet.Cells(rowIndex, _cType.ColumnLength)
			tmpSheet.Range(range1, range2).Select()
			selection = MyBase.sheet.App.Selection
			selection.Copy()

			'�\��t��
			MyBase.sheet.Book.Activate()
			MyBase.sheet.Select()
			range1 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			range2 = MyBase.sheet.Cells(_cType.StartRow, _cType.StartCol)
			MyBase.sheet.Range(range1, range2).Select()
			selection = MyBase.sheet.App.Selection
			selection.PasteSpecial(Paste:=XlPasteType.xlPasteValues, Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=True, Transpose:=False)
		End Sub

	End Class

End Namespace
