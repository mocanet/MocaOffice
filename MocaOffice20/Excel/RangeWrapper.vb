
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' �Z���A�s�A��A1 �ȏ�̃Z���͈͂��܂ޑI��͈́A�܂��� 3-D �͈͂�\���܂�
	''' </summary>
	''' <remarks></remarks>
	Public Class RangeWrapper
		Inherits AbstractExcelWrapper

		''' <summary>�e�̃V�[�g</summary>
		Private _sheet As SheetWrapper

		''' <summary>Excel.Range</summary>
		Private _range As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="sheet">�e�̃V�[�g</param>
		''' <param name="range">Excel.Range</param>
		''' <remarks></remarks>
		Public Sub New(ByVal sheet As SheetWrapper, ByVal range As Object)
			MyBase.New(sheet.ApplicationWrapper)
			_sheet = sheet
			_range = range
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_range)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _range
			End Get
		End Property

#End Region
#Region " �v���p�e�B "

		''' <summary>
		''' Excel.Application �̃��b�p�[
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property App() As ExcelWrapper
			Get
				Return DirectCast(MyBase.xlsWrapper, ExcelWrapper)
			End Get
		End Property

		''' <summary>
		''' �e�̃V�[�g
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Sheet() As SheetWrapper
			Get
				Return _sheet
			End Get
		End Property

		''' <summary>
		''' �Z����
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' Excel.Range.Count
		''' </remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_range, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' �w�肵���Z���͈͂̍ŏ��̗̈�̐擪�s�̔ԍ����擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' Excel.Range.Count
		''' </remarks>
		Public ReadOnly Property Row() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_range, "Row", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' �w�肵���Z���͈͂̒l
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Value() As Object
			' 2003 �ȍ~�̎��͉��L���g�p�\
			'''' <summary>
			'''' �w�肵���Z���͈͂̒l
			'''' </summary>
			'''' <param name="RangeValueDataType">�w�肵�� Range �I�u�W�F�N�g�̃f�[�^�^</param>
			'''' <value></value>
			'''' <returns></returns>
			'''' <remarks></remarks>
			'Public Property Value(Optional ByVal RangeValueDataType As XlRangeValueDataType = XlRangeValueDataType.xlRangeValueDefault) As Object
			'	Get
			'		Return InvokeGetProperty(_range, "Value", New Object() {RangeValueDataType})
			'	End Get
			'	Set(ByVal value As Object)
			'		InvokeSetProperty(_range, "Value", New Object() {RangeValueDataType, value})
			'	End Set
			'End Property
			Get
				Return InvokeGetProperty(_range, "Value", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "Value", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �w�肵���Z���͈͂Ɋ܂܂�� 1 �s�܂��͕����̍s�S�̂�\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property EntireRow() As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "EntireRow", Nothing)

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �Z�����܂܂��̈�̏I�[�̃Z����\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Direction"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' End + �����L�[ (���A���A���A�� �̂����ꂩ) �ɑ������܂��B
		''' </remarks>
		Public ReadOnly Property [End](ByVal Direction As XlDirection) As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "End", New Object() {Direction})

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �I�u�W�F�N�g�̃t�H���g���� (�t�H���g���A�t�H���g �T�C�Y�A�F�Ȃ�) �̑S�̂�\���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Font() As FontWrapper
			Get
				Dim val As Object
				Dim wrap As FontWrapper

				val = InvokeGetProperty(_range, "Font", New Object() {})

				wrap = New FontWrapper(Me, val)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �|�C���g�P�ʂ̃Z���͈͂̕��ł��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Width() As Object
			Get
				Return InvokeGetProperty(_range, "Width", Nothing)
			End Get
		End Property

		''' <summary>
		''' �w�肵���Z���͈͓��̂��ׂĂ̗�̕���ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColumnWidth() As Object
			Get
				Return InvokeGetProperty(_range, "ColumnWidth", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "ColumnWidth", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �w�肵���Z���͈̗͂��\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Columns() As RangeWrapper
			Get
				Dim range As Object
				Dim wrap As RangeWrapper

				range = InvokeGetProperty(_range, "Columns", Nothing)

				wrap = New RangeWrapper(_sheet, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �X�^�C���܂��̓Z���͈� (�����t�������̈ꕔ�Ƃ��Ē�`���ꂽ�͈͂��܂�) �̌r����\�� Borders �R���N�V�������擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property Borders() As BordersWrapper
			Get
				Dim range As Object
				Dim wrap As BordersWrapper

				range = InvokeGetProperty(_range, "Borders", Nothing)

				wrap = New BordersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �X�^�C���܂��̓Z���͈� (�����t�������̈ꕔ�Ƃ��Ē�`���ꂽ�͈͂��܂�) �̌r����\�� Borders �R���N�V�������擾���܂��B
		''' </summary>
		''' <param name="index">�r�������ʂ���l</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property Borders(ByVal index As XlBordersIndex) As BordersWrapper
			Get
				Dim range As Object
				Dim wrap As BordersWrapper

				range = InvokeGetProperty(_range, "Borders", New Object() {index})

				wrap = New BordersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �I�u�W�F�N�g�̕�������̕����͈̔͂�\�� Characters �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Start">���\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�擾���镶����͈͂̍ŏ��̕������w�肵�܂��B���̈����� 1 ���w�肷�邩�A�ȗ�����ƁA�擪��������n�܂镶����͈͂��擾���܂��B</param>
		''' <param name="Length"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Characters( _
		  <InAttribute()> Optional ByVal Start As Object = Nothing, _
		  <InAttribute()> Optional ByVal Length As Object = Nothing _
		 ) As CharactersWrapper
			Get
				Dim range As Object
				Dim wrap As CharactersWrapper

				range = InvokeGetProperty(_range, "Characters", New Object() {Start, Length})

				wrap = New CharactersWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' �w�肵���I�u�W�F�N�g�̐��������̔z�u��ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property HorizontalAlignment() As Object
			Get
				Return InvokeGetProperty(_range, "HorizontalAlignment", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_range, "HorizontalAlignment", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �w�肵���I�u�W�F�N�g�̓�����\�� Interior �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Interior() As InteriorWrapper
			Get
				Dim range As Object
				Dim wrap As InteriorWrapper

				range = InvokeGetProperty(_range, "Interior", Nothing)

				wrap = New InteriorWrapper(Me, range)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

#End Region

		''' <summary>
		''' �����W�w�肵���Z����I����Ԃɂ���
		''' </summary>
		''' <remarks></remarks>
		Public Sub [Select]()
			InvokeMethod(_range, "Select", Nothing)
		End Sub

		''' <summary>
		''' �I�u�W�F�N�g���R�s�[���܂�
		''' </summary>
		''' <param name="destination">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�R�s�[��̃Z���͈͂��w�肵�܂��B���̈������ȗ�����ƁA�N���b�v�{�[�h�ɃR�s�[���܂��B</param>
		''' <remarks></remarks>
		Public Sub Copy(Optional ByVal destination As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If destination IsNot Nothing Then
				argsV.Add(destination)
				argsN.Add("Destination")
			End If

			InvokeMethod(_range, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' �����W�w�肵���Z�����폜����
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_range, "Delete", Nothing)
		End Sub

		''' <summary>
		''' ���e���N���A���܂�
		''' </summary>
		''' <remarks></remarks>
		Public Sub ClearContents()
			InvokeMethod(_range, "ClearContents", Nothing)
		End Sub

		''' <summary>
		''' �N���b�v�{�[�h�ɂ��� Range �I�u�W�F�N�g���A�w�肵���Z���͈͂ɓ\��t���܂��B
		''' </summary>
		''' <param name="Paste">�ȗ��\�ł��B<see cref="XlPasteType" /> �񋓌^�̒萔���w�肵�܂��B�Z���͈͂̒��œ\��t���镔�����w�肵�܂��B</param>
		''' <param name="Operation">�ȗ��\�ł��B<see cref="XlPasteSpecialOperation" /> �񋓌^�̒l���w�肵�܂��B�\��t���̑�����w�肵�܂��B</param>
		''' <param name="SkipBlanks">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ���w�肷��ƁA�N���b�v�{�[�h�Ɋ܂܂��󔒂̃Z����ΏۃZ���͈͂ɓ\��t���܂���B����l�� False �ł��B</param>
		''' <param name="Transpose">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ���w�肷��ƁA�\��t����Ƃ��ɃZ���͈͂̍s�Ɨ�����ւ��܂��B����l�� False �ł��B</param>
		''' <remarks></remarks>
		Public Sub PasteSpecial( _
		 Optional ByVal Paste As XlPasteType = XlPasteType.xlPasteAll, _
		 Optional ByVal Operation As XlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone, _
		 Optional ByVal SkipBlanks As Object = Nothing, _
		 Optional ByVal Transpose As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Paste)
			argsN.Add("Paste")

			argsV.Add(Operation)
			argsN.Add("Operation")

			If SkipBlanks IsNot Nothing Then
				argsV.Add(SkipBlanks)
				argsN.Add("SkipBlanks")
			End If
			If SkipBlanks IsNot Nothing Then
				argsV.Add(SkipBlanks)
				argsN.Add("SkipBlanks")
			End If
			If Transpose IsNot Nothing Then
				argsV.Add(Transpose)
				argsN.Add("Transpose")
			End If

			InvokeMethod(_range, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

	End Class

End Namespace
