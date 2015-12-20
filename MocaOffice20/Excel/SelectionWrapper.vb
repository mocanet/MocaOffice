
Namespace Excel

	''' <summary>
	''' �A�N�e�B�u�ȃE�B���h�E�őI������Ă���I�u�W�F�N�g
	''' </summary>
	''' <remarks></remarks>
	Public Class SelectionWrapper
		Inherits AbstractExcelWrapper

		''' <summary>�e��Excel.Application �̃��b�p�[</summary>
		Private _xls As ExcelWrapper
		''' <summary>Excel.Application Selection �I�u�W�F�N�g</summary>
		Private _selection As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="xls">Excel.Application���b�p�[</param>
		''' <param name="selection">Excel.Range �ȂǑI������Ă���I�u�W�F�N�g</param>
		''' <remarks></remarks>
		Public Sub New(ByVal xls As ExcelWrapper, ByVal selection As Object)
			MyBase.New(xls.ApplicationWrapper)
			_xls = xls
			_selection = selection
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_selection)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _selection
			End Get
		End Property

#End Region

		''' <summary>
		''' �I�u�W�F�N�g���R�s�[���܂�
		''' </summary>
		''' <remarks></remarks>
		Public Sub Copy()
			If _selection Is Nothing Then
				Exit Sub
			End If
			InvokeMethod(_selection, "Copy", Nothing)
		End Sub

		''' <summary>
		''' �I������Ă���I�u�W�F�N�g���C���T�[�g���܂�
		''' </summary>
		''' <remarks></remarks>
		Public Sub Insert()
			If _selection Is Nothing Then
				Exit Sub
			End If
			InvokeMethod(_selection, "Insert", Nothing)
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
			If Transpose IsNot Nothing Then
				argsV.Add(Transpose)
				argsN.Add("Transpose")
			End If

			InvokeMethod(_selection, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		'''' <summary>
		'''' �w�肵���`���ŁA�N���b�v�{�[�h�̓��e���V�[�g�ɓ\��t���܂��B���̃A�v���P�[�V��������f�[�^��\��t������A����̌`���Ńf�[�^��\��t����ꍇ�Ɏg�p���܂��B
		'''' </summary>
		'''' <param name="Format">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�N���b�v�{�[�h�̃f�[�^�̌`���𕶎���Ŏw�肵�܂��B</param>
		'''' <param name="Link">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���̃f�[�^�Ɠ\��t�����f�[�^�̊ԂɃ����N��ݒ肷��ɂ́ATrue ���w�肵�܂��B���̃f�[�^�������N�ɓK���Ȃ��f�[�^�ł���ꍇ��A���̃f�[�^���쐬�����A�v���P�[�V�����������N���T�|�[�g���Ȃ��ꍇ�A���̈����͖�������܂��B����l�� False �ł��B</param>
		'''' <param name="DisplayAsIcon">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�\��t�����f�[�^���A�C�R���Ƃ��ĕ\������ɂ́ATrue ���w�肵�܂��B����l�� False �ł��B</param>
		'''' <param name="IconFileName">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) ���w�肵�܂��BDisplayAsIcon �� True �̏ꍇ�Ɏg�p����A�C�R�����܂ރt�@�C���̖��O���w�肵�܂��B</param>
		'''' <param name="IconIndex">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�A�C�R���̃t�@�C�����̃A�C�R���̃C���f�b�N�X�ԍ����w�肵�܂��B</param>
		'''' <param name="IconLabel">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�A�C�R���̃��x���̕�������w�肵�܂��B</param>
		'''' <param name="NoHTMLFormatting">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BHTML ���珑���ݒ�A�n�C�p�[�����N�A����уC���[�W�����ׂč폜����ɂ́ATrue ���w�肵�܂��BHTML �����̂܂ܓ\��t����ɂ́AFalse ���w�肵�܂��B����l�� False �ł��B</param>
		'''' <remarks>
		'''' ���̃��\�b�h���g�p����O�ɓ\��t����̃Z���͈͂�I������K�v������܂��B<br/>
		'''' ���̃��\�b�h���g�p����ƁA�N���b�v�{�[�h�̓��e�ɂ���Ă͑I��͈͂��ύX�����ꍇ������܂��B
		'''' </remarks>
		'Public Sub PasteSpecial( _
		' Optional ByVal Format As Object = Nothing, _
		' Optional ByVal Link As Object = Nothing, _
		' Optional ByVal DisplayAsIcon As Object = Nothing, _
		' Optional ByVal IconFileName As Object = Nothing, _
		' Optional ByVal IconIndex As Object = Nothing, _
		' Optional ByVal IconLabel As Object = Nothing, _
		' Optional ByVal NoHTMLFormatting As Object = Nothing)
		'	Dim argsV As ArrayList
		'	Dim argsN As ArrayList

		'	argsV = New ArrayList()
		'	argsN = New ArrayList()

		'	If Format Is Nothing Then
		'		argsV.Add(Format)
		'		argsN.Add("Format")
		'	End If
		'	If Link Is Nothing Then
		'		argsV.Add(Link)
		'		argsN.Add("Link")
		'	End If
		'	If DisplayAsIcon Is Nothing Then
		'		argsV.Add(DisplayAsIcon)
		'		argsN.Add("DisplayAsIcon")
		'	End If
		'	If IconFileName Is Nothing Then
		'		argsV.Add(IconFileName)
		'		argsN.Add("IconFileName")
		'	End If
		'	If IconIndex Is Nothing Then
		'		argsV.Add(IconIndex)
		'		argsN.Add("IconIndex")
		'	End If
		'	If IconLabel Is Nothing Then
		'		argsV.Add(IconLabel)
		'		argsN.Add("IconLabel")
		'	End If
		'	If NoHTMLFormatting Is Nothing Then
		'		argsV.Add(NoHTMLFormatting)
		'		argsN.Add("NoHTMLFormatting")
		'	End If

		'	InvokeMethod(_selection, "PasteSpecial", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		'End Sub

	End Class

End Namespace
