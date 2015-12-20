
Namespace Excel

	''' <summary>
	''' �I�u�W�F�N�g�̃t�H���g���� (�t�H���g���A�t�H���g �T�C�Y�A�F�Ȃ�) �̑S�̂�\���܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Class FontWrapper
		Inherits AbstractExcelWrapper

		''' <summary>�e��Excel.Range</summary>
		Private _range As RangeWrapper

		''' <summary>Excel.Font</summary>
		Private _font As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="range">�e��range</param>
		''' <param name="font">Excel.font</param>
		''' <remarks></remarks>
		Public Sub New(ByVal range As RangeWrapper, ByVal font As Object)
			MyBase.New(range.ApplicationWrapper)
			_range = range
			_font = font
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_font)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _font
			End Get
		End Property

#End Region
#Region " �v���p�e�B "

		''' <summary>
		''' �O���t�Ŏg�p���镶����̔w�i�̎�ނ�ݒ肵�܂��BXlBackground �񋓌^�̒萔�̂����ꂩ���g�p�ł��܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Background() As XlBackground
			Get
				Return DirectCast(InvokeGetProperty(_font, "Background", Nothing), XlBackground)
			End Get
			Set(ByVal value As XlBackground)
				InvokeSetProperty(_font, "Background", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�t�H���g�������ɂȂ�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Bold() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Bold", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Bold", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �t�H���g�̊�{�F��ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Color() As Object
			Get
				Return InvokeGetProperty(_font, "Color", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Color", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �t�H���g�̐F��ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ColorIndex() As Object
			Get
				Return InvokeGetProperty(_font, "ColorIndex", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "ColorIndex", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �w�肵���I�u�W�F�N�g�̍쐬���̃A�v���P�[�V���������� 32 �r�b�g�̐����l���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_font, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' �t�H���g �X�^�C����ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property FontStyle() As Object
			Get
				Return InvokeGetProperty(_font, "FontStyle", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "FontStyle", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�t�H���g���Α̂ɂȂ�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Italic() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Italic", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Italic", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �I�u�W�F�N�g�̖��O��ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Name() As Object
			Get
				Return DirectCast(InvokeGetProperty(_font, "Name", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Name", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�t�H���g���A�E�g���C�� �t�H���g�ɂȂ�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property OutlineFont() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "OutlineFont", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "OutlineFont", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�t�H���g���e�t���t�H���g�ɁA�܂��͎w�肵���I�u�W�F�N�g���e�t���ɐݒ肳��܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Shadow() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Shadow", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Shadow", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �t�H���g �T�C�Y��ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Size() As Object
			Get
				Return DirectCast(InvokeGetProperty(_font, "Size", Nothing), Object)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_font, "Size", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�t�H���g�Ɏ����������t�����܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Strikethrough() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Strikethrough", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Strikethrough", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�w�肵���t�H���g�����t�������ɂȂ�܂��B����l�� False �ł��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Subscript() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Subscript", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Subscript", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' True ��ݒ肷��ƁA�w�肵���t�H���g����t�������ɂȂ�܂��B����l�� False �ł��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Superscript() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(_font, "Superscript", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(_font, "Superscript", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �t�H���g�ɓK�p���鉺���̎�ނ�ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Underline() As XlUnderlineStyle
			Get
				Return DirectCast(InvokeGetProperty(_font, "Underline", Nothing), XlUnderlineStyle)
			End Get
			Set(ByVal value As XlUnderlineStyle)
				InvokeSetProperty(_font, "Underline", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
