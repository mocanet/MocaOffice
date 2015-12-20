
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Workbooks �̃��b�p�[�N���X
	''' </summary>
	''' <remarks>
	''' Microsoft Excel �Ō��݊J���Ă��邷�ׂĂ� Workbook �I�u�W�F�N�g�̃R���N�V�����ł��B
	''' </remarks>
	Public Class BooksWrapper
		Inherits AbstractExcelWrapper

		''' <summary>Excel.Workbooks �C���X�^���X</summary>
		Private _workbooks As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="xls"></param>
		''' <remarks></remarks>
		Public Sub New(ByVal xls As ExcelWrapper)
			MyBase.New(xls)
			_init(xls)
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_workbooks)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _workbooks
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
		''' Excel.Workbooks.Count
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_workbooks, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' �w�肵���I�u�W�F�N�g�̍쐬���̃A�v���P�[�V���������� 32 �r�b�g�̐����l���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_workbooks, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' �w�肵�������l��ݒ肵�܂��B
		''' </summary>
		''' <param name="Index">�K���w�肵�܂��B�����^ (Integer) �̒l���w�肵�܂��B�����̃C���f�b�N�X�ԍ����w�肵�܂��B</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public Property Item( _
		  <InAttribute()> ByVal Index As Integer _
		 ) As BookWrapper
			Get
				Dim obj As Object
				Dim wrapper As BookWrapper
				obj = InvokeGetProperty(_workbooks, "Item", New Object() {Index})
				wrapper = New BookWrapper(Me.App, obj)
				addXlsObject(wrapper)
				Return wrapper
			End Get
			Set(ByVal value As BookWrapper)
				InvokeSetProperty(_workbooks, "Item", New Object() {value})
			End Set
		End Property

#End Region
#Region " ���\�b�h "

		''' <summary>
		''' ����������
		''' </summary>
		''' <param name="xls">Excel.Application���b�p�[</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal xls As ExcelWrapper)
			' Books�I�u�W�F�N�g�̍쐬
			_workbooks = InvokeGetProperty(xlsApp, "Workbooks", Nothing)
		End Sub

		''' <summary>
		''' �u�b�N�̐V�K�ǉ�
		''' </summary>
		''' <param name="filename">�u�b�N�̃t�@�C����</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add(ByVal filename As String) As BookWrapper
			Dim xlBook As Object
			Dim book As BookWrapper
			xlBook = InvokeMethod(_workbooks, "Add", Nothing)
			book = New BookWrapper(Me.App, xlBook, True)
			book.Filename = filename
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' True ��ݒ肷��ƁA�T�[�o�[����w�肵���u�b�N���`�F�b�N�A�E�g�ł��܂��B�l�̎擾����ѐݒ肪�\�ł��B�u�[���^ (Boolean) �̒l���g�p���܂��B
		''' </summary>
		''' <param name="Filename">�K���w�肵�܂��B������^ (String) �̒l���w�肵�܂��B�`�F�b�N�A�E�g����t�@�C���̖��O���w�肵�܂��B</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function CanCheckOut( _
		 <InAttribute()> ByVal Filename As String _
		) As Boolean
			Return DirectCast(InvokeMethod(_workbooks, "CanCheckOut", New Object() {Filename}), Boolean)
		End Function

		''' <summary>
		''' �w�肵���u�b�N���T�[�o�[���烍�[�J�� �R���s���[�^�ɃR�s�[���āA�ҏW�ł���悤�ɂ��܂��B
		''' </summary>
		''' <param name="Filename">�K���w�肵�܂��B������^ (String) �̒l���w�肵�܂��B�`�F�b�N�A�E�g����t�@�C���̖��O���w�肵�܂��B</param>
		''' <remarks></remarks>
		Public Sub CheckOut( _
		  <InAttribute()> ByVal Filename As String _
		 )
			InvokeMethod(_workbooks, "CheckOut", New Object() {Filename})
		End Sub

		''' <summary>
		''' �w�肵���I�u�W�F�N�g����܂��B
		''' </summary>
		''' <remarks></remarks>
		Public Sub Close()
			InvokeMethod(_workbooks, "Close", Nothing)
		End Sub

		''' <summary>
		''' �R���N�V�����S�̂ł̌J��Ԃ����T�|�[�g���邽�߂ɁA�񋓌^�̒l���擾���܂��B
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator
			Dim bookEnum As IEnumerator
			Dim result As IList(Of BookWrapper)

			result = New List(Of BookWrapper)

			bookEnum = DirectCast(InvokeMethod(_workbooks, "GetEnumerator", Nothing), IEnumerator)
			While bookEnum.MoveNext()
				Dim wrapper As BookWrapper
				wrapper = New BookWrapper(Me.App, bookEnum.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

		''' <summary>
		''' �u�b�N���J���܂��B
		''' </summary>
		''' <param name="Filename">�K���w�肵�܂��B������^ (String) �̒l���w�肵�܂��B�J���u�b�N�̃t�@�C�������w�肵�܂��B</param>
		''' <param name="UpdateLinks">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�t�@�C�����̃����N�̍X�V���@���w�肵�܂��B���̈������ȗ�����ƁA�����N�̍X�V���@�̎w��𑣂��_�C�A���O �{�b�N�X���\������܂��B�ȗ����Ȃ��ꍇ�́A���̂����ꂩ�̒l���w�肵�܂��B</param>
		''' <param name="ReadOnly">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�u�b�N��ǂݎ���p���[�h�ŊJ���ɂ́ATrue ���w�肵�܂��B</param>
		''' <param name="Format">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��BMicrosoft Excel ���e�L�X�g �t�@�C�����J���Ƃ��ɁA���̈����ɍ��ڂ̋�؂蕶�����w�肵�܂��B�w��ł����؂蕶���͎��̂Ƃ���ł��B���̈������ȗ�����ƁA���ݎw�肳��Ă����؂蕶�����g���܂��B</param>
		''' <param name="Password">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�p�X���[�h�ی삳�ꂽ�u�b�N���J�����߂ɕK�v�ȃp�X���[�h���w�肵�܂��B�p�X���[�h���K�v�ȂƂ��ɂ��̈������ȗ�����ƁA�p�X���[�h�̓��͂𑣂��_�C�A���O �{�b�N�X���\������܂��B</param>
		''' <param name="WriteResPassword">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�������ݕی삳�ꂽ�u�b�N�ɏ������݂����邽�߂ɕK�v�ȃp�X���[�h���w�肵�܂��B�p�X���[�h���K�v�ȂƂ��ɂ��̈������ȗ�����ƁA�p�X���[�h�̓��͂𑣂��_�C�A���O �{�b�N�X���\������܂��B</param>
		''' <param name="IgnoreReadOnlyRecommended">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B[�ǂݎ���p�𐄏�����] �`�F�b�N �{�b�N�X���I���ɂ��ĕۑ����ꂽ�u�b�N���J���Ƃ��ł��A�ǂݎ���p�𐄏����郁�b�Z�[�W���\���ɂ���ɂ́ATrue ���w�肵�܂��B</param>
		''' <param name="Origin">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�w�肵���t�@�C�����e�L�X�g �t�@�C���̂Ƃ��ɁA���ꂪ�ǂ̂悤�Ȍ`���̃e�L�X�g �t�@�C�������w�肵�܂��B�R�[�h �y�[�W�� CR/LF �𐳂����ϊ����邽�߂ɕK�v�ł��B�g�p�ł���萔�́AXlPlatform �񋓌^�� xlMacintosh�AxlWindows�AxlMSDOS �̂����ꂩ�ł��B���̈������ȗ�����ƁA���݂̃I�y���[�e�B���O �V�X�e���̌`�����g���܂��B</param>
		''' <param name="Delimiter">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�w�肵���t�@�C�����e�L�X�g �t�@�C���ł���A���� Format �� 6 ���ݒ肳��Ă���Ƃ��ɁA��؂�L���Ƃ��Ďg���������w�肵�܂��B���Ƃ��΁A�^�u�̏ꍇ�� Chr(9)�A�R���}�̏ꍇ�� ","�A�Z�~�R�����̏ꍇ�� ";" ���w�肵�܂��B�C�ӂ̕������w�肷�邱�Ƃ��ł��܂��B��������w�肵���Ƃ��́A�ŏ��̕����������g���܂��B</param>
		''' <param name="Editable">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�w�肵���t�@�C���� Microsoft Excel 4.0 �̃A�h�C���̏ꍇ�A���̈����� True ���w�肷��ƁA�A�h�C�����E�B���h�E�Ƃ��ĕ\�����܂��B���̈����� False ���w�肷�邩�ȗ�����ƁA�A�h�C���͔�\���̏�ԂŊJ����A�E�B���h�E�Ƃ��ĕ\�����邱�Ƃ͂ł��܂���B���̈����́AMicrosoft Excel 5.0 �ȍ~�̃A�h�C���ɂ͓K�p����܂���B�w�肵���t�@�C���� Excel �̃e���v���[�g�̏ꍇ�ATrue ���w�肷��ƁA�w�肳�ꂽ�e���v���[�g��ҏW�p�ɊJ���܂��BFalse ���w�肷��ƁA�w�肳�ꂽ�e���v���[�g����ɂ����A�V�����u�b�N���J���܂��B����l�� False �ł��B</param>
		''' <param name="Notify">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�w�肵���t�@�C�����ǂݎ��/�������݃��[�h�ŊJ���Ȃ��ꍇ�ɁA�t�@�C����ʒm���X�g�ɒǉ�����ɂ́ATrue ���w�肵�܂��B�t�@�C���͓ǂݎ���p���[�h�ŊJ����Ēʒm���X�g�ɒǉ�����A�u�b�N��ҏW�ł����ԂɂȂ������_�ŁA���[�U�[�ɂ��̎|���ʒm����܂��B�t�@�C�����J���Ȃ��ꍇ�ɁA���̂悤�Ȓʒm���s�킸�ɃG���[�𔭐�������ɂ́AFalse ���w�肷�邩�ȗ����܂��B</param>
		''' <param name="Converter">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�t�@�C�����J���Ƃ��ɁA�ŏ��Ɏg���t�@�C�� �R���o�[�^�̃C���f�b�N�X�ԍ����w�肵�܂��B�w�肵���t�@�C�� �R���o�[�^�Ńt�@�C����ϊ��ł��Ȃ��ꍇ�́A���̂��ׂẴt�@�C�� �R���o�[�^�ł̕ϊ������݂��܂��B�w�肷��C���f�b�N�X�ԍ��́AFileConverters �v���p�e�B�Ŏ擾�����t�@�C�� �R���o�[�^�̍s�ԍ��ł��B</param>
		''' <param name="AddToMru">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ��ݒ肷��ƁA�ŋߎg�p�����t�@�C���̈ꗗ�ɂ��̃u�b�N���ǉ�����܂��B����l�� False �ł��B</param>
		''' <param name="Local">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ��ݒ肷��ƁAMicrosoft Excel �Ŏg�p����Ă��錾��Ńt�@�C�����ۑ�����܂� (�R���g���[�� �p�l���̐ݒ���܂�)�BFalse (����l) ��ݒ肷��ƁAVBA (Visual Basic for Applications) �Ŏg�p����Ă��錾��Ńt�@�C�����ۑ�����܂� (�ʏ�̓A�����J�p��ł��B�������A�Â����۔ł� XL5/95 VBA �v���W�F�N�g���� Workbooks.Open �����s���Ă���ꍇ�������܂�)�B</param>
		''' <param name="CorruptLoad">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B�g�p�ł���萔�́AxlNormalLoad�AxlRepairFile�AxlExtractData �̂����ꂩ�ł��B���̈������ȗ������Ƃ��̊���̓���́A�W���̓ǂݍ��ݏ����ƂȂ�̂����ʂł����A2 ��ڈȍ~�̓Z�[�t ���[�h��f�[�^ ���J�o���ƂȂ邱�Ƃ�����܂��B�܂�A�ŏ��͕W���̓ǂݍ��ݏ��������݂܂��B�t�@�C�����J���Ă���r���ŏ�������~�����Ƃ��́A���ɃZ�[�t ���[�h�����݂܂��B�Ăя�������~�����Ƃ��́A���Ƀf�[�^ ���J�o�������݂܂��B</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Open(ByVal Filename As String, _
		   Optional ByVal UpdateLinks As Object = Nothing, _
		   Optional ByVal [ReadOnly] As Object = Nothing, _
		   Optional ByVal Format As Object = Nothing, _
		   Optional ByVal Password As Object = Nothing, _
		   Optional ByVal WriteResPassword As Object = Nothing, _
		   Optional ByVal IgnoreReadOnlyRecommended As Object = Nothing, _
		   Optional ByVal Origin As Object = Nothing, _
		   Optional ByVal Delimiter As Object = Nothing, _
		   Optional ByVal Editable As Object = Nothing, _
		   Optional ByVal Notify As Object = Nothing, _
		   Optional ByVal Converter As Object = Nothing, _
		   Optional ByVal AddToMru As Object = Nothing, _
		   Optional ByVal Local As Object = Nothing, _
		   Optional ByVal CorruptLoad As Object = Nothing) As BookWrapper
			Dim xlBook As Object
			Dim book As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If UpdateLinks IsNot Nothing Then
				argsV.Add(UpdateLinks)
				argsN.Add("UpdateLinks")
			End If
			If [ReadOnly] IsNot Nothing Then
				argsV.Add([ReadOnly])
				argsN.Add("ReadOnly")
			End If
			If Format IsNot Nothing Then
				argsV.Add(Format)
				argsN.Add("Format")
			End If
			If Password IsNot Nothing Then
				argsV.Add(Password)
				argsN.Add("Password")
			End If
			If WriteResPassword IsNot Nothing Then
				argsV.Add(WriteResPassword)
				argsN.Add("WriteResPassword")
			End If
			If IgnoreReadOnlyRecommended IsNot Nothing Then
				argsV.Add(IgnoreReadOnlyRecommended)
				argsN.Add("IgnoreReadOnlyRecommended")
			End If
			If Origin IsNot Nothing Then
				argsV.Add(Origin)
				argsN.Add("Origin")
			End If
			If Delimiter IsNot Nothing Then
				argsV.Add(Delimiter)
				argsN.Add("Delimiter")
			End If
			If Editable IsNot Nothing Then
				argsV.Add(Editable)
				argsN.Add("Editable")
			End If
			If Notify IsNot Nothing Then
				argsV.Add(Notify)
				argsN.Add("Notify")
			End If
			If Converter IsNot Nothing Then
				argsV.Add(Converter)
				argsN.Add("Converter")
			End If
			If AddToMru IsNot Nothing Then
				argsV.Add(AddToMru)
				argsN.Add("AddToMru")
			End If
			If Local IsNot Nothing Then
				argsV.Add(Local)
				argsN.Add("Local")
			End If
			If CorruptLoad IsNot Nothing Then
				argsV.Add(CorruptLoad)
				argsN.Add("CorruptLoad")
			End If

			xlBook = InvokeMethod(_workbooks, "Open", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' �f�[�^�x�[�X��\�� Workbook �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Filename"></param>
		''' <param name="CommandText"></param>
		''' <param name="CommandType"></param>
		''' <param name="BackgroundQuery"></param>
		''' <param name="ImportDataAs"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function OpenDatabase( _
		 <InAttribute()> ByVal Filename As String, _
		 <InAttribute()> Optional ByVal CommandText As Object = Nothing, _
		 <InAttribute()> Optional ByVal CommandType As Object = Nothing, _
		 <InAttribute()> Optional ByVal BackgroundQuery As Object = Nothing, _
		 <InAttribute()> Optional ByVal ImportDataAs As Object = Nothing _
		) As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If CommandText IsNot Nothing Then
				argsV.Add(CommandText)
				argsN.Add("CommandText")
			End If
			If CommandType IsNot Nothing Then
				argsV.Add(CommandType)
				argsN.Add("CommandType")
			End If
			If BackgroundQuery IsNot Nothing Then
				argsV.Add(BackgroundQuery)
				argsN.Add("BackgroundQuery")
			End If
			If ImportDataAs IsNot Nothing Then
				argsV.Add(ImportDataAs)
				argsN.Add("ImportDataAs")
			End If

			Dim xlBook As Object
			Dim book As BookWrapper

			xlBook = InvokeMethod(_workbooks, "OpenDatabase", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' �e�L�X�g �t�@�C���𕪐͂��ēǂݍ��݂܂��B�e�L�X�g �t�@�C���� 1 ���̃V�[�g�Ƃ��āA������܂ސV�����u�b�N���J���܂��B
		''' </summary>
		''' <param name="Filename">�K���w�肵�܂��B������^ (String) �̒l���g�p���܂��B�ǂݍ��܂��e�L�X�g �t�@�C���̖��O���w�肵�܂��B</param>
		''' <param name="Origin">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�e�L�X�g �t�@�C�����쐬���ꂽ�@����w�肵�܂��B�g�p�ł���萔�́A<see cref="XlPlatform"/> �񋓌^�� xlMacintosh�AxlWindows�AxlMSDOS �̂����ꂩ�ł��B���̑��ɁA�ړI�̃R�[�h �y�[�W�̃R�[�h �y�[�W�ԍ���\���������w��ł��܂��B���Ƃ��΁A"1256" �́A�\�[�X �e�L�X�g �t�@�C���̃G���R�[�h�ŃA���r�A�� (Windows) ���w�肵�܂��B���̈������ȗ�����ƁAText Import Wizard �� [���̃t�@�C��] �I�v�V�����̌��݂̐ݒ肪�g�p����܂��B</param>
		''' <param name="StartRow">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B��荞�݊J�n�s���w�肵�܂��B�ŏ��̍s�� 1 �Ƃ��Đ����܂��B����l�� 1 �ł��B </param>
		''' <param name="DataType">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�t�@�C���Ɋ܂܂��f�[�^�̌`�����w�肵�܂��B�g�p�ł���萔�́A<see cref="XlTextParsingType"/> �񋓌^�� xlDelimited �܂��� xlFixedWidth �ł��B���̈������ȗ�����ƁA�t�@�C�����J�����Ƃ��Ƀf�[�^�̌`���������I�Ɍ��肳��܂��B</param>
		''' <param name="TextQualifier">�ȗ��\�ł��B<see cref="XlTextQualifier"/> �񋓌^�̒萔���g�p���܂��B������̈��p�����w�肵�܂��B�g�p�ł���萔�́A���Ɏ��� XlTextQualifier �񋓌^�̒萔�̂����ꂩ�ł��B</param>
		''' <param name="ConsecutiveDelimiter">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�A��������؂蕶���� 1 �����Ƃ��Ĉ����Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="Tab">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� DataType �� xlDelimited ���w�肵�A��؂蕶���Ƀ^�u���g���Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="Semicolon">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� DataType �� xlDelimited ���w�肵�A��؂蕶���ɃZ�~�R���� (;) ���g���Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="Comma">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� DataType �� xlDelimited ���w�肵�A��؂蕶���ɃR���} (,) ���g���Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="Space">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� DataType �� xlDelimited ���w�肵�A��؂蕶���ɃX�y�[�X���g���Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="Other">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� DataType �� xlDelimited ���w�肵�A��؂蕶���� OtherChar �Ŏw�肵���������g���Ƃ��� True ���w�肵�܂��B����l�� False �ł��B</param>
		''' <param name="OtherChar">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� Other �� True �̂Ƃ��́A�K�����̈����ɋ�؂蕶�����w�肵�܂��B�����̕������w�肵���Ƃ��́A�擪�̕�������؂蕶���ƂȂ�A�c��̕����͖�������܂��B</param>
		''' <param name="FieldInfo">�ȗ��\�ł��B<see cref="XlColumnDataType"/> �񋓌^�̒萔���g�p���܂��B�e��̃f�[�^�`���Ɋւ���������z����w�肵�܂��B�f�[�^�`���̉��߂́A���� DataType �̒l�ɂ���ĈقȂ�܂��B�f�[�^����؂�L���ŋ�؂��Ă���ꍇ�́A���̈����� 2 �v�f�z��̔z��ŁA�e 2 �v�f�z��͓���̗�̕ϊ��I�v�V�������w�肵�܂��B1 �Ԗڂ̗v�f�ɂ� 1 ����n�܂��ԍ����w�肵�A2 �Ԗڂ̗v�f�ɂ͗�̃f�[�^�`�������� XlColumnDataType �񋓌^�̒萔���w�肵�܂��B</param>
		''' <param name="TextVisualLayout">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�e�L�X�g�̎��o�I�Ȕz�u���w�肵�܂��B</param>
		''' <param name="DecimalSeparator">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BMicrosoft Excel �Ő��l��F������ꍇ�Ɏg�������_�̋L���ł��B����̓V�X�e���ݒ�ł��B </param>
		''' <param name="ThousandsSeparator">�ȗ��\�ł��B������^ (Object) �̒l���w�肵�܂��B�����̔F���Ɏg�p����錅��؂蕶�����w�肵�܂��B����l�́A�V�X�e���ݒ�ł��B<br/>���܂��܂ȃC���|�[�g�ݒ�Ńe�L�X�g�� Excel �ɃC���|�[�g���錋�ʂ����Ɏ����܂��B���l�̌��ʂ͉E�l�߂ŕ\�����܂��B</param>
		''' <param name="TrailingMinusNumbers">�ȗ��\�ł��B</param>
		''' <param name="Local">�ȗ��\�ł��B</param>
		''' <remarks></remarks>
		Public Sub OpenText( _
		 ByVal Filename As String, _
		 Optional ByVal Origin As Object = Nothing, _
		 Optional ByVal StartRow As Object = Nothing, _
		 Optional ByVal DataType As Object = Nothing, _
		 Optional ByVal TextQualifier As XlTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, _
		 Optional ByVal ConsecutiveDelimiter As Object = Nothing, _
		 Optional ByVal Tab As Object = Nothing, _
		 Optional ByVal Semicolon As Object = Nothing, _
		 Optional ByVal Comma As Object = Nothing, _
		 Optional ByVal Space As Object = Nothing, _
		 Optional ByVal Other As Object = Nothing, _
		 Optional ByVal OtherChar As Object = Nothing, _
		 Optional ByVal FieldInfo As Object = Nothing, _
		 Optional ByVal TextVisualLayout As Object = Nothing, _
		 Optional ByVal DecimalSeparator As Object = Nothing, _
		 Optional ByVal ThousandsSeparator As Object = Nothing, _
		 Optional ByVal TrailingMinusNumbers As Object = Nothing, _
		 Optional ByVal Local As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If Origin IsNot Nothing Then
				argsV.Add(Origin)
				argsN.Add("Origin")
			End If
			If StartRow IsNot Nothing Then
				argsV.Add(StartRow)
				argsN.Add("StartRow")
			End If
			If DataType IsNot Nothing Then
				argsV.Add(DataType)
				argsN.Add("DataType")
			End If
			argsV.Add(TextQualifier)
			argsN.Add("TextQualifier")
			If ConsecutiveDelimiter IsNot Nothing Then
				argsV.Add(ConsecutiveDelimiter)
				argsN.Add("ConsecutiveDelimiter")
			End If
			If Tab IsNot Nothing Then
				argsV.Add(Tab)
				argsN.Add("Tab")
			End If
			If Semicolon IsNot Nothing Then
				argsV.Add(Semicolon)
				argsN.Add("Semicolon")
			End If
			If Comma IsNot Nothing Then
				argsV.Add(Comma)
				argsN.Add("Comma")
			End If
			If Space IsNot Nothing Then
				argsV.Add(Space)
				argsN.Add("Space")
			End If
			If Other IsNot Nothing Then
				argsV.Add(Other)
				argsN.Add("Other")
			End If
			If OtherChar IsNot Nothing Then
				argsV.Add(OtherChar)
				argsN.Add("OtherChar")
			End If
			If FieldInfo IsNot Nothing Then
				argsV.Add(FieldInfo)
				argsN.Add("FieldInfo")
			End If
			If TextVisualLayout IsNot Nothing Then
				argsV.Add(TextVisualLayout)
				argsN.Add("TextVisualLayout")
			End If
			If DecimalSeparator IsNot Nothing Then
				argsV.Add(DecimalSeparator)
				argsN.Add("DecimalSeparator")
			End If
			If ThousandsSeparator IsNot Nothing Then
				argsV.Add(ThousandsSeparator)
				argsN.Add("ThousandsSeparator")
			End If
			If TrailingMinusNumbers IsNot Nothing Then
				argsV.Add(TrailingMinusNumbers)
				argsN.Add("TrailingMinusNumbers")
			End If
			If Local IsNot Nothing Then
				argsV.Add(Local)
				argsN.Add("Local")
			End If

			InvokeMethod(_workbooks, "OpenText", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' XML �f�[�^ �t�@�C�����J���܂��BWorkbook �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Filename">�K���w�肵�܂��B������^ (String) �̒l���w�肵�܂��B�J���t�@�C�������w�肵�܂��B</param>
		''' <param name="Stylesheets">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�K�p���� XSLT (XSL �ϊ�) �X�^�C���V�[�g�������߂��w�肷��P��̒l�܂��͒l�̔z����w�肵�܂��B</param>
		''' <param name="LoadOption">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BExcel �� XML �f�[�^ �t�@�C�����J�����@���w�肵�܂��B�g�p�ł���萔�́A���Ɏ��� <see cref="XlXmlLoadOption"/> �񋓌^�̒萔�̂����ꂩ�ł��B</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function OpenXML( _
		  <InAttribute()> ByVal Filename As String, _
		  <InAttribute()> Optional ByVal Stylesheets As Object = Nothing, _
		  <InAttribute()> Optional ByVal LoadOption As Object = Nothing _
		 ) As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If Stylesheets IsNot Nothing Then
				argsV.Add(Stylesheets)
				argsN.Add("Stylesheets")
			End If
			If LoadOption IsNot Nothing Then
				argsV.Add(LoadOption)
				argsN.Add("LoadOption")
			End If

			Dim xlBook As Object
			Dim book As BookWrapper

			xlBook = InvokeMethod(_workbooks, "OpenXML", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		Friend Function GetMyBook(ByVal name As String) As BookWrapper
			For Each item As Object In myXlsObject()
				If Not TypeOf item Is BookWrapper Then
					Continue For
				End If
				Dim book As BookWrapper
				book = DirectCast(item, BookWrapper)
				If book.Name <> name Then
					Continue For
				End If
				Return book
			Next
			Return Nothing
		End Function

#End Region

	End Class

End Namespace
