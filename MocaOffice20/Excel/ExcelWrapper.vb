
Imports System.Reflection

Namespace Excel

	''' <summary>
	''' Excel.Application �̃��b�p�[�N���X
	''' </summary>
	''' <remarks>
	''' Excel �����C�g�o�C���f�B���O�ɂāi�Q�Ɛݒ肷�邱�ƂȂ��j����o���܂��B<br/>
	''' Excel �𑀍삷���ŃC���X�^���X�����ꂽ�I�u�W�F�N�g�͓��N���X���J�����邱�ƂőS�ĊJ������悤�ɂȂ��Ă��܂��B<br/>
	''' Excel ���I�����邩�ǂ����́A<see cref="ExcelWrapper.Visible"/> �ɂ���Ď����Ŕ��f���܂��B<br/>
	''' �g�p����Ƃ��́A<c>Using</c>��𗘗p���Ă��������B<br/>
	''' <br/>
	''' <example>
	''' <code lang="vb">
	''' Using xls As ExcelWrapper = New ExcelWrapper()
	''' 	Try
	''' 		Dim book As BookWrapper
	''' 
	''' 		xls.Visible = False
	''' 		xls.DisplayAlerts = False
	''' 		xls.Interactive = False
	''' 		xls.ScreenUpdating = False
	''' 
	''' 		book = xls.Workbooks.Open(Path.Combine(My.Application.Info.DirectoryPath, "test.xls"))
	''' 		book.Save()
	''' 		book.Close(False)
	''' 	Catch ex As Exception
	''' 		xls.Dispose()
	''' 	End Try
	''' End Using
	''' </code>
	''' </example>
	''' </remarks>
	Public Class ExcelWrapper
		Inherits AbstractExcelWrapper

		''' <summary>Excel.Workbooks �C���X�^���X</summary>
		Private _workbooks As BooksWrapper


		Private _quit As Boolean

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �f�t�H���g�R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
			MyBase.New()

			Try

				' Excel�N���X ProgID �Ɋ֘A�t�����Ă���^���擾
				typApplication = Type.GetTypeFromProgID("Excel.Application")

				' Excel�̌^������Ɋ֘A�t�����Ă��Ȃ��B(Excel�����݂��Ȃ�)
				If typApplication Is Nothing Then
					_mylog.Error("Excel�����݂��܂���B�C���X�g�[������Ă��邩�m�F���Ă��������B")
					Throw New NotSupportedException("Excel�����݂��܂���B�C���X�g�[������Ă��邩�m�F���Ă��������B")
				End If

				' �e�평����
				Me.ApplicationWrapper = Me

				' Excel�̃C���X�^���X���쐬���܂��B
				xlsApp = Activator.CreateInstance(typApplication)

				_mylog.DebugFormat("{0} Version:{1} ProductCode:{2}", Me.Name, Me.Version, Me.ProductCode)

				' Books�I�u�W�F�N�g�̍쐬
				_workbooks = New BooksWrapper(Me)
				addXlsObject(_workbooks)
			Catch ex As ExcelException
				Me.MyDispose()
				Throw ex
			Catch ex As Exception
				Me.MyDispose()
				Throw New ExcelException(Me, ex, "ExcelWrapper �̃C���X�^���X�������ɃG���[���������܂����B")
			End Try
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			If _quit Then
				Exit Sub
			End If
			Quit(Not Me.Visible)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return xlsApp
			End Get
		End Property

#End Region

#Region " �v���p�e�B "

		''' <summary>
		''' �A�v���P�[�V������
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Name() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Name", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' Excel�o�[�W����
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Version() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Version", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' �v���_�N�g�R�[�h
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ProductCode() As String
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "ProductCode", Nothing), String)
			End Get
		End Property

		''' <summary>
		''' �A�N�e�B�u�ȃE�B���h�E�őI������Ă���I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Selection() As SelectionWrapper
			Get
				Dim xl As Object
				Dim sel As SelectionWrapper

				sel = Nothing
				xl = InvokeGetProperty(xlsApp, "Selection", Nothing)
				If xl IsNot Nothing Then
					sel = New SelectionWrapper(Me, xl)
					addXlsObject(sel)
				End If

				Return sel
			End Get
		End Property

		''' <summary>
		''' ��ʕ\���L��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Visible() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Visible", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "Visible", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �m�F�_�C�A���O�\���L��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property DisplayAlerts() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "DisplayAlerts", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "DisplayAlerts", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ��ʍX�V��L��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property ScreenUpdating() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "ScreenUpdating", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "ScreenUpdating", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' ���[�U�[����̊��L��
		''' </summary>
		''' <value></value>
		''' <returns>
		''' True ��ݒ肷��ƁAMicrosoft Excel ���Θb���[�h�ɂȂ�܂��B���̃v���p�e�B�͒ʏ� True �ł��B���̃v���p�e�B�� False ��ݒ肷��ƁA�L�[�{�[�h����у}�E�X����̓��͂��󂯕t���Ȃ��Ȃ�܂��B�������A�R�[�h�ɂ���ĕ\�����ꂽ�_�C�A���O �{�b�N�X�ւ̓��͉͂\�ł��B���͂ł��Ȃ���Ԃɂ��Ă����ƁA�R�[�h���g�p���� Microsoft Excel �̃I�u�W�F�N�g���ړ�������A�N�e�B�u�ɂ����肵�Ă���Ƃ��ɁA���[�U�[����̊���h�����Ƃ��ł��܂��B<br/>
		''' ���̃v���p�e�B�� False ��ݒ肵���ꍇ�́ATrue �ɖ߂��̂�Y��Ȃ��悤�ɂ��Ă��������B�R�[�h�̎��s���I�����Ă��A���̃v���p�e�B�͎����I�� True �ɖ߂�܂���B
		''' </returns>
		''' <remarks></remarks>
		Public Property Interactive() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "Interactive", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "Interactive", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �C�x���g�����̗L��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' True �̏ꍇ�A�w�肳�ꂽ�I�u�W�F�N�g�ɑ΂��ăC�x���g���������܂��B�l�̎擾����ѐݒ肪�\�ł��B�u�[���^ (Boolean) �̒l���g�p���܂��B
		''' </remarks>
		Public Property EnableEvents() As Boolean
			Get
				Return DirectCast(InvokeGetProperty(xlsApp, "EnableEvents", Nothing), Boolean)
			End Get
			Set(ByVal value As Boolean)
				InvokeSetProperty(xlsApp, "EnableEvents", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' Excel.Workbooks
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Workbooks() As BooksWrapper
			Get
				Return _workbooks
			End Get
		End Property

		''' <summary>
		''' ���݃A�N�e�B�u�ȃu�b�N
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ActiveWorkbook() As BookWrapper
			Get
				Dim book As Object
				Dim bookWrap As BookWrapper
				bookWrap = Nothing
				book = InvokeGetProperty(xlsApp, "ActiveWorkbook", Nothing)
				If book IsNot Nothing Then
					Dim nm As String
					nm = DirectCast(InvokeGetProperty(book, "Name", Nothing), String)
					bookWrap = _workbooks.GetMyBook(nm)
					If bookWrap Is Nothing Then
						bookWrap = New BookWrapper(Me, book)
					End If
				End If
				Return bookWrap
			End Get
		End Property

		''' <summary>
		''' ���݃A�N�e�B�u�ȃV�[�g
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property ActiveSheet() As SheetWrapper
			Get
				Dim xl As Object
				Dim sheet As SheetWrapper
				sheet = Nothing
				xl = InvokeGetProperty(xlsApp, "ActiveSheet", Nothing)
				If xl IsNot Nothing Then
					sheet = New SheetWrapper(ActiveWorkbook, xl)
					addXlsObject(sheet)
				End If

				Return sheet
			End Get
		End Property

		''' <summary>
		''' Microsoft Excel �ŐV�K�u�b�N�Ɏ����I�ɑ}�������V�[�g�̐���ݒ肵�܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property SheetsInNewWorkbook() As Integer
			Get
				Return CInt(InvokeGetProperty(xlsApp, "SheetsInNewWorkbook", Nothing))
			End Get
			Set(ByVal value As Integer)
				InvokeSetProperty(xlsApp, "SheetsInNewWorkbook", New Object() {value})
			End Set
		End Property

#End Region

		''' <summary>
		''' Excel�I��
		''' </summary>
		''' <param name="windowClose">��ʂ���邩�ǂ���</param>
		''' <remarks></remarks>
		Public Sub Quit(Optional ByVal windowClose As Boolean = False)
			ReleaseExcelObject(xlsApp, windowClose)
			_quit = True
		End Sub

	End Class

End Namespace
