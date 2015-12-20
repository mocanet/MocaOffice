
Imports System.IO

Namespace Excel

	''' <summary>
	''' �u�b�N����p�̃C�x���g����
	''' </summary>
	''' <remarks></remarks>
	Public Class BookContentsEventArgs
		Inherits EventArgs

		''' <summary>���쒆�̃V�[�g��</summary>
		Public SheetName As String
	End Class

	''' <summary>
	''' �i��
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	''' <remarks></remarks>
	Public Delegate Sub PerformStep(ByVal sender As AbstractBookContents, ByVal e As BookContentsEventArgs)

	''' <summary>
	''' Excel�u�b�N�𑀍삷��ׂ̋��ʏ�����ێ����钊�ۃN���X
	''' </summary>
	''' <remarks>
	''' �����ۃN���X�ł́AExcel�o�͂���ۂ̂����܂�̃u�b�N��������Ɏ������Ă���܂��B<br/>
	''' ���N���X���p�����T�u�N���X�����āA�u�b�N�t�@�C�����i<see cref="AbstractBookContents.SaveFilename"/>�A<see cref="AbstractBookContents.TemplateFilename"/>�j��ݒ肵�A
	''' <see cref="AbstractBookContents.Add"/> �ɂĊe�V�[�g�o�̓N���X��ǉ����邱�ƂŁA��r�I�ȒP�� Excel �o�͋@�\�������o����悤�ɂȂ��Ă܂��B<br/>
	''' �V�[�g�o�̓N���X�� <seealso cref="ISheetContents"/>, <seealso cref="ISheetContentsUseTemplate"/>, <seealso cref="ISheetContentsUseTemplateMakeCsv"/>
	''' ���������Ă��������B<br/>
	''' </remarks>
	Public MustInherit Class AbstractBookContents

		'Public Event PerformStep As PerformStep

		''' <summary>Excel�A�v���P�[�V����</summary>
		Private _app As ExcelWrapper

		''' <summary>�t�@�C����</summary>
		Private _saveFilename As String

		''' <summary>�e���v���[�g�ƂȂ�Excel�t�@�C����</summary>
		Protected myTemplateFilename As String

		''' <summary>Excel����ʕ\�����邩�ǂ����𔻒肷��ϐ�</summary>
		Private _display As Boolean
		''' <summary>Excel��ۑ����邩�ǂ����𔻒肷��ϐ�</summary>
		Private _save As Boolean
		''' <summary>Excel��������邩�ǂ����𔻒肷��ϐ�</summary>
		Private _print As Boolean

		Private _performStep As PerformStep

		''' <summary>�V�[�g���e</summary>
		Protected sheetContents As IList(Of ISheetContents)

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �f�t�H���g�R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
			sheetContents = New List(Of ISheetContents)
		End Sub

#End Region

#Region " �v���p�e�B "

		''' <summary>
		''' Excel�A�v���P�[�V����
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected ReadOnly Property App() As ExcelWrapper
			Get
				Return _app
			End Get
		End Property

		''' <summary>�t�@�C����</summary>
		Public Property SaveFilename() As String
			Get
				Return _saveFilename
			End Get
			Set(ByVal value As String)
				_saveFilename = value
			End Set
		End Property

		''' <summary>�e���v���[�g�ƂȂ�Excel�t�@�C����</summary>
		Public Property TemplateFilename() As String
			Get
				Return myTemplateFilename
			End Get
			Set(ByVal value As String)
				myTemplateFilename = value
			End Set
		End Property

		''' <summary>Excel����ʕ\�����邩�ǂ����𔻒肷��ϐ�</summary>
		Public Property Display() As Boolean
			Get
				Return _display
			End Get
			Set(ByVal value As Boolean)
				_display = value
			End Set
		End Property

		''' <summary>Excel��ۑ����邩�ǂ����𔻒肷��ϐ�</summary>
		Public Property Save() As Boolean
			Get
				Return _save
			End Get
			Set(ByVal value As Boolean)
				_save = value
			End Set
		End Property

		''' <summary>Excel��������邩�ǂ����𔻒肷��ϐ�</summary>
		Public Property Print() As Boolean
			Get
				Return _print
			End Get
			Set(ByVal value As Boolean)
				_print = value
			End Set
		End Property

#End Region

		''' <summary>
		''' �V�[�g��ǉ�����
		''' </summary>
		''' <param name="sheet">�V�[�g�R���e���c</param>
		''' <remarks></remarks>
		Public Sub Add(ByVal sheet As ISheetContents)
			sheetContents.Add(sheet)
		End Sub

		''' <summary>
		''' �R���e���c�o��
		''' </summary>
		''' <param name="performStep">�i���i�v���O���X�o�[�Ȃǁj��K�v�Ƃ���Ƃ��́A<see cref="PerformStep"/> �f���Q�[�g���w�肵�Ă��������B</param>
		''' <remarks>
		''' </remarks>
		Public Sub Write(Optional ByVal performStep As PerformStep = Nothing)
			_performStep = performStep
			Using xls As ExcelWrapper = New ExcelWrapper
				Try
					_app = xls
					_writeContents()
				Catch ex As Exception
					xls.Dispose()
					Throw ex
				End Try
			End Using
			''DoSomething���\�b�h��ʂ̃X���b�h�Ŏ��s����
			''Thread�I�u�W�F�N�g���쐬����
			'Dim t As New System.Threading.Thread( _
			' New System.Threading.ThreadStart( _
			' AddressOf DoSomething))
			''�X���b�h���J�n����
			't.Start()
			''t.Join()
		End Sub

		Private Sub DoSomething()
			Using xls As ExcelWrapper = New ExcelWrapper
				_app = xls
				_writeContents()
			End Using
		End Sub

		''' <summary>
		''' �o�̓��W�b�N
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeContents()
			Dim cbValue As String

			' �V�[�g�w�肪�����Ƃ��͏����I��
			If sheetContents.Count = 0 Then
				Throw New ExcelException(Me.App, "�V�[�g������w�肳��Ă��܂���B")
			End If

			' �N���b�v�{�[�h�̑ޔ�
			cbValue = My.Computer.Clipboard.GetText()

			Try
				' Excel�e���v���[�g�t�@�C���̑��݃`�F�b�N
				_xlsTemplateFileExists()

				' Excel�o��
				_writeExcel()
			Finally
				Try
					' �N���b�v�{�[�h�̕���
					My.Computer.Clipboard.SetText(cbValue)
				Catch ex As Exception
				End Try
			End Try
		End Sub

		''' <summary>
		''' Excel�e���v���[�g�t�@�C���̑��݃`�F�b�N
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _xlsTemplateFileExists()
			If TemplateFilename.Length = 0 Then
				Exit Sub
			End If

			' Excel�e���v���[�g�t�@�C���̑��݃`�F�b�N
			If Not File.Exists(TemplateFilename) Then
				Throw New ExcelException(Me.App, String.Format("Excel�e���v���[�g�t�@�C�������݂��܂���B({0})", TemplateFilename))
			End If
		End Sub

		''' <summary>
		''' �o��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Private Sub _writeExcel()
			Dim xlBook As BookWrapper
			Dim xlSheet As SheetWrapper
			Dim e As BookContentsEventArgs

			e = New BookContentsEventArgs()
			e.SheetName = String.Empty

			Try
				Me.App.Visible = False
				Me.App.DisplayAlerts = False
				Me.App.Interactive = False
				Me.App.ScreenUpdating = False

				' �u�b�N���擾
				If TemplateFilename.Length = 0 Then
					' �V�K�u�b�N�쐬
					xlBook = Me.App.Workbooks.Add(Me.SaveFilename)
				Else
					' �e���v���[�g�t�@�C���̓Ǎ�
					xlBook = Me.App.Workbooks.Open(TemplateFilename)
				End If

				' �e�V�[�g����
				For ii As Integer = 0 To sheetContents.Count - 1
					Dim contents As ISheetContents
					Dim contentsWriter As SheetContents

					contents = sheetContents(ii)

					If contents.BaseSheetName = String.Empty Then
						If contents.SaveSheetName = String.Empty Then
							xlSheet = xlBook.Worksheets(ii + 1)
							contents.BaseSheetName = xlSheet.Name
						Else
							xlSheet = xlBook.Worksheets.Add(contents.SaveSheetName)
						End If
					Else
						xlSheet = xlBook.Worksheets(contents.BaseSheetName)
					End If
					xlSheet.Activate()
					e.SheetName = xlSheet.Name

					' �o��
					contentsWriter = SheetContentsFactory.Create(contents, xlSheet)
					Dim tim As Stopwatch = New Stopwatch()
					tim.Start()
					contentsWriter.WriteContents()
					tim.Stop()
					_mylog.DebugFormat("[{0}] Write Time {1}", IIf(contents.SaveSheetName.Length = 0, contents.BaseSheetName, contents.SaveSheetName), tim.ElapsedMilliseconds)

					' �z�[���|�W�V�����Ɉړ�
					If Not xlSheet.IsDeleted Then
						xlSheet.Range("A1").Select()
					End If

					_runPerformStep(e)
				Next

				' �I������
				_endWrite(xlBook)
			Catch chex As ExcelException
				Throw chex
			Catch ex As Exception
				Throw New ExcelException(Me.App, ex)
			End Try
		End Sub

		''' <summary>
		''' Excel�I������
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' </remarks>
		Private Function _endWrite(ByVal xlBook As BookWrapper) As Boolean
			Dim sheet As SheetWrapper

			' �z�[���|�W�V�����Ɉړ�
			sheet = Me.App.ActiveWorkbook.Worksheets(1)
			sheet.Activate()
			sheet.Range("A1").Select()

			' �����ۑ�
			If Save Then
				If Not _autoSave(xlBook) Then
					Exit Function
				End If
			End If
			' �������
			If Print Then
				If Not _autoPrint() Then
					Exit Function
				End If
			End If
			' ��ʕ\��
			If Display Then
				If Not _autoDisplay() Then
					Exit Function
				End If
			End If
		End Function

		''' <summary>
		''' �����ۑ�
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' <see cref="SaveFilename" /> �Ɏw�肳��Ă��閼�O�ŕۑ����܂��B
		''' </remarks>
		Private Function _autoSave(ByVal xlBook As BookWrapper) As Boolean
			' �ۑ��t�@�C�������w�肵�Ă��Ȃ��Ƃ��͖���
			If _saveFilename.Length = 0 Then
				Exit Function
			End If

			Try
				_app.DisplayAlerts = False
				xlBook.SaveAs(_saveFilename)
				_app.DisplayAlerts = True
				Return True
			Catch ex As ExcelException
				Throw ex
			Catch ex As Exception
				Throw New ExcelException(_app, ex, "Excel �����ۑ����ɃG���[���������܂����B")
			End Try
		End Function

		''' <summary>
		''' �������
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' </remarks>
		Private Function _autoPrint() As Boolean
			_app.ActiveWorkbook.PrintOut()	' ���
			Return True
		End Function

		''' <summary>
		''' ��ʂ֕\������
		''' </summary>
		''' <returns></returns>
		''' <remarks>
		''' ���L�̍��ڂ� <c>Ture</c> �ɐݒ肵�܂��B<br/>
		''' <list>
		''' <item><description><see cref="ExcelWrapper.ScreenUpdating"/></description></item>
		''' <item><description><see cref="ExcelWrapper.Interactive"/></description></item>
		''' <item><description><see cref="ExcelWrapper.DisplayAlerts"/></description></item>
		''' <item><description><see cref="ExcelWrapper.Visible"/></description></item>
		''' </list>
		''' </remarks>
		Private Function _autoDisplay() As Boolean
			_app.ScreenUpdating = True
			_app.Interactive = True
			_app.DisplayAlerts = True
			_app.Visible = True
			Return True
		End Function

		''' <summary>
		''' �i����
		''' </summary>
		''' <param name="e"></param>
		''' <remarks></remarks>
		Private Sub _runPerformStep(ByVal e As BookContentsEventArgs)
			Try
				'RaiseEvent PerformStep(Me, e)
				If _performStep Is Nothing Then
					Exit Sub
				End If

				_performStep(Me, e)
				'_performStep.Target.GetType.InvokeMember(_performStep.GetType().Name, Reflection.BindingFlags.InvokeMethod, Nothing, _performStep.Target, New Object() {Me, e})
			Catch ex As Exception
				_mylog.ErrorFormat(ex.Message)
			End Try
		End Sub

	End Class

End Namespace
