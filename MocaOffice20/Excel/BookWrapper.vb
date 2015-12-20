
Imports System.Reflection

Namespace Excel

    ''' <summary>
    ''' Excel.Workbook �̃��b�p�[�N���X
    ''' </summary>
    ''' <remarks>
    ''' Microsoft Excel �̃u�b�N��\���܂��B
    ''' </remarks>
    Public Class BookWrapper
        Inherits AbstractExcelWrapper

        ''' <summary>Excel.Workbook</summary>
        Private _book As Object

        ''' <summary>Excel.Sheets</summary>
        Private _sheets As SheetsWrapper
        ''' <summary>Excel.Worksheets</summary>
        Private _worksheets As WorksheetsWrapper

        ''' <summary>�t�@�C����</summary>
        Private _filename As String

        ''' <summary>�V�K�t�@�C��</summary>
        Private _new As Boolean
        ''' <summary>�������t���O</summary>
        Private _close As Boolean

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

        ''' <summary>
        ''' �R���X�g���N�^
        ''' </summary>
        ''' <param name="xls">Excel.Application���b�p�[</param>
        ''' <param name="workbook">Excel.Workbook</param>
        ''' <param name="newBook">�V�K�u�b�N���ǂ���</param>
        ''' <remarks>
        ''' ���ɊJ���Ă���u�b�N�𑀍삷��Ƃ��Ɏg�p���܂��B
        ''' </remarks>
        Friend Sub New(ByVal xls As ExcelWrapper, ByVal workbook As Object, Optional ByVal newBook As Boolean = False)
            MyBase.New(xls)
            _init(xls)

            _book = workbook

            _new = newBook

            If _new Then
                Me.Saved = False
            Else
                _filename = FullName
            End If
        End Sub

#End Region
#Region " Overrides "

        ''' <summary>
        ''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub MyDispose()
            ReleaseExcelObject(_book)
        End Sub

        ''' <summary>
        ''' �擾���� Excel �C���X�^���X
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Overrides ReadOnly Property OrigianlInstance() As Object
            Get
                Return _book
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
        ''' �t�@�C����
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Filename() As String
            Get
                Return _filename
            End Get
            Set(ByVal value As String)
                _filename = value
            End Set
        End Property

        ''' <summary>
        ''' �ۑ��ς�
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Saved() As Boolean
            Get
                Return DirectCast(InvokeGetProperty(_book, "Saved", Nothing), Boolean)
            End Get
            Set(ByVal value As Boolean)
                InvokeSetProperty(_book, "Saved", New Object() {value})
            End Set
        End Property

        ''' <summary>
        ''' �u�b�N���̑S�V�[�g
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Sheets() As SheetsWrapper
            Get
                If _sheets Is Nothing Then
                    _sheets = New SheetsWrapper(Me)
                    addXlsObject(_sheets)
                End If
                Return _sheets
            End Get
        End Property

        ''' <summary>
        ''' �u�b�N���̑S���[�N�V�[�g
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Worksheets() As WorksheetsWrapper
            Get
                If _worksheets Is Nothing Then
                    _worksheets = New WorksheetsWrapper(Me)
                    addXlsObject(_worksheets)
                End If
                Return _worksheets
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
                xl = InvokeGetProperty(_book, "ActiveSheet", Nothing)
                If xl IsNot Nothing Then
                    sheet = New SheetWrapper(Me, xl)
                    addXlsObject(sheet)
                End If

                Return sheet
            End Get
        End Property

        ''' <summary>
        ''' �I�u�W�F�N�g�̖��O��������������擾���܂��B���O�ɂ̓f�B�X�N��̃p�X���܂܂�܂��B�l�̎擾�̂݉\�ł��B������^ (String) �̒l���g�p���܂��B 
        ''' </summary>
        ''' <value></value>
        ''' <returns>
        ''' ���̃v���p�e�B���g�p����ƁA<see cref="Path"/> �v���p�e�B�A���݂̃t�@�C�� �V�X�e���̋�؂蕶���A<see cref="Name"/> �v���p�e�B�𑱂��ċL�q����̂Ɠ������ʂ������܂��B
        ''' </returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FullName() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "FullName", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' �A�v���P�[�V�����܂ł̐�΃p�X���擾���܂��B���̃p�X�ł́A�Ō�̋�؂蕶���ƃA�v���P�[�V���������Ȃ���܂��B�l�̎擾�̂݉\�ł��B������^ (String) �̒l���g�p���܂��B
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Path() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "Path", Nothing), String)
            End Get
        End Property

        ''' <summary>
        ''' �I�u�W�F�N�g�̖��O���擾���܂��B�l�̎擾�̂݉\�ł��B������^ (String) �̒l���g�p���܂��B
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name() As String
            Get
                Return DirectCast(InvokeGetProperty(_book, "Name", Nothing), String)
            End Get
        End Property

#End Region
#Region " ���\�b�h "

        ''' <summary>
        ''' ����������
        ''' </summary>
        ''' <param name="xls">Excel.Application���b�p�[</param>
        ''' <remarks></remarks>
        Private Sub _init(ByVal xls As ExcelWrapper)
            _new = False
            _close = False
        End Sub

        ''' <summary>
        ''' �V�[�g���A�N�e�B�u�ɂ���
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Activate()
            InvokeMethod(_book, "Activate", Nothing)
        End Sub

        ''' <summary>
        ''' �u�b�N�����
        ''' </summary>
        ''' <remarks>
        ''' ����Ɠ����ɕۑ����܂��B
        ''' </remarks>
        Public Sub Close()
            Close(True)
        End Sub

        ''' <summary>
        ''' �u�b�N�����
        ''' </summary>
        ''' <param name="save">�ۑ����邩�ǂ���</param>
        ''' <remarks></remarks>
        Public Sub Close(ByVal save As Boolean)
            If _book Is Nothing Then
                Exit Sub
            End If
            If _close Then
                Exit Sub
            End If

            If save Then
                Me.Save()
            End If

            InvokeMethod(_book, "Close", New Object() {save})
            _close = True
        End Sub

        '''' <summary>
        '''' �u�b�N��ۑ�
        '''' </summary>
        '''' <remarks></remarks>
        'Friend Sub SaveTitle()
        '	If Not _new Then
        '		Exit Sub
        '	End If
        '	Dim flg As Boolean
        '	flg = Saved
        '	Saved = False
        '	SaveAs(_filename)
        '	Saved = flg
        '	System.IO.File.Delete(_filename)
        'End Sub

        ''' <summary>
        ''' �u�b�N��ۑ�
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Save()
            If Saved Then
                Exit Sub
            End If

            If _new Then
                SaveAs(_filename)
                Exit Sub
            End If

            InvokeMethod(_book, "Save", Nothing)
        End Sub

        ''' <summary>
        ''' �u�b�N�Ƀt�@�C������t���ĕۑ�
        ''' </summary>
        ''' <param name="filename">�t�@�C����</param>
        ''' <remarks></remarks>
        Public Sub SaveAs(ByVal filename As String)
            If Saved Then
                Exit Sub
            End If

            InvokeMethod(_book, "SaveAs", New Object() {filename})
        End Sub

        ''' <summary>
        ''' �u�b�N�����
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub PrintOut()
            InvokeMethod(_book, "PrintOut", Nothing)
        End Sub

        ''' <summary>
        ''' �������s�}�N���̎��s
        ''' </summary>
        ''' <param name="Which">���s���鎩�����s�}�N�����w��<br/>
        ''' XlRunAutoMacro�Œ�`���ꂽ�萔���g�p���܂��B</param>
        ''' <remarks></remarks>
        Public Sub RunAutoMacros(ByVal which As XlRunAutoMacro)
            InvokeMethod(_book, "RunAutoMacros", New Object() {which})
        End Sub

        ''' <summary>
        ''' �w�肵���`���̃t�@�C���ɃG�N�X�|�[�g
        ''' </summary>
        ''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff198122.aspx</remarks>
        Public Sub ExportAsFixedFormat(ByVal xlType As FixedFormatType,
                                       Optional ByVal filename As String = Nothing,
                                       Optional ByVal quality As FixedFormatQuality = FixedFormatQuality.QualityStandard,
                                       Optional ByVal includeDocProperties As Boolean = False,
                                       Optional ByVal ignorePrintAreas As Boolean = False,
                                       Optional ByVal [from] As Integer = 0,
                                       Optional ByVal [to] As Integer = 0,
                                       Optional ByVal openAfterPublish As Boolean = False,
                                       Optional ByVal fixedFormatExtClassPtr As Object = Nothing)
            Dim argsV As New List(Of Object)
            Dim argsN As New List(Of String)

            argsV.Add(xlType)
            argsN.Add("Type")
            If filename IsNot Nothing Then
                argsV.Add(filename)
                argsN.Add("Filename")
            End If
            argsV.Add(quality)
            argsN.Add("Quality")
            argsV.Add(includeDocProperties)
            argsN.Add("IncludeDocProperties")
            argsV.Add(ignorePrintAreas)
            argsN.Add("IgnorePrintAreas")
            If [from] > 0 Then
                argsV.Add([from])
                argsN.Add("From")
            End If
            If [to] > 0 Then
                argsV.Add([to])
                argsN.Add("To")
            End If
            argsV.Add(openAfterPublish)
            argsN.Add("OpenAfterPublish")
            If fixedFormatExtClassPtr IsNot Nothing Then
                argsV.Add(fixedFormatExtClassPtr)
                argsN.Add("FixedFormatExtClassPtr")
            End If

            'args = New Object() {xlType, filename, quality, includeDocProperties, ignorePrintAreas, [from], [to], openAfterPublish, fixedFormatExtClassPtr}

            InvokeMethod(_book, "ExportAsFixedFormat", argsV.ToArray, argsN.ToArray)
        End Sub

#End Region

    End Class

End Namespace
