
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Worksheet �̃��b�p�[�N���X
	''' </summary>
	''' <remarks></remarks>
	Public Class SheetWrapper
		Inherits AbstractExcelWrapper

		''' <summary>�e��Excel.Workbook �̃��b�p�[</summary>
		Private _book As BookWrapper

		''' <summary>Excel.Worksheet</summary>
		Private _sheet As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <remarks>
		''' �V�[�g��V�K�Œǉ�����Ƃ��Ɏg��
		''' </remarks>
		Private Sub New(ByVal book As BookWrapper)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = InvokeGetProperty(_book.Worksheets.OrigianlInstance, "Add", Nothing)
		End Sub

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <param name="sheetName">�J���V�[�g��</param>
		''' <remarks>
		''' �V�[�g�����w�肵�ĊJ���Ƃ��Ɏg���B
		''' </remarks>
		Public Sub New(ByVal book As BookWrapper, ByVal sheetName As String)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = InvokeGetProperty(_book.Worksheets.OrigianlInstance, "Item", New Object() {sheetName})
		End Sub

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <param name="xlSheet">Excel.Worksheet</param>
		''' <remarks>
		''' ���ɊJ�����V�[�g���Ǘ�����Ƃ��Ɏg���B
		''' </remarks>
		Public Sub New(ByVal book As BookWrapper, ByVal xlSheet As Object)
			MyBase.New(book.ApplicationWrapper)
			_init(book)

			_sheet = xlSheet
		End Sub

#End Region

#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_sheet)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _sheet
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
		''' �e�̃u�b�N
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Book() As BookWrapper
			Get
				Return _book
			End Get
		End Property

		''' <summary>
		''' �V�[�g��
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Name() As String
			Get
				Return DirectCast(InvokeGetProperty(_sheet, "Name", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_sheet, "Name", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �P��̃Z���܂��̓Z���͈͂�\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Cell1"></param>
		''' <param name="Cell2"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Range(ByVal Cell1 As RangeWrapper, Optional ByVal Cell2 As RangeWrapper = Nothing) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				If Cell2 Is Nothing Then
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1.OrigianlInstance})
				Else
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1.OrigianlInstance, Cell2.OrigianlInstance})
				End If
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' �P��̃Z���܂��̓Z���͈͂�\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <param name="Cell1"></param>
		''' <param name="Cell2"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Range(ByVal Cell1 As String, Optional ByVal Cell2 As String = Nothing) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				If Cell2 Is Nothing Then
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1})
				Else
					rangeBuf = InvokeGetProperty(_sheet, "Range", New Object() {Cell1, Cell2})
				End If
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' �A�N�e�B�u�ȃ��[�N�V�[�g�ɂ��邷�ׂẴZ����\�� Range �I�u�W�F�N�g���擾���܂��B�A�N�e�B�u�ȕ��������[�N�V�[�g�łȂ��ꍇ�A���̃v���p�e�B�͎��s���܂��B
		''' </summary>
		''' <param name="row"></param>
		''' <param name="col"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Cells(ByVal row As Integer, ByVal col As Integer) As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				rangeBuf = InvokeGetProperty(_sheet, "Cells", New Object() {row, col})
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' �V�[�g���폜����Ă��邩
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property IsDeleted() As Boolean
			Get
				Return (_sheet Is Nothing)
			End Get
		End Property

		''' <summary>
		''' ���[�N�V�[�g��̂��ׂĂ̗��\�� Range �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Columns() As RangeWrapper
			Get
				Dim rangeBuf As Object
				Dim rangeWrap As RangeWrapper
				rangeBuf = InvokeGetProperty(_sheet, "Columns", New Object() {})
				rangeWrap = New RangeWrapper(Me, rangeBuf)
				addXlsObject(rangeWrap)
				Return rangeWrap
			End Get
		End Property

		''' <summary>
		''' ���[�N�V�[�g�̂��ׂẴy�[�W�ݒ���܂� <see cref="PageSetupWrapper"/>  ���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overridable ReadOnly Property PageSetup() As PageSetupWrapper
			Get
				Dim obj As Object
				Dim wrap As PageSetupWrapper
				obj = InvokeGetProperty(_sheet, "PageSetup", New Object() {})
				wrap = New PageSetupWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ���[�N�V�[�g��̂��ׂĂ̐}�`��\�� <see cref="ShapesWrapper"/> �I�u�W�F�N�g���擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Overridable ReadOnly Property Shapes() As ShapesWrapper
			Get
				Dim obj As Object
				Dim wrap As ShapesWrapper
				obj = InvokeGetProperty(_sheet, "Shapes", New Object() {})
				wrap = New ShapesWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ���[�N�V�[�g��̐��������̉��y�[�W��\�� <see cref="HPageBreaksWrapper"/> �R���N�V�������擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property HPageBreaks() As HPageBreaksWrapper
			Get
				Dim obj As Object
				Dim wrap As HPageBreaksWrapper
				obj = InvokeGetProperty(_sheet, "HPageBreaks", New Object() {})
				wrap = New HPageBreaksWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

		''' <summary>
		''' ���[�N�V�[�g��̐��������̉��y�[�W��\�� <see cref="VPageBreaksWrapper"/> �R���N�V�������擾���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property VPageBreaks() As VPageBreaksWrapper
			Get
				Dim obj As Object
				Dim wrap As VPageBreaksWrapper
				obj = InvokeGetProperty(_sheet, "VPageBreaks", New Object() {})
				wrap = New VPageBreaksWrapper(Me, obj)
				addXlsObject(wrap)
				Return wrap
			End Get
		End Property

#End Region

		''' <summary>
		''' ������
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal book As BookWrapper)
			_book = book
		End Sub

		''' <summary>
		''' �V�[�g���A�N�e�B�u�ɂ���
		''' </summary>
		''' <remarks></remarks>
		Public Sub Activate()
			InvokeMethod(_sheet, "Activate", Nothing)
		End Sub

		''' <summary>
		''' �V�[�g�S�̂�I���ɂ���
		''' </summary>
		''' <remarks></remarks>
		Public Sub [Select]()
			InvokeMethod(_sheet, "Select", Nothing)
		End Sub

		''' <summary>
		''' �V�[�g���폜����
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_sheet, "Delete", Nothing)
			MyDispose()
		End Sub

		''' <summary>
		''' �N���b�v�{�[�h����\�t��
		''' </summary>
		''' <remarks></remarks>
		Public Sub Paste()
			InvokeMethod(_sheet, "Paste", Nothing)
		End Sub

		''' <summary>
		''' �V�[�g���u�b�N���̑��̏ꏊ�ɃR�s�[���܂��B
		''' </summary>
		''' <param name="Before">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�R�s�[����V�[�g�����̃V�[�g�̒��O�̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BAfter ���w�肳��Ă���ꍇ�ABefore �͎w��ł��܂���B</param>
		''' <param name="After">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�R�s�[����V�[�g�����̃V�[�g�̒���̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BBefore ���w�肳��Ă���ꍇ�AAfter �͎w��ł��܂���B</param>
		''' <remarks>���� Before �ƈ��� After �����ɏȗ������ꍇ�́A�V�K�u�b�N�������I�ɍ쐬����A�V�[�g�͂��̃u�b�N���ɑ}������܂��B</remarks>
		Public Sub Copy(<InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, <InAttribute()> Optional ByVal After As SheetWrapper = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Before IsNot Nothing Then
				argsV.Add(Before.OrigianlInstance)
				argsN.Add("Before")
			End If
			If After IsNot Nothing Then
				argsV.Add(After.OrigianlInstance)
				argsN.Add("After")
			End If

			InvokeMethod(_sheet, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

	End Class

End Namespace
