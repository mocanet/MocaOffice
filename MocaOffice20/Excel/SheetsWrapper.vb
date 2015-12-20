
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Sheets �̃��b�p�[�N���X
	''' </summary>
	''' <remarks></remarks>
	Public Class SheetsWrapper
		Inherits AbstractExcelWrapper

		''' <summary>�e��Excel.Workbook �̃��b�p�[</summary>
		Private _book As BookWrapper

		''' <summary>Excel.Worksheets</summary>
		Private _sheets As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="book">�e�u�b�N</param>
		''' <remarks></remarks>
		Public Sub New(ByVal book As BookWrapper)
			MyBase.New(book.ApplicationWrapper)
			_init(book)
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_sheets)
		End Sub

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _sheets
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
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_sheets, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' ��ƒ��̃u�b�N�̂��ׂẴV�[�g��\�� Sheets �R���N�V��������w�肳�ꂽ�V�[�g���擾���܂��B 
		''' </summary>
		''' <param name="name">�V�[�g��</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(ByVal name As String) As SheetWrapper
			Get
				Dim sheet As Object
				Dim wrapper As SheetWrapper
				sheet = InvokeGetProperty(_sheets, "Item", New Object() {name})
				wrapper = New SheetWrapper(_book, sheet)
				addXlsObject(wrapper)
				Return wrapper
			End Get
		End Property

		''' <summary>
		''' ��ƒ��̃u�b�N�̂��ׂẴV�[�g��\�� Sheets �R���N�V��������w�肳�ꂽ�V�[�g���擾���܂��B 
		''' </summary>
		''' <param name="index">�V�[�g�ԍ�</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public ReadOnly Property Item(ByVal index As Integer) As SheetWrapper
			Get
				Dim sheet As Object
				Dim wrapper As SheetWrapper
				sheet = InvokeGetProperty(_sheets, "Item", New Object() {index})
				wrapper = New SheetWrapper(_book, sheet)
				addXlsObject(wrapper)
				Return wrapper
			End Get
		End Property

		''' <summary>
		''' �I�u�W�F�N�g��\�����邩�A��\���ɂ��邩�����肵�܂��B�l�̎擾����ѐݒ肪�\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���g�p���܂��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Visible() As Object
			Get
				Return InvokeGetProperty(_sheets, "Visible", Nothing)
			End Get
			Set(ByVal value As Object)
				InvokeSetProperty(_sheets, "Visible", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' �V�[�g�̐��������̉��y�[�W��\�� VPageBreaks �R���N�V�������擾���܂��B�l�̎擾�̂݉\�ł��B
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property VPageBreaks() As VPageBreaksWrapper
			Get
				Return DirectCast(InvokeGetProperty(_sheets, "VPageBreaks", Nothing), VPageBreaksWrapper)
			End Get
		End Property

#End Region
#Region " ���\�b�h "

		''' <summary>
		''' ������
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal book As BookWrapper)
			_book = book
			' Worksheets�I�u�W�F�N�g�̍쐬
			_sheets = InvokeGetProperty(_book.OrigianlInstance, "Sheets", Nothing)
		End Sub

		''' <summary>
		''' �V�������[�N�V�[�g�A�O���t �V�[�g�A�}�N�� �V�[�g�̂����ꂩ���쐬���܂��B�쐬�������[�N�V�[�g�̓A�N�e�B�u�ɂȂ�܂��B
		''' </summary>
		''' <param name="Before">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�����V�[�g�����̃V�[�g�̒��O�̈ʒu�ɒǉ�����Ƃ��ɁA���̃V�[�g���w�肵�܂��B</param>
		''' <param name="After">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�����V�[�g�����̃V�[�g�̒���̈ʒu�ɒǉ�����Ƃ��ɁA���̃V�[�g���w�肵�܂��B</param>
		''' <param name="Count">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�ǉ�����V�[�g�̐����w�肵�܂��B����l�� 1 �ł��B</param>
		''' <param name="Type">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�[�g�̎�ނ��w�肵�܂��B�g�p�ł���萔�́AXlSheetType �񋓌^�� xlWorksheet�AxlChart�AxlExcel4MacroSheet�AxlExcel4IntlMacroSheet �̂����ꂩ�ł��B�����̃e���v���[�g����ɂ����V�[�g��}������ꍇ�́A���̃e���v���[�g�ւ̃p�X���w�肵�܂��B����l�� xlWorksheet �ł��B</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add( _
		  <InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal After As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal Count As Object = 1, _
		  <InAttribute()> Optional ByVal Type As XlSheetType = XlSheetType.xlWorksheet _
		 ) As SheetWrapper
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
			If Count IsNot Nothing Then
				argsV.Add(Count)
				argsN.Add("Count")
			End If
			argsV.Add(Type)
			argsN.Add("Type")

			Dim obj As Object
			Dim wrapper As SheetWrapper
			obj = InvokeMethod(_sheets, "Add", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			wrapper = New SheetWrapper(_book, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' �V�[�g���u�b�N���̑��̏ꏊ�ɃR�s�[���܂��B
		''' </summary>
		''' <param name="Before">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�R�s�[����V�[�g�����̃V�[�g�̒��O�̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BAfter ���w�肳��Ă���ꍇ�ABefore �͎w��ł��܂���B</param>
		''' <param name="After">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�R�s�[����V�[�g�����̃V�[�g�̒���̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BBefore ���w�肳��Ă���ꍇ�AAfter �͎w��ł��܂���B</param>
		''' <remarks>
		''' ���� Before �ƈ��� After �����ɏȗ������ꍇ�́A�V�K�u�b�N�������I�ɍ쐬����A�V�[�g�͂��̃u�b�N���ɑ}������܂��B
		''' </remarks>
		Public Sub Copy( _
		  <InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		  <InAttribute()> Optional ByVal After As SheetWrapper = Nothing _
		 )
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

			InvokeMethod(_sheets, "Copy", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' �I�u�W�F�N�g���폜���܂��B
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_sheets, "Delete", Nothing)
		End Sub

		''' <summary>
		''' �w�肳�ꂽ�Z���͈͂��A�R���N�V�������̑��̂��ׂẴ��[�N�V�[�g�̓����̈�ɃR�s�[���܂��B
		''' </summary>
		''' <param name="Range">�K���w�肵�܂��BRange �I�u�W�F�N�g���w�肵�܂��B�R���N�V�����ɑ����邷�ׂẴ��[�N�V�[�g�̃t�B���Ɏg�p����Z���͈͂��w�肵�܂��B���̃Z���͈͂ɂ́A�R���N�V�������̃��[�N�V�[�g���w�肷��K�v������܂��B</param>
		''' <param name="Type">�ȗ��\�ł��BXlFillWith �̒l���w�肵�܂��B�w�肵���Z���͈͂��R�s�[������@���w�肵�܂��B</param>
		''' <remarks></remarks>
		Public Sub FillAcrossSheets( _
		  <InAttribute()> ByVal Range As RangeWrapper, _
		  <InAttribute()> Optional ByVal Type As XlFillWith = XlFillWith.xlFillWithAll _
		 )
			InvokeMethod(_sheets, "FillAcrossSheets", New Object() {Range.OrigianlInstance, Type})
		End Sub

		''' <summary>
		''' Excel.Sheets.GetEnumerator��Ԃ�
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator(Of SheetWrapper)
			Dim sheetEnum As IEnumerator
			Dim result As IList(Of SheetWrapper)

			result = New List(Of SheetWrapper)

			sheetEnum = DirectCast(InvokeMethod(_sheets, "GetEnumerator", Nothing), IEnumerator)
			While sheetEnum.MoveNext()
				Dim wrapper As SheetWrapper
				wrapper = New SheetWrapper(_book, sheetEnum.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

		''' <summary>
		''' �V�[�g���u�b�N���̑��̏ꏊ�Ɉړ����܂��B
		''' </summary>
		''' <param name="Before">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�ړ�����V�[�g�����̃V�[�g�̒��O�̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BAfter ���w�肳��Ă���ꍇ�ABefore �͎w��ł��܂���B</param>
		''' <param name="After">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�ړ�����V�[�g�����̃V�[�g�̒���̈ʒu�ɑ}������Ƃ��ɁA���̃V�[�g���w�肵�܂��BBefore ���w�肳��Ă���ꍇ�AAfter �͎w��ł��܂���B</param>
		''' <remarks></remarks>
		Public Sub Move( _
		<InAttribute()> Optional ByVal Before As SheetWrapper = Nothing, _
		<InAttribute()> Optional ByVal After As SheetWrapper = Nothing _
		  )
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

			InvokeMethod(_sheets, "Move", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' �I�u�W�F�N�g��������܂��B
		''' </summary>
		''' <param name="From">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B������J�n����y�[�W�ԍ����w�肵�܂��B���̈������ȗ�����ƁA�ŏ��̃y�[�W����������܂��B</param>
		''' <param name="To">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B������I������y�[�W�ԍ����w�肵�܂��B���̈������ȗ�����ƁA����͍Ō�̃y�[�W�ŏI�����܂��B</param>
		''' <param name="Copies">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B������镔�����w�肵�܂��B���̈������ȗ�����ƁA1 �����������܂��B</param>
		''' <param name="Preview">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ��ݒ肷��ƁA�I�u�W�F�N�g���������O�Ɉ���v���r���[�����s����܂��BFalse ��ݒ肷�邩�A�܂��͈������ȗ�����ƁA�I�u�W�F�N�g�͒����Ɉ������܂��B</param>
		''' <param name="ActivePrinter">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���ݎg�p���Ă���v�����^�̖��O��ݒ肵�܂��B</param>
		''' <param name="PrintToFile">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ��ݒ肷��ƁA�t�@�C����������܂��B���� PrToFileName ���w�肵�Ȃ��ƁA�o�̓t�@�C�����̓��͂𑣂��_�C�A���O �{�b�N�X���\������܂��B</param>
		''' <param name="Collate">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��BTrue ��ݒ肷��ƁA�������P�ʂň������܂��B</param>
		''' <param name="PrToFileName">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B���� PrintToFile �� True ��ݒ肵���ꍇ�A�������t�@�C���̖��O�����̈����Ɏw�肵�܂��B</param>
		''' <remarks></remarks>
		Public Sub PrintOut( _
		  <InAttribute()> Optional ByVal From As Object = Nothing, _
		  <InAttribute()> Optional ByVal [To] As Object = Nothing, _
		  <InAttribute()> Optional ByVal Copies As Object = Nothing, _
		  <InAttribute()> Optional ByVal Preview As Object = Nothing, _
		  <InAttribute()> Optional ByVal ActivePrinter As Object = Nothing, _
		  <InAttribute()> Optional ByVal PrintToFile As Object = Nothing, _
		  <InAttribute()> Optional ByVal Collate As Object = Nothing, _
		  <InAttribute()> Optional ByVal PrToFileName As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If From IsNot Nothing Then
				argsV.Add(From)
				argsN.Add("From")
			End If
			If [To] IsNot Nothing Then
				argsV.Add([To])
				argsN.Add("To")
			End If
			If Copies IsNot Nothing Then
				argsV.Add(Copies)
				argsN.Add("Copies")
			End If
			If Preview IsNot Nothing Then
				argsV.Add(Preview)
				argsN.Add("Preview")
			End If
			If ActivePrinter IsNot Nothing Then
				argsV.Add(ActivePrinter)
				argsN.Add("ActivePrinter")
			End If
			If PrintToFile IsNot Nothing Then
				argsV.Add(PrintToFile)
				argsN.Add("PrintToFile")
			End If
			If Collate IsNot Nothing Then
				argsV.Add(Collate)
				argsN.Add("Collate")
			End If
			If PrToFileName IsNot Nothing Then
				argsV.Add(PrToFileName)
				argsN.Add("PrToFileName")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' �I�u�W�F�N�g�̈���v���r���[ (������̃C���[�W) ��\�����܂��B
		''' </summary>
		''' <param name="EnableChanges">�I�u�W�F�N�g�̕ύX���\�ɂ��܂��B</param>
		''' <remarks></remarks>
		Public Sub PrintPreview( _
		  <InAttribute()> Optional ByVal EnableChanges As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If EnableChanges IsNot Nothing Then
				argsV.Add(EnableChanges)
				argsN.Add("EnableChanges")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' �I�u�W�F�N�g��I�����܂��B
		''' </summary>
		''' <param name="Replace">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�w��̃I�u�W�F�N�g��I������ۂɁA���ɑI�����Ă���I�u�W�F�N�g�̑I�����������邩�ǂ������w�肵�܂��B</param>
		''' <remarks></remarks>
		Public Sub [Select]( _
		  <InAttribute()> Optional ByVal Replace As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If Replace IsNot Nothing Then
				argsV.Add(Replace)
				argsN.Add("Replace")
			End If

			InvokeMethod(_sheets, "PrintOut", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

#End Region

	End Class

End Namespace
