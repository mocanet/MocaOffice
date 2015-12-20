
Namespace Excel

	''' <summary>
	''' Excel.Worksheets �̃��b�p�[�N���X
	''' </summary>
	''' <remarks></remarks>
	Public Class WorksheetsWrapper
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

#End Region

		''' <summary>
		''' ������
		''' </summary>
		''' <param name="book">�e�̃u�b�N</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal book As BookWrapper)
			_book = book
			' Worksheets�I�u�W�F�N�g�̍쐬
			_sheets = InvokeGetProperty(_book.OrigianlInstance, "Worksheets", Nothing)
		End Sub

		''' <summary>
		''' �V�������[�N�V�[�g���쐬���܂��B�쐬�������[�N�V�[�g�̓A�N�e�B�u�ɂȂ�܂��B 
		''' </summary>
		''' <param name="sheetName">�V�[�g��</param>
		''' <param name="Before">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�����V�[�g�����̃V�[�g�̒��O�̈ʒu�ɒǉ�����Ƃ��ɁA���̃V�[�g���w�肵�܂��B</param>
		''' <param name="After">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�����V�[�g�����̃V�[�g�̒���̈ʒu�ɒǉ�����Ƃ��ɁA���̃V�[�g���w�肵�܂��B</param>
		''' <param name="Count">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�ǉ�����V�[�g�̐����w�肵�܂��B����l�� 1 �ł��B</param>
		''' <param name="Type">�ȗ��\�ł��B�I�u�W�F�N�g�^ (Object) �̒l���w�肵�܂��B�V�[�g�̎�ނ��w�肵�܂��B�g�p�ł���萔�́AXlSheetType �񋓌^�� xlWorksheet�AxlChart�AxlExcel4MacroSheet�AxlExcel4IntlMacroSheet �̂����ꂩ�ł��B�����̃e���v���[�g����ɂ����V�[�g��}������ꍇ�́A���̃e���v���[�g�ւ̃p�X���w�肵�܂��B����l�� xlWorksheet �ł��B</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add(ByVal sheetName As String _
		, Optional ByVal Before As SheetWrapper = Nothing _
		, Optional ByVal After As SheetWrapper = Nothing _
		, Optional ByVal Count As Integer = 1 _
		, Optional ByVal Type As Object = Nothing) As SheetWrapper
			Dim xls As Object
			Dim sheet As SheetWrapper

			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If After IsNot Nothing Then
				argsV.Add(After)
				argsN.Add("After")
			End If
			If Before IsNot Nothing Then
				argsV.Add(Before)
				argsN.Add("Before")
			End If
			If argsV.Count = 0 Then
				argsV.Add(Me.App.ActiveSheet.OrigianlInstance)
				argsN.Add("After")
			End If

			argsV.Add(Count)
			argsN.Add("Count")
			If Type IsNot Nothing Then
				argsV.Add(Type)
				argsN.Add("Type")
			End If

			xls = InvokeMethod(_sheets, "Add", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))

			sheet = New SheetWrapper(_book, xls)
			sheet.Name = sheetName
			addXlsObject(sheet)
			Return sheet
		End Function

		''' <summary>
		''' Excel.Worksheets.GetEnumerator��Ԃ�
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
		''' �f�t�H���g�V�[�g�폜
		''' </summary>
		''' <remarks>
		''' Excel��V�K�ō쐬����Ƃ��ɏo����f�t�H���g�̃V�[�g�iSheet1...�j���폜���܂��B
		''' </remarks>
		Public Sub ClearDefaultSheet()
			Dim sheetEnum As IEnumerator(Of SheetWrapper)
			Dim xlSheet As SheetWrapper = Nothing

			sheetEnum = GetEnumerator()
			While sheetEnum.MoveNext()
				xlSheet = sheetEnum.Current()

				If xlSheet.Name.StartsWith("Sheet") Then
					xlSheet.Delete()
				End If
			End While
		End Sub

		''' <summary>
		''' ����̃V�[�g�������݂���ꍇ�̌������Z�o����
		''' </summary>
		''' <param name="sheetName">�V�[�g��</param>
		''' <returns></returns>
		''' <remarks>
		''' �u�w�肳�ꂽ���́{�h�Q�h�{�ԍ��Ȃǁv�̖��̃V�[�g�����݂�������Ԃ�
		''' </remarks>
		Public Function MultiSheetCount(ByVal sheetName As String) As Integer
			Return MultiSheetCount(sheetName, "_")
		End Function

		''' <summary>
		''' ����̃V�[�g�������݂���ꍇ�̌������Z�o����
		''' </summary>
		''' <param name="sheetName">�V�[�g��</param>
		''' <param name="delim">��؂蕶��</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function MultiSheetCount(ByVal sheetName As String, ByVal delim As String) As Integer
			Dim sheetEnum As IEnumerator(Of SheetWrapper)
			Dim xlSheet As SheetWrapper = Nothing
			Dim sheetCount As Integer

			sheetCount = 0
			sheetEnum = GetEnumerator()
			While sheetEnum.MoveNext()
				xlSheet = sheetEnum.Current()
				Try
					Dim aryName() As String

					aryName = xlSheet.Name.Split(CChar(delim))
					If aryName(0) = sheetName Then
						sheetCount += 1
					End If
				Finally
					xlSheet.Dispose()
				End Try
			End While

			Return sheetCount
		End Function

	End Class

End Namespace
