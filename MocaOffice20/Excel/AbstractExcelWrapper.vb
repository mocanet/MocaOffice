
Imports System.Reflection

Namespace Excel

	''' <summary>
	''' Excel����֌W�̒��ۃN���X
	''' </summary>
	''' <remarks></remarks>
	Public MustInherit Class AbstractExcelWrapper
		Implements IDisposable

		''' <summary>Excel.Application �̃^�C�v</summary>
		Protected typApplication As Type
		''' <summary>Excel�A�v���P�[�V����</summary>
		Protected xlsApp As Object
		''' <summary>Excel�A�v���P�[�V�������b�p�[</summary>
		Protected xlsWrapper As AbstractExcelWrapper
		''' <summary>�ꎞ�I�ɃC���X�^���X�����ꂽExcel�֌W�̃I�u�W�F�N�g�B</summary>
		Private _xlsObjects As IList(Of Object)

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �f�t�H���g�R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Protected Sub New()
			_xlsObjects = New List(Of Object)
		End Sub

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <param name="xls">Excel�A�v���P�[�V����</param>
		''' <remarks></remarks>
		Protected Sub New(ByVal xls As AbstractExcelWrapper)
			Me.New()
			xlsWrapper = xls
			xlsApp = xlsWrapper.Application
		End Sub

#End Region

#Region " IDisposable Support "

		Private disposedValue As Boolean = False		' �d������Ăяo�������o����ɂ�

		' IDisposable
		Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Me.disposedValue Then
				If disposing Then
					' TODO: �����I�ɌĂяo���ꂽ�Ƃ��Ƀ}�l�[�W ���\�[�X��������܂�
				End If

				' TODO: ���L�̃A���}�l�[�W ���\�[�X��������܂�

				' �ꎞ�I�ɃC���X�^���X�����ꂽExcel�֌W�̃I�u�W�F�N�g�̃������J��
				AllDispose()
				' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
				MyDispose()
			End If
			Me.disposedValue = True
		End Sub

		' ���̃R�[�h�́A�j���\�ȃp�^�[���𐳂��������ł���悤�� Visual Basic �ɂ���Ēǉ�����܂����B
		Public Sub Dispose() Implements IDisposable.Dispose
			' ���̃R�[�h��ύX���Ȃ��ł��������B�N���[���A�b�v �R�[�h����� Dispose(ByVal disposing As Boolean) �ɋL�q���܂��B
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

#End Region

#Region " MustOverride "

		''' <summary>
		''' �������g�ŊǗ����Ă���Excel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Public MustOverride Sub MyDispose()

		''' <summary>
		''' �擾���� Excel �C���X�^���X
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend MustOverride ReadOnly Property OrigianlInstance() As Object

#End Region

#Region " �v���p�e�B "

		''' <summary>
		''' Excel�A�v���P�[�V����
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Property Application() As Object
			Get
				Return xlsApp
			End Get
			Set(ByVal value As Object)
				xlsApp = value
			End Set
		End Property

		''' <summary>
		''' Excel�A�v���P�[�V�������b�p�[
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Property ApplicationWrapper() As AbstractExcelWrapper
			Get
				Return xlsWrapper
			End Get
			Set(ByVal value As AbstractExcelWrapper)
				xlsWrapper = value
				Me.Application = xlsWrapper.Application
			End Set
		End Property

#End Region

		''' <summary>
		''' �C���X�^���X������ Excel �I�u�W�F�N�g��ǉ�����
		''' </summary>
		''' <param name="obj"></param>
		''' <remarks></remarks>
		Protected Sub addXlsObject(ByVal obj As Object)
			_xlsObjects.Add(obj)
		End Sub

		''' <summary>
		''' �C���X�^���X������ Excel �I�u�W�F�N�g
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function myXlsObject() As IList(Of Object)
			Return _xlsObjects
		End Function

		''' <summary>
		''' �ꎞ�I�ɃC���X�^���X�����ꂽExcel�֌W�̃I�u�W�F�N�g�̃������J��
		''' </summary>
		''' <remarks></remarks>
		Private Sub AllDispose()
			For Each obj As Object In _xlsObjects
				If TypeOf obj Is AbstractExcelWrapper Then
					DirectCast(obj, AbstractExcelWrapper).Dispose()
				Else
					ReleaseExcelObject(obj)
				End If
			Next
		End Sub

		''' <summary>
		''' ���t���N�V�����ɂ�郁�\�b�h�̎��s
		''' </summary>
		''' <param name="target"></param>
		''' <param name="name"></param>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function InvokeMethod(ByVal target As Object, ByVal name As String, ByVal args() As Object, Optional ByVal argNames() As String = Nothing) As Object
			Try
                _mylog.DebugFormat(_makeMsg("Method", target, name, args, argNames))
                Return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, Nothing, target, args, Nothing, Nothing, argNames)
			Catch ex As Exception
				Throw New ExcelException(xlsWrapper, ex, _makeMsg("Method", target, name, args) & vbCrLf & ex.InnerException.Message)
			End Try
		End Function

		''' <summary>
		''' ���t���N�V�����ɂ��v���p�e�B�擾�̎��s
		''' </summary>
		''' <param name="target"></param>
		''' <param name="name"></param>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function InvokeGetProperty(ByVal target As Object, ByVal name As String, ByVal args() As Object, Optional ByVal argNames() As String = Nothing) As Object
			Try
				_mylog.DebugFormat(_makeMsg("GetProperty", target, name, args, argNames))
				Return target.GetType().InvokeMember(name, BindingFlags.GetProperty, Nothing, target, args, Nothing, Nothing, argNames)
			Catch ex As Exception
				Throw New ExcelException(xlsWrapper, ex, _makeMsg("GetProperty", target, name, args) & vbCrLf & ex.InnerException.Message)
			End Try
		End Function

		''' <summary>
		''' ���t���N�V�����ɂ��v���p�e�B�ݒ�̎��s
		''' </summary>
		''' <param name="target"></param>
		''' <param name="name"></param>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function InvokeSetProperty(ByVal target As Object, ByVal name As String, ByVal args() As Object) As Object
			Try
				_mylog.DebugFormat(_makeMsg("SetProperty", target, name, args))
				Return target.GetType().InvokeMember(name, BindingFlags.SetProperty, Nothing, target, args)
			Catch ex As Exception
				Throw New ExcelException(xlsWrapper, ex, _makeMsg("SetProperty", target, name, args) & vbCrLf & ex.InnerException.Message)
			End Try
		End Function

		''' <summary>
		''' Excel�I�u�W�F�N�g���J������
		''' </summary>
		''' <param name="excelObject">Com�I�u�W�F�N�g</param>
		''' <param name="quitExcelApplication">Excel�A�v���P�[�V�������̂��I�����邩�ǂ���</param>
		''' <remarks>
		''' </remarks>
		Protected Sub ReleaseExcelObject(ByRef excelObject As Object, Optional ByVal quitExcelApplication As Boolean = False)
			If excelObject Is Nothing Then
				Exit Sub
			End If

			' Excel�A�v���P�[�V�������̂��I��
			If quitExcelApplication Then
				If excelObject.GetType.Equals(typApplication) Then
					InvokeMethod(excelObject, "Quit", Nothing)
				End If
			End If

			'COM �I�u�W�F�N�g�̎g�p��A�����I�� COM �I�u�W�F�N�g�ւ̎Q�Ƃ��������
			Try
				'�񋟂��ꂽ�����^�C���Ăяo���\���b�p�[�̎Q�ƃJ�E���g���f�N�������g���܂�
				If System.Runtime.InteropServices.Marshal.IsComObject(excelObject) Then
					System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelObject)
				End If
				' .NET1.1
				'If Not excelObject Is Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(excelObject) Then
				'	Dim ii As Integer
				'	Do
				'		ii = System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObject)
				'	Loop Until ii <= 0
				'End If
			Finally
				'�Q�Ƃ���������
				excelObject = Nothing
			End Try
		End Sub

		''' <summary>
		''' ���ۃN���X�փL���X�g
		''' </summary>
		''' <param name="wrapper">�ΏۂƂȂ�I�u�W�F�N�g</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function castAbstractExcelWrapper(ByVal wrapper As Object) As AbstractExcelWrapper
			Return DirectCast(wrapper, AbstractExcelWrapper)
		End Function

		''' <summary>
		''' Invoke���̃��b�Z�[�W�쐬
		''' </summary>
		''' <param name="binding"></param>
		''' <param name="target"></param>
		''' <param name="name"></param>
		''' <param name="args"></param>
		''' <param name="argNames"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _makeMsg(ByVal binding As String, ByVal target As Object, ByVal name As String, ByVal args() As Object, Optional ByVal argNames() As String = Nothing) As String
			Return String.Format("Invoke Binding:{0} Name:{1} Target:{2} Args:{3} ({4})({5})", binding, name, target.GetType(), _argsLength(args), _joinArgs(args), _joinArgsName(argNames))
		End Function

		''' <summary>
		''' �����̐���Ԃ�
		''' </summary>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _argsLength(ByVal args() As Object) As Integer
			If args Is Nothing Then
				Return 0
			Else
				Return args.Length
			End If
		End Function

		''' <summary>
		''' �����̃^�C�v�z����쐬
		''' </summary>
		''' <param name="args"></param>
		''' <param name="propertyArgs"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _getArgsType(ByVal args() As Object, Optional ByVal propertyArgs As Boolean = False) As Type()
			Dim typ As ArrayList
			typ = New ArrayList
			If args IsNot Nothing Then
				For Each item As Object In args
					If propertyArgs Then
						propertyArgs = False
						Continue For
					End If
					typ.Add(item.GetType())
				Next
			End If
			Return DirectCast(typ.ToArray(GetType(Type)), Type())
		End Function

		''' <summary>
		''' �����̃C���X�^���X�����̌^���J���}��؂�ŕ�����ɂ���
		''' </summary>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _joinArgs(ByVal args() As Object) As String
			Dim msg As ArrayList
			msg = New ArrayList
			If args Is Nothing Then
				Return String.Empty
			End If
            For Each item As Object In args
                If item Is Nothing Then
                    msg.Add("Nothing")
                Else
                    msg.Add(item.GetType().Name)
                End If
            Next
			Return Join(msg.ToArray(), ",")
		End Function

		''' <summary>
		''' �����̖��̔z����쐬
		''' </summary>
		''' <param name="args"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function _joinArgsName(ByVal args() As Object) As String
			Dim msg As ArrayList
			msg = New ArrayList
			If args Is Nothing Then
				Return String.Empty
			End If
			For Each item As Object In args
				msg.Add(item)
			Next
			Return Join(msg.ToArray(), ",")
		End Function

	End Class

End Namespace
