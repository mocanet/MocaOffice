
Imports System.Reflection

Namespace Excel

	''' <summary>
	''' Excel操作関係の抽象クラス
	''' </summary>
	''' <remarks></remarks>
	Public MustInherit Class AbstractExcelWrapper
		Implements IDisposable

		''' <summary>Excel.Application のタイプ</summary>
		Protected typApplication As Type
		''' <summary>Excelアプリケーション</summary>
		Protected xlsApp As Object
		''' <summary>Excelアプリケーションラッパー</summary>
		Protected xlsWrapper As AbstractExcelWrapper
		''' <summary>一時的にインスタンス化されたExcel関係のオブジェクト達</summary>
		Private _xlsObjects As IList(Of Object)

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' デフォルトコンストラクタ
		''' </summary>
		''' <remarks></remarks>
		Protected Sub New()
			_xlsObjects = New List(Of Object)
		End Sub

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="xls">Excelアプリケーション</param>
		''' <remarks></remarks>
		Protected Sub New(ByVal xls As AbstractExcelWrapper)
			Me.New()
			xlsWrapper = xls
			xlsApp = xlsWrapper.Application
		End Sub

#End Region

#Region " IDisposable Support "

		Private disposedValue As Boolean = False		' 重複する呼び出しを検出するには

		' IDisposable
		Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Me.disposedValue Then
				If disposing Then
					' TODO: 明示的に呼び出されたときにマネージ リソースを解放します
				End If

				' TODO: 共有のアンマネージ リソースを解放します

				' 一時的にインスタンス化されたExcel関係のオブジェクトのメモリ開放
				AllDispose()
				' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
				MyDispose()
			End If
			Me.disposedValue = True
		End Sub

		' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
		Public Sub Dispose() Implements IDisposable.Dispose
			' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

#End Region

#Region " MustOverride "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public MustOverride Sub MyDispose()

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend MustOverride ReadOnly Property OrigianlInstance() As Object

#End Region

#Region " プロパティ "

		''' <summary>
		''' Excelアプリケーション
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
		''' Excelアプリケーションラッパー
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
		''' インスタンス化した Excel オブジェクトを追加する
		''' </summary>
		''' <param name="obj"></param>
		''' <remarks></remarks>
		Protected Sub addXlsObject(ByVal obj As Object)
			_xlsObjects.Add(obj)
		End Sub

		''' <summary>
		''' インスタンス化した Excel オブジェクト
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function myXlsObject() As IList(Of Object)
			Return _xlsObjects
		End Function

		''' <summary>
		''' 一時的にインスタンス化されたExcel関係のオブジェクトのメモリ開放
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
		''' リフレクションによるメソッドの実行
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
		''' リフレクションによるプロパティ取得の実行
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
		''' リフレクションによるプロパティ設定の実行
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
		''' Excelオブジェクトを開放する
		''' </summary>
		''' <param name="excelObject">Comオブジェクト</param>
		''' <param name="quitExcelApplication">Excelアプリケーション自体を終了するかどうか</param>
		''' <remarks>
		''' </remarks>
		Protected Sub ReleaseExcelObject(ByRef excelObject As Object, Optional ByVal quitExcelApplication As Boolean = False)
			If excelObject Is Nothing Then
				Exit Sub
			End If

			' Excelアプリケーション自体を終了
			If quitExcelApplication Then
				If excelObject.GetType.Equals(typApplication) Then
					InvokeMethod(excelObject, "Quit", Nothing)
				End If
			End If

			'COM オブジェクトの使用後、明示的に COM オブジェクトへの参照を解放する
			Try
				'提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
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
				'参照を解除する
				excelObject = Nothing
			End Try
		End Sub

		''' <summary>
		''' 抽象クラスへキャスト
		''' </summary>
		''' <param name="wrapper">対象となるオブジェクト</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Protected Function castAbstractExcelWrapper(ByVal wrapper As Object) As AbstractExcelWrapper
			Return DirectCast(wrapper, AbstractExcelWrapper)
		End Function

		''' <summary>
		''' Invoke時のメッセージ作成
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
		''' 引数の数を返す
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
		''' 引数のタイプ配列を作成
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
		''' 引数のインスタンスたちの型をカンマ区切りで文字列にする
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
		''' 引数の名称配列を作成
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
