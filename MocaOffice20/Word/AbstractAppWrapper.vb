
Imports System.Reflection

Namespace Word

    ''' <summary>
    ''' Office 操作関係の抽象クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class AbstractAppWrapper
        Implements IDisposable

#Region " Declare "

        ''' <summary>Application のタイプ</summary>
        Protected typApplication As Type
        ''' <summary>アプリケーション</summary>
        Protected officeApp As Object
        ''' <summary>アプリケーションラッパー</summary>
        Protected appWrapper As AbstractAppWrapper
        ''' <summary>一時的にインスタンス化されたオブジェクト達</summary>
        Private _objects As IList(Of Object)

        ''' <summary>log4net logger</summary>
        Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#End Region

#Region " コンストラクタ "

        ''' <summary>
        ''' デフォルトコンストラクタ
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub New()
            _objects = New List(Of Object)
        End Sub

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <param name="appWrapper">アプリケーション</param>
        ''' <remarks></remarks>
        Protected Sub New(ByVal appWrapper As AbstractAppWrapper)
            Me.New()
            Me.appWrapper = appWrapper
            officeApp = Me.appWrapper.Application
        End Sub

#End Region

#Region " IDisposable Support "

        Private disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: 明示的に呼び出されたときにマネージ リソースを解放します
                End If

                ' TODO: 共有のアンマネージ リソースを解放します

                ' 一時的にインスタンス化されたオブジェクトのメモリ開放
                _allDispose()
                ' 自分自身で管理しているオブジェクトのメモリ開放
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
        ''' 自分自身で管理しているオブジェクトのメモリ開放
        ''' </summary>
        ''' <remarks></remarks>
        Public MustOverride Sub MyDispose()

        ''' <summary>
        ''' 取得した アプリケーション インスタンス
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend MustOverride ReadOnly Property OrigianlInstance() As Object

#End Region

#Region " プロパティ "

        ''' <summary>
        ''' アプリケーション
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Property Application() As Object
            Get
                Return officeApp
            End Get
            Set(ByVal value As Object)
                officeApp = value
            End Set
        End Property

        ''' <summary>
        ''' アプリケーションラッパー
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Property ApplicationWrapper() As AbstractAppWrapper
            Get
                Return appWrapper
            End Get
            Set(ByVal value As AbstractAppWrapper)
                appWrapper = value
                Me.Application = appWrapper.Application
            End Set
        End Property

#End Region
#Region " Method "

        ''' <summary>
        ''' インスタンス化したオブジェクトを追加する
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <remarks></remarks>
        Protected Sub addObject(ByVal obj As Object)
            _objects.Add(obj)
        End Sub

        ''' <summary>
        ''' インスタンス化したオブジェクト
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function myObject() As IList(Of Object)
            Return _objects
        End Function

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
                Throw New OfficeException(appWrapper, ex, _makeMsg("Method", target, name, args) & vbCrLf & ex.InnerException.Message)
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
                Throw New OfficeException(appWrapper, ex, _makeMsg("GetProperty", target, name, args) & vbCrLf & ex.InnerException.Message)
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
                Throw New OfficeException(appWrapper, ex, _makeMsg("SetProperty", target, name, args) & vbCrLf & ex.InnerException.Message)
            End Try
        End Function

        ''' <summary>
        ''' オブジェクトを開放する
        ''' </summary>
        ''' <param name="officeObject">Comオブジェクト</param>
        ''' <param name="quitOfficeApplication">アプリケーション自体を終了するかどうか</param>
        ''' <remarks>
        ''' </remarks>
        Protected Sub ReleaseOfficeObject(ByRef officeObject As Object, Optional ByVal quitOfficeApplication As Boolean = False)
            If officeObject Is Nothing Then
                Exit Sub
            End If

            ' アプリケーション自体を終了
            If quitOfficeApplication Then
                If officeObject.GetType.Equals(typApplication) Then
                    InvokeMethod(officeObject, "Quit", Nothing)
                End If
            End If

            'COM オブジェクトの使用後、明示的に COM オブジェクトへの参照を解放する
            Try
                '提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
                If System.Runtime.InteropServices.Marshal.IsComObject(officeObject) Then
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(officeObject)
                End If
                ' .NET1.1
                'If Not officeObject Is Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(officeObject) Then
                '	Dim ii As Integer
                '	Do
                '		ii = System.Runtime.InteropServices.Marshal.ReleaseComObject(officeObject)
                '	Loop Until ii <= 0
                'End If
            Finally
                '参照を解除する
                officeObject = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' 抽象クラスへキャスト
        ''' </summary>
        ''' <param name="wrapper">対象となるオブジェクト</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function castAbstractOffcieWrapper(ByVal wrapper As Object) As AbstractAppWrapper
            Return DirectCast(wrapper, AbstractAppWrapper)
        End Function

        ''' <summary>
        ''' 一時的にインスタンス化されたオブジェクトのメモリ開放
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub _allDispose()
            For Each obj As Object In _objects
                If TypeOf obj Is AbstractAppWrapper Then
                    DirectCast(obj, AbstractAppWrapper).Dispose()
                Else
                    ReleaseOfficeObject(obj)
                End If
            Next
        End Sub

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

#End Region

        Protected argsV As ArrayList
        Protected argsN As ArrayList

        Protected Sub argsClear()
            If argsV Is Nothing Then
                argsV = New ArrayList()
                argsN = New ArrayList()
            End If

            argsV.Clear()
            argsN.Clear()
        End Sub

        Protected Overloads Sub argsAdd(ByVal name As String, ByVal value As Object)
            argsAdd(name, value, False)
        End Sub

        Protected Overloads Sub argsAdd(ByVal name As String, ByVal value As Object, ByVal required As Boolean)
            If Not required Then
                If value Is Nothing Then
                    Return
                End If
            End If

            argsV.Add(value)
            argsN.Add(name)
        End Sub

    End Class

End Namespace
